#!/usr/bin/env python3
"""GBS Laptop Diagnostics v2.0 - run from an Ubuntu Live USB session.

Python 3 standard library only. Internal drives are NEVER written to or erased
(the SSD check only reads the NVMe SMART log).

Automatic : inventory, battery, SSD/HDD SMART, controllers, Wi-Fi, Bluetooth, memory screen
Interactive: keyboard + touchpad, speakers + microphone, webcam, USB ports

Usage:  bash run_diagnostics.sh            (full test)
        bash run_diagnostics.sh --auto     (automatic tests only, no questions)
        bash run_diagnostics.sh --only keyboard,usb   (re-test specific items)
"""
import argparse, array, configparser, csv, ctypes, datetime, fcntl, getpass, glob, html, json, math
import mmap, os, pathlib, pwd, re, select, shutil, signal, struct, subprocess, sys, tempfile, termios, time, wave, zlib
import base64, urllib.error, urllib.request

VERSION = "2.1"

# ---------------------------------------------------------------- thresholds
CFG = {
    "battery_pass": 25,       # health >  25%  -> PASS, otherwise FAIL
    "ssd_warn_used": 80,      # NVMe "percentage used" >= 80 (life <= 20%) -> WARN
    "ssd_temp_warn": 70,      # deg C
    "mem_fraction": 0.40,     # share of available RAM used by the memory screen
    "mem_max_gib": 8,         # upper bound for the memory screen
    "keyboard_timeout": 180,  # seconds
    "touchpad_timeout": 60,
    "usb_wait": 15,
    "bt_scan": 15,
    "mic_seconds": 4,
    "mic_min_dbfs": -50,
    "timezone": "America/New_York",   # time zone shown on the reports
}

BASE = pathlib.Path(__file__).resolve().parent
ORIGIN = os.environ.get("GBS_ORIGIN", "USB")
OUT = BASE / "reports"
AUTO = False
checks = []          # (name, result, detail)
PARTS = {}           # partial results that are combined into one line
INFO = {}

C = dict(g="\033[32m", r="\033[31m", y="\033[33m", b="\033[1m", d="\033[2m", c="\033[36m", x="\033[0m")
SCOLOR = {"PASS": C["g"], "WARN": C["y"], "FAIL": C["r"], "NOT_TESTED": C["d"], "NOT_INSTALLED": C["d"]}
RANK = {"PASS": 0, "NOT_INSTALLED": 0, "WARN": 1, "NOT_TESTED": 2, "FAIL": 3}


# ---------------------------------------------------------------- helpers
def run(cmd, timeout=30):
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return p.returncode, (p.stdout or "").strip(), (p.stderr or "").strip()
    except (OSError, subprocess.TimeoutExpired) as e:
        return 127, "", str(e)


def read(path):
    try:
        return pathlib.Path(path).read_text(errors="replace").strip()
    except OSError:
        return ""


def cmdexists(x):
    return shutil.which(x) is not None


def dmi(key):
    return read("/sys/class/dmi/id/" + key)


def safe(s):
    return re.sub(r"[^A-Za-z0-9_.-]", "_", s or "UNKNOWN")[:80]


def worst(*statuses):
    return max(statuses, key=lambda s: RANK.get(s, 2))


def add(name, result, detail):
    detail = str(detail).replace("\n", " ")[:700]
    checks.append((name, result, detail))
    print(f"   {SCOLOR.get(result, '')}[{result}]{C['x']} {name}: {detail}")


def title(t):
    print(f"\n{C['b']}{C['c']}== {t} =={C['x']}")


def info(t):
    print(f"   {t}")


def flush_stdin():
    try:
        termios.tcflush(sys.stdin, termios.TCIFLUSH)
    except (termios.error, OSError, ValueError):
        pass


def pause(msg):
    if AUTO:
        return
    flush_stdin()
    input(f"   {C['b']}>> {msg}{C['x']} ")


def choice(msg, opts, default=None):
    if AUTO:
        return default
    flush_stdin()
    while True:
        a = input(f"   {C['b']}>> {msg}{C['x']} ").strip().lower()
        if not a and default:
            return default
        if a and a[0] in opts:
            return a[0]
        print(f"   Answer with: {'/'.join(opts)}")


# Commands that must run inside the logged-in desktop session (audio, image viewer)
TARGET_USER = os.environ.get("GBS_USER") or os.environ.get("SUDO_USER") or ""


def user_cmd(cmd):
    if os.geteuid() != 0 or not TARGET_USER or TARGET_USER == "root":
        return cmd
    try:
        pw = pwd.getpwnam(TARGET_USER)
    except KeyError:
        return cmd
    env = [f"XDG_RUNTIME_DIR=/run/user/{pw.pw_uid}",
           f"DBUS_SESSION_BUS_ADDRESS=unix:path=/run/user/{pw.pw_uid}/bus", f"HOME={pw.pw_dir}"]
    for k in ("DISPLAY", "WAYLAND_DISPLAY", "XAUTHORITY"):
        if os.environ.get(k):
            env.append(f"{k}={os.environ[k]}")
    if not os.environ.get("WAYLAND_DISPLAY"):
        env.append("WAYLAND_DISPLAY=wayland-0")
    if cmdexists("runuser"):
        return ["runuser", "-u", TARGET_USER, "--", "env", *env, *cmd]
    return ["sudo", "-u", TARGET_USER, "env", *env, *cmd]


def _IOC(d, t, nr, size):
    return (d << 30) | (size << 16) | (ord(t) << 8) | nr


VIDIOC_QUERYCAP = _IOC(2, "V", 0, 104)
VIDIOC_S_FMT = _IOC(3, "V", 5, 208)
VIDIOC_REQBUFS = _IOC(3, "V", 8, 20)
VIDIOC_QUERYBUF = _IOC(3, "V", 9, 88)
VIDIOC_QBUF = _IOC(3, "V", 15, 88)
VIDIOC_DQBUF = _IOC(3, "V", 17, 88)
VIDIOC_STREAMON = _IOC(1, "V", 18, 4)
VIDIOC_STREAMOFF = _IOC(1, "V", 19, 4)
NVME_IOCTL_ADMIN_CMD = _IOC(3, "N", 0x41, 72)
EVIOCGRAB = _IOC(1, "E", 0x90, 4)


# ---------------------------------------------------------------- inventory
def test_inventory():
    title("Inventory")
    tag = dmi("product_serial") or dmi("board_serial")
    if tag.lower() in ("", "to be filled by o.e.m.", "default string", "none"):
        tag = ""
    INFO.update(service_tag=tag or "UNKNOWN", manufacturer=dmi("sys_vendor"), model=dmi("product_name"),
                bios=dmi("bios_version"))
    cpu = ""
    for line in read("/proc/cpuinfo").splitlines():
        if line.startswith("model name"):
            cpu = line.split(":", 1)[-1].strip()
            break
    mem_kb = 0
    for line in read("/proc/meminfo").splitlines():
        if line.startswith("MemTotal:"):
            mem_kb = int(re.search(r"\d+", line).group())
    INFO.update(cpu=cpu, ram_gib=round(mem_kb / 1024 / 1024, 2))
    add("CPU detected", "PASS" if cpu else "FAIL", cpu or "CPU not identified")
    add("RAM detected", "PASS" if mem_kb >= 1024 * 1024 else "FAIL", f"{mem_kb / 1024 / 1024:.2f} GiB visible to OS")
    add("BIOS identified", "PASS" if INFO["bios"] else "WARN", INFO["bios"] or "Unavailable")
    add("Service tag", "PASS" if tag else "WARN", tag or "Not readable (run as root / blank on board)")


# ---------------------------------------------------------------- battery
def test_battery():
    title("Battery")
    bats = [b for b in sorted(glob.glob("/sys/class/power_supply/*")) if read(b + "/type").lower() == "battery"]
    if not bats:
        add("Battery health", "FAIL", "No battery detected by the kernel")
        return
    for b in bats:
        name = "Battery health" if len(bats) == 1 else f"Battery health ({os.path.basename(b)})"
        energy = bool(read(b + "/energy_full"))
        full = read(b + "/energy_full") or read(b + "/charge_full")
        design = read(b + "/energy_full_design") or read(b + "/charge_full_design")
        health = None
        try:
            if int(design) > 0:
                health = round(int(full) / int(design) * 100, 1)
        except (TypeError, ValueError):
            pass
        unit, div = ("Wh", 1e6) if energy else ("Ah", 1e6)
        cap = read(b + "/capacity") or "?"
        cycles = read(b + "/cycle_count") or "?"
        parts = [f"health={health if health is not None else '?'}%", f"charge={cap}%", f"cycles={cycles}"]
        try:
            parts.append(f"full={int(full) / div:.1f}{unit} / design={int(design) / div:.1f}{unit}")
        except (TypeError, ValueError):
            pass
        parts.append(f"state={read(b + '/status') or '?'}")
        mfr = " ".join(x for x in (read(b + "/manufacturer"), read(b + "/model_name")) if x)
        if mfr:
            parts.append(mfr)
        if health is None:
            st = "NOT_TESTED"
        elif health > CFG["battery_pass"]:
            st = "PASS"
        else:
            st = "FAIL"
        INFO.setdefault("battery_health", health)
        INFO.setdefault("battery_cycles", cycles)
        add(name, st, "; ".join(parts))
    ac = [p for p in glob.glob("/sys/class/power_supply/*") if read(p + "/type").lower() in ("mains", "usb", "usb_c")]
    if ac:
        online = [read(p + "/online") for p in ac]
        add("External power (charger)", "PASS" if "1" in online else "WARN",
            "charger connected" if "1" in online else "charger not connected - plug in to verify charging port")
    else:
        add("External power (charger)", "NOT_TESTED", "AC adapter telemetry unavailable")


# ---------------------------------------------------------------- storage
def internal_disks():
    out = []
    for disk in sorted(glob.glob("/sys/block/*")):
        name = os.path.basename(disk)
        if not re.fullmatch(r"(sd[a-z]+|nvme\d+n\d+|mmcblk\d+)", name):
            continue
        if "/usb" in os.path.realpath(disk).lower() or read(disk + "/removable") == "1":
            continue
        try:
            sectors = int(read(disk + "/size") or "0")
        except ValueError:
            continue
        if sectors > 0:
            out.append((name, sectors * 512 / 1e9, read(disk + "/device/model")))
    return out


def nvme_smart(ns):
    ctrl = "/dev/" + re.match(r"(nvme\d+)", ns).group(1)
    buf = ctypes.create_string_buffer(512)
    cmd = bytearray(72)
    struct.pack_into("<BBHIIIQQII6III", cmd, 0, 0x02, 0, 0, 0xFFFFFFFF, 0, 0, 0, ctypes.addressof(buf),
                     0, 512, (127 << 16) | 0x02, 0, 0, 0, 0, 0, 0, 0)
    fd = os.open(ctrl, os.O_RDONLY)
    try:
        fcntl.ioctl(fd, NVME_IOCTL_ADMIN_CMD, cmd, True)
    finally:
        os.close(fd)
    raw = buf.raw

    def le(a, b):
        return int.from_bytes(raw[a:b], "little")
    t = le(1, 3)
    return dict(critical_warning=raw[0], temp_c=(t - 273) if t else None, spare=raw[3], spare_thr=raw[4],
                pct_used=raw[5], tb_written=le(48, 64) * 512000 / 1e12, power_cycles=le(112, 128),
                power_on_hours=le(128, 144), unsafe_shutdowns=le(144, 160), media_errors=le(160, 176))


def test_storage():
    title("Storage (SSD/HDD)")
    disks = internal_disks()
    add("Internal storage", "PASS" if disks else "NOT_INSTALLED",
        ", ".join(f"{n} {g:.0f} GB {m}".strip() for n, g, m in disks) or "No internal drive detected")
    INFO["storage"] = ", ".join(f"{n} {g:.0f}GB" for n, g, _ in disks) or "none"
    if not disks:
        add("SSD/HDD health", "NOT_INSTALLED", "No internal drive")
        return
    for n, g, m in disks:
        name = "SSD/HDD health" if len(disks) == 1 else f"SSD/HDD health ({n})"
        if n.startswith("nvme"):
            try:
                s = nvme_smart(n)
            except OSError as e:
                add(name, "NOT_TESTED", f"NVMe SMART read failed: {e}")
                continue
            life = max(0, 100 - s["pct_used"])
            INFO.setdefault("ssd_life", life)
            det = (f"life remaining={life}%; critical_warning={s['critical_warning']}; media_errors={s['media_errors']}; "
                   f"spare={s['spare']}% (threshold {s['spare_thr']}%); temp={s['temp_c']}C; "
                   f"power_on_hours={s['power_on_hours']}; written={s['tb_written']:.1f} TB; "
                   f"unsafe_shutdowns={s['unsafe_shutdowns']}")
            if s["critical_warning"] or s["media_errors"] or s["spare"] < s["spare_thr"] or s["pct_used"] >= 100:
                st = "FAIL"
            elif s["pct_used"] >= CFG["ssd_warn_used"] or (s["temp_c"] or 0) >= CFG["ssd_temp_warn"]:
                st = "WARN"
            else:
                st = "PASS"
            add(name, st, det)
        elif cmdexists("smartctl"):
            rc, out, _ = run(["smartctl", "-H", "-A", "/dev/" + n], timeout=60)
            if "PASSED" in out or "OK" in out:
                add(name, "PASS", "SMART overall-health: PASSED")
            elif "FAILED" in out:
                add(name, "FAIL", "SMART overall-health: FAILED")
            else:
                add(name, "NOT_TESTED", "SMART status not reported")
        else:
            add(name, "NOT_TESTED", "SATA disk: install smartmontools for SMART (sudo apt install smartmontools)")


# ---------------------------------------------------------------- controllers / display
def test_devices():
    title("Controllers and display")
    rc, lspci, _ = run(["lspci", "-nn"]) if cmdexists("lspci") else (127, "", "")
    for name, needle in [("Wi-Fi controller", r"network controller|wireless"),
                         ("Graphics controller", r"vga compatible controller|3d controller|display controller"),
                         ("Audio controller", r"audio device|multimedia audio")]:
        if not lspci:
            add(name, "NOT_TESTED", "lspci unavailable")
        else:
            m = [l.split(": ", 1)[-1] for l in lspci.splitlines() if re.search(needle, l, re.I)]
            add(name, "PASS" if m else "FAIL", (m[0][:120] if m else "not found on PCI bus"))
    modes = []
    for p in glob.glob("/sys/class/drm/card*-eDP*/modes") or glob.glob("/sys/class/drm/card*-*/modes"):
        modes.extend(read(p).splitlines())
    add("Display panel detected", "PASS" if modes else "WARN", ", ".join(sorted(set(modes))[:3]) or "Not exposed")


# ---------------------------------------------------------------- Wi-Fi
def load_config():
    cp = configparser.ConfigParser()
    cp.read(BASE / "config.ini")
    return cp


def wifi_connected(iface):
    rc, out, _ = run(["nmcli", "-t", "-f", "DEVICE,STATE", "dev"])
    return f"{iface}:connected" in out.splitlines()


def test_wifi():
    title("Wi-Fi")
    ifaces = [p.name for p in pathlib.Path("/sys/class/net").glob("*") if (p / "wireless").exists() or (p / "phy80211").exists()]
    if not ifaces:
        add("Wi-Fi connectivity", "FAIL", "No wireless interface")
        return
    iface = ifaces[0]
    run(["rfkill", "unblock", "wifi"])
    if not cmdexists("nmcli"):
        add("Wi-Fi connectivity", "NOT_TESTED", "nmcli unavailable")
        return
    run(["nmcli", "radio", "wifi", "on"])
    time.sleep(2)
    info("Scanning for Wi-Fi networks...")
    rc, out, err = run(["nmcli", "-t", "-f", "SSID,SIGNAL", "dev", "wifi", "list", "ifname", iface, "--rescan", "yes"], 40)
    nets = {}
    for line in out.splitlines():
        ssid, _, sig = line.rpartition(":")
        ssid = ssid.replace("\\:", ":")
        if ssid:
            nets[ssid] = max(nets.get(ssid, 0), int(sig) if sig.isdigit() else 0)
    if not nets:
        add("Wi-Fi connectivity", "FAIL", f"Radio found no networks on {iface} (rfkill/antenna?) {err[:100]}")
        return
    best = max(nets.values())
    connected = wifi_connected(iface)
    if not connected:
        cp = load_config()
        ssid = cp.get("wifi", "ssid", fallback="").strip()
        pw = cp.get("wifi", "password", fallback="")
        if not ssid and not AUTO:
            flush_stdin()
            ssid = input(f"   {C['b']}>> Wi-Fi network name to test connection (Enter = radio test only): {C['x']}").strip()
            if ssid:
                pw = getpass.getpass("      Password: ")
        if ssid:
            info(f"Connecting to '{ssid}'...")
            cmd = ["nmcli", "dev", "wifi", "connect", ssid, "ifname", iface] + (["password", pw] if pw else [])
            run(cmd, 60)
            time.sleep(3)
            connected = wifi_connected(iface)
    base = f"{len(nets)} networks visible (best signal {best}%)"
    if not connected:
        add("Wi-Fi connectivity", "WARN", base + "; radio OK, not connected (no internet test)")
        return
    rc, out, _ = run(["ping", "-c", "3", "-W", "3", "-I", iface, "1.1.1.1"], 20)
    if rc == 0:
        m = re.search(r"= [\d.]+/([\d.]+)/", out)
        add("Wi-Fi connectivity", "PASS", base + f"; connected, internet OK (ping {m.group(1) if m else '?'} ms)")
    else:
        add("Wi-Fi connectivity", "WARN", base + "; connected but internet ping failed")


# ---------------------------------------------------------------- Bluetooth
def test_bluetooth():
    title("Bluetooth")
    if not glob.glob("/sys/class/bluetooth/hci*"):
        add("Bluetooth functionality", "FAIL", "No Bluetooth controller detected")
        return
    run(["rfkill", "unblock", "bluetooth"])
    if not cmdexists("bluetoothctl"):
        add("Bluetooth functionality", "NOT_TESTED", "bluetoothctl unavailable")
        return
    pause("Place a PHONE with Bluetooth ON (on its Bluetooth screen) near the laptop and press Enter")
    run(["bluetoothctl", "power", "on"], 15)
    time.sleep(1)
    rc, show, _ = run(["bluetoothctl", "show"], 10)
    if "Powered: yes" not in show:
        add("Bluetooth functionality", "FAIL", "Controller present but failed to power on")
        return
    info(f"Scanning for {CFG['bt_scan']} s...")
    run(["bluetoothctl", "--timeout", str(CFG["bt_scan"]), "scan", "on"], CFG["bt_scan"] + 15)
    rc, out, _ = run(["bluetoothctl", "devices"], 10)
    devs = [l.split(" ", 2)[-1] for l in out.splitlines() if l.startswith("Device ")]
    if devs:
        add("Bluetooth functionality", "PASS", f"Controller on; {len(devs)} device(s) found: " + ", ".join(devs[:5]))
    else:
        add("Bluetooth functionality", "WARN", "Controller on, but no nearby device found (put a phone in pairing screen)")


# ---------------------------------------------------------------- input devices (keyboard/touchpad)
EV_FMT = "llHHi"
EV_SZ = struct.calcsize(EV_FMT)


def input_devices():
    devs, cur = [], {}
    for line in read("/proc/bus/input/devices").splitlines() + [""]:
        if not line.strip():
            if cur:
                devs.append(cur)
            cur = {}
            continue
        if line.startswith("I:"):
            m = re.search(r"Bus=(\w+)", line)
            cur["bus"] = m.group(1) if m else ""
        elif line.startswith("N:"):
            cur["name"] = line.split("=", 1)[1].strip().strip('"')
        elif line.startswith("H:"):
            h = line.split("=", 1)[1].split()
            cur["handlers"] = h
            cur["event"] = next((x for x in h if x.startswith("event")), None)
    return devs


def open_inputs(events):
    fds = []
    for e in events:
        try:
            fd = os.open("/dev/input/" + e, os.O_RDONLY | os.O_NONBLOCK)
        except OSError:
            continue
        try:
            fcntl.ioctl(fd, EVIOCGRAB, 1)   # keys go only to this test, not to the desktop
        except OSError:
            pass
        fds.append(fd)
    return fds


def close_inputs(fds):
    for fd in fds:
        try:
            fcntl.ioctl(fd, EVIOCGRAB, 0)
        except OSError:
            pass
        os.close(fd)


def read_events(fds, timeout):
    out = []
    r, _, _ = select.select(fds, [], [], timeout)
    for fd in r:
        try:
            data = os.read(fd, EV_SZ * 64)
        except (BlockingIOError, OSError):
            continue
        for i in range(0, len(data) - EV_SZ + 1, EV_SZ):
            _, _, t, c, v = struct.unpack_from(EV_FMT, data, i)
            out.append((t, c, v))
    return out


KEY_ROWS = [
    [("Esc", {1}), ("F1", {59, 113}), ("F2", {60, 114}), ("F3", {61, 115}), ("F4", {62, 248}),
     ("F5", {63, 228, 229, 230}), ("F6", {64, 224}), ("F7", {65, 225}), ("F8", {66, 227, 431}),
     ("F9", {67, 120, 204, 217, 240}), ("F10", {68, 99, 210}), ("F11", {87, 102}), ("F12", {88, 107}), ("Del", {111})],
    [("`", {41})] + [(str((i + 1) % 10), {2 + i}) for i in range(10)] + [("-", {12}), ("=", {13}), ("Backspace", {14})],
    [("Tab", {15})] + [(k, {16 + i}) for i, k in enumerate("QWERTYUIOP")] + [("[", {26}), ("]", {27}), ("\\", {43})],
    [("Caps", {58})] + [(k, {30 + i}) for i, k in enumerate("ASDFGHJKL")] + [(";", {39}), ("'", {40}), ("Enter", {28})],
    [("Shift", {42})] + [(k, {44 + i}) for i, k in enumerate("ZXCVBNM")] + [(",", {51}), (".", {52}), ("/", {53}), ("Shift-R", {54})],
    [("Ctrl", {29}), ("Win", {125}), ("Alt", {56}), ("Space", {57}), ("Alt-R", {100}), ("Ctrl-R", {97}),
     ("Left", {105}), ("Up", {103}), ("Down", {108}), ("Right", {106})],
]
ALL_KEYS = [l for row in KEY_ROWS for l, _ in row]
FROW = {f"F{i}" for i in range(1, 13)}
CODE2LABEL = {c: l for row in KEY_ROWS for l, codes in row for c in codes}


def kb_render(done, na, last, left):
    out = ["\033[H\033[2J",
           f"{C['b']}KEYBOARD TEST{C['x']} - press EACH key once (F1-F12 with or without Fn).",
           f"Green = OK.  Esc 3x in a row = finish.  Time left: {left}s", ""]
    for row in KEY_ROWS:
        cells = []
        for l, _ in row:
            col = C["g"] if l in done else (C["d"] if l in na else C["x"])
            cells.append(f"{col}[{l}]{C['x']}")
        out.append("  " + " ".join(cells))
    rem = [l for l in ALL_KEYS if l not in done and l not in na]
    out += ["", f"{len(rem)} keys left.  Last key: {last}", "(The keyboard is captured by the test - the desktop does not respond right now.)"]
    sys.stdout.write("\n".join(out) + "\n")
    sys.stdout.flush()


def keyboard_loop(fds, done, na):
    start = time.time()
    esc, last, last_draw = [], "-", 0
    while True:
        left = int(CFG["keyboard_timeout"] - (time.time() - start))
        if left <= 0:
            return "timeout"
        changed = False
        for t, c, v in read_events(fds, 0.3):
            if t != 1 or v != 1:
                continue
            if c == 1:
                now = time.time()
                esc = [x for x in esc if now - x < 3] + [now]
                if len(esc) >= 3:
                    done.add("Esc")
                    return "esc"
            lbl = CODE2LABEL.get(c)
            if lbl:
                done.add(lbl)
                last = lbl
            else:
                last = f"(code {c})"
            changed = True
        if all(l in done or l in na for l in ALL_KEYS):
            kb_render(done, na, last, left)
            return "complete"
        if changed or time.time() - last_draw > 1:
            kb_render(done, na, last, left)
            last_draw = time.time()


def test_keyboard():
    title("Keyboard")
    devs = [d for d in input_devices() if d.get("event") and "kbd" in d.get("handlers", [])
            and d.get("bus") not in ("0003", "0005")
            and not re.search(r"power button|sleep button|lid switch", d.get("name", ""), re.I)]
    if not devs:
        PARTS["keyboard"] = ("FAIL", "Internal keyboard not detected")
        return
    info("Devices: " + ", ".join(d.get("name", "?") for d in devs))
    info("Unplug any external USB keyboards before the test.")
    pause("Press Enter to start the keyboard test")
    done, na = set(), set()
    while True:
        fds = open_inputs([d["event"] for d in devs])
        try:
            reason = keyboard_loop(fds, done, na)
        finally:
            close_inputs(fds)
            time.sleep(0.3)
            flush_stdin()
        missing = [l for l in ALL_KEYS if l not in done and l not in na]
        print(f"\n   Test finished ({reason}).")
        if not missing:
            break
        print(f"   {C['y']}Keys that did NOT register: {' '.join(missing)}{C['x']}")
        a = choice("[d]efect confirmed / [n]ot present on this model / [r]etry only these keys", "dnr", "d")
        if a == "r":
            continue
        if a == "n":
            flush_stdin()
            t = input("      Type the keys that do NOT exist (space separated; Enter = all above): ").split()
            na.update(missing if not t else [x for x in t if x in missing])
            if all(l in done or l in na for l in ALL_KEYS):
                break
            continue
        break
    missing = [l for l in ALL_KEYS if l not in done and l not in na]
    main_missing = [m for m in missing if m not in FROW]
    f_missing = [m for m in missing if m in FROW]
    det = f"{len(done)}/{len(ALL_KEYS) - len(na)} keys OK"
    if main_missing:
        det += "; FAILED keys: " + " ".join(main_missing)
    if f_missing:
        det += "; F-row not registering: " + " ".join(f_missing)
    if na:
        det += "; not present on model: " + " ".join(sorted(na))
    st = "FAIL" if main_missing else ("WARN" if f_missing else "PASS")
    PARTS["keyboard"] = (st, det)
    print(f"   {SCOLOR[st]}Keyboard: {st}{C['x']} - {det}")


def test_touchpad():
    title("Touchpad")
    tps = [d for d in input_devices() if d.get("event") and "touchpad" in d.get("name", "").lower()]
    if not tps:
        PARTS["touchpad"] = ("FAIL", "Touchpad not detected")
        return
    pause("Press Enter, then: (1) SLIDE your finger, (2) do a physical CLICK, (3) touch with TWO fingers and scroll")
    goals = {"move": 0, "click": False, "two": False}
    while True:
        fds = open_inputs([d["event"] for d in tps])
        start = time.time()
        try:
            while time.time() - start < CFG["touchpad_timeout"]:
                for t, c, v in read_events(fds, 0.3):
                    if t == 3:
                        goals["move"] += 1
                    elif t == 1 and c == 272 and v == 1:
                        goals["click"] = True
                    elif t == 1 and c == 333 and v == 1:
                        goals["two"] = True
                ok = lambda b: f"{C['g']}OK{C['x']}" if b else "..."
                sys.stdout.write(f"\r   Movement: {ok(goals['move'] > 40)}   Click: {ok(goals['click'])}   "
                                 f"Two fingers: {ok(goals['two'])}   ({int(CFG['touchpad_timeout'] - (time.time() - start))}s) ")
                sys.stdout.flush()
                if goals["move"] > 40 and goals["click"] and goals["two"]:
                    break
        finally:
            close_inputs(fds)
            print()
            flush_stdin()
        miss = [k for k, okv in (("movement", goals["move"] > 40), ("click", goals["click"]), ("two-finger", goals["two"])) if not okv]
        if not miss:
            PARTS["touchpad"] = ("PASS", "movement, physical click and two-finger touch OK")
            print(f"   {C['g']}Touchpad: PASS{C['x']}")
            return
        a = choice(f"Not registered: {', '.join(miss)}. [r]etry / [f]ailed", "rf", "f")
        if a != "r":
            PARTS["touchpad"] = ("FAIL", "not registering: " + ", ".join(miss))
            return


# ---------------------------------------------------------------- audio
def find_tool(*names):
    return next((n for n in names if cmdexists(n)), None)


def make_tone(path, side):
    rate, n = 44100, 44100
    frames = bytearray()
    for i in range(n):
        env = min(1.0, i / 2000, (n - i) / 2000)
        beep = 1 if (i // 8820) % 2 == 0 else 0          # 3 short beeps
        s = int(22000 * env * beep * math.sin(2 * math.pi * 880 * i / rate))
        frames += struct.pack("<hh", s if side != "right" else 0, s if side != "left" else 0)
    with wave.open(path, "wb") as w:
        w.setnchannels(2)
        w.setsampwidth(2)
        w.setframerate(rate)
        w.writeframes(bytes(frames))
    os.chmod(path, 0o666)


def play(path):
    tool = find_tool("pw-play", "paplay", "aplay")
    if not tool:
        return False, "no audio player found"
    rc, _, err = run(user_cmd([tool, path]), 30)
    return rc == 0, err


def record(path, secs):
    if cmdexists("arecord"):
        rc, _, err = run(user_cmd(["arecord", "-q", "-d", str(secs), "-f", "S16_LE", "-r", "16000", "-c", "1", path]), secs + 15)
        if rc == 0 and os.path.exists(path):
            return True, ""
    tool = find_tool("pw-record", "parecord")
    if not tool:
        return False, "no recorder found"
    cmd = [tool, "--rate", "16000", "--channels", "1", path] if tool == "pw-record" else \
        ["parecord", "--rate=16000", "--channels=1", "--format=s16le", "--file-format=wav", path]
    p = subprocess.Popen(user_cmd(cmd), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(secs)
    p.send_signal(signal.SIGINT)
    try:
        p.wait(5)
    except subprocess.TimeoutExpired:
        p.kill()
    return os.path.exists(path) and os.path.getsize(path) > 1000, ""


def wav_level(path):
    data, sw = b"", 2
    try:
        with wave.open(path, "rb") as w:
            sw = w.getsampwidth()
            data = w.readframes(w.getnframes())
    except (wave.Error, EOFError, OSError):
        pass
    if not data:
        try:
            data = pathlib.Path(path).read_bytes()[44:]
            sw = 2
        except OSError:
            return None
    if sw != 2 or len(data) < 3200:
        return None
    a = array.array("h")
    a.frombytes(data[3200:len(data) // 2 * 2])     # skip first 0.1 s
    if not a:
        return None
    rms = math.sqrt(sum(x * x for x in a) / len(a))
    return -120.0 if rms < 1 else 20 * math.log10(rms / 32768)


def audio_prepare():
    if cmdexists("wpctl"):
        for c in (["set-mute", "@DEFAULT_AUDIO_SINK@", "0"], ["set-volume", "@DEFAULT_AUDIO_SINK@", "0.7"],
                  ["set-mute", "@DEFAULT_AUDIO_SOURCE@", "0"], ["set-volume", "@DEFAULT_AUDIO_SOURCE@", "1.0"]):
            run(user_cmd(["wpctl", *c]), 10)
    elif cmdexists("pactl"):
        for c in (["set-sink-mute", "@DEFAULT_SINK@", "0"], ["set-sink-volume", "@DEFAULT_SINK@", "70%"],
                  ["set-source-mute", "@DEFAULT_SOURCE@", "0"], ["set-source-volume", "@DEFAULT_SOURCE@", "100%"]):
            run(user_cmd(["pactl", *c]), 10)


def test_audio(tmp):
    title("Speakers and microphone")
    audio_prepare()
    res = {}
    for side, label in (("left", "LEFT"), ("right", "RIGHT")):
        f = os.path.join(tmp, f"tone_{side}.wav")
        make_tone(f, side)
        while True:
            info(f"Playing 3 beeps on the {label} speaker only...")
            ok, err = play(f)
            if not ok:
                info(f"{C['y']}Playback error: {err[:150]}{C['x']}")
            a = choice(f"Did you hear the beeps on the {label} side? [y]es / [n]o / [r]etry / [s]kip", "ynrs", "s")
            if a != "r":
                res[side] = a
                break
    spk = {"y": "OK", "n": "FAIL", "s": "skipped"}
    spk_st = "FAIL" if "n" in res.values() else ("NOT_TESTED" if "s" in res.values() else "PASS")
    spk_det = f"left speaker {spk[res['left']]}, right speaker {spk[res['right']]}"

    mic_st, mic_det = "NOT_TESTED", "skipped"
    while True:
        f = os.path.join(tmp, "mic.wav")
        if os.path.exists(f):
            os.remove(f)
        a = choice(f"Microphone: press Enter and SPEAK out loud for {CFG['mic_seconds']} s (or [s] to skip)", "s", "go")
        if a == "s":
            break
        info("Recording... speak now!")
        ok, err = record(f, CFG["mic_seconds"])
        lvl = wav_level(f) if os.path.exists(f) else None
        if lvl is None:
            info(f"{C['y']}Recording failed {err[:150]}{C['x']}")
            a = choice("[r]etry / [f]ailed / [s]kip", "rfs", "f")
            if a == "r":
                continue
            mic_st, mic_det = ("FAIL", "recording failed") if a == "f" else ("NOT_TESTED", "skipped")
            break
        low = lvl < CFG["mic_min_dbfs"]
        info(f"Captured level: {lvl:.1f} dBFS" + (f"  {C['y']}(very low - microphone muted? F4 key/LED){C['x']}" if low else ""))
        info("Playing back the recording...")
        play(f)
        a = choice("Did you hear your voice in the playback? [y]es / [n]o / [r]etry / [s]kip", "ynrs", "s")
        if a == "r":
            continue
        if a == "s":
            break
        if a == "y":
            mic_st = "WARN" if low else "PASS"
        else:
            mic_st = "FAIL"
        mic_det = f"mic level {lvl:.1f} dBFS, playback {'heard' if a == 'y' else 'NOT heard'}"
        break
    add("Speaker/microphone functionality", worst(spk_st, mic_st), f"{spk_det}; {mic_det}")


# ---------------------------------------------------------------- webcam
def v4l2_capture(dev, frames=15):
    fd = os.open(dev, os.O_RDWR | os.O_NONBLOCK)
    maps = []
    try:
        cap = bytearray(104)
        fcntl.ioctl(fd, VIDIOC_QUERYCAP, cap, True)
        card = cap[16:48].split(b"\0")[0].decode(errors="replace")
        caps, dcaps = struct.unpack_from("<II", cap, 84)
        if caps & 0x80000000:
            caps = dcaps
        if not (caps & 0x1) or not (caps & 0x04000000):
            return None, card, "not a capture node"
        chosen = None
        for four in (b"YUYV", b"MJPG"):
            want = int.from_bytes(four, "little")
            fmt = bytearray(208)
            struct.pack_into("<I", fmt, 0, 1)
            struct.pack_into("<IIII", fmt, 8, 640, 480, want, 1)
            try:
                fcntl.ioctl(fd, VIDIOC_S_FMT, fmt, True)
            except OSError:
                continue
            w, h, pf, _, bpl, _ = struct.unpack_from("<IIIIII", fmt, 8)
            if pf == want:
                chosen = (four.decode(), w, h, bpl)
                break
        if not chosen:
            return None, card, "no YUYV/MJPG format (IR camera?)"
        req = bytearray(20)
        struct.pack_into("<III", req, 0, 4, 1, 1)
        fcntl.ioctl(fd, VIDIOC_REQBUFS, req, True)
        for i in range(struct.unpack_from("<I", req, 0)[0]):
            b = bytearray(88)
            struct.pack_into("<II", b, 0, i, 1)
            struct.pack_into("<I", b, 60, 1)
            fcntl.ioctl(fd, VIDIOC_QUERYBUF, b, True)
            off, length = struct.unpack_from("<I", b, 64)[0], struct.unpack_from("<I", b, 72)[0]
            maps.append(mmap.mmap(fd, length, mmap.MAP_SHARED, mmap.PROT_READ | mmap.PROT_WRITE, offset=off))
            fcntl.ioctl(fd, VIDIOC_QBUF, b, True)
        fcntl.ioctl(fd, VIDIOC_STREAMON, struct.pack("<I", 1))
        data = None
        for _ in range(frames):
            r, _, _ = select.select([fd], [], [], 3)
            if not r:
                break
            b = bytearray(88)
            struct.pack_into("<II", b, 0, 0, 1)
            struct.pack_into("<I", b, 60, 1)
            fcntl.ioctl(fd, VIDIOC_DQBUF, b, True)
            idx, _, used = struct.unpack_from("<III", b, 0)
            data = bytes(maps[idx][:used])
            fcntl.ioctl(fd, VIDIOC_QBUF, b, True)
        try:
            fcntl.ioctl(fd, VIDIOC_STREAMOFF, struct.pack("<I", 1))
        except OSError:
            pass
        if not data:
            return None, card, "no frames received"
        return (chosen, data), card, ""
    finally:
        for m in maps:
            m.close()
        os.close(fd)


def write_png(path, w, h, raw):
    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    png = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(bytes(raw), 6)) + chunk(b"IEND", b""))
    pathlib.Path(path).write_bytes(png)


def yuyv_to_png(data, w, h, bpl, path):
    clip = lambda x: 0 if x < 0 else (255 if x > 255 else x)
    raw, ysum = bytearray(), 0
    for y in range(h):
        raw.append(0)
        row = data[y * bpl:y * bpl + w * 2]
        for x in range(0, len(row) - 3, 4):
            y0, u, y1, v = row[x], row[x + 1], row[x + 2], row[x + 3]
            d, e = u - 128, v - 128
            ysum += y0 + y1
            for yy in (y0, y1):
                c = 298 * (yy - 16)
                raw += bytes((clip((c + 409 * e + 128) >> 8), clip((c - 100 * d - 208 * e + 128) >> 8),
                              clip((c + 516 * d + 128) >> 8)))
    write_png(path, w, h, raw)
    return ysum / max(1, w * h)


def test_webcam(tmp):
    title("Webcam")
    nodes = sorted(glob.glob("/dev/video*"), key=lambda p: int(re.sub(r"\D", "", p) or 0))
    if not nodes:
        add("Webcam functionality", "FAIL", "No camera device found (camera disabled in BIOS / not connected)")
        return
    pause("OPEN the camera privacy shutter (above the screen) and press Enter")
    while True:
        errors, result = [], None
        for n in nodes:
            try:
                res, card, err = v4l2_capture(n)
            except OSError as e:
                errors.append(f"{n}: {e}")
                continue
            if res:
                result = (n, card, res)
                break
            errors.append(f"{n} ({card}): {err}")
        if not result:
            if any("ipu" in e.lower() for e in errors):
                add("Webcam functionality", "NOT_TESTED",
                    "MIPI (Intel IPU) camera - not supported in Linux live; test with Windows Camera app")
                return
            info(f"{C['y']}Could not capture: {'; '.join(errors)[:300]}{C['x']}")
            a = choice("[r]etry / [f]ailed / [s]kip", "rfs", "f")
            if a == "r":
                continue
            add("Webcam functionality", "FAIL" if a == "f" else "NOT_TESTED",
                "capture failed: " + "; ".join(errors)[:300] if a == "f" else "skipped")
            return
        n, card, ((fmt, w, h, bpl), data) = result
        luma = None
        if fmt == "YUYV":
            path = os.path.join(tmp, "webcam.png")
            luma = yuyv_to_png(data, w, h, bpl, path)
        else:
            path = os.path.join(tmp, "webcam.jpg")
            pathlib.Path(path).write_bytes(data)
        os.chmod(path, 0o666)
        subprocess.Popen(user_cmd(["xdg-open", path]), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        info(f"Photo captured ({card}, {w}x{h}) and opened in a window: {path}")
        info("Check the image, CLOSE the photo window and come back to this Terminal.")
        if luma is not None and luma < 15:
            info(f"{C['y']}Image almost black - is the shutter closed?{C['x']}")
        a = choice("Is the image OK (sharp, in color)? [y]es / [n]o / [r]etry / [s]kip", "ynrs", "s")
        if a == "r":
            continue
        st = {"y": "PASS", "n": "FAIL", "s": "NOT_TESTED"}[a]
        det = f"{card} {w}x{h} {fmt}; image " + {"y": "approved by technician", "n": "REJECTED by technician", "s": "not reviewed"}[a]
        if luma is not None:
            det += f"; brightness {luma:.0f}/255"
        add("Webcam functionality", st, det)
        return


# ---------------------------------------------------------------- USB ports
def usb_snapshot():
    out = {}
    for d in pathlib.Path("/sys/bus/usb/devices").glob("*"):
        if not (d / "idVendor").exists() or read(d / "bDeviceClass") == "09":
            continue
        name = read(d / "product") or f"{read(d / 'idVendor')}:{read(d / 'idProduct')}"
        out[d.name] = (name, read(d / "speed"))
    return out


def test_usb():
    title("USB ports")
    info("Use a test device (USB mouse, flash drive, phone with cable). Do NOT remove the Ubuntu boot drive.")
    flush_stdin()
    a = input(f"   {C['b']}>> How many free USB ports will you test? (Enter = 3, 0 = skip): {C['x']}").strip()
    n = int(a) if a.isdigit() else 3
    if n == 0:
        add("USB port functionality", "NOT_TESTED", "skipped by technician")
        return
    results = []
    for i in range(1, n + 1):
        while True:
            base = usb_snapshot()
            info(f"Port {i}/{n}: PLUG IN the device now (waiting {CFG['usb_wait']} s)...")
            found, t0 = None, time.time()
            while time.time() - t0 < CFG["usb_wait"] and not found:
                time.sleep(0.5)
                new = {k: v for k, v in usb_snapshot().items() if k not in base}
                if new:
                    time.sleep(1)
                    new = {k: v for k, v in usb_snapshot().items() if k not in base} or new
                    found = next(iter(new.values()))
            if found:
                info(f"{C['g']}Port {i} OK{C['x']}: {found[0]} ({found[1]} Mbps). Now REMOVE the device.")
                t0 = time.time()
                while time.time() - t0 < CFG["usb_wait"] and any(k not in base for k in usb_snapshot()):
                    time.sleep(0.5)
                results.append((i, "OK", f"{found[0]} {found[1]}Mbps"))
                break
            a = choice(f"Nothing detected on port {i}. [r]etry / [f]ailed / [s]kip", "rfs", "f")
            if a == "r":
                continue
            results.append((i, "FAIL" if a == "f" else "SKIP", "no device detected" if a == "f" else ""))
            break
    if any(r[1] == "FAIL" for r in results):
        st = "FAIL"
    elif all(r[1] == "OK" for r in results):
        st = "PASS"
    else:
        st = "NOT_TESTED"
    add("USB port functionality", st, "; ".join(f"port {i}: {s} {d}".strip() for i, s, d in results)
        + "; (port holding the boot USB not tested)")


# ---------------------------------------------------------------- memory
def test_memory(mem_gib=None):
    title("Memory (quick pattern test)")
    avail = 0
    for line in read("/proc/meminfo").splitlines():
        if line.startswith("MemAvailable:"):
            avail = int(re.search(r"\d+", line).group()) * 1024
    MB = 1 << 20
    size = int(mem_gib * (1 << 30)) if mem_gib else int(min(avail * CFG["mem_fraction"], CFG["mem_max_gib"] * (1 << 30)))
    size = size // MB * MB
    if size < 256 * MB:
        add("Memory stress test", "NOT_TESTED", f"not enough free RAM ({avail / 2**30:.1f} GiB available)")
        return
    nb = size // MB
    pats = [b"\x00" * MB, b"\xff" * MB, b"\x55" * MB, b"\xaa" * MB, os.urandom(MB)]
    try:
        buf = bytearray(size)
    except MemoryError:
        add("Memory stress test", "NOT_TESTED", "could not allocate test buffer")
        return
    errors = 0
    t0 = time.time()
    for pi, pat in enumerate(pats, 1):
        for i in range(nb):
            buf[i * MB:(i + 1) * MB] = pat
            if i % 256 == 0:
                sys.stdout.write(f"\r   Pattern {pi}/{len(pats)}: writing {i * 100 // nb:3d}%   ")
                sys.stdout.flush()
        for i in range(nb):
            if buf[i * MB:(i + 1) * MB] != pat:
                errors += 1
            if i % 256 == 0:
                sys.stdout.write(f"\r   Pattern {pi}/{len(pats)}: verifying {i * 100 // nb:3d}%")
                sys.stdout.flush()
    del buf
    print()
    rc, dm, _ = run(["dmesg"], 15)
    hw = [l for l in dm.splitlines() if re.search(r"machine check|hardware error|edac.*(ce|ue)|mce:", l, re.I)]
    det = (f"{size / 2**30:.1f} GiB tested, {len(pats)} patterns, {errors} bad blocks, {time.time() - t0:.0f}s"
           f"; kernel HW error messages: {len(hw)} (quick screen - Memtest86+ is the full test)")
    add("Memory stress test", "FAIL" if errors else ("WARN" if hw else "PASS"), det)


# ---------------------------------------------------------------- reports
REQUIRED = ["Battery health", "Memory stress test", "Keyboard/touchpad functionality", "Webcam functionality",
            "Speaker/microphone functionality", "USB port functionality", "Wi-Fi connectivity",
            "Bluetooth functionality", "SSD/HDD health"]
FLEET_FIELDS = ["timestamp_utc", "internal_uid", "service_tag", "manufacturer", "model", "cpu", "ram_gib", "storage",
                "battery_health_pct", "battery_cycles", "ssd_life_pct", "verdict", "fail_items", "warn_items",
                "not_tested_items", "technician", "report_file"]
VERDICT_TXT = {"PASS": "PASSED", "PASS_WITH_WARNINGS": "PASSED WITH NOTES", "FAIL": "FAILED",
               "INCOMPLETE": "INCOMPLETE (items not tested)"}


def fmt_local(dt):
    return f"{dt:%B} {dt.day}, {dt.year} - {dt.hour % 12 or 12}:{dt:%M} {'AM' if dt.hour < 12 else 'PM'} {dt.tzname()}"


def resolve_test_time():
    """Correct UTC time for the report. Ubuntu Live reads the BIOS clock as UTC, but laptops
    that ran Windows keep LOCAL time in the BIOS - so try network time first, then compensate."""
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo(CFG["timezone"])
    except Exception:
        tz = None
    utc = datetime.timezone.utc
    synced = False
    if cmdexists("timedatectl"):
        run(["timedatectl", "set-ntp", "true"], 10)
        for _ in range(12):
            if run(["timedatectl", "show", "-p", "NTPSynchronized", "--value"], 5)[1].strip() == "yes":
                synced = True
                break
            time.sleep(1)
    now = datetime.datetime.now(utc)
    if synced:
        INFO["clock_source"] = "network time (NTP)"
    elif tz is not None and not os.environ.get("GBS_RTC_UTC"):
        # BIOS clock holds local time (Windows default): reinterpret the wall clock as local time
        now = now.replace(tzinfo=None).replace(tzinfo=tz).astimezone(utc)
        INFO["clock_source"] = "laptop BIOS clock (no network; assumed local time)"
    else:
        INFO["clock_source"] = "laptop BIOS clock (no network)"
    INFO["test_date"] = fmt_local(now.astimezone(tz)) if tz is not None else now.strftime("%Y-%m-%d %H:%M UTC")
    return now.strftime("%Y%m%dT%H%M%SZ")


def verdict():
    res = [r for _, r, _ in checks]
    if "FAIL" in res:
        return "FAIL"
    if "NOT_TESTED" in res:
        return "INCOMPLETE"
    if "WARN" in res:
        return "PASS_WITH_WARNINGS"
    return "PASS"


def write_reports(stamp):
    OUT.mkdir(exist_ok=True)
    tag = INFO.get("service_tag", "UNKNOWN")
    stem = safe(INFO.get("internal_uid") or tag) + "_" + stamp
    v = verdict()
    fields = ["timestamp_utc", "internal_uid", "service_tag", "manufacturer", "model", "bios_version", "cpu",
              "ram_gib", "test", "result", "details"]
    rows = [dict(zip(fields, [stamp, INFO.get("internal_uid", ""), tag, INFO.get("manufacturer", ""),
                              INFO.get("model", ""), INFO.get("bios", ""), INFO.get("cpu", ""),
                              INFO.get("ram_gib", ""), n, s, d])) for n, s, d in checks]
    with (OUT / (stem + ".csv")).open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)
    with (OUT / (stem + ".json")).open("w", encoding="utf-8") as f:
        json.dump({"version": VERSION, "origin": "UBUNTU-" + ORIGIN, "stamp": stamp, "verdict": v, "device": INFO,
                   "pdf_file": stem + ".pdf", "checks": rows}, f, indent=2)

    def items(s):
        return " | ".join(n for n, r, _ in checks if r == s)
    fp = OUT / "fleet_summary.csv"
    if fp.exists():
        with fp.open(encoding="utf-8") as f:
            header = f.readline().strip()
        if header != ",".join(FLEET_FIELDS):
            fp.rename(OUT / f"fleet_summary_v1_{stamp}.csv")
    new = not fp.exists()
    with fp.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(FLEET_FIELDS)
        w.writerow([stamp, INFO.get("internal_uid", ""), tag, INFO.get("manufacturer", ""), INFO.get("model", ""),
                    INFO.get("cpu", ""), INFO.get("ram_gib", ""), INFO.get("storage", ""),
                    INFO.get("battery_health", ""), INFO.get("battery_cycles", ""), INFO.get("ssd_life", ""), v,
                    items("FAIL"), items("WARN"), items("NOT_TESTED"), INFO.get("technician", ""), stem + ".html"])
    # HTML (printable)
    col = {"PASS": "#1e8e3e", "WARN": "#b06000", "FAIL": "#c5221f", "NOT_TESTED": "#5f6368", "NOT_INSTALLED": "#5f6368"}
    vcol = {"PASS": "#1e8e3e", "PASS_WITH_WARNINGS": "#b06000", "FAIL": "#c5221f", "INCOMPLETE": "#5f6368"}[v]
    trs = "".join(f"<tr><td>{html.escape(n)}</td><td style='color:{col.get(s, '#000')};font-weight:700'>{s}</td>"
                  f"<td>{html.escape(d)}</td></tr>" for n, s, d in checks)
    e = lambda k: html.escape(str(INFO.get(k, "")))
    page = f"""<!doctype html><html><head><meta charset="utf-8"><title>GBS Diagnostics {e('service_tag')}</title>
<style>body{{font:14px/1.4 Arial,sans-serif;margin:24px;color:#202124}}h1{{font-size:20px;margin:0}}
.v{{padding:12px 16px;color:#fff;background:{vcol};font-size:22px;font-weight:700;margin:16px 0;border-radius:6px}}
table{{border-collapse:collapse;width:100%}}td,th{{border:1px solid #dadce0;padding:6px 8px;text-align:left;vertical-align:top}}
th{{background:#f1f3f4}}.meta td{{border:none;padding:2px 12px 2px 0}}</style></head><body>
<h1>GBS - Global Business Solution &middot; Laptop Diagnostic Report</h1>
<div class="v">{VERDICT_TXT[v]} ({v})</div>
<table class="meta"><tr><td><b>Internal UID</b></td><td>{e('internal_uid')}</td><td><b>Service tag</b></td><td>{e('service_tag')}</td></tr>
<tr><td><b>Model</b></td><td>{e('manufacturer')} {e('model')}</td><td><b>BIOS</b></td><td>{e('bios')}</td></tr>
<tr><td><b>CPU</b></td><td>{e('cpu')}</td><td><b>RAM</b></td><td>{e('ram_gib')} GiB</td></tr>
<tr><td><b>Storage</b></td><td>{e('storage')}</td><td><b>Test date</b></td><td>{e('test_date')}</td></tr>
<tr><td><b>Technician</b></td><td>{e('technician')}</td><td><b>Tool</b></td><td>v{VERSION}</td></tr></table>
<h2 style="font-size:16px">Results</h2><table><tr><th>Test</th><th>Result</th><th>Details</th></tr>{trs}</table>
<p style="color:#5f6368;font-size:12px">PASS = OK &middot; WARN = works with caveat &middot; FAIL = defect &middot; NOT_TESTED = not verified.
Battery: PASS &gt;{CFG['battery_pass']}% health, FAIL &le;{CFG['battery_pass']}%.
Memory test is a quick pattern screen, not a full Memtest86+ certification. Internal drives were not modified. Time source: {e('clock_source')}.</p></body></html>"""
    (OUT / (stem + ".html")).write_text(page, encoding="utf-8")
    write_pdf(OUT / (stem + ".pdf"), stamp, v)
    return stem, v


# ---------------------------------------------------------------- PDF (stdlib only) and upload
def _pdf_txt(s):
    s = str(s).encode("latin-1", "replace").decode("latin-1")
    return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _wrap(text, n):
    out, cur = [], ""
    for w in str(text).split():
        while len(w) > n:
            if cur:
                out.append(cur)
                cur = ""
            out.append(w[:n])
            w = w[n:]
        if not cur:
            cur = w
        elif len(cur) + 1 + len(w) <= n:
            cur += " " + w
        else:
            out.append(cur)
            cur = w
    if cur:
        out.append(cur)
    return out or [""]


def write_pdf(path, stamp, v):
    W, H, M = 612, 792, 40
    RGB = {"PASS": (0.12, 0.56, 0.24), "WARN": (0.69, 0.38, 0.0), "FAIL": (0.77, 0.13, 0.12),
           "NOT_TESTED": (0.37, 0.39, 0.41), "NOT_INSTALLED": (0.37, 0.39, 0.41),
           "PASS_WITH_WARNINGS": (0.69, 0.38, 0.0), "INCOMPLETE": (0.37, 0.39, 0.41)}
    pages = []
    st = {"ops": None, "y": 0}

    def new_page():
        st["ops"] = []
        pages.append(st["ops"])
        st["y"] = H - M

    def text(x, y, s, size=9, bold=False, rgb=(0, 0, 0)):
        st["ops"].append(f"{rgb[0]:.3f} {rgb[1]:.3f} {rgb[2]:.3f} rg BT /{'F2' if bold else 'F1'} {size} Tf "
                         f"{x:.1f} {y:.1f} Td ({_pdf_txt(s)}) Tj ET")

    def rect(x, y, w, h, rgb):
        st["ops"].append(f"{rgb[0]:.3f} {rgb[1]:.3f} {rgb[2]:.3f} rg {x:.1f} {y:.1f} {w:.1f} {h:.1f} re f")

    cols = [(M, 36), (M + 178, 14), (M + 250, 72)]       # x, wrap chars

    def header_row():
        rect(M, st["y"] - 14, W - 2 * M, 16, (0.93, 0.94, 0.95))
        for (x, _), t in zip(cols, ("Test", "Result", "Details")):
            text(x + 3, st["y"] - 10, t, 9, True)
        st["y"] -= 18

    new_page()
    text(M, st["y"] - 14, "GBS - Global Business Solution", 15, True)
    text(M, st["y"] - 30, "Laptop Diagnostic Report / Certificate of Testing", 11, False, (0.3, 0.3, 0.3))
    st["y"] -= 46
    rect(M, st["y"] - 30, W - 2 * M, 30, RGB.get(v, (0.3, 0.3, 0.3)))
    text(M + 10, st["y"] - 20, f"{VERDICT_TXT[v]}  ({v})", 15, True, (1, 1, 1))
    st["y"] -= 46
    meta = [("Internal UID", INFO.get("internal_uid", "")), ("Service tag", INFO.get("service_tag", "")),
            ("Model", f"{INFO.get('manufacturer', '')} {INFO.get('model', '')}"), ("BIOS", INFO.get("bios", "")),
            ("CPU", INFO.get("cpu", "")), ("RAM", f"{INFO.get('ram_gib', '')} GiB"),
            ("Storage", INFO.get("storage", "")), ("Battery health", f"{INFO.get('battery_health', '')}%"),
            ("Test date", INFO.get("test_date", stamp)), ("Technician", INFO.get("technician", ""))]
    for i in range(0, len(meta), 2):
        for j, (k, val) in enumerate(meta[i:i + 2]):
            x = M + j * 270
            text(x, st["y"], k + ":", 9, True)
            text(x + 78, st["y"], str(val)[:40], 9)
        st["y"] -= 14
    st["y"] -= 8
    header_row()
    for n, (name, res, det) in enumerate(checks):
        cells = [_wrap(name, cols[0][1]), [res], _wrap(det, cols[2][1])]
        h = max(len(c) for c in cells) * 10 + 4
        if st["y"] - h < M + 30:
            new_page()
            header_row()
        if n % 2:
            rect(M, st["y"] - h + 2, W - 2 * M, h, (0.97, 0.97, 0.98))
        for ci, ((x, _), lines) in enumerate(zip(cols, cells)):
            for li, ln in enumerate(lines):
                text(x + 3, st["y"] - 9 - li * 10, ln, 8, ci == 1, RGB.get(res, (0, 0, 0)) if ci == 1 else (0.1, 0.1, 0.1))
        st["y"] -= h
    st["y"] -= 10
    if st["y"] < M + 60:
        new_page()
    note = (f"PASS = OK; WARN = works with caveat; FAIL = defect; NOT_TESTED = not verified. Battery: PASS "
            f"> {CFG['battery_pass']}%, FAIL <= {CFG['battery_pass']}%. Memory test is a quick pattern screen. "
            f"Internal drives were not modified. Time source: {INFO.get('clock_source', '?')}. Tool v{VERSION}.")
    for i, ln in enumerate(_wrap(note, 110)):
        text(M, st["y"] - i * 10, ln, 7, False, (0.37, 0.39, 0.41))
    for i, ops in enumerate(pages, 1):
        st["ops"] = ops
        text(M, 22, f"GBS Diagnostic Report - {INFO.get('service_tag', '')} - {INFO.get('test_date', stamp)} - page {i}/{len(pages)}", 7, False, (0.5, 0.5, 0.5))
    # assemble PDF objects
    objs = ["<< /Type /Catalog /Pages 2 0 R >>", None,
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"]
    kids = []
    for ops in pages:
        stream = zlib.compress("\n".join(ops).encode("latin-1", "replace"))
        objs.append(f"<< /Length {len(stream)} /Filter /FlateDecode >>".encode() + b"\nstream\n" + stream + b"\nendstream")
        objs.append(f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {W} {H}] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
                    f"/Contents {len(objs)} 0 R >>")
        kids.append(f"{len(objs)} 0 R")
    objs[1] = f"<< /Type /Pages /Kids [{' '.join(kids)}] /Count {len(kids)} >>"
    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + (o if isinstance(o, bytes) else o.encode()) + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    out += "".join(f"{o:010d} 00000 n \n" for o in offs).encode()
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    pathlib.Path(path).write_bytes(bytes(out))


def upload(server, token, origin, stamp, v, pdf_path):
    payload = {"tool_version": VERSION, "origin": origin, "stamp": stamp, "verdict": v,
               "device": {k: INFO.get(k, "") for k in ("internal_uid", "service_tag", "manufacturer", "model", "bios", "cpu",
                                                       "ram_gib", "storage", "battery_health", "battery_cycles",
                                                       "ssd_life", "technician")},
               "checks": [{"test": n, "result": r, "details": d} for n, r, d in checks],
               "test_date_local": INFO.get("test_date", ""), "clock_source": INFO.get("clock_source", ""),
               "pdf_filename": os.path.basename(pdf_path),
               "pdf_base64": base64.b64encode(pathlib.Path(pdf_path).read_bytes()).decode()}
    req = urllib.request.Request(server.rstrip("/") + "/api/diagnostico", data=json.dumps(payload, default=str).encode(),
                                 headers={"Content-Type": "application/json", "X-GBS-Token": token}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return True, r.read().decode(errors="replace")[:200]
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}: {e.read().decode(errors='replace')[:200]}"
    except (urllib.error.URLError, OSError) as e:
        return False, str(e)


def print_summary(stem, v):
    print(f"\n{C['b']}{'=' * 70}\nSUMMARY - {INFO.get('model', '')} - {INFO.get('service_tag', '')}{C['x']}")
    for n, s, d in checks:
        if n in REQUIRED or n.startswith("Battery health") or n.startswith("SSD/HDD health") or s in ("FAIL", "WARN"):
            print(f"  {SCOLOR.get(s, '')}{s:<13}{C['x']} {n}")
    vc = {"PASS": C["g"], "PASS_WITH_WARNINGS": C["y"], "FAIL": C["r"], "INCOMPLETE": C["d"]}[v]
    print(f"\n  {C['b']}{vc}FINAL RESULT: {VERDICT_TXT[v]}{C['x']}")
    print(f"\n  Reports saved to: {OUT}/{stem}.pdf (.html/.csv/.json) and fleet_summary.csv\n")


# ---------------------------------------------------------------- main
def ensure_root():
    if os.geteuid() == 0:
        return
    user = pwd.getpwuid(os.getuid()).pw_name
    env = [f"GBS_USER={user}"] + [f"{k}={os.environ[k]}" for k in
                                  ("DISPLAY", "WAYLAND_DISPLAY", "XAUTHORITY", "XDG_RUNTIME_DIR",
                                   "GBS_SERVER", "GBS_TOKEN", "GBS_ORIGIN") if os.environ.get(k)]
    print("Requesting administrator permission (sudo)...")
    os.execvp("sudo", ["sudo", "env", *env, sys.executable, os.path.abspath(__file__), *sys.argv[1:]])


def main():
    global AUTO
    keys = ["battery", "storage", "devices", "wifi", "bluetooth", "keyboard", "touchpad", "audio", "webcam", "usb", "memory"]
    interactive = {"keyboard", "touchpad", "audio", "webcam", "usb"}
    ap = argparse.ArgumentParser(description="GBS laptop diagnostics")
    ap.add_argument("--auto", action="store_true", help="automatic tests only, no questions")
    ap.add_argument("--only", help="comma list: " + ",".join(keys))
    ap.add_argument("--skip", help="comma list of tests to skip")
    ap.add_argument("--mem-gib", type=float, help="size of memory screen in GiB")
    ap.add_argument("--server", help="GBS Diag Receiver URL, e.g. http://192.168.1.50:8090")
    ap.add_argument("--token", help="receiver token")
    ap.add_argument("--origin", default=None, help="UBUNTU-USB or UBUNTU-PXE (default from GBS_ORIGIN)")
    ap.add_argument("--no-root", action="store_true", help=argparse.SUPPRESS)
    args = ap.parse_args()
    if not args.no_root:
        ensure_root()
    AUTO = args.auto or not sys.stdin.isatty()
    sel = [k for k in keys if (not args.only or k in args.only.split(",")) and (not args.skip or k not in args.skip.split(","))]
    stamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    print(f"{C['b']}{C['c']}GBS Laptop Diagnostics v{VERSION}{C['x']}  (Ctrl+C cancels and saves what was already tested)")
    if not AUTO:
        INFO["internal_uid"] = input(f"{C['b']}Internal UID (GBS label) [Enter = skip]: {C['x']}").strip()
        INFO["technician"] = input(f"{C['b']}Technician name [Enter = skip]: {C['x']}").strip()
    tmp = tempfile.mkdtemp(prefix="gbsdiag_")
    os.chmod(tmp, 0o777)
    fns = {"battery": test_battery, "storage": test_storage, "devices": test_devices, "wifi": test_wifi,
           "bluetooth": test_bluetooth, "keyboard": test_keyboard, "touchpad": test_touchpad,
           "audio": lambda: test_audio(tmp), "webcam": lambda: test_webcam(tmp), "usb": test_usb,
           "memory": lambda: test_memory(args.mem_gib)}
    try:
        test_inventory()
        for k in sel:
            if AUTO and k in interactive:
                continue
            try:
                fns[k]()
            except KeyboardInterrupt:
                raise
            except Exception as e:  # never lose the whole run because of one test
                if k in ("keyboard", "touchpad"):
                    PARTS[k] = ("NOT_TESTED", f"error: {e}")
                else:
                    add(f"{k} test", "NOT_TESTED", f"error: {type(e).__name__}: {e}")
    except KeyboardInterrupt:
        print(f"\n{C['y']}Cancelled by technician - saving partial results.{C['x']}")
    finally:
        flush_stdin()
        if PARTS or "keyboard" in sel or "touchpad" in sel:
            kb = PARTS.get("keyboard", ("NOT_TESTED", "not run"))
            tp = PARTS.get("touchpad", ("NOT_TESTED", "not run"))
            add("Keyboard/touchpad functionality", worst(kb[0], tp[0]), f"Keyboard: {kb[0]} - {kb[1]} | Touchpad: {tp[0]} - {tp[1]}")
        for r in REQUIRED:
            if not any(n.startswith(r) for n, _, _ in checks):
                add(r, "NOT_TESTED", "not run" + (" (--auto mode)" if AUTO else ""))
        stamp = resolve_test_time()
        stem, v = write_reports(stamp)
        print_summary(stem, v)
        cp = load_config()
        server = args.server or os.environ.get("GBS_SERVER") or cp.get("server", "url", fallback="").strip()
        token = args.token or os.environ.get("GBS_TOKEN") or cp.get("server", "token", fallback="").strip()
        if server:
            while True:
                info(f"Sending result + PDF to {server} ...")
                ok, msg = upload(server, token, args.origin or ("UBUNTU-" + ORIGIN), stamp, v, OUT / (stem + ".pdf"))
                if ok:
                    print(f"   {C['g']}SENT and saved to the GBS database.{C['x']} {msg}")
                    break
                print(f"   {C['r']}Could NOT send: {msg}{C['x']}")
                if choice("[r]etry sending / [s]kip (the PDF stays saved locally)", "rs", "s") != "r":
                    break
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
