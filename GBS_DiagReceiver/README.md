# GBS_DiagReceiver

Console VB.NET (.NET Framework 4.8) que recebe os resultados do diagnóstico de laptops
(pen drive Ubuntu, boot PXE ou Windows) e grava no Oracle XE 10g local
(`PACK_DIAGNOSTICO.PROC_INSERT` → `TBL_DIAGNOSTICO`). Os resultados aparecem na aba
**Diagnostics** do GBS Inventory (Desktop).

Pré-requisito no banco: `Database\32_DIAGNOSTICO.sql`.

## Driver Oracle

Usa o **mesmo** `Oracle.ManagedDataAccess` 21.x do Desktop (`packages\Oracle.ManagedDataAccess.21.13.0`).
Não atualizar para 23.x.

Limitação comprovada no XE 10.2 com o 21.x: **colunas LOB em SELECT/cursor falham com
`TTC Error`**; LOB como parâmetro (IN e OUT) funciona. Por isso `PROC_LIST` não traz
`RESULT_JSON`/`PDF_BLOB`, e PDF/JSON são lidos por `PROC_GET_PDF` / `PROC_GET_JSON` (OUT).

## Configuração (`GBS_DiagReceiver.exe.config`)

| Chave | Padrão | Descrição |
|---|---|---|
| `connectionStrings/OracleDB` | `localhost:1521/XE` gbs_owner | mesma do Desktop |
| `Port` | `8090` | porta HTTP |
| `Token` | (vazio) | valor exigido no header `X-GBS-Token` — definir em `secrets.config` |
| `DiagFolder` | `diag` | pasta servida em `/diag/` (relativa ao .exe ou absoluta) |
| `MaxBodyMB` | `50` | limite do corpo do POST |

O token real fica em `secrets.config` (fora do git; copiado para o `bin` no build).
Sem ele o receptor não inicia. Formato:

```xml
<?xml version="1.0" encoding="utf-8"?>
<appSettings>
  <add key="Token" value="SEU_TOKEN_AQUI"/>
</appSettings>
```

## Liberar acesso pela rede (uma vez, como Administrador)

```bat
netsh http add urlacl url=http://+:8090/ user=Everyone
netsh advfirewall firewall add rule name="GBS DiagReceiver 8090" dir=in action=allow protocol=TCP localport=8090
```

Sem o `urlacl` (e sem rodar como admin) o receptor cai para `http://localhost:8090/` e
avisa no console — funciona só nesta máquina.

Para desfazer:

```bat
netsh http delete urlacl url=http://+:8090/
netsh advfirewall firewall delete rule name="GBS DiagReceiver 8090"
```

## Executar

```bat
GBS_DiagReceiver.exe              :: console (Ctrl+C encerra)
GBS_DiagReceiver.exe --selftest   :: grava e lê de volta BLOB + CLOB no Oracle e sai (exit 0 = ok)
```

## Serviço do Windows ("GBS DiagReceiver")

O próprio .exe é o serviço (ServiceBase): aberto à mão roda como console; iniciado pelo
Windows roda como serviço (conta LocalSystem, mesmo `http://+:8090/`). Instalação, num
PowerShell **como Administrador** (feche antes qualquer janela console do receptor):

```powershell
taskkill /IM GBS_DiagReceiver.exe /F
sc.exe create GBSDiagReceiver binPath= "C:\GBS\GBS_Inventory\GBS_DiagReceiver\bin\Release\GBS_DiagReceiver.exe" start= auto DisplayName= "GBS DiagReceiver"
sc.exe description GBSDiagReceiver "Recebe os diagnosticos de laptops (porta 8090) e grava no Oracle XE"
sc.exe failure GBSDiagReceiver reset= 86400 actions= restart/5000/restart/10000/restart/60000
sc.exe failureflag GBSDiagReceiver 1
sc.exe start GBSDiagReceiver
sc.exe query GBSDiagReceiver
```

Recompilar o receptor: `sc.exe stop GBSDiagReceiver` antes (o serviço trava o .exe), depois `sc.exe start GBSDiagReceiver`.
Remover: `sc.exe stop GBSDiagReceiver` e `sc.exe delete GBSDiagReceiver`.

## Log

`C:\GBS\Logs\DiagReceiver_yyyyMMdd.log` (`LogFolder`), um arquivo por dia, apagados após
90 dias (`LogRetentionDays`). Cada request: IP, UID, tag, verdict, origin, ID gravado ou erro
completo (stack trace). O 200 só é enviado depois do COMMIT confirmado; falha no banco = 500.

## Endpoints

### `POST /api/diagnostico`

Header obrigatório `X-GBS-Token` (401 se ausente/inválido). Corpo JSON:

```json
{ "tool_version": "1.0.0", "origin": "UBUNTU-USB|UBUNTU-PXE|WINDOWS",
  "stamp": "20261015T143000Z", "verdict": "PASS|PASS_WITH_WARNINGS|FAIL|INCOMPLETE",
  "device": { "internal_uid": "", "service_tag": "", "manufacturer": "", "model": "",
              "bios": "", "cpu": "", "ram_gib": 16, "storage": "",
              "battery_health": 87.5, "battery_cycles": 300, "ssd_life": 95, "technician": "" },
  "checks": [ { "test": "Battery", "result": "WARN", "details": "..." } ],
  "pdf_filename": "x.pdf", "pdf_base64": "JVBERi0..." }
```

- `FAIL_ITEMS` / `WARN_ITEMS` / `NOT_TESTED_ITEMS` = nomes (`test`) dos checks com
  `result` FAIL / WARN / NOT_TESTED, separados por `" | "` (máx. 1000 caracteres).
- `RESULT_JSON` = JSON inteiro sem `pdf_base64`.
- Equipamento é vinculado por `INTERNAL_UID` (ignorando zeros à esquerda) e, se não achar,
  por `SERIAL_NUMBER = service_tag`. Se não achar, grava com `ID_EQUIPAMENTO` nulo.
- Exige `internal_uid` ou `service_tag`.

Resposta 200: `{"id":N,"equipamento_encontrado":true|false}`.
Erros: 400 (JSON/verdict/origin/base64 inválido), 401, 405, 413, 500 — sempre `{"error":"..."}`.

Teste:

```bat
curl -X POST http://localhost:8090/api/diagnostico -H "X-GBS-Token: <token>" -H "Content-Type: application/json" -d @teste.json
```

### `GET /diag/<arquivo>`

Serve arquivos da pasta `diag` (subpastas permitidas) para o boot pela rede. Suporta `HEAD`.

### `GET /health`

`{"status":"ok"}` — para checar se o serviço está no ar.
