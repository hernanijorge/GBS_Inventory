<#
    GBS Inventory - Weekly Oracle schema backup (Data Pump export)

    - Reads the DB user/password from the app's own App.config at run time,
      so no credential is ever hardcoded in this script or checked into git.
    - Builds a throwaway Data Pump parameter file (.par) to pass USERID to
      expdp, instead of putting user/password on the process command line
      (avoids exposing the password via Task Scheduler history / Get-Process).
      The .par file is deleted immediately after the export finishes.
    - Prunes dumps/logs older than $RetentionWeeks from $BackupDir.
    - Logs every run (manual or scheduled) into TBL_BACKUP_LOG, so the app's
      Backup tab and the Task Scheduler path both feed the same history.

    Note: expdp/impdp are version-strict (unlike SQL*Plus), so $ExpdpExe must
    match the actual DB engine version, not just any Oracle client on the box.
    This instance is Oracle 10g XE - the 11g client's expdp fails with
    UDE-00018 against it.

    Scheduled via Task Scheduler, Sundays 02:00 (see registration command in
    the task notes). Safe to also run manually (also used by the app's
    "Run Backup Now" button on the Backup tab).
#>

param(
    [string]$BackupDir       = 'C:\GBS_Backups',
    [int]   $RetentionWeeks  = 8,
    [string]$OracleDirectory = 'GBS_BACKUP_DIR',
    [string]$SchemaName      = 'GBS_OWNER',
    [string]$ExpdpExe        = 'C:\oraclexe\app\oracle\product\10.2.0\server\BIN\expdp.exe',
    [string]$SqlplusExe      = 'C:\oraclexe\app\oracle\product\10.2.0\server\BIN\sqlplus.exe',
    [string]$AppConfigPath   = (Join-Path $PSScriptRoot '..\GBS_Inventory\App.config')
)

$ErrorActionPreference = 'Stop'
$timestamp  = Get-Date -Format 'yyyyMMdd'
$runLogPath = Join-Path $BackupDir "backup_run_$timestamp.log"

function Write-RunLog {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Output $line
    Add-Content -Path $runLogPath -Value $line
}

function Get-OracleCredentialFromAppConfig {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        throw "App.config not found at '$Path'."
    }

    [xml]$xml = Get-Content -Path $Path -Raw
    $connStr = $xml.configuration.connectionStrings.add |
        Where-Object { $_.name -eq 'OracleDB' } |
        Select-Object -ExpandProperty connectionString

    if (-not $connStr) {
        throw "Could not find a connectionString named 'OracleDB' in '$Path'."
    }

    $parts = @{}
    foreach ($pair in $connStr.Split(';')) {
        if ([string]::IsNullOrWhiteSpace($pair)) { continue }
        $kv = $pair.Split('=', 2)
        if ($kv.Length -eq 2) { $parts[$kv[0].Trim()] = $kv[1].Trim() }
    }

    if (-not $parts.ContainsKey('User Id') -or -not $parts.ContainsKey('Password') -or -not $parts.ContainsKey('Data Source')) {
        throw "OracleDB connection string is missing User Id / Password / Data Source."
    }

    return [pscustomobject]@{
        UserId     = $parts['User Id']
        Password   = $parts['Password']
        DataSource = $parts['Data Source']
    }
}

function Write-BackupLog {
    param(
        [Parameter(Mandatory)][pscustomobject]$Credential,
        [Parameter(Mandatory)][string]$Status,
        [string]$DumpFile,
        [string]$Message
    )

    # VARCHAR2(4000) on MENSAGEM; keep well under that.
    if ($Message -and $Message.Length -gt 3900) { $Message = $Message.Substring(0, 3900) }

    $safeMsg = if ($Message) { "'" + ($Message -replace "'", "''") + "'" } else { 'NULL' }
    $safeDmp = if ($DumpFile) { "'" + ($DumpFile -replace "'", "''") + "'" } else { 'NULL' }

    $sql = @"
INSERT INTO TBL_BACKUP_LOG (ID_BACKUP_LOG, ARQUIVO_DMP, STATUS, MENSAGEM)
VALUES (SEQ_BACKUP_LOG.NEXTVAL, $safeDmp, '$Status', $safeMsg);
COMMIT;
EXIT;
"@

    $sqlFile = Join-Path $env:TEMP "gbs_backup_log_$timestamp.sql"
    try {
        Set-Content -Path $sqlFile -Value $sql -Encoding ASCII
        $connect = "$($Credential.UserId)/$($Credential.Password)@$($Credential.DataSource)"
        & $SqlplusExe "-S" $connect "@$sqlFile" | ForEach-Object { Write-RunLog "sqlplus(log): $_" }
    }
    finally {
        if (Test-Path $sqlFile) { Remove-Item $sqlFile -Force }
    }
}

if (-not (Test-Path $BackupDir)) {
    New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
}

Write-RunLog "=== Backup started ==="

$cred         = $null
$logStatus    = 'ERROR'
$logMessage   = 'Unknown error before completion.'
$dumpFileName = $null

try {
    $cred = Get-OracleCredentialFromAppConfig -Path $AppConfigPath

    if (-not (Test-Path $ExpdpExe)) {
        throw "expdp.exe not found at '$ExpdpExe'."
    }

    $dumpFile = "gbs_backup_$timestamp.dmp"
    $logFile  = "gbs_backup_$timestamp.log"
    $parFile  = Join-Path $env:TEMP "gbs_expdp_$timestamp.par"

    $parContent = @(
        "USERID=$($cred.UserId)/$($cred.Password)@$($cred.DataSource)"
        "SCHEMAS=$SchemaName"
        "DIRECTORY=$OracleDirectory"
        "DUMPFILE=$dumpFile"
        "LOGFILE=$logFile"
    )

    try {
        Set-Content -Path $parFile -Value $parContent -Encoding ASCII

        Write-RunLog "Running expdp (schema=$SchemaName, dumpfile=$dumpFile)..."
        & $ExpdpExe "parfile=$parFile" | ForEach-Object { Write-RunLog "expdp: $_" }
        $expdpExitCode = $LASTEXITCODE
    }
    finally {
        if (Test-Path $parFile) { Remove-Item $parFile -Force }
    }

    $dumpFullPath = Join-Path $BackupDir $dumpFile
    if ($expdpExitCode -ne 0 -or -not (Test-Path $dumpFullPath) -or (Get-Item $dumpFullPath).Length -eq 0) {
        throw "Export failed or produced an empty/missing dump file (expdp exit code: $expdpExitCode)."
    }

    $sizeKb       = [math]::Round((Get-Item $dumpFullPath).Length / 1KB, 1)
    $logStatus    = 'SUCCESS'
    $logMessage   = "Dump created: $dumpFile ($sizeKb KB)"
    $dumpFileName = $dumpFile
    Write-RunLog "OK: dump created - $dumpFullPath ($sizeKb KB)"
}
catch {
    $logStatus  = 'ERROR'
    $logMessage = $_.Exception.Message
    Write-RunLog "ERROR: $logMessage"
}

if ($cred) {
    try {
        Write-BackupLog -Credential $cred -Status $logStatus -DumpFile $dumpFileName -Message $logMessage
    }
    catch {
        Write-RunLog "WARN: could not write to TBL_BACKUP_LOG - $($_.Exception.Message)"
    }
}
else {
    Write-RunLog "WARN: no Oracle credentials available - skipped writing to TBL_BACKUP_LOG."
}

# ── Retention: remove dumps/logs/run-logs older than $RetentionWeeks ────────
$cutoff = (Get-Date).AddDays(-7 * $RetentionWeeks)
$removed = 0
Get-ChildItem -Path $BackupDir -Filter 'gbs_backup_*.dmp' -File |
    Where-Object { $_.LastWriteTime -lt $cutoff } |
    ForEach-Object { Write-RunLog "Retention: removing $($_.Name)"; Remove-Item $_.FullName -Force; $removed++ }
Get-ChildItem -Path $BackupDir -Filter 'gbs_backup_*.log' -File |
    Where-Object { $_.LastWriteTime -lt $cutoff } |
    ForEach-Object { Write-RunLog "Retention: removing $($_.Name)"; Remove-Item $_.FullName -Force; $removed++ }
Get-ChildItem -Path $BackupDir -Filter 'backup_run_*.log' -File |
    Where-Object { $_.LastWriteTime -lt $cutoff } |
    ForEach-Object { Remove-Item $_.FullName -Force; $removed++ }

Write-RunLog "Retention sweep complete ($removed file(s) older than $RetentionWeeks weeks removed)."

if ($logStatus -eq 'SUCCESS') {
    Write-RunLog "=== Backup finished OK ==="
    exit 0
}
else {
    Write-RunLog "=== Backup finished with ERROR ==="
    exit 1
}
