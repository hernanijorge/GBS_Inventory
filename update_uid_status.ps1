# ============================================================
# update_uid_status.ps1
# Marca UIDs da planilha como IN_STOCK, o restante como SOLD
# Banco: Oracle (GBS_Inventory)
# ============================================================

# ---------- CONFIGURACAO ----------
$DataSource = "localhost:1521/XE"
$UserId     = "gbs_owner"
$Password   = "gbs_pass"
$Table      = "TBL_EQUIPAMENTO"
$UidCol     = "INTERNAL_UID"
$StatusCol  = "STATUS"

$UidFile = Join-Path $PSScriptRoot "uids.txt"
# ----------------------------------

$oracleDll = "C:\Users\GBS\GBS_Inventory\packages\Oracle.ManagedDataAccess.21.13.0\lib\net462\Oracle.ManagedDataAccess.dll"
Add-Type -Path $oracleDll

$uids = Get-Content $UidFile |
        Where-Object { $_ -match '^\d+$' } |
        ForEach-Object { $_.Trim() }

if ($uids.Count -eq 0) { Write-Error "Nenhum UID encontrado em $UidFile"; exit 1 }

Write-Host "UIDs carregados: $($uids.Count)"

$connStr = "User Id=$UserId;Password=$Password;Data Source=$DataSource;"
$conn = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connStr)
$conn.Open()

try {
    $tx = $conn.BeginTransaction()

    # 1) Marca TODOS como SOLD
    $cmdSold = New-Object Oracle.ManagedDataAccess.Client.OracleCommand(
        "UPDATE $Table SET $StatusCol = 'SOLD'", $conn)
    $cmdSold.Transaction = $tx
    $affectedSold = $cmdSold.ExecuteNonQuery()
    Write-Host "sold atualizados     : $affectedSold"

    # 2) Marca os UIDs da lista como IN_STOCK (lotes de 999 - limite Oracle IN)
    $batchSize = 999
    $affectedIn = 0
    for ($i = 0; $i -lt $uids.Count; $i += $batchSize) {
        $batch  = $uids[$i..([Math]::Min($i + $batchSize - 1, $uids.Count - 1))]
        $inList = ($batch | ForEach-Object { "'$_'" }) -join ","
        $sql    = "UPDATE $Table SET $StatusCol = 'IN_STOCK' WHERE $UidCol IN ($inList)"
        $cmd    = New-Object Oracle.ManagedDataAccess.Client.OracleCommand($sql, $conn)
        $cmd.Transaction = $tx
        $affectedIn += $cmd.ExecuteNonQuery()
    }
    Write-Host "in_stock atualizados : $affectedIn"

    $tx.Commit()
    Write-Host "Commit realizado com sucesso."

} catch {
    $tx.Rollback()
    Write-Error "Erro - rollback feito: $_"
} finally {
    $conn.Close()
}
