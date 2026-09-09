param(
    [switch]$Force
)

$processos = Get-Process EXCEL -ErrorAction SilentlyContinue

if (-not $processos) {
    Write-Host "OK: nenhum processo EXCEL.EXE aberto."
    exit 0
}

$invisiveis = $processos | Where-Object { [string]::IsNullOrWhiteSpace($_.MainWindowTitle) }
$visiveis = $processos | Where-Object { -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle) }

if ($visiveis) {
    Write-Host "Excel visivel encontrado. Nao sera encerrado por seguranca:"
    $visiveis | Select-Object Id, ProcessName, MainWindowTitle | Format-Table -AutoSize
}

if (-not $invisiveis) {
    Write-Host "OK: nao ha Excel invisivel para limpar."
    exit 0
}

Write-Host "Excel invisivel encontrado:"
$invisiveis | Select-Object Id, ProcessName, MainWindowTitle | Format-Table -AutoSize

if (-not $Force) {
    Write-Host ""
    Write-Host "Modo diagnostico: nada foi encerrado."
    Write-Host "Para encerrar somente os invisiveis, rode:"
    Write-Host "powershell -ExecutionPolicy Bypass -File tools\\limpar_excel_invisivel.ps1 -Force"
    exit 2
}

foreach ($p in $invisiveis) {
    try {
        Stop-Process -Id $p.Id -Force -ErrorAction Stop
        Write-Host "Encerrado: EXCEL.EXE PID $($p.Id)"
    } catch {
        Write-Host "Falha ao encerrar PID $($p.Id): $($_.Exception.Message)"
    }
}

$restantes = Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { [string]::IsNullOrWhiteSpace($_.MainWindowTitle) }
if ($restantes) {
    Write-Host "Ainda restam processos invisiveis:"
    $restantes | Select-Object Id, ProcessName, MainWindowTitle | Format-Table -AutoSize
    exit 1
}

Write-Host "OK: processos invisiveis do Excel encerrados."
