# =============================================================================
# Script de Health Check para el Self-Hosted Runner
# Verifica el estado del runner y del sistema
# =============================================================================

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Health Check - Self-Hosted Runner" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Verificar sistema
Write-Host "📌 SISTEMA" -ForegroundColor Yellow
Write-Host "Computer: $env:COMPUTERNAME"
Write-Host "User: $env:USERNAME"
Write-Host "OS: $((Get-CimInstance Win32_OperatingSystem).Caption)"
Write-Host "PowerShell: $($PSVersionTable.PSVersion)"
Write-Host ""

# Verificar Git
Write-Host "📌 GIT" -ForegroundColor Yellow
try {
    $gitVersion = git --version
    Write-Host "✅ Git: $gitVersion"
} catch {
    Write-Host "❌ Git no encontrado"
}
Write-Host ""

# Verificar conectividad
Write-Host "📌 CONECTIVIDAD" -ForegroundColor Yellow
try {
    $response = Invoke-WebRequest -Uri "https://api.github.com" -UseBasicParsing -TimeoutSec 5
    Write-Host "✅ GitHub API: Status $($response.StatusCode)"
} catch {
    Write-Host "❌ GitHub API: No accesible"
}
Write-Host ""

# Verificar recursos
Write-Host "📌 RECURSOS" -ForegroundColor Yellow
$os = Get-CimInstance Win32_OperatingSystem
$totalGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
$freeGB = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
Write-Host "Memoria Total: ${totalGB} GB"
Write-Host "Memoria Libre: ${freeGB} GB"

$disk = Get-PSDrive C
$diskFreeGB = [math]::Round($disk.Free / 1GB, 2)
Write-Host "Disco C: Libre: ${diskFreeGB} GB"
Write-Host ""

Write-Host "==========================================" -ForegroundColor Green
Write-Host " ✅ Health Check Completado" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
