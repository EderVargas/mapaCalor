# Genera un mapa de calor por cada hoja (Hoja1 a Hoja14)
# del archivo input/informe_resultados_SMS.xlsx
#
# Uso:
#   .\generar_mapas.ps1
#   .\generar_mapas.ps1 -Formato png
#   .\generar_mapas.ps1 -Formato pkl
#   .\generar_mapas.ps1 -Formato pdf
#   .\generar_mapas.ps1 -Formato show

param(
    [ValidateSet("png", "pdf", "svg", "jpg", "pkl", "show")]
    [string]$Formato = "png"
)

$root   = $PSScriptRoot
$excel  = Join-Path $root "input/informe_resultados_SMS.xlsx"
$python = Join-Path $root ".venv/Scripts/python.exe"
$script = Join-Path $root "mapa_calor.py"

$extensiones = @{
    png = ".png"
    pdf = ".pdf"
    svg = ".svg"
    jpg = ".jpg"
    pkl = ".pkl"
}

# Configuracion de cada reporte (titulo y decimales)
$reportes = @(
    @{ Titulo = "Responsabilidad Administrativa y Penal"; Decimales = 2 }
    @{ Titulo = "Investigacion concluidas con IPRA"; Decimales = 0 }
    @{ Titulo = "Investigacion en tramite"; Decimales = 0 }
    @{ Titulo = "Denuncias - SIDEC"; Decimales = 0 }
    @{ Titulo = "Sanciones administrativas impuestas"; Decimales = 0 }
    @{ Titulo = "Sancionados"; Decimales = 0 }
    @{ Titulo = "Sanciones graves"; Decimales = 0 }
    @{ Titulo = "Direccion General de asuntos penales y estrategicos"; Decimales = 0 }
    @{ Titulo = "Direccion general Anticorrupcion"; Decimales = 0 }
    @{ Titulo = "Delitos denunciados"; Decimales = 0 }
    @{ Titulo = "Delitos denunciados"; Decimales = 0 }
    @{ Titulo = "Delitos por sector"; Decimales = 0 }
    @{ Titulo = "Delitos denunciados"; Decimales = 0 }
    @{ Titulo = "Principales delitos denunciados en los 4 sectores"; Decimales = 0 }
)

$carpeta = $null
if ($Formato -ne "show") {
    $salidaDir = Join-Path $root "output"
    New-Item -ItemType Directory -Path $salidaDir -Force | Out-Null
    $carpeta = Join-Path $salidaDir "mapas_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -ItemType Directory -Path $carpeta -Force | Out-Null
}

for ($i = 1; $i -le 14; $i++) {
    $config    = $reportes[$i - 1]
    $hoja      = "Hoja$i"
    $titulo    = $config.Titulo
    $decimales = $config.Decimales

    $argsPython = @(
        $script, $excel,
        "--hoja", $hoja,
        "--titulo", $titulo,
        "--decimales", $decimales,
        "--formato", $Formato
    )

    if ($Formato -eq "show") {
        Write-Host "Mostrando $hoja -> $titulo (cierra la ventana para seguir) ..." -ForegroundColor Cyan
    } else {
        $ext = $extensiones[$Formato]
        $salida = Join-Path $carpeta "mapa$i$ext"
        $argsPython += @("--salida", $salida)
        Write-Host "Generando $salida | $hoja -> $titulo [$decimales decimales] ..." -ForegroundColor Cyan
    }

    & $python @argsPython

    if ($LASTEXITCODE -eq 0) {
        if ($Formato -eq "show") {
            Write-Host "  OK -> $hoja" -ForegroundColor Green
        } else {
            Write-Host "  OK -> $salida" -ForegroundColor Green
        }
    } else {
        Write-Host "  ERROR en $hoja (codigo $LASTEXITCODE)" -ForegroundColor Red
    }
}

Write-Host ""
if ($Formato -eq "show") {
    Write-Host "Proceso terminado. Ventanas cerradas; no se guardo archivo." -ForegroundColor Yellow
} elseif ($Formato -eq "pkl") {
    Write-Host "Proceso terminado. Figuras en $carpeta" -ForegroundColor Yellow
    Write-Host "Abrir una: $python $script --mostrar `"$carpeta\mapa1.pkl`"" -ForegroundColor Yellow
    Get-ChildItem -Path $carpeta -Filter "mapa*.pkl" | Select-Object Name, LastWriteTime
} else {
    $ext = $extensiones[$Formato]
    Write-Host "Proceso terminado. Mapas generados:" -ForegroundColor Yellow
    Get-ChildItem -Path $carpeta -Filter "mapa*$ext" | Select-Object Name, LastWriteTime
}
