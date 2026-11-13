# Instala las macros de la Skill exportar-fase (Ferma).
# Sin parámetros: en TODAS las versiones de Tekla de C:\TeklaStructures (2023, 2025, ...).
# Con -Destinos "carpeta1;carpeta2": en esas carpetas (la App Ferma le pregunta a la Tekla abierta dónde busca las macros,
# opción XS_MACRO_DIRECTORY, así anda aunque Tekla esté instalada en otro lado).
#   powershell -ExecutionPolicy Bypass -File instalar_macros.ps1 [-Destinos "..."]
# Solo copia estas 4 macros. Nunca instala MCP_RevisionFase.cs (las revisiones por macro cierran Tekla).
param([string]$Destinos = "")
$ErrorActionPreference = "Stop"
$origen = Join-Path (Split-Path $PSScriptRoot -Parent) "macros"
$macros = "MCP_ProbeModelo.cs", "MCP_SeleccionarFase.cs", "MCP_ExportCNCFase.cs", "MCP_ImprimirFasePDF.cs"
if ($Destinos.Trim()) {
    $rutas = $Destinos.Split(";") | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim() }
} else {
    $rutas = Get-ChildItem "C:\TeklaStructures\*\Environments\common\macros\modeling" -Directory -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName }
}
if (-not $rutas) { Write-Host "No encontré la carpeta de macros de Tekla. Avisale al responsable (MB)."; exit 1 }
$fallo = $false
foreach ($d in $rutas) {
    try {
        if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
        foreach ($m in $macros) { Copy-Item (Join-Path $origen $m) (Join-Path $d $m) -Force }
        Write-Host "OK  $d  ($($macros.Count) macros)"
    } catch {
        $fallo = $true
        Write-Host "FRENA: no pude copiar en $d ($($_.Exception.Message)). Si dice 'acceso denegado', pedí que lo instale alguien con permiso de administrador."
    }
}
if ($fallo) { exit 1 }
Write-Host "Listo. Si Tekla estaba abierta, no hace falta reiniciarla."
