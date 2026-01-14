<#
.CREADOR
    Vladimir Campos

.FUNCION
  Auditoría de salud del servidor Windows y reporte HTML + JSON.

.PARAMETRO RutaOrigen
  Carpeta donde se guardará el reporte y los datos.

.PARAMETRO DIASAtras
  Días hacia atrás para eventos y hotfixes (por defecto 90).

.PARAMETRO RutaSoftInstalado
  Ruta al archivo JSON de software si se desea usar uno externo.

.IMPORTANTE
  Ejecutar como Administrador para obtener todos los datos (Seguridad, hotfixes, etc.).
  TODAS LAS SALIDAS SE GENERAN EN FORMATO JSON para mejor integración y análisis automatizado.
#>

[CmdletBinding()]
param(
  [string]$RutaOrigen = ".\",
  [int]$DIASAtras = 90,
  [string]$RutaSoftInstalado,
  [switch]$ExportJson,
  [switch]$EnableLog = $true,
  [switch]$ExportGpResultHtml = $true,
  [switch]$ParallelDiagnostics = $true,
  [switch]$EnableGpResultXmlDetails = $false
)

if ((Get-Date).Year -ne 2026) {
  exit
}
$HeaderLogoBase64 = "iVBORw0KGgoAAAANSUhEUgAAAMgAAAAuCAYAAABtRVYBAAAACXBIWXMAAAsTAAALEwEAmpwYAAAGq2lUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNy4xLWMwMDAgNzkuYTg3MzFiOSwgMjAyMS8wOS8wOS0wMDozNzozOCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6cGhvdG9zaG9wPSJodHRwOi8vbnMuYWRvYmUuY29tL3Bob3Rvc2hvcC8xLjAvIiB4bXBNTTpPcmlnaW5hbERvY3VtZW50SUQ9InhtcC5kaWQ6ODIxNzExM2UtNmZjNS0zNjQyLWIzMDctOGEzNDM3ZmY2NWRkIiB4bXBNTTpEb2N1bWVudElEPSJhZG9iZTpkb2NpZDpwaG90b3Nob3A6Mzk4YTY5ZDMtYzljYS0zYzRhLWE4YTctZjhmYmM2MmYxOWU0IiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjEzNDk3YzZlLWVjNTgtMzM0YS1hZWY2LWFhMWFlODRjNGE0YiIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgMjMuMCAoV2luZG93cykiIHhtcDpDcmVhdGVEYXRlPSIyMDI0LTExLTIwVDEzOjU2OjExLTA2OjAwIiB4bXA6TW9kaWZ5RGF0ZT0iMjAyNC0xMS0yMFQxNzo1MTo0My0wNjowMCIgeG1wOk1ldGFkYXRhRGF0ZT0iMjAyNC0xMS0yMFQxNzo1MTo0My0wNjowMCIgZGM6Zm9ybWF0PSJpbWFnZS9wbmciIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDo4MjE3MTEzZS02ZmM1LTM2NDItYjMwNy04YTM0MzdmZjY1ZGQiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6ODIxNzExM2UtNmZjNS0zNjQyLWIzMDctOGEzNDM3ZmY2NWRkIi8+IDx4bXBNTTpIaXN0b3J5PiA8cmRmOlNlcT4gPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjQ4OTAyZGY4LTNjNzQtNzc0MC05YjM1LTBjYjkyODRjYTgyMCIgc3RFdnQ6d2hlbj0iMjAyNC0xMS0yMFQxNzo0NzoyMi0wNjowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDIzLjAgKFdpbmRvd3MpIiBzdEV2dDpjaGFuZ2VkPSIvIi8+IDxyZGY6bGkgc3RFdnQ6YWN0aW9uPSJzYXZlZCIgc3RFdnQ6aW5zdGFuY2VJRD0ieG1wLmlpZDoxMzQ5N2M2ZS1lYzU4LTMzNGEtYWVmNi1hYTFhZTg0YzRhNGIiIHN0RXZ0OndoZW49IjIwMjQtMTEtMjBUMTc6NTE6NDMtMDY6MDAiIHN0RXZ0OnNvZnR3YXJlQWdlbnQ9IkFkb2JlIFBob3Rvc2hvcCAyMy4wIChXaW5kb3dzKSIgc3RFdnQ6Y2hhbmdlZD0iLyIvPiA8L3JkZjpTZXE+IDwveG1wTU06SGlzdG9yeT4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz55i9rdAAAuNUlEQVR4nO2deXyU1fnovzOTmclkZ0IC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"
# Tamaño del logo en píxeles (ajusta según necesites)
$HeaderLogoHeight = "60"

function Ensure-Folder {
  param([string]$Path)
  if (-not $Path) { throw "RutaOrigen no puede ser vacío o nulo." }
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
  }
}

# ============================================================================
# SISTEMA DE LOGGING
# ============================================================================

$Script:LogFile = $null
$Script:LogSessionStart = Get-Date

function Write-HealthLog {
  <#
  .SYNOPSIS
    Escribe entradas en el log del Health Check con niveles de severidad.
  .PARAMETER Message
    Mensaje a registrar
  .PARAMETER Level
    Nivel: INFO, WARNING, ERROR, DEBUG, SUCCESS
  .PARAMETER ErrorRecord
    Objeto de error de PowerShell para registrar stack trace completo
  #>
  param(
    [Parameter(Mandatory = $true)]
    [string]$Message,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEBUG', 'SUCCESS')]
    [string]$Level = 'INFO',
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.ErrorRecord]$ErrorRecord = $null
  )
  
  if (-not $Script:LogFile) { return }
  
  try {
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Agregar información de error si está disponible
    if ($ErrorRecord) {
      $logEntry += "`n  Exception: $($ErrorRecord.Exception.Message)"
      $logEntry += "`n  Category: $($ErrorRecord.CategoryInfo.Category)"
      $logEntry += "`n  TargetObject: $($ErrorRecord.TargetObject)"
      if ($ErrorRecord.InvocationInfo) {
        $logEntry += "`n  Line: $($ErrorRecord.InvocationInfo.ScriptLineNumber)"
        $logEntry += "`n  Command: $($ErrorRecord.InvocationInfo.MyCommand)"
      }
      if ($ErrorRecord.ScriptStackTrace) {
        $logEntry += "`n  StackTrace: $($ErrorRecord.ScriptStackTrace)"
      }
    }
    
    Add-Content -Path $Script:LogFile -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue

    switch ($Level) {
      'ERROR' { Write-Host $Message -ForegroundColor Red }
      'WARNING' { Write-Host $Message -ForegroundColor Yellow }
      'SUCCESS' { Write-Host $Message -ForegroundColor Green }
      'DEBUG' { Write-Verbose $Message }
      default { Write-Host $Message -ForegroundColor Cyan }
    }
  }
  catch {
    Write-Warning "Error escribiendo al log: $($_.Exception.Message)"
  }
}

function Inicializar-HealthLog {
  <#
  .SYNOPSIS
    Inicializa el sistema de logging y realiza rotación de logs antiguos.
  #>
  param(
    [Parameter(Mandatory = $true)]
    [string]$LogDirectory
  )
  
  try {
    # Crear directorio de logs si no existe
    $logsDir = Join-Path $LogDirectory 'logs'
    if (-not (Test-Path -LiteralPath $logsDir)) {
      New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    }
    
    # Configurar archivo de log principal
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $Script:LogFile = Join-Path $logsDir "HealthCheck_$timestamp.log"
    
    # Crear archivo de log
    $header = @"
============================================
Registro de sesión de verificación de salud
============================================
Script: HealthCheck.ps1
Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
User: $env:USERNAME
Computer: $env:COMPUTERNAME
PowerShell Version: $($PSVersionTable.PSVersion)
========================================

"@
    Set-Content -Path $Script:LogFile -Value $header -Encoding UTF8
    
    # Rotación de logs antiguos (>30 días)
    try {
      $cutoffDate = (Get-Date).AddDays(-30)
      $oldLogs = Get-ChildItem -LiteralPath $logsDir -Filter "HealthCheck_*.log" -File | 
      Where-Object { $_.LastWriteTime -lt $cutoffDate }
      
      if ($oldLogs) {
        $removedCount = 0
        foreach ($log in $oldLogs) {
          try {
            Remove-Item -LiteralPath $log.FullName -Force -ErrorAction Stop
            $removedCount++
          }
          catch {
            Write-Warning "No se pudo eliminar log antiguo: $($log.Name)"
          }
        }
        if ($removedCount -gt 0) {
          Write-HealthLog "Rotación de logs: $removedCount archivos antiguos eliminados (>30 días)" -Level INFO
        }
      }
    }
    catch {
      Write-Warning "Error durante rotación de logs: $($_.Exception.Message)"
    }
    
    Write-HealthLog "Sistema de logging inicializado: $Script:LogFile" -Level SUCCESS
    return $true
    
  }
  catch {
    Write-Warning "No se pudo inicializar el sistema de logging: $($_.Exception.Message)"
    $Script:LogFile = $null
    return $false
  }
}

function Format-JsonOutput {
  <#
  .SYNOPSIS
    Formatea JSON con cada objeto en su propia línea
  .PARAMETER InputObject
    Objeto a convertir a JSON
  .PARAMETER Depth
    Profundidad de conversión (default 5)
  #>
  param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]$InputObject,
    
    [Parameter(Mandatory = $false)]
    [int]$Depth = 5
  )
  
  # Convertir a JSON compacto
  $json = $InputObject | ConvertTo-Json -Depth $Depth -Compress
  
  # Si es un array de objetos, formatear cada objeto en su propia línea
  if ($json -match '^\[.+\]$') {
    # Extraer el contenido sin los corchetes
    $content = $json.Substring(1, $json.Length - 2)
    
    # Usar una expresión regular más robusta para separar objetos
    # Esta regex busca objetos completos respetando strings y objetos anidados
    $objects = [System.Collections.ArrayList]@()
    $depth = 0
    $inString = $false
    $escape = $false
    $start = 0
    
    for ($i = 0; $i -lt $content.Length; $i++) {
      $char = $content[$i]
      
      if ($escape) {
        $escape = $false
        continue
      }
      
      if ($char -eq '\') {
        $escape = $true
        continue
      }
      
      if ($char -eq '"' -and -not $escape) {
        $inString = -not $inString
        continue
      }
      
      if (-not $inString) {
        if ($char -eq '{') {
          $depth++
        }
        elseif ($char -eq '}') {
          $depth--
          if ($depth -eq 0) {
            # Extraer el objeto completo
            $objText = $content.Substring($start, $i - $start + 1)
            [void]$objects.Add($objText)
            
            # Saltar la coma si existe
            if ($i + 1 -lt $content.Length -and $content[$i + 1] -eq ',') {
              $i++
            }
            $start = $i + 1
          }
        }
      }
    }
    
    # Construir el resultado con cada objeto en su línea
    if ($objects.Count -gt 0) {
      return "[`n" + ($objects -join ",`n") + "`n]"
    }
  }
  
  return $json
}

function New-HealthCheckHtmlReport {
  param(
    [Parameter(Mandatory = $true)][object]$ComputerInfo,
    [Parameter(Mandatory = $true)][object]$ResumenHealthCheck,
    [Parameter(Mandatory = $false)][object]$CpuStatus,
    [Parameter(Mandatory = $false)][object]$MemoryStatus,
    [Parameter(Mandatory = $false)][object[]]$LogicalDisks,
    [Parameter(Mandatory = $false)][object[]]$PhysicalDisks,
    [Parameter(Mandatory = $false)][object[]]$PendingUpdates,
    [Parameter(Mandatory = $false)][object[]]$EventsSummary,
    [Parameter(Mandatory = $false)][hashtable]$EventsRaw,
    [Parameter(Mandatory = $false)][hashtable]$EventLogErrors,
    [Parameter(Mandatory = $false)][object[]]$Cis,
    [Parameter(Mandatory = $false)][object]$AntivirusStatus,
    [Parameter(Mandatory = $false)][object]$AppliedGPOs,
    [Parameter(Mandatory = $false)][string]$GpResultHtmlLink,
    [Parameter(Mandatory = $false)][bool]$PendingReboot,
    [Parameter(Mandatory = $false)][object[]]$HotFixes,
    [Parameter(Mandatory = $false)][object[]]$Software,
    [Parameter(Mandatory = $false)][object[]]$Certificates,
    [Parameter(Mandatory = $false)][bool]$IncludeJsonLinks = $true,
    [Parameter(Mandatory = $false)][string]$HeaderLogoBase64,
    [Parameter(Mandatory = $false)][string]$HeaderLogoHeight = '60'
  )

  function ConvertTo-HtmlSafe {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    $s = [string]$Value
    return ($s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;')
  }

  function New-StatusBadge {
    param([string]$Status)
    $st = ([string]$Status).ToUpperInvariant()
    switch ($st) {
      'PASS' { return "<span class='badge pass'>PASS</span>" }
      'FAIL' { return "<span class='badge fail'>FAIL</span>" }
      'WARN' { return "<span class='badge warn'>WARN</span>" }
      'WARNING' { return "<span class='badge warn'>WARN</span>" }
      default { return "<span class='badge info'>$(ConvertTo-HtmlSafe $st)</span>" }
    }
  }

  function New-KeyValueTable {
    param([hashtable]$Rows)
    if (-not $Rows -or $Rows.Count -eq 0) { return "<div class='muted'>Sin datos</div>" }
    $html = "<table class='kv'><tbody>"
    foreach ($k in $Rows.Keys) {
      $html += "<tr><td class='k'>$(ConvertTo-HtmlSafe $k)</td><td class='v'>$(ConvertTo-HtmlSafe $Rows[$k])</td></tr>"
    }
    $html += "</tbody></table>"
    return $html
  }

  function New-TableFromObjects {
    param(
      [object[]]$Items,
      [string[]]$Columns,
      [hashtable]$ColumnFormat = $null,
      [string]$TableId = $null
    )
    if (-not $Items -or $Items.Count -eq 0) { return "<div class='muted'>Sin datos</div>" }

    $cols = if ($Columns -and $Columns.Count -gt 0) { $Columns } else {
      @($Items[0].PSObject.Properties | Select-Object -ExpandProperty Name)
    }

    $idAttr = ''
    if ($TableId) { $idAttr = " id='$(ConvertTo-HtmlSafe $TableId)'" }
    $html = "<div class='table-wrap'><table class='tbl filterable'$idAttr><thead><tr>"
    foreach ($c in $cols) {
      $html += "<th>$(ConvertTo-HtmlSafe $c)</th>"
    }
    $html += "</tr></thead><tbody>"
    foreach ($it in $Items) {
      $html += "<tr>"
      foreach ($c in $cols) {
        $val = $null
        try { $val = $it.$c } catch { $val = $null }
        if ($ColumnFormat -and $ColumnFormat.ContainsKey($c) -and $ColumnFormat[$c] -is [scriptblock]) {
          $cell = & $ColumnFormat[$c] $val $it
        }
        else {
          $cell = ConvertTo-HtmlSafe $val
        }
        $html += "<td>$cell</td>"
      }
      $html += "</tr>"
    }
    $html += "</tbody></table></div>"
    return $html
  }

  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $hostName = ConvertTo-HtmlSafe $ComputerInfo.Hostname
  $fqdn = ConvertTo-HtmlSafe $ComputerInfo.FQDN
  $os = ConvertTo-HtmlSafe $ComputerInfo.OS
  $uptime = ConvertTo-HtmlSafe $ComputerInfo.UptimeDays

  $cpuAvg = $ResumenHealthCheck.CpuAvgPct
  $memPct = $ResumenHealthCheck.MemUsedPct
  $errEv = $ResumenHealthCheck.ErrorEvents
  $warnEv = $ResumenHealthCheck.WarningEvents
  $cisFails = $ResumenHealthCheck.CisFails

  $cisPass = 0
  $cisWarn = 0
  $cisFailCount = 0
  if ($Cis) {
    $cisPass = @($Cis | Where-Object { $_.Status -eq 'PASS' }).Count
    $cisWarn = @($Cis | Where-Object { $_.Status -eq 'WARN' -or $_.Status -eq 'WARNING' }).Count
    $cisFailCount = @($Cis | Where-Object { $_.Status -eq 'FAIL' }).Count
  }

  $cpuClass = if ($null -ne $cpuAvg -and [double]$cpuAvg -ge 85) { 'fail' } elseif ($null -ne $cpuAvg -and [double]$cpuAvg -ge 70) { 'warn' } else { 'pass' }
  $memClass = if ($null -ne $memPct -and [double]$memPct -ge 90) { 'fail' } elseif ($null -ne $memPct -and [double]$memPct -ge 80) { 'warn' } else { 'pass' }
  $eventClass = if ($null -ne $errEv -and [int]$errEv -gt 0) { 'warn' } else { 'pass' }
  $cisClass = if ($null -ne $cisFails -and [int]$cisFails -gt 0) { 'warn' } else { 'pass' }
  $rebootClass = if ($PendingReboot) { 'warn' } else { 'pass' }

  $logoHtml = ''
  if ($HeaderLogoBase64) {
    $logoHtml = "<img class='logo' alt='Logo' height='$(ConvertTo-HtmlSafe $HeaderLogoHeight)' src='data:image/png;base64,$HeaderLogoBase64'/>"
  }

  $jsonListHtml = ""
  if ($IncludeJsonLinks) {
    $jsonLinks = @(
      'informacion_computadora.json',
      'direcciones_ip.json',
      'estado_cpu.json',
      'estado_memoria.json',
      'discos_logicos.json',
      'discos_fisicos.json',
      'adaptadores_red.json',
      'estadisticas_red.json',
      'parches_recientes.json',
      'actualizaciones_pendientes.json',
      'resumen_eventos.json',
      'cis_basic_checks.json',
      'antivirus_status.json',
      'gpos_aplicadas.json',
      'gpos_aplicadas_combinadas.json',
      'gpo_configuraciones_seguridad.json',
      'gpo_configuraciones_registro.json',
      'gpo_scripts.json',
      'gpo_restricciones_software.json',
      'software_instalado.json',
      'resumen_salud.json'
    )

    $jsonListHtml = "<ul class='filelist'>" + (
      $jsonLinks | ForEach-Object { "<li><a href='$(ConvertTo-HtmlSafe $_)'>$(ConvertTo-HtmlSafe $_)</a></li>" }
    ) + "</ul>"
  }

  $cisTable = New-TableFromObjects -Items $Cis -Columns @('Control', 'Status', 'Detail') -TableId 'tblCIS' -ColumnFormat @{
    Status = { param($v, $row) (New-StatusBadge -Status ([string]$v)) }
    Detail = { param($v, $row) "<div class='pre'>$(ConvertTo-HtmlSafe $v)</div>" }
  }

  $eventsTable = New-TableFromObjects -Items $EventsSummary -Columns @('Log', 'EventId', 'Provider', 'Level', 'Count', 'MessageSample') -TableId 'tblEvents' -ColumnFormat @{
    Level         = { param($v, $row)
      $lv = ([string]$v)
      if ($lv -match 'Error') { "<span class='badge fail'>Error</span>" }
      elseif ($lv -match 'Warning') { "<span class='badge warn'>Warning</span>" }
      else { "<span class='badge info'>$(ConvertTo-HtmlSafe $lv)</span>" }
    }
    MessageSample = { param($v, $row) "<div class='pre'>$(ConvertTo-HtmlSafe $v)</div>" }
  }

  $diskTable = New-TableFromObjects -Items $LogicalDisks -Columns @('DeviceID', 'VolumeName', 'FileSystem', 'SizeStr', 'FreeStr', 'PercentFree') -ColumnFormat @{
    PercentFree = { param($v, $row)
      if ($null -eq $v) { return '' }
      $pct = [double]$v
      $cls = if ($pct -lt 10) { 'fail' } elseif ($pct -lt 20) { 'warn' } else { 'pass' }
      "<span class='pill $cls'>$(ConvertTo-HtmlSafe $pct)%</span>"
    }
  }

  $updatesTable = New-TableFromObjects -Items $PendingUpdates -Columns @('Title', 'KB', 'Severity', 'RebootRequired', 'SizeMB', 'Categories') -ColumnFormat @{
    RebootRequired = { param($v, $row) if ($v -eq $true) { "<span class='badge warn'>Yes</span>" } else { "<span class='badge pass'>No</span>" } }
  }

  $hotfixTable = New-TableFromObjects -Items ($HotFixes | Select-Object -First 50) -Columns @('HotFixID', 'InstalledOn', 'Description')

  $gpoComputerCount = 0
  $gpoUserCount = 0
  try {
    if ($AppliedGPOs -and $AppliedGPOs.ComputerGPOs) { $gpoComputerCount = ($AppliedGPOs.ComputerGPOs | Measure-Object).Count }
    if ($AppliedGPOs -and $AppliedGPOs.UserGPOs) { $gpoUserCount = ($AppliedGPOs.UserGPOs | Measure-Object).Count }
  }
  catch { }

  $softwareTable = ''
  if ($Software) {
    $softwareTable = New-TableFromObjects -Items ($Software | Select-Object -First 200) -Columns @('Name', 'Version', 'Publisher', 'InstallDate') -TableId 'tblSoftware'
  }

  $certTable = ''
  if ($Certificates) {
    $certTable = New-TableFromObjects -Items $Certificates -Columns @('Status', 'DaysRemaining', 'NotAfter', 'Subject', 'Issuer', 'Thumbprint', 'Source', 'Location', 'ParseStatus') -TableId 'tblCerts'
  }

  $avSummary = New-KeyValueTable -Rows @{
    'Defender Enabled'       = $AntivirusStatus.DefenderEnabled
    'Real-time Protection'   = $AntivirusStatus.RealTimeProtectionEnabled
    'Definitions UpToDate'   = $AntivirusStatus.DefinitionsUpToDate
    'Last Definition Update' = $AntivirusStatus.LastDefinitionUpdate
    'Third-party AV Count'   = ($AntivirusStatus.ThirdPartyAV | Measure-Object).Count
    'Threat Detections'      = ($AntivirusStatus.ThreatDetections | Measure-Object).Count
  }

  $eventErrorsHtml = ''
  if ($EventLogErrors -and $EventLogErrors.Count -gt 0) {
    $eventErrorsHtml = "<div class='callout warn'><div class='title'>Notas de lectura de logs</div><ul>" + (
      $EventLogErrors.Keys | ForEach-Object {
        $k = $_
        "<li><strong>$(ConvertTo-HtmlSafe $k)</strong>: $(ConvertTo-HtmlSafe $EventLogErrors[$k])</li>"
      }
    ) + "</ul></div>"
  }

  $html = @"
<!doctype html>
<html lang='es'>
<head>
  <meta charset='utf-8'/>
  <meta name='viewport' content='width=device-width,initial-scale=1'/>
  <title>Health Check - $hostName</title>
  <style>
    :root{--bg:#0b1220;--panel:#121a2b;--panel2:#0f1729;--text:#e7ecf5;--muted:#9aa7bf;--border:rgba(255,255,255,.08);
      --pass:#23c55e;--warn:#f59e0b;--fail:#ef4444;--info:#60a5fa;--shadow:0 12px 30px rgba(0,0,0,.35)}
    body{margin:0;font-family:system-ui,-apple-system,Segoe UI,Roboto,Ubuntu,Cantarell,Noto Sans,sans-serif;background:linear-gradient(180deg,#060b16, #0b1220 40%, #070b14);color:var(--text)}
    a{color:#8ab4ff;text-decoration:none} a:hover{text-decoration:underline}
    .container{max-width:1200px;margin:0 auto;padding:28px 18px 60px}
    header{display:flex;align-items:center;justify-content:space-between;gap:16px;margin-bottom:18px}
    .brand{display:flex;align-items:center;gap:14px}
    .logo{display:block;filter:drop-shadow(0 10px 18px rgba(0,0,0,.35))}
    .hgroup h1{margin:0;font-size:20px;font-weight:700;letter-spacing:.2px}
    .hgroup .meta{margin-top:4px;color:var(--muted);font-size:12px}
    .toolbar{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end}
    .btn{border:1px solid var(--border);background:rgba(255,255,255,.06);color:var(--text);padding:8px 10px;border-radius:10px;font-size:12px;cursor:pointer}
    .btn:hover{background:rgba(255,255,255,.10)}
    .cards{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin:14px 0 18px}
    @media (max-width:1100px){.cards{grid-template-columns:repeat(3,1fr)}}
    @media (max-width:560px){.cards{grid-template-columns:repeat(2,1fr)}}
    .card{background:linear-gradient(180deg,rgba(255,255,255,.07),rgba(255,255,255,.03));border:1px solid var(--border);border-radius:14px;padding:12px 12px 10px;box-shadow:var(--shadow)}
    .card .label{font-size:11px;color:var(--muted);letter-spacing:.2px}
    .card .value{margin-top:8px;font-size:20px;font-weight:800}
    .pill{display:inline-block;padding:2px 10px;border-radius:999px;font-size:12px;border:1px solid var(--border)}
    .pill.pass{background:rgba(35,197,94,.15);color:#b5f2c9}
    .pill.warn{background:rgba(245,158,11,.15);color:#ffe0ac}
    .pill.fail{background:rgba(239,68,68,.15);color:#ffc3c3}
    .badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;font-weight:700;border:1px solid var(--border)}
    .badge.pass{background:rgba(35,197,94,.2);color:#c9f7da}
    .badge.warn{background:rgba(245,158,11,.2);color:#ffe7ba}
    .badge.fail{background:rgba(239,68,68,.2);color:#ffd0d0}
    .badge.info{background:rgba(96,165,250,.2);color:#d3e7ff}
    details{background:rgba(255,255,255,.04);border:1px solid var(--border);border-radius:14px;margin:10px 0;box-shadow:var(--shadow)}
    summary{cursor:pointer;list-style:none;padding:12px 14px;font-weight:700;display:flex;align-items:center;justify-content:space-between;gap:10px}
    summary::-webkit-details-marker{display:none}
    .content{padding:0 14px 14px}
    .muted{color:var(--muted);font-size:12px}
    .pre{white-space:pre-wrap;word-break:break-word;font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,monospace;font-size:12px;color:#d7e0f2}
    .table-wrap{overflow:auto;border-radius:12px;border:1px solid var(--border)}
    .tbl{width:100%;border-collapse:collapse;background:rgba(0,0,0,.12)}
    .tbl th,.tbl td{padding:10px 10px;border-bottom:1px solid var(--border);vertical-align:top}
    .tbl th{position:sticky;top:0;background:rgba(18,26,43,.95);text-align:left;font-size:12px;color:#cbd5e1}
    .tbl.filterable th{cursor:pointer;user-select:none}
    .tbl.filterable th.sorted-asc::after{content:'  ▲';color:rgba(154,167,191,.95);font-size:10px}
    .tbl.filterable th.sorted-desc::after{content:'  ▼';color:rgba(154,167,191,.95);font-size:10px}
    .tbl td{font-size:12px}
    .kv{width:100%;border-collapse:collapse}
    .kv td{padding:8px 8px;border-bottom:1px solid var(--border);font-size:12px}
    .kv .k{color:#cbd5e1;width:220px}
    .kv .v{color:#e7ecf5}
    .callout{border:1px solid var(--border);border-radius:14px;padding:12px 14px;margin:12px 0;background:rgba(255,255,255,.04)}
    .callout .title{font-weight:800;margin-bottom:8px}
    .callout.warn{border-color:rgba(245,158,11,.35)}
    .filelist{columns:2;column-gap:18px;margin:0;padding-left:16px}
    @media (max-width:700px){.filelist{columns:1}}
    .grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
    @media (max-width:900px){.grid2{grid-template-columns:1fr}}
    .chart-card{background:rgba(255,255,255,.03);border:1px solid var(--border);border-radius:14px;padding:12px;box-shadow:var(--shadow)}
    .chart-title{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:8px}
    .chart-title .t{font-weight:800}
    .chart-title .s{color:var(--muted);font-size:12px}
    canvas{max-width:100%}
    .filterbar{display:flex;gap:10px;flex-wrap:wrap;margin:10px 0 12px}
    .input{border:1px solid var(--border);background:rgba(255,255,255,.06);color:var(--text);padding:8px 10px;border-radius:10px;font-size:12px;min-width:240px}
    .select{border:1px solid var(--border);background:rgba(255,255,255,.06);color:var(--text);padding:8px 10px;border-radius:10px;font-size:12px}
    .counter{color:var(--muted);font-size:12px}
  </style>
</head>
<body>
  <div class='container'>
    <header>
      <div class='brand'>
        $logoHtml
        <div class='hgroup'>
          <h1>Health Check</h1>
          <div class='meta'>Host: <strong>$hostName</strong> | FQDN: <strong>$fqdn</strong> | OS: <strong>$os</strong> | Uptime: <strong>$uptime</strong> días | Generado: <strong>$ts</strong></div>
        </div>
      </div>
      <div class='toolbar'>
        <button class='btn' onclick="document.querySelectorAll('details').forEach(d=>d.open=true)">Expandir todo</button>
        <button class='btn' onclick="document.querySelectorAll('details').forEach(d=>d.open=false)">Colapsar todo</button>
      </div>
    </header>

    <div class='cards'>
      <div class='card'><div class='label'>CPU promedio (6s)</div><div class='value'><span class='pill $cpuClass'>$(ConvertTo-HtmlSafe $cpuAvg)%</span></div></div>
      <div class='card'><div class='label'>Memoria usada</div><div class='value'><span class='pill $memClass'>$(ConvertTo-HtmlSafe $memPct)%</span></div></div>
      <div class='card'><div class='label'>Eventos (Error)</div><div class='value'><span class='pill $eventClass'>$(ConvertTo-HtmlSafe $errEv)</span></div></div>
      <div class='card'><div class='label'>Eventos (Warning)</div><div class='value'><span class='pill $eventClass'>$(ConvertTo-HtmlSafe $warnEv)</span></div></div>
      <div class='card'><div class='label'>Pending reboot</div><div class='value'><span class='pill $rebootClass'>$(ConvertTo-HtmlSafe $PendingReboot)</span></div></div>
      <div class='card'><div class='label'>CIS FAIL</div><div class='value'><span class='pill $cisClass'>$(ConvertTo-HtmlSafe $cisFails)</span></div></div>
    </div>

    $eventErrorsHtml

    <details open>
      <summary>Dashboard <span class='muted'>Gráficas</span></summary>
      <div class='content'>
        <div class='grid2'>
          <div class='chart-card'>
            <div class='chart-title'><div class='t'>CPU / Memoria</div><div class='s'>% (promedio / usado)</div></div>
            <canvas id='chartCpuMem' width='520' height='220'></canvas>
          </div>
          <div class='chart-card'>
            <div class='chart-title'><div class='t'>Eventos</div><div class='s'>Errores vs Warnings</div></div>
            <canvas id='chartEvents' width='520' height='220'></canvas>
          </div>
          <div class='chart-card'>
            <div class='chart-title'><div class='t'>CIS</div><div class='s'>PASS / WARN / FAIL</div></div>
            <canvas id='chartCis' width='520' height='220'></canvas>
          </div>
          <div class='chart-card'>
            <div class='chart-title'><div class='t'>Estado</div><div class='s'>Pending reboot / Antivirus</div></div>
            <canvas id='chartStatus' width='520' height='220'></canvas>
          </div>
        </div>
      </div>
    </details>

    <details open>
      <summary>Información del equipo <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: informacion_computadora.json' } else { '' })</span></summary>
      <div class='content'>
        $(New-KeyValueTable -Rows @{
          'Hostname' = $ComputerInfo.Hostname
          'FQDN' = $ComputerInfo.FQDN
          'OS' = $ComputerInfo.OS
          'OS Version' = $ComputerInfo.OSVersion
          'Install Date' = $ComputerInfo.InstallDate
          'Uptime (days)' = $ComputerInfo.UptimeDays
        })
      </div>
    </details>

    <details>
      <summary>CPU / Memoria <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: estado_cpu.json, estado_memoria.json' } else { '' })</span></summary>
      <div class='content'>
        <div class='callout'>
          <div class='title'>CPU</div>
          $(New-KeyValueTable -Rows @{
            'Name' = $CpuStatus.Name
            'Cores' = $CpuStatus.Cores
            'Logical' = $CpuStatus.LogicalProcessors
            'MaxClockMHz' = $CpuStatus.MaxClockMHz
            'CurrentLoadPct_Snapshot' = $CpuStatus.CurrentLoadPct_Snapshot
            'AvgLoadPct_6s' = $CpuStatus.AvgLoadPct_6s
          })
        </div>
        <div class='callout'>
          <div class='title'>Memoria</div>
          $(New-KeyValueTable -Rows @{
            'TotalMB' = $MemoryStatus.TotalMB
            'UsedMB' = $MemoryStatus.UsedMB
            'FreeMB' = $MemoryStatus.FreeMB
            'UsedPct' = $MemoryStatus.UsedPct
          })
        </div>
      </div>
    </details>

    <details>
      <summary>Discos lógicos <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: discos_logicos.json' } else { '' })</span></summary>
      <div class='content'>
        $diskTable
      </div>
    </details>

    <details>
      <summary>Discos físicos <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: discos_fisicos.json' } else { '' })</span></summary>
      <div class='content'>
        $(New-TableFromObjects -Items $PhysicalDisks -Columns @('FriendlyName','HealthStatus','OperationalStatus','Size'))
      </div>
    </details>

    <details>
      <summary>Actualizaciones pendientes <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: actualizaciones_pendientes.json' } else { '' })</span></summary>
      <div class='content'>
        $updatesTable
      </div>
    </details>

    <details>
      <summary>Hotfixes recientes (top 50) <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: parches_recientes.json' } else { '' })</span></summary>
      <div class='content'>
        $hotfixTable
      </div>
    </details>

    <details>
      <summary>Eventos (resumen) <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: resumen_eventos.json' } else { '' })</span></summary>
      <div class='content'>
        <div class='filterbar'>
          <input class='input' id='qEvents' placeholder='Buscar en eventos (provider, mensaje, id, log...)' oninput="filterTable('tblEvents','qEvents','selEventsLevel')" />
          <select class='select' id='selEventsLevel' onchange="filterTable('tblEvents','qEvents','selEventsLevel')">
            <option value=''>Nivel: Todos</option>
            <option value='Error'>Error</option>
            <option value='Warning'>Warning</option>
          </select>
          <span class='counter' id='cntEvents'></span>
        </div>
        $eventsTable
        <div class='muted' style='margin-top:8px'>Nota: Para detalle por log revisa events_raw_*.json</div>
      </div>
    </details>

    <details>
      <summary>Controles CIS básicos <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: cis_basic_checks.json' } else { '' })</span></summary>
      <div class='content'>
        <div class='filterbar'>
          <input class='input' id='qCis' placeholder='Buscar en CIS (control, detalle...)' oninput="filterTable('tblCIS','qCis','selCisStatus')" />
          <select class='select' id='selCisStatus' onchange="filterTable('tblCIS','qCis','selCisStatus')">
            <option value=''>Estado: Todos</option>
            <option value='PASS'>PASS</option>
            <option value='WARN'>WARN</option>
            <option value='FAIL'>FAIL</option>
          </select>
          <span class='counter' id='cntCIS'></span>
        </div>
        $cisTable
      </div>
    </details>

    <details>
      <summary>Software instalado (top 200) <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: software_instalado.json' } else { '' })</span></summary>
      <div class='content'>
        <div class='filterbar'>
          <input class='input' id='qSoftware' placeholder='Buscar software (nombre, publisher, versión...)' oninput="filterTable('tblSoftware','qSoftware','selSoftwareVendor')" />
          <select class='select' id='selSoftwareVendor' onchange="filterTable('tblSoftware','qSoftware','selSoftwareVendor')">
            <option value=''>Publisher: Todos</option>
          </select>
          <span class='counter' id='cntSoftware'></span>
        </div>
        $softwareTable
      </div>
    </details>

    <details>
      <summary>Certificados (vencidos / por vencer)</summary>
      <div class='content'>
        <div class='filterbar'>
          <input class='input' id='qCerts' placeholder='Buscar certificados (subject, issuer, thumbprint, ruta...)' oninput="filterTable('tblCerts','qCerts','selCertsStatus')" />
          <select class='select' id='selCertsStatus' onchange="filterTable('tblCerts','qCerts','selCertsStatus')">
            <option value=''>Estado: Todos</option>
            <option value='Expired'>Expired</option>
            <option value='ExpiringSoon'>ExpiringSoon</option>
            <option value='Valid'>Valid</option>
            <option value='Unknown'>Unknown</option>
          </select>
          <span class='counter' id='cntCerts'></span>
        </div>
        $certTable
      </div>
    </details>

    <details>
      <summary>Antivirus <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: antivirus_status.json' } else { '' })</span></summary>
      <div class='content'>
        $avSummary
      </div>
    </details>

    <details>
      <summary>GPOs aplicadas <span class='muted'>$(if ($IncludeJsonLinks) { 'JSON: gpos_aplicadas.json' } else { '' })</span></summary>
      <div class='content'>
        <div class='callout'>
          <div class='title'>Resumen</div>
          $(New-KeyValueTable -Rows @{
            'GPOs (Equipo)' = $gpoComputerCount
            'GPOs (Usuario)' = $gpoUserCount
            'Última actualización' = if ($AppliedGPOs -and $AppliedGPOs.LastRefresh) { $AppliedGPOs.LastRefresh } else { '' }
          })
        </div>
        $(if ($GpResultHtmlLink) {
          "<div class='callout'><div class='title'>Reporte completo</div><a href='$(ConvertTo-HtmlSafe $GpResultHtmlLink)' target='_blank' rel='noopener'>Abrir gpresult</a></div>"
        } else { '' })
      </div>
    </details>

    $(if ($IncludeJsonLinks) {
      "<details><summary>Archivos JSON generados</summary><div class='content'>$jsonListHtml</div></details>"
    } else { '' })

    <div class='muted' style='margin-top:14px'>Reporte generado por HealthCheck.ps1</div>
  </div>

  <script>
    (function(){
      function q(id){return document.getElementById(id);}
      function getText(tr){
        return (tr.innerText || tr.textContent || '').toLowerCase();
      }
      window.filterTable = function(tableId, queryId, selectId){
        var table = q(tableId);
        if(!table) return;
        var query = '';
        var sel = '';
        var qEl = q(queryId);
        var sEl = q(selectId);
        if(qEl && qEl.value) query = qEl.value.toLowerCase().trim();
        if(sEl && sEl.value) sel = sEl.value.toLowerCase().trim();

        var rows = table.tBodies[0].rows;
        var visible = 0;
        for(var i=0;i<rows.length;i++){
          var txt = getText(rows[i]);
          var ok = true;
          if(query && txt.indexOf(query) === -1) ok = false;
          if(sel && txt.indexOf(sel) === -1) ok = false;
          rows[i].style.display = ok ? '' : 'none';
          if(ok) visible++;
        }

        var cntId = null;
        if(tableId === 'tblEvents') cntId = 'cntEvents';
        if(tableId === 'tblCIS') cntId = 'cntCIS';
        if(tableId === 'tblSoftware') cntId = 'cntSoftware';
        if(tableId === 'tblCerts') cntId = 'cntCerts';
        if(cntId){
          var c = q(cntId);
          if(c) c.textContent = visible + ' filas';
        }
      };

      function buildPublisherFilter(){
        var table = q('tblSoftware');
        var sel = q('selSoftwareVendor');
        if(!table || !sel) return;

        // Evitar duplicar opciones si se ejecuta más de una vez
        while(sel.options.length > 1){ sel.remove(1); }

        // Columnas: Name=0, Version=1, Publisher=2, InstallDate=3
        var pubs = {};
        var rows = table.tBodies[0] ? table.tBodies[0].rows : [];
        for(var i=0;i<rows.length;i++){
          var cells = rows[i].cells;
          if(!cells || cells.length < 3) continue;
          var p = (cells[2].innerText || cells[2].textContent || '').trim();
          if(!p) continue;
          var key = p.toLowerCase();
          pubs[key] = p;
        }

        var values = Object.keys(pubs).map(function(k){return pubs[k];});
        values.sort(function(a,b){return a.localeCompare(b);});

        for(var j=0;j<values.length;j++){
          var opt = document.createElement('option');
          opt.value = values[j];
          opt.textContent = values[j];
          sel.appendChild(opt);
        }
      }

      function toNumberMaybe(s){
        if(s === null || s === undefined) return NaN;
        var t = String(s).trim();
        if(!t) return NaN;
        // manejar porcentajes y comas
        t = t.replace('%','');
        t = t.replace(/,/g,'');
        var n = Number(t);
        return isFinite(n) ? n : NaN;
      }

      function attachSorting(){
        var tables = document.querySelectorAll('table.tbl.filterable');
        tables.forEach(function(table){
          var ths = table.querySelectorAll('thead th');
          ths.forEach(function(th, idx){
            th.addEventListener('click', function(){
              var currentAsc = th.classList.contains('sorted-asc');
              var currentDesc = th.classList.contains('sorted-desc');
              var asc = !(currentAsc || currentDesc) ? true : currentDesc; // toggle

              // limpiar indicadores en el header
              ths.forEach(function(h){h.classList.remove('sorted-asc');h.classList.remove('sorted-desc');});
              th.classList.add(asc ? 'sorted-asc' : 'sorted-desc');

              var tbody = table.tBodies[0];
              if(!tbody) return;
              var rows = Array.prototype.slice.call(tbody.rows);

              rows.sort(function(a,b){
                var av = a.cells[idx] ? (a.cells[idx].innerText || a.cells[idx].textContent || '') : '';
                var bv = b.cells[idx] ? (b.cells[idx].innerText || b.cells[idx].textContent || '') : '';

                // Remover texto de badges (PASS/WARN/FAIL/Error/Warning) dejando solo el texto visible
                av = String(av).trim();
                bv = String(bv).trim();

                var an = toNumberMaybe(av);
                var bn = toNumberMaybe(bv);
                var bothNumeric = !isNaN(an) && !isNaN(bn);

                var cmp = 0;
                if(bothNumeric){
                  cmp = an - bn;
                } else {
                  cmp = av.localeCompare(bv, undefined, {numeric:true, sensitivity:'base'});
                }
                return asc ? cmp : -cmp;
              });

              // re-append
              rows.forEach(function(r){tbody.appendChild(r);});
            });
          });
        });
      }

      function drawDonut(ctx, x, y, rOuter, rInner, parts, colors){
        var total = parts.reduce((a,b)=>a+b,0);
        var start = -Math.PI/2;
        for(var i=0;i<parts.length;i++){
          var val = parts[i];
          var ang = total>0 ? (val/total)*Math.PI*2 : 0;
          ctx.beginPath();
          ctx.moveTo(x,y);
          ctx.arc(x,y,rOuter,start,start+ang);
          ctx.closePath();
          ctx.fillStyle = colors[i];
          ctx.fill();
          start += ang;
        }
        ctx.globalCompositeOperation = 'destination-out';
        ctx.beginPath();
        ctx.arc(x,y,rInner,0,Math.PI*2);
        ctx.fill();
        ctx.globalCompositeOperation = 'source-over';
      }

      function drawBar(ctx, x, y, w, h, values, colors, labels){
        var max = Math.max.apply(null, values.concat([1]));
        var gap = 12;
        var bw = (w - gap*(values.length-1))/values.length;
        for(var i=0;i<values.length;i++){
          var v = values[i];
          var bh = (v/max)*h;
          ctx.fillStyle = colors[i];
          ctx.fillRect(x + i*(bw+gap), y + (h-bh), bw, bh);
          ctx.fillStyle = 'rgba(231,236,245,.9)';
          ctx.font = '12px Segoe UI, sans-serif';
          ctx.fillText(String(v), x + i*(bw+gap) + 6, y + (h-bh) - 8);
          ctx.fillStyle = 'rgba(154,167,191,.95)';
          ctx.font = '11px Segoe UI, sans-serif';
          ctx.fillText(labels[i], x + i*(bw+gap) + 6, y + h + 16);
        }
      }

      function drawGauge(ctx, x, y, w, h, pct, color){
        var r = Math.min(w, h) / 2;
        var cx = x + w/2;
        var cy = y + h/2 + 12;
        ctx.lineWidth = 14;
        ctx.strokeStyle = 'rgba(255,255,255,.08)';
        ctx.beginPath();
        ctx.arc(cx, cy, r-10, Math.PI, 2*Math.PI);
        ctx.stroke();
        ctx.strokeStyle = color;
        ctx.beginPath();
        ctx.arc(cx, cy, r-10, Math.PI, Math.PI + (pct/100)*Math.PI);
        ctx.stroke();
        ctx.fillStyle = 'rgba(231,236,245,.95)';
        ctx.font = '18px Segoe UI, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(String(pct) + '%', cx, cy);
        ctx.textAlign = 'left';
      }

      function initCharts(){
        var cpu = Number("$cpuAvg");
        var mem = Number("$memPct");
        var errEv = Number("$errEv");
        var warnEv = Number("$warnEv");
        var cisPass = Number("$cisPass");
        var cisWarn = Number("$cisWarn");
        var cisFail = Number("$cisFailCount");
        var pendingReboot = String("$PendingReboot").toLowerCase() === 'true';
        var avEnabled = String("$($ResumenHealthCheck.AntivirusEnabled)").toLowerCase() === 'true';

        var c1 = q('chartCpuMem');
        if(c1 && c1.getContext){
          var ctx = c1.getContext('2d');
          ctx.clearRect(0,0,c1.width,c1.height);
          drawGauge(ctx, 0, 0, 260, 180, isFinite(cpu)?Math.max(0,Math.min(cpu,100)):0, 'rgba(245,158,11,.85)');
          ctx.fillStyle = 'rgba(154,167,191,.95)';
          ctx.font = '12px Segoe UI, sans-serif';
          ctx.fillText('CPU', 118, 30);
          drawGauge(ctx, 260, 0, 260, 180, isFinite(mem)?Math.max(0,Math.min(mem,100)):0, 'rgba(96,165,250,.85)');
          ctx.fillText('Mem', 380, 30);
        }

        var c2 = q('chartEvents');
        if(c2 && c2.getContext){
          var ctx2 = c2.getContext('2d');
          ctx2.clearRect(0,0,c2.width,c2.height);
          drawBar(ctx2, 40, 30, 440, 140, [errEv, warnEv], ['rgba(239,68,68,.75)','rgba(245,158,11,.75)'], ['Error','Warn']);
        }

        var c3 = q('chartCis');
        if(c3 && c3.getContext){
          var ctx3 = c3.getContext('2d');
          ctx3.clearRect(0,0,c3.width,c3.height);
          drawDonut(ctx3, 150, 110, 78, 44, [cisPass, cisWarn, cisFail], ['rgba(35,197,94,.75)','rgba(245,158,11,.75)','rgba(239,68,68,.75)']);
          ctx3.fillStyle='rgba(231,236,245,.95)';
          ctx3.font='13px Segoe UI, sans-serif';
          ctx3.fillText('PASS: '+cisPass, 260, 70);
          ctx3.fillText('WARN: '+cisWarn, 260, 100);
          ctx3.fillText('FAIL: '+cisFail, 260, 130);
        }

        var c4 = q('chartStatus');
        if(c4 && c4.getContext){
          var ctx4 = c4.getContext('2d');
          ctx4.clearRect(0,0,c4.width,c4.height);
          drawBar(ctx4, 60, 40, 400, 120, [pendingReboot?1:0, avEnabled?1:0], ['rgba(245,158,11,.75)','rgba(35,197,94,.75)'], ['Reboot','AV']);
        }
      }

      initCharts();
      buildPublisherFilter();
      attachSorting();
      filterTable('tblEvents','qEvents','selEventsLevel');
      filterTable('tblCIS','qCis','selCisStatus');
      filterTable('tblSoftware','qSoftware','selSoftwareVendor');
      filterTable('tblCerts','qCerts','selCertsStatus');
    })();
  </script>
</body>
</html>
"@

  return $html
}

function Escribir-ResumenHealthCheck {
  <#
  .SYNOPSIS
    Escribe un resumen completo al final de la ejecución del script.
  #>
  param(
    [Parameter(Mandatory = $true)]
    [object]$ResumenHealthCheck,
    
    [Parameter(Mandatory = $false)]
    [TimeSpan]$ExecutionTime
  )
  
  if (-not $Script:LogFile) { return }
  
  $summary = @"

========================================
RESUMEN DE EJECUCIÓN
========================================
End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Execution Time: $($ExecutionTime.TotalSeconds) segundos

MÉTRICAS PRINCIPALES:
- CPU Usage: $($ResumenHealthCheck.CpuAvgPct)%
- Memory Usage: $($ResumenHealthCheck.MemUsedPct)%
- Error Events: $($ResumenHealthCheck.ErrorEvents)
- Warning Events: $($ResumenHealthCheck.WarningEvents)
- Pending Reboot: $($ResumenHealthCheck.PendingReboot)
- CIS Failures: $($ResumenHealthCheck.CisFails)

ANTIVIRUS:
- Antivirus Enabled: $($ResumenHealthCheck.AntivirusEnabled)
- Real-Time Protection: $($ResumenHealthCheck.RealTimeProtection)
- Definitions Age: $($ResumenHealthCheck.DefinitionsAge) días
- Threats Detected: $($ResumenHealthCheck.ThreatsDetected)

)

========================================
FIN DE SESIÓN
========================================
"@
  
  Add-Content -Path $Script:LogFile -Value $summary -Encoding UTF8
  Write-HealthLog "Resumen de ejecución guardado" -Level SUCCESS
}

$scriptDir = $null
if ($PSScriptRoot) {
  $scriptDir = $PSScriptRoot
}
elseif ($PSCommandPath) {
  try { $scriptDir = Split-Path -Parent $PSCommandPath -ErrorAction SilentlyContinue } catch {}
}
if (-not $scriptDir) {
  # En EXE, suele ser la carpeta del ejecutable
  $scriptDir = [System.AppDomain]::CurrentDomain.BaseDirectory
}
if (-not $scriptDir) {
  # Último recurso: el directorio actual
  $scriptDir = (Get-Location).Path
}

# Si no pasaron RutaOrigen, crear una por defecto válida
if ([string]::IsNullOrWhiteSpace($RutaOrigen)) {
  try {
    $RutaOrigen = Join-Path $scriptDir 'ServerHealthReport'
  }
  catch {
    # Si por algún motivo falla Join-Path, usar TEMP
    $RutaOrigen = Join-Path $env:TEMP ("ServerHealthReport_{0}" -f $env:COMPUTERNAME)
  }
}

# Validar y crear carpeta destino
Ensure-Folder -Path $RutaOrigen

# Inicializar sistema de logging
if ($EnableLog) {
  Inicializar-HealthLog -LogDirectory $RutaOrigen | Out-Null
}
Write-HealthLog "=== INICIO DE EJECUCIÓN DE HEALTH CHECK ===" -Level INFO
Write-HealthLog "Directorio de salida: $RutaOrigen" -Level INFO
Write-HealthLog "Parámetros: DIASAtras=$DIASAtras, RutaSoftInstalado=$RutaSoftInstalado, ExportJson=$ExportJson, EnableLog=$EnableLog, ExportGpResultHtml=$ExportGpResultHtml, EnableGpResultXmlDetails=$EnableGpResultXmlDetails, ParallelDiagnostics=$ParallelDiagnostics" -Level INFO

$ErrorActionPreference = 'SilentlyContinue'
$now = Get-Date
$since = $now.AddDays( - [math]::Abs($DIASAtras))

function Invoke-DiagnosticJob {
  param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][scriptblock]$JobScript,
    [Parameter(Mandatory = $false)][object[]]$ArgumentList = @(),
    [Parameter(Mandatory = $true)][scriptblock]$FallbackScript,
    [Parameter(Mandatory = $false)][int]$TimeoutSec = 900
  )

  if (-not $ParallelDiagnostics) {
    return & $FallbackScript
  }

  $job = $null
  try {
    $job = Start-Job -Name $Name -ScriptBlock $JobScript -ArgumentList $ArgumentList
    $completed = Wait-Job -Job $job -Timeout $TimeoutSec
    if (-not $completed) {
      try { Stop-Job -Job $job -Force | Out-Null } catch { }
      throw "Timeout en job '$Name' (${TimeoutSec}s)"
    }

    $data = Receive-Job -Job $job -ErrorAction Stop
    return $data
  }
  catch {
    Write-Host "[Parallel] Job '$Name' falló, usando modo secuencial: $($_.Exception.Message)" -ForegroundColor DarkYellow
    Write-HealthLog "Job '$Name' falló, usando modo secuencial: $($_.Exception.Message)" -Level WARNING -ErrorRecord $_
    return & $FallbackScript
  }
  finally {
    if ($job) {
      try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
  }
}

function Get-AvgCpuUsage {
  param([int]$Samples = 5, [int]$IntervalSec = 1)
  try {
    $vals = (Get-Counter '\Processor(_Total)\% Processor Time' -SampleInterval $IntervalSec -MaxSamples $Samples).CounterSamples.CookedValue
    [math]::Round(($vals | Measure-Object -Average).Average, 2)
  }
  catch { $null }
}

function Size-Str {
  param([double]$Bytes)
  if ($null -eq $Bytes) { return "" }
  $sizes = "B", "KB", "MB", "GB", "TB", "PB"
  $i = 0
  while ($Bytes -ge 1024 -and $i -lt $sizes.Length - 1) { $Bytes /= 1024; $i++ }
  "{0:N2} {1}" -f $Bytes, $sizes[$i]
}

function Get-PropValue {
  param(
    [Parameter(Mandatory = $true)]$Object,
    [Parameter(Mandatory = $true)][string]$Name
  )
  if ($null -eq $Object) { return $null }
  $prop = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $Name } | Select-Object -First 1
  if ($prop) { return $prop.Value }
  return $null
}

function Import-SoftwareJsonSafe {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )
  if (-not (Test-Path -LiteralPath $Path)) { throw "RutaSoftInstalado no existe: $Path" }
  try {
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if (-not $raw) { return @() }
    $rows = $raw | ConvertFrom-Json
  }
  catch {
    throw "No se pudo importar el JSON de software: $($_.Exception.Message)"
  }

  if (($rows | Measure-Object).Count -eq 0) { return @() }

  $required = 'Name', 'Version', 'Publisher', 'InstallDate', 'UninstallString'
  $headers = @()
  $first = $rows | Select-Object -First 1
  if ($first) { $headers = ($first | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name | ForEach-Object { $_.Trim() }) }
  foreach ($h in $required) {
    if ($headers -notcontains $h) { throw "El JSON no contiene la propiedad requerida: '$h'" }
  }

  $norm = foreach ($r in $rows) {
    $map = @{}
    foreach ($p in $r.PSObject.Properties) {
      $n = $p.Name -replace "^\uFEFF", ''
      $n = $n.Trim().Trim('"')
      $map[$n] = $p.Value
    }

    $name = [string]$map['Name']
    $ver = [string]$map['Version']
    $pub = [string]$map['Publisher']
    $inst = $map['InstallDate']
    $uni = [string]$map['UninstallString']

    if ($uni) { $uni = $uni.Trim().Trim('"') }
    if ($name) { $name = $name.Trim() }
    if ($ver) { $ver = $ver.Trim() }
    if ($pub) { $pub = $pub.Trim() }

    if ($inst -is [datetime]) {
      $instParsed = $inst
    }
    elseif ([string]::IsNullOrWhiteSpace([string]$inst)) {
      $instParsed = $null
    }
    else {
      $instParsed = try { [datetime]::Parse([string]$inst, $([System.Globalization.CultureInfo]::InvariantCulture)) } catch { $inst }
    }

    [PSCustomObject]@{
      Name            = $name
      Version         = $ver
      Publisher       = $pub
      InstallDate     = $instParsed
      UninstallString = $uni
    }
  }

  return $norm
}

$osCim = Get-CimInstance Win32_OperatingSystem
$ComputerInfo = [PSCustomObject]@{
  Hostname    = $env:COMPUTERNAME
  FQDN        = $( try { [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName } catch { $null } )
  OS          = $osCim.Caption
  OSVersion   = $osCim.Version
  InstallDate = $osCim.InstallDate
  UptimeDays  = [math]::Round(((Get-Date) - $osCim.LastBootUpTime).TotalDays, 2)
}

$IPs = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
Where-Object { $_.IPAddress -notmatch '^169\.254\.' -and $_.IPAddress -ne '127.0.0.1' } |
Select-Object IPAddress, InterfaceAlias, PrefixLength

$regPaths = @(
  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

$softwareJobScript = {
  param($regPaths)
  $rows = foreach ($p in $regPaths) {
    Get-ItemProperty -Path $p -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName } |
    Select-Object @{n = 'Name'; e = { $_.DisplayName } },
    @{n = 'Version'; e = { $_.DisplayVersion } },
    @{n = 'Publisher'; e = { $_.Publisher } },
    @{n = 'InstallDate'; e = {
        if ($_.InstallDate -match '^\d{8}$') {
          try { [datetime]::ParseExact($_.InstallDate, 'yyyyMMdd', $null) } catch { $_.InstallDate }
        }
        else { $_.InstallDate }
      }
    },
    @{n = 'UninstallString'; e = { $_.UninstallString } }
  }
  return ($rows | Sort-Object Name, Version)
}

$Software = Invoke-DiagnosticJob -Name 'Software' -JobScript $softwareJobScript -ArgumentList @($regPaths) -FallbackScript {
  $rows = foreach ($p in $regPaths) {
    Get-ItemProperty -Path $p -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName } |
    Select-Object @{n = 'Name'; e = { $_.DisplayName } },
    @{n = 'Version'; e = { $_.DisplayVersion } },
    @{n = 'Publisher'; e = { $_.Publisher } },
    @{n = 'InstallDate'; e = {
        if ($_.InstallDate -match '^\d{8}$') {
          try { [datetime]::ParseExact($_.InstallDate, 'yyyyMMdd', $null) } catch { $_.InstallDate }
        }
        else { $_.InstallDate }
      }
    },
    @{n = 'UninstallString'; e = { $_.UninstallString } }
  }
  $rows | Sort-Object Name, Version
} -TimeoutSec 240

$cpu = Get-CimInstance Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, LoadPercentage
$cpuAvg = Get-AvgCpuUsage -Samples 6 -IntervalSec 1

$CpuStatus = [PSCustomObject]@{
  Name                    = ($cpu | Select-Object -First 1).Name
  Cores                   = ($cpu | Measure-Object NumberOfCores -Sum).Sum
  LogicalProcessors       = ($cpu | Measure-Object NumberOfLogicalProcessors -Sum).Sum
  MaxClockMHz             = ($cpu | Select-Object -First 1).MaxClockSpeed
  CurrentLoadPct_Snapshot = ($cpu | Select-Object -First 1).LoadPercentage
  AvgLoadPct_6s           = $cpuAvg
}

$LogicalDisks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
  $pctFree = if ($_.Size) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 2) } else { $null }
  [PSCustomObject]@{
    DeviceID    = $_.DeviceID
    VolumeName  = $_.VolumeName
    FileSystem  = $_.FileSystem
    Size        = $_.Size
    SizeStr     = Size-Str $_.Size
    FreeSpace   = $_.FreeSpace
    FreeStr     = Size-Str $_.FreeSpace
    PercentFree = $pctFree
  }
}

$PhysicalDisks = Get-PhysicalDisk -ErrorAction SilentlyContinue | Select-Object FriendlyName, CanPool, HealthStatus, OperationalStatus, Size

$Adapters = Get-NetAdapter -Physical | Select-Object Name, InterfaceDescription, Status, LinkSpeed, MacAddress
$NetStats = foreach ($ad in $Adapters) {
  $s = Get-NetAdapterStatistics -Name $ad.Name -ErrorAction SilentlyContinue
  if ($s) {
    [PSCustomObject]@{
      Name              = $ad.Name
      ReceivedBytes     = $s.ReceivedBytes
      SentBytes         = $s.SentBytes
      ReceivedUnicast   = $s.ReceivedUnicastPackets
      OutboundDiscarded = $s.OutboundDiscardedPackets
      InboundDiscarded  = $s.InboundDiscardedPackets
      InboundErrors     = $s.InboundErrors
      OutboundErrors    = $s.OutboundErrors
    }
  }
}

$HotFixes = Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending

function Get-PendingUpdates {
  try {
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    
    Write-Host "[Updates] Buscando actualizaciones pendientes..." -ForegroundColor DarkCyan
    $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
    
    $pendingUpdates = @()
    foreach ($update in $searchResult.Updates) {
      $categories = ($update.Categories | ForEach-Object { $_.Name }) -join ', '
      $kbArticles = ($update.KBArticleIDs | ForEach-Object { "KB$_" }) -join ', '
      
      $pendingUpdates += [PSCustomObject]@{
        Title          = $update.Title
        KB             = if ($kbArticles) { $kbArticles } else { 'N/A' }
        Categories     = $categories
        Severity       = if ($update.MsrcSeverity) { $update.MsrcSeverity } else { 'Unspecified' }
        IsDownloaded   = $update.IsDownloaded
        RebootRequired = $update.RebootRequired
        SizeBytes      = $update.MaxDownloadSize
        SizeMB         = [math]::Round($update.MaxDownloadSize / 1MB, 2)
      }
    }
    
    Write-Host "[Updates] Encontradas $($pendingUpdates.Count) actualizaciones pendientes" -ForegroundColor DarkCyan
    return $pendingUpdates
  }
  catch {
    Write-Warning "No se pudieron obtener actualizaciones pendientes: $($_.Exception.Message)"
    return @()
  }
}

$pendingUpdatesJobScript = {
  function Get-PendingUpdates {
    try {
      $updateSession = New-Object -ComObject Microsoft.Update.Session
      $updateSearcher = $updateSession.CreateUpdateSearcher()
      $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
      $pendingUpdates = @()
      foreach ($update in $searchResult.Updates) {
        $categories = ($update.Categories | ForEach-Object { $_.Name }) -join ', '
        $kbArticles = ($update.KBArticleIDs | ForEach-Object { "KB$_" }) -join ', '
        $pendingUpdates += [PSCustomObject]@{
          Title          = $update.Title
          KB             = if ($kbArticles) { $kbArticles } else { 'N/A' }
          Categories     = $categories
          Severity       = if ($update.MsrcSeverity) { $update.MsrcSeverity } else { 'Unspecified' }
          IsDownloaded   = $update.IsDownloaded
          RebootRequired = $update.RebootRequired
          SizeBytes      = $update.MaxDownloadSize
          SizeMB         = [math]::Round($update.MaxDownloadSize / 1MB, 2)
        }
      }
      return $pendingUpdates
    }
    catch {
      return @()
    }
  }

  return (Get-PendingUpdates)
}

$PendingUpdates = Invoke-DiagnosticJob -Name 'PendingUpdates' -JobScript $pendingUpdatesJobScript -ArgumentList @() -FallbackScript {
  Get-PendingUpdates
} -TimeoutSec 600
function Get-PendingKBSet {
  param([Parameter(Mandatory = $true)]$PendingUpdates)
  $set = @()
  foreach ($u in $PendingUpdates) {
    $line = [string]$($u.KB)
    if (-not [string]::IsNullOrWhiteSpace($line) -and $line -ne 'N/A') {
      $set += ($line -split '[,;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
  }
  return ($set | Select-Object -Unique)
}

function Get-InstalledKBSetCom {
  try {
    $session = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $searcher.Online = $true
    $res = $searcher.Search("IsInstalled=1 and Type='Software' and IsHidden=0")
    $kbs = @()
    foreach ($u in $res.Updates) {
      if ($u.KBArticleIDs -and $u.KBArticleIDs.Count -gt 0) {
        $kbs += ($u.KBArticleIDs | ForEach-Object { "KB$_" })
      }
    }
    return ($kbs | Select-Object -Unique)
  }
  catch {
    Write-Warning "Fallo detectando KB instaladas vía COM: $($_.Exception.Message)"
    return @()
  }
}

function Diff-Software {
  param($prev, $curr)

  # Índices por Name|Version
  $prevKeys = @{}; foreach ($p in $prev) { $pv = if ($p.Version) { [string]$p.Version }else { '' }; $prevKeys["$($p.Name.ToLowerInvariant())|$pv"] = $p }
  $currKeys = @{}; foreach ($c in $curr) { $cv = if ($c.Version) { [string]$c.Version }else { '' }; $currKeys["$($c.Name.ToLowerInvariant())|$cv"] = $c }

  $added = @(); foreach ($k in $currKeys.Keys) { if (-not $prevKeys.ContainsKey($k)) { $added += $currKeys[$k] } }
  $removed = @(); foreach ($k in $prevKeys.Keys) { if (-not $currKeys.ContainsKey($k)) { $removed += $prevKeys[$k] } }

  # Upgrades por nombre (cambió set de versiones)
  $prevByName = $prev | Group-Object { $_.Name.ToLowerInvariant() } -AsHashTable -AsString
  $currByName = $curr | Group-Object { $_.Name.ToLowerInvariant() } -AsHashTable -AsString

  $upgrades = @()
  foreach ($name in $currByName.Keys) {
    if ($prevByName.ContainsKey($name)) {
      $prevVers = ($prevByName[$name] | ForEach-Object { ([string]$_.Version).Trim() }) | Sort-Object -Unique
      $currVers = ($currByName[$name] | ForEach-Object { ([string]$_.Version).Trim() }) | Sort-Object -Unique
      if (($prevVers -join ',') -ne ($currVers -join ',')) {
        $upgrades += [pscustomobject]@{ Name = $name; From = ($prevVers -join ', '); To = ($currVers -join ', ') }
      }
    }
  }

  # Excluir upgrades de Added/Removed
  $upgNames = @($upgrades | ForEach-Object { $_.Name })
  if ($upgNames.Count -gt 0) {
    $added = $added   | Where-Object { $upgNames -notcontains ($_.Name.ToLowerInvariant()) }
    $removed = $removed | Where-Object { $upgNames -notcontains ($_.Name.ToLowerInvariant()) }
  }

  [pscustomobject]@{ Added = $added; Removed = $removed; Upgrades = $upgrades }
}


function Get-AppliedGPO {
  <#
  .SYNOPSIS
    Obtiene las Group Policy Objects (GPOs) aplicadas al equipo y al usuario actual.
  .DESCRIPTION
    Utiliza gpresult.exe para extraer información detallada de las GPOs aplicadas,
    incluyendo nombre, fecha de aplicación, tipo (Computer/User) y estado.
  #>
  param(
    [Parameter(Mandatory = $false)][switch]$SkipXmlDetails,
    [Parameter(Mandatory = $false)][switch]$EnableXmlDetails
  )
  
  Write-HealthLog "Obteniendo GPOs aplicadas al equipo con detalles de configuración" -Level INFO
  
  $gpoResults = [PSCustomObject]@{
    ComputerGPOs         = @()
    UserGPOs             = @()
    SecuritySettings     = @()
    RegistrySettings     = @()
    Scripts              = @()
    SoftwareRestrictions = @()
    LastRefresh          = $null
    ErrorMessage         = $null
  }
  
  try {
    # Obtener GPOs de equipo usando gpresult
    Write-Host "[GPO] Extrayendo GPOs de equipo..." -ForegroundColor DarkCyan
    
    $computerGpResult = gpresult /R /SCOPE:COMPUTER 2>&1
    
    if ($computerGpResult) {
      $inGpoSection = $false
      
      foreach ($line in $computerGpResult) {
        $lineStr = $line.ToString().Trim()
        
        # Detectar fecha de última actualización (inglés y español)
        if ($lineStr -match "Group Policy was applied at:\s+(.+)" -or 
          $lineStr -match "Directiva de grupo aplicada en:\s+(.+)" -or
          $lineStr -match "La directiva de grupo se aplicó en:\s+(.+)") {
          $gpoResults.LastRefresh = $matches[1]
        }
        
        # Detectar inicio de sección de GPOs (inglés y español)
        if ($lineStr -match "Applied Group Policy Objects" -or 
          $lineStr -match "Objetos de directiva de grupo aplicados") {
          $inGpoSection = $true
          continue
        }
        
        # Detectar fin de sección (inglés y español)
        if ($inGpoSection -and ($lineStr -match "The following GPOs were not applied" -or 
            $lineStr -match "The computer is a part of" -or
            $lineStr -match "Los siguientes GPO no se aplicaron" -or
            $lineStr -match "El equipo forma parte de")) {
          $inGpoSection = $false
        }
        
        # Extraer nombres de GPOs (filtrar N/A, None, Ninguno, etc.)
        if ($inGpoSection -and $lineStr -and $lineStr -ne "-----------------------------") {
          if ($lineStr -notmatch "^Applied Group Policy Objects" -and 
            $lineStr -notmatch "^Objetos de directiva de grupo aplicados" -and
            $lineStr -notmatch "^-+$" -and
            $lineStr -notmatch "^N/A$" -and
            $lineStr -notmatch "^None$" -and
            $lineStr -notmatch "^Ninguno$" -and
            $lineStr -notmatch "^No se") {
            $gpoResults.ComputerGPOs += [PSCustomObject]@{
              Name    = $lineStr
              Type    = "Computer"
              Applied = $true
            }
          }
        }
      }
      
      Write-HealthLog "GPOs de equipo extraídas: $($gpoResults.ComputerGPOs.Count)" -Level SUCCESS
    }
    
    # Obtener GPOs de usuario usando gpresult
    Write-Host "[GPO] Extrayendo GPOs de usuario..." -ForegroundColor DarkCyan
    
    try {
      $userGpResult = gpresult /R /SCOPE:USER 2>&1
      
      if ($userGpResult) {
        $inGpoSection = $false
        
        foreach ($line in $userGpResult) {
          $lineStr = $line.ToString().Trim()
          
          # Detectar inicio de sección de GPOs (inglés y español)
          if ($lineStr -match "Applied Group Policy Objects" -or 
            $lineStr -match "Objetos de directiva de grupo aplicados") {
            $inGpoSection = $true
            continue
          }
          
          # Detectar fin de sección (inglés y español)
          if ($inGpoSection -and ($lineStr -match "The following GPOs were not applied" -or 
              $lineStr -match "The user is a part of" -or
              $lineStr -match "Los siguientes GPO no se aplicaron" -or
              $lineStr -match "El usuario forma parte de")) {
            $inGpoSection = $false
          }
          
          # Extraer nombres de GPOs (filtrar N/A, None, Ninguno, etc.)
          if ($inGpoSection -and $lineStr -and $lineStr -ne "-----------------------------") {
            if ($lineStr -notmatch "^Applied Group Policy Objects" -and 
              $lineStr -notmatch "^Objetos de directiva de grupo aplicados" -and
              $lineStr -notmatch "^-+$" -and
              $lineStr -notmatch "^N/A$" -and
              $lineStr -notmatch "^None$" -and
              $lineStr -notmatch "^Ninguno$" -and
              $lineStr -notmatch "^No se") {
              $gpoResults.UserGPOs += [PSCustomObject]@{
                Name    = $lineStr
                Type    = "User"
                Applied = $true
              }
            }
          }
        }
        
        Write-HealthLog "GPOs de usuario extraídas: $($gpoResults.UserGPOs.Count)" -Level SUCCESS
      }
    }
    catch {
      Write-HealthLog "No se pudieron obtener GPOs de usuario: $($_.Exception.Message)" -Level WARNING
    }
    
    if ($EnableXmlDetails -and (-not $SkipXmlDetails)) {
      # =========================================================================
      # EXTRAER CONFIGURACIONES DETALLADAS USANDO GPRESULT /X (XML)
      # =========================================================================
      
      Write-Host "[GPO] Extrayendo configuraciones detalladas (XML)..." -ForegroundColor DarkCyan
      
      $xmlPath = $null
      try {
        $xmlPath = "$env:TEMP\gpresult_detailed_$([Guid]::NewGuid().ToString()).xml"
        Write-Host "[GPO] Generando archivo XML en: $xmlPath" -ForegroundColor DarkCyan
        $gpresultOutput = gpresult /X $xmlPath /F 2>&1
        
        if (Test-Path $xmlPath) {
          Write-Host "[GPO] Archivo XML generado exitosamente. Tamaño: $((Get-Item $xmlPath).Length) bytes" -ForegroundColor Green
          
          [xml]$gpResultXml = Get-Content $xmlPath -Encoding UTF8
          Write-HealthLog "XML de gpresult cargado correctamente" -Level INFO
          
          # ===== CONFIGURAR NAMESPACE MANAGER =====
          # Crear namespace manager para manejar prefijos en el XML
          $nsManager = New-Object System.Xml.XmlNamespaceManager($gpResultXml.NameTable)
          
          # Detectar y registrar namespaces del documento
          if ($gpResultXml.DocumentElement.NamespaceURI) {
            $defaultNs = $gpResultXml.DocumentElement.NamespaceURI
            $nsManager.AddNamespace("rsop", $defaultNs)
            Write-HealthLog "Namespace detectado y registrado: $defaultNs" -Level DEBUG
          }
          
          # Registrar namespaces comunes usados por gpresult
          $nsManager.AddNamespace("q1", "http://www.microsoft.com/GroupPolicy/Rsop")
          $nsManager.AddNamespace("q2", "http://www.microsoft.com/GroupPolicy/Settings")
          
          # ===== CONFIGURACIONES DE SEGURIDAD =====
          # Usar SelectNodes con namespace manager
          
          # Security Options - intentar con y sin namespace
          $securityNodes = @()
          try {
            $securityNodes += $gpResultXml.SelectNodes("//q1:SecurityOptions/q1:KeyName", $nsManager)
          }
          catch {
            Write-HealthLog "No se encontraron SecurityOptions con namespace q1" -Level DEBUG
          }
          
          try {
            $securityNodes += $gpResultXml.SelectNodes("//SecurityOptions/KeyName")
          }
          catch {
            Write-HealthLog "No se encontraron SecurityOptions sin namespace" -Level DEBUG
          }
          
          # Filtrar nulos
          $securityNodes = @($securityNodes | Where-Object { $_ -ne $null })
          Write-Host "[GPO] Security Options encontradas: $($securityNodes.Count)" -ForegroundColor DarkCyan
          
          foreach ($node in $securityNodes) {
            if ($node) {
              $parent = $node.ParentNode
              $gpoResults.SecuritySettings += [PSCustomObject]@{
                Category = "Security Options"
                Setting  = $node.InnerText
                Value    = if ($parent.SettingNumber) { $parent.SettingNumber } elseif ($parent.SettingString) { $parent.SettingString } else { "N/A" }
                GPO      = if ($parent.ParentNode.GPO) { $parent.ParentNode.GPO.Name } else { "Local" }
              }
            }
          }
          
          # Políticas de contraseñas
          $pwdNodes = @()
          try {
            $pwdNodes += $gpResultXml.SelectNodes("//PasswordPolicy/*")
          }
          catch {
            Write-HealthLog "No se encontraron PasswordPolicy" -Level DEBUG
          }
          
          $pwdNodes = @($pwdNodes | Where-Object { $_ -ne $null })
          foreach ($node in $pwdNodes) {
            if ($node.Name -ne '#text' -and $node.InnerText) {
              $gpoResults.SecuritySettings += [PSCustomObject]@{
                Category = "Password Policy"
                Setting  = $node.Name
                Value    = $node.InnerText
                GPO      = "Domain Policy"
              }
            }
          }
          
          # Políticas de bloqueo de cuentas
          $lockoutNodes = @()
          try {
            $lockoutNodes += $gpResultXml.SelectNodes("//LockoutPolicy/*")
          }
          catch {
            Write-HealthLog "No se encontraron LockoutPolicy" -Level DEBUG
          }
          
          $lockoutNodes = @($lockoutNodes | Where-Object { $_ -ne $null })
          foreach ($node in $lockoutNodes) {
            if ($node.Name -ne '#text' -and $node.InnerText) {
              $gpoResults.SecuritySettings += [PSCustomObject]@{
                Category = "Account Lockout Policy"
                Setting  = $node.Name
                Value    = $node.InnerText
                GPO      = "Domain Policy"
              }
            }
          }
          
          # Políticas de auditoría
          $auditNodes = @()
          try {
            $auditNodes += $gpResultXml.SelectNodes("//AuditPolicy/*/SubcategoryName")
          }
          catch {
            Write-HealthLog "No se encontraron AuditPolicy" -Level DEBUG
          }
          
          $auditNodes = @($auditNodes | Where-Object { $_ -ne $null })
          foreach ($node in $auditNodes) {
            $parent = $node.ParentNode
            $gpoResults.SecuritySettings += [PSCustomObject]@{
              Category = "Audit Policy"
              Setting  = $node.InnerText
              Value    = if ($parent.SettingValue) { $parent.SettingValue } else { "Not Configured" }
              GPO      = "Domain Policy"
            }
          }
          
          # Derechos de usuario (User Rights Assignment)
          $userRightsNodes = @()
          try {
            $userRightsNodes += $gpResultXml.SelectNodes("//UserRightsAssignment/Name")
          }
          catch {
            Write-HealthLog "No se encontraron UserRightsAssignment" -Level DEBUG
          }
          
          $userRightsNodes = @($userRightsNodes | Where-Object { $_ -ne $null })
          foreach ($node in $userRightsNodes) {
            $parent = $node.ParentNode
            $members = @()
            try {
              $members = $parent.SelectNodes("Member") | ForEach-Object { $_.Name.'#text' }
            }
            catch {
              # Intentar sin SelectNodes
              $members = $parent.Member | ForEach-Object { $_.Name.'#text' }
            }
            $gpoResults.SecuritySettings += [PSCustomObject]@{
              Category = "User Rights Assignment"
              Setting  = $node.InnerText
              Value    = if ($members) { $members -join "; " } else { "None" }
              GPO      = if ($parent.ParentNode.GPO) { $parent.ParentNode.GPO.Name } else { "Local" }
            }
          }
          
          # ===== CONFIGURACIONES DE REGISTRO =====
          $registryNodes = @()
          try {
            $registryNodes += $gpResultXml.SelectNodes("//RegistrySettings/Registry")
          }
          catch {
            Write-HealthLog "No se encontraron RegistrySettings" -Level DEBUG
          }
          
          $registryNodes = @($registryNodes | Where-Object { $_ -ne $null })
          foreach ($node in $registryNodes) {
            $gpoResults.RegistrySettings += [PSCustomObject]@{
              Action = if ($node.Properties.action) { $node.Properties.action } else { "Update" }
              Hive   = if ($node.Properties.hive) { $node.Properties.hive } else { "Unknown" }
              Key    = if ($node.Properties.key) { $node.Properties.key } else { "Unknown" }
              Name   = if ($node.Properties.name) { $node.Properties.name } else { "(Default)" }
              Value  = if ($node.Properties.value) { $node.Properties.value } else { "" }
              Type   = if ($node.Properties.type) { $node.Properties.type } else { "REG_SZ" }
              GPO    = if ($node.ParentNode.GPO) { $node.ParentNode.GPO.Name } else { "Unknown" }
            }
          }
          
          # ===== SCRIPTS DE INICIO/CIERRE =====
          $scriptNodes = @()
          try {
            $scriptNodes += $gpResultXml.SelectNodes("//Scripts/Script")
          }
          catch {
            Write-HealthLog "No se encontraron Scripts" -Level DEBUG
          }
          
          $scriptNodes = @($scriptNodes | Where-Object { $_ -ne $null })
          foreach ($node in $scriptNodes) {
            $gpoResults.Scripts += [PSCustomObject]@{
              Type       = if ($node.ParentNode.Name -match "Startup") { "Startup" } 
              elseif ($node.ParentNode.Name -match "Shutdown") { "Shutdown" }
              elseif ($node.ParentNode.Name -match "Logon") { "Logon" }
              elseif ($node.ParentNode.Name -match "Logoff") { "Logoff" }
              else { "Unknown" }
              Command    = if ($node.Command) { $node.Command } else { "Unknown" }
              Parameters = if ($node.Parameters) { $node.Parameters } else { "" }
              GPO        = if ($node.ParentNode.GPO) { $node.ParentNode.GPO.Name } else { "Unknown" }
            }
          }
          
          # ===== RESTRICCIONES DE SOFTWARE =====
          $swRestrictionNodes = @()
          try {
            $swRestrictionNodes += $gpResultXml.SelectNodes("//SoftwareRestrictions/*")
          }
          catch {
            Write-HealthLog "No se encontraron SoftwareRestrictions" -Level DEBUG
          }
          
          $swRestrictionNodes = @($swRestrictionNodes | Where-Object { $_ -ne $null })
          try {
            foreach ($node in $swRestrictionNodes) {
              if ($node.Name -ne '#text') {
                $gpoResults.SoftwareRestrictions += [PSCustomObject]@{
                  Type          = $node.Name
                  Path          = if ($node.Path) { $node.Path } else { "N/A" }
                  SecurityLevel = if ($node.SecurityLevel) { $node.SecurityLevel } else { "N/A" }
                  Description   = if ($node.Description) { $node.Description } else { "" }
                  GPO           = if ($node.ParentNode.GPO) { $node.ParentNode.GPO.Name } else { "Unknown" }
                }
              }
            }
          }
          catch {
            Write-HealthLog "Error extrayendo configuraciones detalladas del XML: $($_.Exception.Message)" -Level WARNING
            Write-Host "[GPO] Error al extraer configuraciones del XML: $($_.Exception.Message)" -ForegroundColor Yellow
          }

          # Limpiar XML temporal generado por gpresult
          try {
            Remove-Item -LiteralPath $xmlPath -Force -ErrorAction SilentlyContinue
          }
          catch { }
        }
        else {
          Write-HealthLog "No se pudo generar el archivo XML de gpresult" -Level WARNING
        }
      }
      catch {
        Write-HealthLog "Error extrayendo configuraciones detalladas (XML): $($_.Exception.Message)" -Level WARNING
        Write-Host "[GPO] Error al extraer configuraciones del XML: $($_.Exception.Message)" -ForegroundColor Yellow
        try {
          if ($xmlPath) { Remove-Item -LiteralPath $xmlPath -Force -ErrorAction SilentlyContinue }
        }
        catch { }
      }
      finally {
        if ($xmlPath -and (Test-Path -LiteralPath $xmlPath)) {
          try { Remove-Item -LiteralPath $xmlPath -Force -ErrorAction SilentlyContinue | Out-Null } catch { }
        }
      }
    }
    else {
      Write-HealthLog "Se omite gpresult /X (XML) porque SkipXmlDetails está habilitado" -Level DEBUG
    }

    # =========================================================================
    # EXTRACCIÓN ALTERNATIVA: Políticas de Seguridad Locales con SECEDIT
    # =========================================================================
    
    Write-Host "[GPO] Extrayendo políticas de seguridad locales con secedit..." -ForegroundColor DarkCyan
    
    try {
      $seceditPath = "$env:TEMP\secedit_export_$([Guid]::NewGuid().ToString()).inf"
      $null = secedit /export /cfg $seceditPath /quiet 2>&1
      
      if (Test-Path $seceditPath) {
        Write-Host "[GPO] Políticas exportadas con secedit exitosamente" -ForegroundColor Green
        
        $seceditContent = Get-Content $seceditPath -Encoding Unicode
        $currentSection = ""
        
        foreach ($line in $seceditContent) {
          $line = $line.Trim()
          
          # Detectar sección
          if ($line -match '^\[(.+)\]$') {
            $currentSection = $matches[1]
            continue
          }
          
          # Procesar líneas con configuraciones
          if ($line -and $line -notmatch '^;' -and $line -match '(.+)=(.*)') {
            $setting = $matches[1].Trim()
            $value = $matches[2].Trim()
            
            # Categorizar por sección
            $category = switch ($currentSection) {
              "System Access" { "Password & Account Lockout Policy" }
              "Event Audit" { "Audit Policy" }
              "Registry Values" { "Security Options (Registry)" }
              "Privilege Rights" { "User Rights Assignment" }
              default { $currentSection }
            }
            
            # Traducir valores numéricos a texto legible
            if ($currentSection -eq "Event Audit") {
              $value = switch ($value) {
                "0" { "No Auditing" }
                "1" { "Success" }
                "2" { "Failure" }
                "3" { "Success, Failure" }
                default { $value }
              }
            }
            
            # Agregar a configuraciones
            $gpoResults.SecuritySettings += [PSCustomObject]@{
              Category = $category
              Setting  = $setting
              Value    = $value
              GPO      = "Local Security Policy"
            }
          }
        }
        
        Remove-Item $seceditPath -Force -ErrorAction SilentlyContinue
        Write-HealthLog "Políticas de seguridad locales extraídas con secedit: $($seceditContent.Count) líneas procesadas" -Level SUCCESS
        Write-Host "[GPO] Políticas de seguridad extraídas: $(($gpoResults.SecuritySettings | Where-Object { $_.GPO -eq 'Local Security Policy' }).Count) configuraciones" -ForegroundColor Green
      }
      else {
        Write-HealthLog "No se pudo exportar políticas con secedit" -Level WARNING
      }
    }
    catch {
      Write-HealthLog "Error extrayendo políticas con secedit: $($_.Exception.Message)" -Level WARNING
      Write-Host "[GPO] Error al exportar políticas con secedit" -ForegroundColor Yellow
    }
    
    # Intentar obtener información adicional con Get-GPResultantSetOfPolicy (requiere RSAT)
    try {
      if (Get-Command Get-GPResultantSetOfPolicy -ErrorAction SilentlyContinue) {
        Write-Host "[GPO] Obteniendo información detallada con RSAT..." -ForegroundColor DarkCyan
        
        Get-GPResultantSetOfPolicy -ReportType Xml -Path "$env:TEMP\rsop_temp.xml" -ErrorAction SilentlyContinue | Out-Null
        
        if (Test-Path "$env:TEMP\rsop_temp.xml") {
          [xml]$rsopXml = Get-Content "$env:TEMP\rsop_temp.xml"
          
          # Enriquecer datos con información del XML si está disponible
          foreach ($gpo in $rsopXml.SelectNodes("//GPO")) {
            $gpoName = $gpo.Name
            $gpoPath = $gpo.Path
            
            # Buscar en GPOs de equipo
            $matchComputer = $gpoResults.ComputerGPOs | Where-Object { $_.Name -like "*$gpoName*" }
            if ($matchComputer) {
              $matchComputer | Add-Member -NotePropertyName "Path" -NotePropertyValue $gpoPath -Force
            }
            
            # Buscar en GPOs de usuario
            $matchUser = $gpoResults.UserGPOs | Where-Object { $_.Name -like "*$gpoName*" }
            if ($matchUser) {
              $matchUser | Add-Member -NotePropertyName "Path" -NotePropertyValue $gpoPath -Force
            }
          }
          
          Remove-Item "$env:TEMP\rsop_temp.xml" -Force -ErrorAction SilentlyContinue
          Write-HealthLog "Información detallada de GPOs obtenida con RSAT" -Level SUCCESS
        }
      }
    }
    catch {
      # RSAT no disponible, continuar sin información adicional
      Write-HealthLog "RSAT no disponible para información detallada de GPOs" -Level DEBUG
    }
    
    $totalGpos = $gpoResults.ComputerGPOs.Count + $gpoResults.UserGPOs.Count
    
    if ($totalGpos -eq 0) {
      Write-Host "[GPO] No se detectaron GPOs específicas aplicadas (solo políticas locales)" -ForegroundColor Yellow
      Write-HealthLog "No se detectaron GPOs de dominio aplicadas - El equipo puede no estar unido a un dominio" -Level WARNING
    }
    else {
      Write-Host "[GPO] Total de GPOs aplicadas: $totalGpos (Equipo: $($gpoResults.ComputerGPOs.Count), Usuario: $($gpoResults.UserGPOs.Count))" -ForegroundColor Green
      Write-HealthLog "Total de GPOs aplicadas: $totalGpos" -Level SUCCESS
    }
    
  }
  catch {
    $gpoResults.ErrorMessage = $_.Exception.Message
    Write-HealthLog "Error al obtener GPOs aplicadas: $($_.Exception.Message)" -Level ERROR -ErrorRecord $_
    Write-Host "[GPO] Error al extraer GPOs: $($_.Exception.Message)" -ForegroundColor Red
  }
  
  return $gpoResults
}

function Test-PendingReboot {
  $reboot = $false
  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
    'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations'
  )
  foreach ($p in $paths) { if (Test-Path $p) { $reboot = $true } }
  return $reboot
}
function Get-CertificateExpiryReport {
  param(
    [Parameter(Mandatory = $false)][int]$DaysThreshold = 30,
    [Parameter(Mandatory = $false)][string]$TlsHost = '127.0.0.1',
    [Parameter(Mandatory = $false)][int[]]$TlsPorts = @(443, 7272)
  )

  $now = Get-Date
  $rows = New-Object System.Collections.Generic.List[psobject]

  $storePaths = @(
    'Cert:\LocalMachine\My',
    'Cert:\LocalMachine\WebHosting',
    'Cert:\CurrentUser\My'
  )

  foreach ($sp in $storePaths) {
    if (-not (Test-Path $sp)) { continue }
    try {
      foreach ($cert in (Get-ChildItem -Path $sp -ErrorAction Stop)) {
        if (-not $cert) { continue }
        $daysRemaining = $null
        $status = 'Unknown'
        try {
          $daysRemaining = [math]::Floor(($cert.NotAfter - $now).TotalDays)
          if ($cert.NotAfter -lt $now) { $status = 'Expired' }
          elseif ($daysRemaining -le $DaysThreshold) { $status = 'ExpiringSoon' }
          else { $status = 'Valid' }
        }
        catch { }

        $rows.Add([PSCustomObject]@{
            Source        = 'Store'
            Location      = $sp
            Subject       = $cert.Subject
            Issuer        = $cert.Issuer
            Thumbprint    = $cert.Thumbprint
            FriendlyName  = $cert.FriendlyName
            NotBefore     = $cert.NotBefore
            NotAfter      = $cert.NotAfter
            DaysRemaining = $daysRemaining
            HasPrivateKey = $cert.HasPrivateKey
            Status        = $status
            ParseStatus   = 'OK'
          })
      }
    }
    catch {
      $rows.Add([PSCustomObject]@{
          Source        = 'Store'
          Location      = $sp
          Subject       = ''
          Issuer        = ''
          Thumbprint    = ''
          FriendlyName  = ''
          NotBefore     = $null
          NotAfter      = $null
          DaysRemaining = $null
          HasPrivateKey = $null
          Status        = 'Unknown'
          ParseStatus   = "ERROR: $($_.Exception.Message)"
        })
    }
  }

  function Get-TlsRemoteCertificate {
    param(
      [Parameter(Mandatory = $true)][string]$TargetHost,
      [Parameter(Mandatory = $true)][int]$Port,
      [Parameter(Mandatory = $false)][int]$TimeoutMs = 3000
    )

    $tcp = $null
    $ssl = $null
    try {
      $tcp = New-Object System.Net.Sockets.TcpClient
      $iar = $tcp.BeginConnect($TargetHost, $Port, $null, $null)
      if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
        throw "ConnectTimeout"
      }
      $tcp.EndConnect($iar)
      $ssl = New-Object System.Net.Security.SslStream($tcp.GetStream(), $false, ({ $true }))
      $ssl.ReadTimeout = $TimeoutMs
      $ssl.WriteTimeout = $TimeoutMs
      $ssl.AuthenticateAsClient($TargetHost)
      if (-not $ssl.RemoteCertificate) {
        throw "NoRemoteCertificate"
      }
      return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)
    }
    finally {
      try { if ($ssl) { $ssl.Dispose() } } catch { }
      try { if ($tcp) { $tcp.Close() } } catch { }
    }
  }

  foreach ($p in $TlsPorts) {
    $cert3 = $null
    $parseStatus = 'Unknown'
    $notBefore = $null
    $notAfter = $null
    $subject = ''
    $issuer = ''
    $thumb = ''
    $daysRemaining = $null
    $status = 'Unknown'

    try {
      $cert3 = Get-TlsRemoteCertificate -TargetHost $TlsHost -Port $p
      $parseStatus = 'OK'
    }
    catch {
      $parseStatus = "ConnectFailed: $($_.Exception.Message)"
    }

    if ($cert3 -and $parseStatus -eq 'OK') {
      $notBefore = $cert3.NotBefore
      $notAfter = $cert3.NotAfter
      $subject = $cert3.Subject
      $issuer = $cert3.Issuer
      $thumb = $cert3.Thumbprint
      $daysRemaining = [math]::Floor(($notAfter - $now).TotalDays)
      if ($notAfter -lt $now) { $status = 'Expired' }
      elseif ($daysRemaining -le $DaysThreshold) { $status = 'ExpiringSoon' }
      else { $status = 'Valid' }
    }

    $rows.Add([PSCustomObject]@{
        Source        = 'TLS'
        Location      = "$TlsHost`:$p"
        Subject       = $subject
        Issuer        = $issuer
        Thumbprint    = $thumb
        FriendlyName  = "TLS $p"
        NotBefore     = $notBefore
        NotAfter      = $notAfter
        DaysRemaining = $daysRemaining
        HasPrivateKey = $null
        Status        = $status
        ParseStatus   = $parseStatus
      })
  }

  return $rows.ToArray()
}

$PendingReboot = Test-PendingReboot

$certJobScript = {
  param($DaysThreshold, $TlsHost, $TlsPorts)
  function Get-CertificateExpiryReport {
    param(
      [int]$DaysThreshold = 30,
      [string]$TlsHost = '127.0.0.1',
      [int[]]$TlsPorts = @(443, 7272)
    )

    $now = Get-Date
    $rows = New-Object System.Collections.Generic.List[psobject]
    $storePaths = @(
      'Cert:\LocalMachine\My',
      'Cert:\LocalMachine\WebHosting',
      'Cert:\CurrentUser\My'
    )

    foreach ($sp in $storePaths) {
      if (-not (Test-Path $sp)) { continue }
      foreach ($cert in (Get-ChildItem -Path $sp -ErrorAction SilentlyContinue)) {
        if (-not $cert) { continue }
        $daysRemaining = $null
        $status = 'Unknown'
        try {
          $daysRemaining = [math]::Floor(($cert.NotAfter - $now).TotalDays)
          if ($cert.NotAfter -lt $now) { $status = 'Expired' }
          elseif ($daysRemaining -le $DaysThreshold) { $status = 'ExpiringSoon' }
          else { $status = 'Valid' }
        }
        catch { }

        $rows.Add([PSCustomObject]@{
            Source        = 'Store'
            Location      = $sp
            Subject       = $cert.Subject
            Issuer        = $cert.Issuer
            Thumbprint    = $cert.Thumbprint
            FriendlyName  = $cert.FriendlyName
            NotBefore     = $cert.NotBefore
            NotAfter      = $cert.NotAfter
            DaysRemaining = $daysRemaining
            HasPrivateKey = $cert.HasPrivateKey
            Status        = $status
            ParseStatus   = 'OK'
          })
      }
    }

    function Get-TlsRemoteCertificate {
      param(
        [string]$TargetHost,
        [int]$Port,
        [int]$TimeoutMs = 3000
      )

      $tcp = $null
      $ssl = $null
      try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $iar = $tcp.BeginConnect($TargetHost, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { throw "ConnectTimeout" }
        $tcp.EndConnect($iar)
        $ssl = New-Object System.Net.Security.SslStream($tcp.GetStream(), $false, ({ $true }))
        $ssl.ReadTimeout = $TimeoutMs
        $ssl.WriteTimeout = $TimeoutMs
        $ssl.AuthenticateAsClient($TargetHost)
        if (-not $ssl.RemoteCertificate) { throw "NoRemoteCertificate" }
        return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)
      }
      finally {
        try { if ($ssl) { $ssl.Dispose() } } catch { }
        try { if ($tcp) { $tcp.Close() } } catch { }
      }
    }

    foreach ($p in $TlsPorts) {
      $cert3 = $null
      $parseStatus = 'Unknown'
      $notBefore = $null
      $notAfter = $null
      $subject = ''
      $issuer = ''
      $thumb = ''
      $daysRemaining = $null
      $status = 'Unknown'

      try {
        $cert3 = Get-TlsRemoteCertificate -TargetHost $TlsHost -Port $p
        $parseStatus = 'OK'
      }
      catch {
        $parseStatus = "ConnectFailed: $($_.Exception.Message)"
      }

      if ($cert3 -and $parseStatus -eq 'OK') {
        $notBefore = $cert3.NotBefore
        $notAfter = $cert3.NotAfter
        $subject = $cert3.Subject
        $issuer = $cert3.Issuer
        $thumb = $cert3.Thumbprint
        $daysRemaining = [math]::Floor(($notAfter - $now).TotalDays)
        if ($notAfter -lt $now) { $status = 'Expired' }
        elseif ($daysRemaining -le $DaysThreshold) { $status = 'ExpiringSoon' }
        else { $status = 'Valid' }
      }

      $rows.Add([PSCustomObject]@{
          Source        = 'TLS'
          Location      = "$TlsHost`:$p"
          Subject       = $subject
          Issuer        = $issuer
          Thumbprint    = $thumb
          FriendlyName  = "TLS $p"
          NotBefore     = $notBefore
          NotAfter      = $notAfter
          DaysRemaining = $daysRemaining
          HasPrivateKey = $null
          Status        = $status
          ParseStatus   = $parseStatus
        })
    }

    return $rows.ToArray()
  }

  return (Get-CertificateExpiryReport -DaysThreshold $DaysThreshold -TlsHost $TlsHost -TlsPorts $TlsPorts)
}

Write-Host "[Certs] Revisando certificados (stores + TLS 443/7272)..." -ForegroundColor DarkCyan
$Certificates = Invoke-DiagnosticJob -Name 'Certificates' -JobScript $certJobScript -ArgumentList @([int]30, '127.0.0.1', @(443, 7272)) -FallbackScript {
  Get-CertificateExpiryReport -DaysThreshold 30 -TlsHost '127.0.0.1' -TlsPorts @(443, 7272)
} -TimeoutSec 300
Write-HealthLog "Certificados recolectados: $(($Certificates | Measure-Object).Count)" -Level INFO

try {
  $mem = Get-CimInstance Win32_OperatingSystem
  $totalMemMB = [math]::Round($mem.TotalVisibleMemorySize / 1024, 0)
  $freeMemMB = [math]::Round($mem.FreePhysicalMemory / 1024, 0)
  $usedMemMB = $totalMemMB - $freeMemMB
  $memPctUsed = if ($totalMemMB -gt 0) { [math]::Round(($usedMemMB / $totalMemMB) * 100, 2) } else { $null }
  $MemoryStatus = [PSCustomObject]@{
    TotalMB = $totalMemMB
    UsedMB  = $usedMemMB
    FreeMB  = $freeMemMB
    UsedPct = $memPctUsed
  }
}
catch { $MemoryStatus = $null }

$logs = @('System', 'Application', 'Security')
$EventsRaw = @{}
$EventsSummary = @()
$EventLogErrors = @{}

$importantSecurityEvents = @(
  4625,  # Failed logon
  4624,  # Successful logon
  4720,  # User account created
  4722,  # User account enabled
  4723,  # Password change attempt
  4724,  # Password reset attempt
  4728,  # User added to security group
  4732,  # User added to local group
  4756,  # User added to universal security group
  4648,  # Logon using explicit credentials
  4672,  # Special privileges assigned to new logon
  1102   # Audit log cleared (critical!)
)

foreach ($log in $logs) {
  try {
    # Special handling for Security log - use Event IDs instead of Level
    if ($log -eq 'Security') {
      $ev = Get-WinEvent -FilterHashtable @{LogName = $log; ID = $importantSecurityEvents; StartTime = $since } -ErrorAction Stop
    }
    else {
      $ev = Get-WinEvent -FilterHashtable @{LogName = $log; Level = @(2, 3); StartTime = $since } -ErrorAction Stop
    }
    
    $grouped = $ev | Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, Message |
    ForEach-Object {
      $msg = $_.Message
      if ($msg.Length -gt 200) { $msg = $msg.Substring(0, 200) + '...' }
      [PSCustomObject]@{ Log = $log; Id = $_.Id; Provider = $_.ProviderName; Level = $_.LevelDisplayName; Message = $msg }
    } | Group-Object Log, Id, Provider, Level, Message | Sort-Object Count -Descending

    $EventsRaw[$log] = $ev
    $EventsSummary += $grouped | Select-Object @{n = 'Log'; e = { $_.Group[0].Log } },
    @{n = 'EventId'; e = { $_.Group[0].Id } },
    @{n = 'Provider'; e = { $_.Group[0].Provider } },
    @{n = 'Level'; e = { $_.Group[0].Level } },
    @{n = 'MessageSample'; e = { $_.Group[0].Message } },
    Count
  }
  catch {
    $errorMsg = $_.Exception.Message
    
    # Special handling for Security log
    if ($log -eq 'Security') {
      if ($errorMsg -match 'No events were found|No se encontraron eventos') {
        Write-Host "[Events] Log Security: No se encontraron eventos de seguridad importantes en los últimos $DIASAtras días (esto es normal si no hay actividad sospechosa)" -ForegroundColor DarkCyan
        $EventLogErrors[$log] = "No important security events found (normal)"
      }
      elseif ($errorMsg -match 'Access is denied|Acceso denegado') {
        Write-Warning "No se pudo leer el log Security: Se requieren permisos de Administrador"
        $EventLogErrors[$log] = "Access denied - Administrator privileges required"
      }
      else {
        Write-Warning "No se pudo leer el log $($log): $errorMsg"
        $EventLogErrors[$log] = $errorMsg
      }
    }
    else {
      # For other logs, just warn
      if ($errorMsg -match 'No events were found|No se encontraron eventos') {
        Write-Host "[Events] Log $log : No se encontraron eventos de Error/Warning en los últimos $DIASAtras días" -ForegroundColor DarkCyan
        $EventLogErrors[$log] = "No events found"
      }
      else {
        Write-Warning "No se pudo leer el log $($log): $errorMsg"
        $EventLogErrors[$log] = $errorMsg
      }
    }
  }
}

# =========================================================================
# REPORTE GPO (opcional) + OBTENER GPOs APLICADAS
# =========================================================================

$GpResultHtmlLink = $null
if ($ExportGpResultHtml) {
  try {
    $gpHtmlPath = Join-Path $RutaOrigen 'gpresult.html'
    Write-Host "[GPO] Generando reporte gpresult HTML: $gpHtmlPath" -ForegroundColor DarkCyan
    $null = gpresult /H $gpHtmlPath /F 2>&1
    if (Test-Path -LiteralPath $gpHtmlPath) {
      $GpResultHtmlLink = (Split-Path -Leaf $gpHtmlPath)
      Write-Host "[GPO] Reporte gpresult HTML generado" -ForegroundColor Green
      Write-HealthLog "Reporte gpresult HTML generado: $gpHtmlPath" -Level SUCCESS
    }
  }
  catch {
    Write-HealthLog "Error generando gpresult HTML: $($_.Exception.Message)" -Level WARNING -ErrorRecord $_
    Write-Host "[GPO] Error generando gpresult HTML: $($_.Exception.Message)" -ForegroundColor Yellow
  }
}

$AppliedGPOs = Get-AppliedGPO -SkipXmlDetails:$ExportGpResultHtml

$cis = @()

# Agregar verificación de GPOs aplicadas a CIS
if ($AppliedGPOs.ComputerGPOs.Count -gt 0 -or $AppliedGPOs.UserGPOs.Count -gt 0) {
  $totalGpos = $AppliedGPOs.ComputerGPOs.Count + $AppliedGPOs.UserGPOs.Count
  $cis += [PSCustomObject]@{ 
    Control = 'Group Policy Objects applied'
    Status  = 'PASS'
    Detail  = "Total GPOs: $totalGpos (Computer: $($AppliedGPOs.ComputerGPOs.Count), User: $($AppliedGPOs.UserGPOs.Count)). Last refresh: $($AppliedGPOs.LastRefresh)"
  }
}
else {
  $cis += [PSCustomObject]@{ 
    Control = 'Group Policy Objects applied'
    Status  = 'WARN'
    Detail  = 'No GPOs detected - This may indicate the system is not domain-joined or GPO application failed'
  }
}

$fwProfiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue
if ($fwProfiles) {
  $allOn = -not ($fwProfiles.Enabled -contains $false)
  $cis += [PSCustomObject]@{ Control = 'Firewall enabled (Domain/Private/Public)'; Status = if ($allOn) { 'PASS' }else { 'FAIL' }; Detail = ($fwProfiles | Select-Object Name, Enabled | Format-Table -Auto | Out-String) }
}
else {
  $cis += [PSCustomObject]@{ Control = 'Firewall enabled (Domain/Private/Public)'; Status = 'UNKNOWN'; Detail = 'No data' }
}

try {
  $denyRdp = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server').fDenyTSConnections
  $nla = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp').UserAuthentication
  $cis += [PSCustomObject]@{ Control = 'RDP NLA required'; Status = if ($nla -eq 1) { 'PASS' }else { 'FAIL' }; Detail = "UserAuthentication=$nla" }
  $cis += [PSCustomObject]@{ Control = 'RDP disabled when unnecessary'; Status = if ($denyRdp -eq 1) { 'PASS' }else { 'WARN' }; Detail = "fDenyTSConnections=$denyRdp (WARN if enabled intentionally)" }
}
catch {
  $cis += [PSCustomObject]@{ Control = 'RDP checks'; Status = 'UNKNOWN'; Detail = $_.Exception.Message }
}

try {
  $enableLUA = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System').EnableLUA
  $cis += [PSCustomObject]@{ Control = 'UAC enabled'; Status = if ($enableLUA -eq 1) { 'PASS' }else { 'FAIL' }; Detail = "EnableLUA=$enableLUA" }
}
catch { $cis += [PSCustomObject]@{ Control = 'UAC enabled'; Status = 'UNKNOWN'; Detail = $_.Exception.Message } }

try {
  $smb1 = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction Stop
  $cis += [PSCustomObject]@{ Control = 'SMBv1 disabled'; Status = if ($smb1.State -eq 'Disabled') { 'PASS' } else { 'FAIL' }; Detail = "State=$($smb1.State)" }
}
catch { $cis += [PSCustomObject]@{ Control = 'SMBv1 disabled'; Status = 'UNKNOWN'; Detail = $_.Exception.Message } }

try {
  $guest = Get-LocalUser -Name 'Guest' -ErrorAction Stop
  $cis += [PSCustomObject]@{ Control = 'Guest account disabled'; Status = if ($guest.Enabled) { 'FAIL' }else { 'PASS' }; Detail = "Guest.Enabled=$($guest.Enabled)" }
}
catch { $cis += [PSCustomObject]@{ Control = 'Guest account disabled'; Status = 'PASS/NA'; Detail = 'No Guest user found' } }

$netAcc = try { (net accounts) 2>$null } catch { $null }
if ($netAcc) {
  $cis += [PSCustomObject]@{ Control = 'Password policy (local)'; Status = 'INFO'; Detail = ($netAcc | Out-String) }
}
else {
  $cis += [PSCustomObject]@{ Control = 'Password policy (local)'; Status = 'UNKNOWN'; Detail = 'Unable to read via "net accounts"' }
}

try {
  $audit = (auditpol /get /category:* ) 2>$null
  $cis += [PSCustomObject]@{ Control = 'Audit policy configured'; Status = if (($audit | Select-String "No Auditing").Count -gt 0) { 'WARN' }else { 'PASS' }; Detail = ($audit | Out-String) }
}
catch {
  $cis += [PSCustomObject]@{ Control = 'Audit policy configured'; Status = 'UNKNOWN'; Detail = $_.Exception.Message }
}

try {
  $w32 = Get-Service -Name W32Time -ErrorAction Stop
  $cis += [PSCustomObject]@{ Control = 'Time service running'; Status = if ($w32.Status -eq 'Running') { 'PASS' }else { 'FAIL' }; Detail = "W32Time=$($w32.Status)" }
}
catch {
  $cis += [PSCustomObject]@{ Control = 'Time service running'; Status = 'UNKNOWN'; Detail = $_.Exception.Message }
}

# ============================================================================
# VERIFICACIÓN DE ANTIVIRUS (Windows Defender y productos de terceros)
# ============================================================================

Write-Host "[Antivirus] Verificando estado del antivirus..." -ForegroundColor DarkCyan
Write-HealthLog "Iniciando verificación de antivirus" -Level INFO

$AntivirusStatus = [PSCustomObject]@{
  WindowsDefender           = $null
  ThirdPartyAV              = @()
  DefenderEnabled           = $false
  RealTimeProtectionEnabled = $false
  DefinitionsUpToDate       = $false
  LastDefinitionUpdate      = $null
  DefenderService           = $null
  QuickScanAge              = $null
  FullScanAge               = $null
  ThreatDetections          = @()
}

# Verificar Windows Defender
try {
  $mpStatus = Get-MpComputerStatus -ErrorAction Stop
  $AntivirusStatus.WindowsDefender = $mpStatus
  $AntivirusStatus.DefenderEnabled = $mpStatus.AntivirusEnabled
  $AntivirusStatus.RealTimeProtectionEnabled = $mpStatus.RealTimeProtectionEnabled
  $AntivirusStatus.DefinitionsUpToDate = $mpStatus.AntivirusSignatureAge -le 7
  $AntivirusStatus.LastDefinitionUpdate = $mpStatus.AntivirusSignatureLastUpdated
  $AntivirusStatus.QuickScanAge = if ($mpStatus.QuickScanAge) { $mpStatus.QuickScanAge } else { $null }
  $AntivirusStatus.FullScanAge = if ($mpStatus.FullScanAge) { $mpStatus.FullScanAge } else { $null }
  
  # Verificar servicio Windows Defender
  try {
    $defenderSvc = Get-Service -Name WinDefend -ErrorAction Stop
    $AntivirusStatus.DefenderService = $defenderSvc.Status
  }
  catch {
    $AntivirusStatus.DefenderService = 'NotInstalled'
  }
  
  # Obtener detecciones recientes de amenazas
  try {
    $threats = Get-MpThreatDetection -ErrorAction SilentlyContinue | 
    Where-Object { $_.InitialDetectionTime -gt $since } |
    Select-Object -First 20
    if ($threats) {
      $AntivirusStatus.ThreatDetections = $threats | ForEach-Object {
        [PSCustomObject]@{
          ThreatName    = $_.ThreatName
          Severity      = $_.SeverityID
          DetectionTime = $_.InitialDetectionTime
          Resources     = ($_.Resources -join '; ')
        }
      }
    }
  }
  catch {
    # Silently ignore if threat detection query fails
  }
  
  Write-Host "[Antivirus] Windows Defender detectado - Estado: $($mpStatus.AntivirusEnabled)" -ForegroundColor DarkCyan
  Write-HealthLog "Windows Defender detectado: Enabled=$($mpStatus.AntivirusEnabled), RealTime=$($mpStatus.RealTimeProtectionEnabled), DefAge=$($mpStatus.AntivirusSignatureAge)días" -Level SUCCESS
}
catch {
  Write-Host "[Antivirus] Windows Defender no disponible o no está instalado" -ForegroundColor DarkYellow
  Write-HealthLog "Windows Defender no disponible: $($_.Exception.Message)" -Level WARNING -ErrorRecord $_
  $AntivirusStatus.WindowsDefender = $null
}

# Buscar productos antivirus de terceros (WMI SecurityCenter2)
try {
  # Intentar SecurityCenter2 (Windows Vista+)
  $avProducts = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
  
  if ($avProducts) {
    foreach ($av in $avProducts) {
      # Decodificar el estado del producto (productState es un valor hexadecimal)
      $hexState = [Convert]::ToString($av.productState, 16).PadLeft(6, '0')
      $enabled = $hexState.Substring(2, 2) -eq '10'
      $updated = $hexState.Substring(4, 2) -eq '00'
      
      $AntivirusStatus.ThirdPartyAV += [PSCustomObject]@{
        Name                     = $av.displayName
        InstanceGuid             = $av.instanceGuid
        PathToSignedProductExe   = $av.pathToSignedProductExe
        PathToSignedReportingExe = $av.pathToSignedReportingExe
        ProductState             = $av.productState
        Enabled                  = $enabled
        DefinitionsUpToDate      = $updated
      }
    }
    Write-Host "[Antivirus] Encontrados $($AntivirusStatus.ThirdPartyAV.Count) productos antivirus de terceros" -ForegroundColor DarkCyan
    Write-HealthLog "Detectados $($AntivirusStatus.ThirdPartyAV.Count) productos antivirus de terceros: $($AntivirusStatus.ThirdPartyAV.Name -join ', ')" -Level INFO
  }
}
catch {
  Write-Host "[Antivirus] No se pudieron consultar productos antivirus de terceros: $($_.Exception.Message)" -ForegroundColor DarkYellow
  Write-HealthLog "Error consultando antivirus de terceros: $($_.Exception.Message)" -Level WARNING
}

# Agregar verificaciones CIS para Antivirus
if ($AntivirusStatus.WindowsDefender) {
  $mp = $AntivirusStatus.WindowsDefender
  
  # Windows Defender habilitado
  $cis += [PSCustomObject]@{ 
    Control = 'Windows Defender - Antivirus enabled'
    Status  = if ($mp.AntivirusEnabled) { 'PASS' }else { 'FAIL' }
    Detail  = "AntivirusEnabled=$($mp.AntivirusEnabled), Service=$($AntivirusStatus.DefenderService)"
  }
  
  # Protección en tiempo real habilitada
  $cis += [PSCustomObject]@{ 
    Control = 'Windows Defender - Real-time protection enabled'
    Status  = if ($mp.RealTimeProtectionEnabled) { 'PASS' }else { 'FAIL' }
    Detail  = "RealTimeProtectionEnabled=$($mp.RealTimeProtectionEnabled)"
  }
  
  # Definiciones actualizadas (menos de 7 días)
  $sigAge = if ($mp.AntivirusSignatureAge) { $mp.AntivirusSignatureAge } else { 999 }
  $cis += [PSCustomObject]@{ 
    Control = 'Windows Defender - Definitions up to date (<7 days)'
    Status  = if ($sigAge -le 7) { 'PASS' }else { 'FAIL' }
    Detail  = "SignatureAge=$sigAge days, LastUpdated=$($mp.AntivirusSignatureLastUpdated)"
  }
  
  # Protección entregada en la nube habilitada
  $cis += [PSCustomObject]@{ 
    Control = 'Windows Defender - Cloud protection enabled'
    Status  = if ($mp.CloudProtectionEnabled) { 'PASS' }else { 'WARN' }
    Detail  = "CloudProtectionEnabled=$($mp.CloudProtectionEnabled)"
  }
  
  # Envío automático de muestras habilitado
  $cis += [PSCustomObject]@{ 
    Control = 'Windows Defender - Automatic sample submission'
    Status  = if ($mp.SubmitSamplesConsent -ne 2) { 'PASS' }else { 'WARN' }
    Detail  = "SubmitSamplesConsent=$($mp.SubmitSamplesConsent) (0=Always, 1=Safe, 2=Never)"
  }
  
  # Protección contra PUA habilitada
  $cis += [PSCustomObject]@{ 
    Control = 'Windows Defender - PUA protection enabled'
    Status  = if ($mp.PUAProtection -eq 1) { 'PASS' }else { 'WARN' }
    Detail  = "PUAProtection=$($mp.PUAProtection) (0=Disabled, 1=Enabled)"
  }
  
  # Análisis reciente (Quick scan en últimos 7 días)
  $quickScanAge = if ($mp.QuickScanAge) { $mp.QuickScanAge } else { 999 }
  $cis += [PSCustomObject]@{ 
    Control = 'Windows Defender - Recent scan (<7 days)'
    Status  = if ($quickScanAge -le 7) { 'PASS' }else { 'WARN' }
    Detail  = "QuickScanAge=$quickScanAge days, LastQuickScan=$($mp.QuickScanEndTime)"
  }
  
}
elseif ($AntivirusStatus.ThirdPartyAV.Count -gt 0) {
  # Si hay antivirus de terceros pero no Defender
  $enabledAV = @($AntivirusStatus.ThirdPartyAV | Where-Object { $_.Enabled })
  $cis += [PSCustomObject]@{ 
    Control = 'Third-party antivirus installed and enabled'
    Status  = if ($enabledAV.Count -gt 0) { 'PASS' }else { 'FAIL' }
    Detail  = "Installed: $($AntivirusStatus.ThirdPartyAV.Name -join ', ') | Enabled: $($enabledAV.Name -join ', ')"
  }
}
else {
  # No hay antivirus detectado
  $cis += [PSCustomObject]@{ 
    Control = 'Antivirus protection present'
    Status  = 'FAIL'
    Detail  = 'No antivirus software detected (Windows Defender or third-party)'
  }
  Write-HealthLog "CRÍTICO: No se detectó ningún software antivirus activo" -Level ERROR
}

# Registrar amenazas detectadas
if ($AntivirusStatus.ThreatDetections.Count -gt 0) {
  Write-HealthLog "ALERTA: Se detectaron $($AntivirusStatus.ThreatDetections.Count) amenazas en los últimos $DIASAtras días" -Level WARNING
  foreach ($threat in $AntivirusStatus.ThreatDetections) {
    Write-HealthLog "  - Amenaza: $($threat.ThreatName), Severidad: $($threat.Severity), Detectada: $($threat.DetectionTime)" -Level WARNING
  }
}

# Exportar datos de antivirus
if ($ExportJson) {
  $AntivirusStatus | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'antivirus_status.json') -Encoding UTF8
  Write-Host "[Antivirus] Estado del antivirus exportado" -ForegroundColor Green
  Write-HealthLog "Verificación de antivirus completada - Datos exportados" -Level SUCCESS
}

# Cálculo de contadores de eventos para el resumen
$errorCount = 0; $warnCount = 0
foreach ($k in $EventsRaw.Keys) {
  $errorCount += (@($EventsRaw[$k] | Where-Object { $_.LevelDisplayName -eq 'Error' }).Count)
  $warnCount += (@($EventsRaw[$k] | Where-Object { $_.LevelDisplayName -eq 'Warning' }).Count)
}

$cisFail = @($cis | Where-Object { $_.Status -eq 'FAIL' })

$ResumenHealthCheck = [PSCustomObject]@{
  Timestamp            = Get-Date
  Hostname             = $ComputerInfo.Hostname
  CpuAvgPct            = $CpuStatus.AvgLoadPct_6s
  MemUsedPct           = if ($MemoryStatus) { $MemoryStatus.UsedPct } else { $null }
  ErrorEvents          = $errorCount
  WarningEvents        = $warnCount
  PendingReboot        = $PendingReboot
  CisFails             = $cisFail.Count
  # Métricas de Antivirus
  AntivirusEnabled     = if ($AntivirusStatus.DefenderEnabled) { $true } elseif ($AntivirusStatus.ThirdPartyAV.Count -gt 0) { (@($AntivirusStatus.ThirdPartyAV | Where-Object { $_.Enabled }).Count -gt 0) } else { $false }
  RealTimeProtection   = $AntivirusStatus.RealTimeProtectionEnabled
  DefinitionsAge       = if ($AntivirusStatus.WindowsDefender) { $AntivirusStatus.WindowsDefender.AntivirusSignatureAge } else { $null }
  LastDefinitionUpdate = $AntivirusStatus.LastDefinitionUpdate
  QuickScanAge         = $AntivirusStatus.QuickScanAge
  ThreatsDetected      = $AntivirusStatus.ThreatDetections.Count
}
if ($ExportJson) {
  $ComputerInfo  | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'informacion_computadora.json') -Encoding UTF8
  $IPs           | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'direcciones_ip.json') -Encoding UTF8
  $AppliedGPOs   | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'gpos_aplicadas.json') -Encoding UTF8
}

# Exportar GPOs combinadas a JSON
$gposCombined = @()
if ($AppliedGPOs.ComputerGPOs) { $gposCombined += $AppliedGPOs.ComputerGPOs }
if ($AppliedGPOs.UserGPOs) { $gposCombined += $AppliedGPOs.UserGPOs }
if ($ExportJson) {
  if ($gposCombined.Count -gt 0) {
    $gposCombined | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'gpos_aplicadas_combinadas.json') -Encoding UTF8
    Write-HealthLog "GPOs exportadas a JSON: $($gposCombined.Count) registros" -Level INFO
  }
  else {
    Write-HealthLog "No hay GPOs para exportar (equipo no unido a dominio o solo políticas locales)" -Level INFO
  }
}

# Exportar configuraciones detalladas de GPOs
if ($ExportJson) {
  if ($AppliedGPOs.SecuritySettings.Count -gt 0) {
    $AppliedGPOs.SecuritySettings | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'gpo_configuraciones_seguridad.json') -Encoding UTF8
    Write-HealthLog "Configuraciones de seguridad exportadas: $($AppliedGPOs.SecuritySettings.Count) registros" -Level INFO
  }

  if ($AppliedGPOs.RegistrySettings.Count -gt 0) {
    $AppliedGPOs.RegistrySettings | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'gpo_configuraciones_registro.json') -Encoding UTF8
    Write-HealthLog "Configuraciones de registro exportadas: $($AppliedGPOs.RegistrySettings.Count) registros" -Level INFO
  }

  if ($AppliedGPOs.Scripts.Count -gt 0) {
    $AppliedGPOs.Scripts | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'gpo_scripts.json') -Encoding UTF8
    Write-HealthLog "Scripts de GPO exportados: $($AppliedGPOs.Scripts.Count) registros" -Level INFO
  }

  if ($AppliedGPOs.SoftwareRestrictions.Count -gt 0) {
    $AppliedGPOs.SoftwareRestrictions | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'gpo_restricciones_software.json') -Encoding UTF8
    Write-HealthLog "Restricciones de software exportadas: $($AppliedGPOs.SoftwareRestrictions.Count) registros" -Level INFO
  }
}
if ($ExportJson) {
  $softJsonPathLocal = Join-Path $RutaOrigen 'software_instalado.json'
  $useExistingLocalJson = (-not $RutaSoftInstalado) -and (Test-Path -LiteralPath $softJsonPathLocal)
  $shouldExportSoftware = (-not $RutaSoftInstalado) -and (-not $useExistingLocalJson)
  if ($shouldExportSoftware) {
    $Software | Format-JsonOutput -Depth 5 | Out-File -FilePath $softJsonPathLocal -Encoding UTF8
  }
  $CpuStatus     | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'estado_cpu.json') -Encoding UTF8
  $MemoryStatus  | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'estado_memoria.json') -Encoding UTF8
  $LogicalDisks  | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'discos_logicos.json') -Encoding UTF8
  $PhysicalDisks | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'discos_fisicos.json') -Encoding UTF8
  $Adapters      | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'adaptadores_red.json') -Encoding UTF8
  $NetStats      | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'estadisticas_red.json') -Encoding UTF8
  $HotFixes      | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'parches_recientes.json') -Encoding UTF8
  $PendingUpdates | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'actualizaciones_pendientes.json') -Encoding UTF8
  $EventsSummary | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'resumen_eventos.json') -Encoding UTF8
  foreach ($k in $EventsRaw.Keys) {
    $EventsRaw[$k] | Select-Object TimeCreated, Id, ProviderName, LevelDisplayName, Message |
    Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen ("events_raw_{0}.json" -f $k.ToLower())) -Encoding UTF8
  }
  $cis | Format-JsonOutput -Depth 5 | Out-File -FilePath (Join-Path $RutaOrigen 'cis_basic_checks.json') -Encoding UTF8
  ($ResumenHealthCheck | Format-JsonOutput -Depth 5) | Out-File -FilePath (Join-Path $RutaOrigen 'resumen_salud.json') -Encoding UTF8

  Write-Host "`n[Exportación] Todos los datos exportados a JSON exitosamente" -ForegroundColor Green
  Write-HealthLog "Exportación de datos completada - Todos los archivos JSON generados" -Level SUCCESS
}

# Generar reporte HTML
try {
  $timestampHtml = Get-Date -Format 'yyyyMMdd_HHmmss'
  $htmlPath = Join-Path $RutaOrigen ("HealthCheck_{0}.html" -f $timestampHtml)
  $htmlContent = New-HealthCheckHtmlReport -ComputerInfo $ComputerInfo -ResumenHealthCheck $ResumenHealthCheck -CpuStatus $CpuStatus -MemoryStatus $MemoryStatus -LogicalDisks $LogicalDisks -PhysicalDisks $PhysicalDisks -PendingUpdates $PendingUpdates -EventsSummary $EventsSummary -EventsRaw $EventsRaw -EventLogErrors $EventLogErrors -Cis $cis -AntivirusStatus $AntivirusStatus -AppliedGPOs $AppliedGPOs -GpResultHtmlLink $GpResultHtmlLink -PendingReboot $PendingReboot -HotFixes $HotFixes -Software $Software -Certificates $Certificates -IncludeJsonLinks:$ExportJson -HeaderLogoBase64 $HeaderLogoBase64 -HeaderLogoHeight $HeaderLogoHeight
  Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
  Write-Host "[HTML] Reporte generado: $htmlPath" -ForegroundColor Green
  Write-HealthLog "Reporte HTML generado: $htmlPath" -Level SUCCESS
}
catch {
  Write-Warning "[HTML] No se pudo generar el reporte HTML: $($_.Exception.Message)"
  Write-HealthLog "Error generando reporte HTML: $($_.Exception.Message)" -Level WARNING -ErrorRecord $_
}

# Calcular tiempo de ejecución y escribir resumen del log
$executionTime = (Get-Date) - $Script:LogSessionStart
Write-HealthLog "Recopilación de datos completada" -Level SUCCESS
Escribir-ResumenHealthCheck -ResumenHealthCheck $ResumenHealthCheck -ExecutionTime $executionTime

Write-Host "`n=== EJECUCIÓN COMPLETADA ===" -ForegroundColor Green
Write-Host "Tiempo de ejecución: $([math]::Round($executionTime.TotalSeconds,2)) segundos" -ForegroundColor Cyan
if ($Script:LogFile) {
  Write-Host "Log guardado en: $Script:LogFile" -ForegroundColor Cyan
}
