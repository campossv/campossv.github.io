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

$HeaderLogoBase64 = "iVBORw0KGgoAAAANSUhEUgAAAJwAAABJCAYAAADIS0/RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAHzcAAB83AeXxie8AAAgCaVRYdFhNTDpjb20uYWRvYmUueG1wAAAAAAA8P3hwYWNrZXQgYmVnaW49Iu+7vyIgaWQ9Ilc1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCI/Pg0KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgOS4xLWMwMDIgNzkuYTFjZDEyZiwgMjAyNC8xMS8xMS0xOTowODo0NiAgICAgICAgIj4NCgk8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPg0KCQk8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczpBdHRyaWI9Imh0dHA6Ly9ucy5hdHRyaWJ1dGlvbi5jb20vYWRzLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczpwaG90b3Nob3A9Imh0dHA6Ly9ucy5hZG9iZS5jb20vcGhvdG9zaG9wLzEuMC8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiBkYzpmb3JtYXQ9ImltYWdlL3BuZyIgeG1wOkNyZWF0b3JUb29sPSJDYW52YSBkb2M9REFHbVNQOVF5VDQgdXNlcj1VQUZzbDYwbHp0ZyBicmFuZD1DQU5WQSBQUk8gMyB0ZW1wbGF0ZT1Db2xsYWdlIGRlIEZlbGljaXRhY2lvbiBDdW1wbGVhw7FvcyBGb3RvcyBNb2Rlcm5vIFBhc3RlbCIgeG1wOkNyZWF0ZURhdGU9IjIwMjUtMDUtMTlUMTE6NTg6MTctMDY6MDAiIHhtcDpNb2RpZnlEYXRlPSIyMDI1LTA1LTE5VDEyOjAyOjQyLTA2OjAwIiB4bXA6TWV0YWRhdGFEYXRlPSIyMDI1LTA1LTE5VDEyOjAyOjQyLTA2OjAwIiBwaG90b3Nob3A6Q29sb3JNb2RlPSIzIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOmQ0YjlhMjRhLWZiM2MtYTg0NS04MTUzLWNlYjBjZWI1ZTNiMiIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDpkNGI5YTI0YS1mYjNjLWE4NDUtODE1My1jZWIwY2ViNWUzYjIiIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDpkNGI5YTI0YS1mYjNjLWE4NDUtODE1My1jZWIwY2ViNWUzYjIiPg0KCQkJPEF0dHJpYjpBZHM+DQoJCQkJPHJkZjpTZXE+DQoJCQkJCTxyZGY6bGkgQXR0cmliOkNyZWF0ZWQ9IjIwMjUtMDUtMTkiIEF0dHJpYjpFeHRJZD0iOTIyNGJiODItZTg3Zi00N2Q3LTg2N2MtYjhkYzYzOTM4NzIzIiBBdHRyaWI6RmJJZD0iNTI1MjY1OTE0MTc5NTgwIiBBdHRyaWI6VG91Y2hUeXBlPSIyIi8+DQoJCQkJPC9yZGY6U2VxPg0KCQkJPC9BdHRyaWI6QWRzPg0KCQkJPGRjOnRpdGxlPg0KCQkJCTxyZGY6QWx0Pg0KCQkJCQk8cmRmOmxpIHhtbDpsYW5nPSJ4LWRlZmF1bHQiPsKhQmllbnZlbmlkb3MgYWwgRXF1aXBvIERldmVsISAtIEx1bmVzIE1vdGl2YWNpb25hbDwvcmRmOmxpPg0KCQkJCTwvcmRmOkFsdD4NCgkJCTwvZGM6dGl0bGU+DQoJCQk8ZGM6Y3JlYXRvcj4NCgkJCQk8cmRmOlNlcT4NCgkJCQkJPHJkZjpsaT5XZWJtYXN0ZXIgRGV2ZWw8L3JkZjpsaT4NCgkJCQk8L3JkZjpTZXE+DQoJCQk8L2RjOmNyZWF0b3I+DQoJCQk8eG1wTU06SGlzdG9yeT4NCgkJCQk8cmRmOlNlcT4NCgkJCQkJPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOmQ0YjlhMjRhLWZiM2MtYTg0NS04MTUzLWNlYjBjZWI1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"

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
