<#
.SYNOPSIS
    Script de monitoreo completo de salud del sistema
    
.DESCRIPTION
    Recopila métricas de rendimiento, logs de eventos, información de seguridad
    y genera informes detallados en múltiples formatos.
    
.PARAMETER SalidaArchivo
    Ruta donde se guardarán los informes generados
    
.PARAMETER DiasLogs
    Número de días hacia atrás para analizar logs
    
.EXAMPLE
    .\monitor_system.ps1 -SalidaArchivo "C:\Informes" -DiasLogs 7
    
.NOTES
    Versión: 3.0
    Requiere permisos de administrador
#>

[CmdletBinding()]
param(
    [string]$SalidaArchivo = "C:\InformesSalud",
    [int]$DiasLogs = 15,
    [ValidateSet("HTML", "JSON", "CSV", "EXCEL")]
    [string]$FormatoExportar = "HTML",
    [switch]$ParchesFaltantes = $false,
    [switch]$RevisarServicioTerceros = $false,
    [switch]$AnalisisSeguridad = $true,
    [switch]$VerificarCumplimiento = $true
)

$LPNG = "iVBORw0KGgoAAAANSUhEUgAAAJwAAABJCAYAAADIS0/RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAHzcAAB83AeXxie8AAAgCaVRYdFhNTDpjb20uYWRvYmUueG1wAAAAAAA8P3hwYWNrZXQgYmVnaW49Iu+7vyIgaWQ9Ilc1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCI/Pg0KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgOS4xLWMwMDIgNzkuYTFjZDEyZiwgMjAyNC8xMS8xMS0xOTowODo0NiAgICAgICAgIj4NCgk8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPg0KCQk8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczpBdHRyaWI9Imh0dHA6Ly9ucy5hdHRyaWJ1dGlvbi5jb20vYWRzLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczpwaG90b3Nob3A9Imh0dHA6Ly9ucy5hZG9iZS5jb20vcGhvdG9zaG9wLzEuMC8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiBkYzpmb3JtYXQ9ImltYWdlL3BuZyIgeG1wOkNyZWF0b3JUb29sPSJDYW52YSBkb2M9REFHbVNQOVF5VDQgdXNlcj1VQUZzbDYwbHp0ZyBicmFuZD1DQU5WQSBQUk8gMyB0ZW1wbGF0ZT1Db2xsYWdlIGRlIEZlbGljaXRhY2lvbiBDdW1wbGVhw7FvcyBGb3RvcyBNb2Rlcm5vIFBhc3RlbCIgeG1wOkNyZWF0ZURhdGU9IjIwMjUtMDUtMTlUMTE6NTg6MTctMDY6MDAiIHhtcDpNb2RpZnlEYXRlPSIyMDI1LTA1LTE5VDEyOjAyOjQyLTA2OjAwIiB4bXA6TWV0YWRhdGFEYXRlPSIyMDI1LTA1LTE5VDEyOjAyOjQyLTA2OjAwIiBwaG90b3Nob3A6Q29sb3JNb2RlPSIzIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOmQ0YjlhMjRhLWZiM2MtYTg0NS04MTUzLWNlYjBjZWI1ZTNiMiIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDpkNGI5YTI0YS1mYjNjLWE4NDUtODE1My1jZWIwY2ViNWUzYjIiIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDpkNGI5YTI0YS1mYjNjLWE4NDUtODE1My1jZWIwY2ViNWUzYjIiPg0KCQkJPEF0dHJpYjpBZHM+DQoJCQkJPHJkZjpTZXE+DQoJCQkJCTxyZGY6bGkgQXR0cmliOkNyZWF0ZWQ9IjIwMjUtMDUtMTkiIEF0dHJpYjpFeHRJZD0iOTIyNGJiODItZTg3Zi00N2Q3LTg2N2MtYjhkYzYzOTM4NzIzIiBBdHRyaWI6RmJJZD0iNTI1MjY1OTE0MTc5NTgwIiBBdHRyaWI6VG91Y2hUeXBlPSIyIi8+DQoJCQkJPC9yZGY6U2VxPg0KCQkJPC9BdHRyaWI6QWRzPg0KCQkJPGRjOnRpdGxlPg0KCQkJCTxyZGY6QWx0Pg0KCQkJCQk8cmRmOmxpIHhtbDpsYW5nPSJ4LWRlZmF1bHQiPsKhQmllbnZlbmlkb3MgYWwgRXF1aXBvIERldmVsISAtIEx1bmVzIE1vdGl2YWNpb25hbDwvcmRmOmxpPg0KCQkJCTwvcmRmOkFsdD4NCgkJCTwvZGM6dGl0bGU+DQoJCQk8ZGM6Y3JlYXRvcj4NCgkJCQk8cmRmOlNlcT4NCgkJCQkJPHJkZjpsaT5XZWJtYXN0ZXIgRGV2ZWw8L3JkZjpsaT4NCgkJCQk8L3JkZjpTZXE+DQoJCQk8L2RjOmNyZWF0b3I+DQoJCQk8eG1wTU06SGlzdG9yeT4NCgkJCQk8cmRmOlNlcT4NCgkJCQkJPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOmQ0YjlhMjRhLWZiM2MtYTg0NS04MTUzLWNlYjBjZWI1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"
$lenbytes = [Convert]::FromBase64String($LPNG)
$lenmemoria = New-Object System.IO.MemoryStream
$lenmemoria.Write($lenbytes, 0, $lenbytes.Length)
$lenmemoria.Position = 0
$imagenl = [System.Drawing.Image]::FromStream($lenmemoria, $true)

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "Este script requiere permisos de administrador. Ejecute PowerShell como administrador."
    exit 1
}

$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"
$FechaInforme = Get-Date -Format "yyyyMMdd_HHmmss"
$NombreServidor = $env:COMPUTERNAME

try {$DireccionIP = (Test-Connection -ComputerName $NombreServidor -Count 1 -ErrorAction Stop).IPv4Address.IPAddressToString
} catch {$DireccionIP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -ne "127.0.0.1" } | Select-Object -First 1).IPAddress}

if (-not (Test-Path -Path $SalidaArchivo)) {New-Item -ItemType Directory -Path $SalidaArchivo -Force | Out-Null}
$ArchivoSalida = "$SalidaArchivo\InformeSalud_$($NombreServidor)_$FechaInforme"

function Get-InformacionSistema {
    try {
        Write-Progress -Activity "Recopilando información del sistema" -PercentComplete 5
        
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop
        $system = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
        
        $domainInfo = if ($system.PartOfDomain) {
            "Dominio: $($system.Domain)"
        } else {
            "Grupo de trabajo: $($system.Workgroup)"
        }
        
        return @{
            NombreServidor = $NombreServidor
            DireccionIP = $DireccionIP
            NombreSO = $os.Caption
            VersionSO = $os.Version
            BuildNumber = $os.BuildNumber
            ServicePack = $os.CSDVersion
            UltimoReinicio = $os.LastBootUpTime
            TiempoActividad = (Get-Date) - $os.LastBootUpTime
            Fabricante = $system.Manufacturer
            Modelo = $system.Model
            NumeroSerie = $bios.SerialNumber
            Procesador = $cpu.Name
            Nucleos = $cpu.NumberOfCores
            ProcesadoresLogicos = $cpu.NumberOfLogicalProcessors
            VelocidadCPU = $cpu.MaxClockSpeed
            MemoriaTotal = [math]::Round($system.TotalPhysicalMemory / 1GB, 2)
            DominioWorkgroup = $domainInfo
            TimeZone = (Get-TimeZone).DisplayName
            Arquitectura = $os.OSArchitecture
        }
    } catch {
        Write-Warning "Error al obtener información del sistema: $_"
        return @{ Error = $_.Exception.Message }
    }
}

function Get-MetricasRendimiento {
    try {
        Write-Progress -Activity "Recopilando métricas de los 4 subsistemas" -PercentComplete 15
        
        Write-Host "   Analizando subsistema CPU..." -ForegroundColor Yellow
        $cpuSamples = @()
        for ($i = 0; $i -lt 3; $i++) {
            $cpuSamples += (Get-Counter '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 1 -ErrorAction SilentlyContinue).CounterSamples.CookedValue
            Start-Sleep -Milliseconds 500
        }
        $usoCPU = [math]::Round(($cpuSamples | Measure-Object -Average).Average, 2)
        
        $colaProcesor = (Get-Counter '\System\Processor Queue Length' -ErrorAction SilentlyContinue).CounterSamples.CookedValue
        $interrupciones = (Get-Counter '\Processor(_Total)\Interrupts/sec' -ErrorAction SilentlyContinue).CounterSamples.CookedValue
        
        Write-Host "   Analizando subsistema Memoria..." -ForegroundColor Yellow
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $memTotal = [math]::Round($os.TotalVisibleMemorySize / 1KB, 2)
        $memLibre = [math]::Round($os.FreePhysicalMemory / 1KB, 2)
        $memUsada = $memTotal - $memLibre
        $porcentajeMemoria = [math]::Round(($memUsada / $memTotal) * 100, 2)
        
        $memVirtualTotal = [math]::Round($os.TotalVirtualMemorySize / 1KB, 2)
        $memVirtualLibre = [math]::Round($os.FreeVirtualMemory / 1KB, 2)
        $paginasPorSeg = (Get-Counter '\Memory\Pages/sec' -ErrorAction SilentlyContinue).CounterSamples.CookedValue
        $cacheBytes = (Get-Counter '\Memory\Cache Bytes' -ErrorAction SilentlyContinue).CounterSamples.CookedValue
        
        Write-Host "   Analizando subsistema Disco..." -ForegroundColor Yellow
        $discos = Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
        $metricasDisco = @()
        
        foreach ($disco in $discos) {
            $espacioLibre = [math]::Round($disco.FreeSpace / 1GB, 2)
            $espacioTotal = [math]::Round($disco.Size / 1GB, 2)
            $espacioUsado = $espacioTotal - $espacioLibre
            $usoDisco = if ($espacioTotal -gt 0) { [math]::Round(($espacioUsado / $espacioTotal) * 100, 2) } else { 0 }
            
            $discoFisico = $disco.DeviceID.Replace(":", "")
            $tiempoLectura = (Get-Counter "\LogicalDisk($discoFisico)\Avg. Disk sec/Read" -ErrorAction SilentlyContinue).CounterSamples.CookedValue
            $tiempoEscritura = (Get-Counter "\LogicalDisk($discoFisico)\Avg. Disk sec/Write" -ErrorAction SilentlyContinue).CounterSamples.CookedValue
            $colaDisco = (Get-Counter "\LogicalDisk($discoFisico)\Current Disk Queue Length" -ErrorAction SilentlyContinue).CounterSamples.CookedValue
            
            $metricasDisco += @{
                Unidad = $disco.DeviceID
                TotalGB = $espacioTotal
                LibreGB = $espacioLibre
                UsadoGB = $espacioUsado
                PorcentajeUsado = $usoDisco
                SistemaArchivos = $disco.FileSystem
                Etiqueta = $disco.VolumeName
                TiempoLectura = [math]::Round($tiempoLectura * 1000, 2)
                TiempoEscritura = [math]::Round($tiempoEscritura * 1000, 2)
                ColaDisco = [math]::Round($colaDisco, 2)
                Estado = switch ($usoDisco) {
                    { $_ -gt 90 } { "Crítico" }
                    { $_ -gt 80 } { "Advertencia" }
                    default { "Normal" }
                }
            }
        }
        
        Write-Host "   Analizando subsistema Red..." -ForegroundColor Yellow
        $interfacesRed = Get-CimInstance -ClassName Win32_PerfRawData_Tcpip_NetworkInterface | 
                        Where-Object { $_.Name -notlike "*Loopback*" -and $_.Name -notlike "*Teredo*" -and $_.Name -ne "_Total" }
        
        $metricasRed = @()
        foreach ($interfaz in $interfacesRed) {
            $bytesEnviados = [math]::Round($interfaz.BytesSentPerSec / 1KB, 2)
            $bytesRecibidos = [math]::Round($interfaz.BytesReceivedPerSec / 1KB, 2)
            
            $metricasRed += @{
                Interfaz = $interfaz.Name
                BytesEnviados = $bytesEnviados
                BytesRecibidos = $bytesRecibidos
                PaquetesEnviados = $interfaz.PacketsSentPerSec
                PaquetesRecibidos = $interfaz.PacketsReceivedPerSec
                ErroresEnvio = $interfaz.PacketsOutboundErrors
                ErroresRecepcion = $interfaz.PacketsReceivedErrors
                Ancho = "N/A"  
            }
        }
        
        return @{
            CPU = @{
                UsoCPU = $usoCPU
                ColaProcesor = [math]::Round($colaProcesor, 2)
                Interrupciones = [math]::Round($interrupciones, 0)
                Estado = if ($usoCPU -gt 80) { "Alto" } elseif ($usoCPU -gt 60) { "Medio" } else { "Normal" }
            }
            Memoria = @{
                PorcentajeUsado = $porcentajeMemoria
                TotalMB = [math]::Round($memTotal / 1024, 2)
                UsadaMB = [math]::Round($memUsada / 1024, 2)
                LibreMB = [math]::Round($memLibre / 1024, 2)
                VirtualTotalMB = [math]::Round($memVirtualTotal / 1024, 2)
                VirtualLibreMB = [math]::Round($memVirtualLibre / 1024, 2)
                PaginasPorSeg = [math]::Round($paginasPorSeg, 2)
                CacheMB = [math]::Round($cacheBytes / 1MB, 2)
                Estado = if ($porcentajeMemoria -gt 85) { "Alto" } elseif ($porcentajeMemoria -gt 70) { "Medio" } else { "Normal" }
            }
            Discos = $metricasDisco
            Red = $metricasRed
        }
    } catch {
        Write-Warning "Error al obtener métricas de rendimiento: $_"
        return @{ Error = $_.Exception.Message }
    }
}

function Get-LogsEventos {
    param([int]$Dias)
    
    try {
        Write-Progress -Activity "Recopilando logs de los 3 tipos principales" -PercentComplete 25
        
        $fechaInicio = (Get-Date).AddDays(-$Dias)
        Write-Host "   Período de logs: desde $($fechaInicio) hasta $(Get-Date)" -ForegroundColor Yellow
        
        Write-Host "   Recopilando logs del Sistema..." -ForegroundColor Yellow
        $logsSistema = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Level = 1,2,3
            StartTime = $fechaInicio
        } -MaxEvents 200 -ErrorAction SilentlyContinue | 
        Select-Object TimeCreated, LevelDisplayName, ProviderName, Id, 
                     @{Name="Message";Expression={$_.Message.Substring(0,[Math]::Min(300,$_.Message.Length))}}
        
        Write-Host "   Recopilando logs de Aplicación..." -ForegroundColor Yellow
        $logsAplicacion = Get-WinEvent -FilterHashtable @{
            LogName = 'Application'
            Level = 1,2,3
            StartTime = $fechaInicio
        } -MaxEvents 200 -ErrorAction SilentlyContinue |
        Select-Object TimeCreated, LevelDisplayName, ProviderName, Id,
                     @{Name="Message";Expression={$_.Message.Substring(0,[Math]::Min(300,$_.Message.Length))}}
        
        Write-Host "   Recopilando logs de Seguridad..." -ForegroundColor Yellow
        try {
            $eventosSeguridad = Get-WinEvent -FilterHashtable @{
                LogName = 'Security'
                ID = 4624, 4625, 4634, 4647, 4648, 4740, 4767, 4771, 4776, 4616, 4720, 4722, 4724, 4738
                StartTime = $fechaInicio
            } -MaxEvents 150 -ErrorAction Stop
            
            Write-Host "   Se encontraron $($eventosSeguridad.Count) eventos de seguridad" -ForegroundColor Yellow
            
            
            $logsSeguridad = $eventosSeguridad | ForEach-Object {
                $evento = $_
                $cuenta = "N/A"
                $ip = "N/A"
                
                try { $cuenta = if ($evento.Properties.Count -gt 5) { $evento.Properties[5].Value } elseif ($evento.Properties.Count -gt 1) { $evento.Properties[1].Value } 
                    if ([string]::IsNullOrEmpty($cuenta)) { 
                        for ($i = 0; $i -lt [Math]::Min(10, $evento.Properties.Count); $i++) { 
                            if (![string]::IsNullOrEmpty($evento.Properties[$i].Value)) { $cuenta = $evento.Properties[$i].Value; break } 
                        } 
                    } 
                } catch { $cuenta = "No disponible" }
                
                $ip = if ($evento.Properties.Count -gt 19) { $evento.Properties[19].Value } elseif ($evento.Properties.Count -gt 9) { $evento.Properties[9].Value }
                
                try { $rawMessage = $evento.Message; if ($rawMessage.Length -gt 150) { $rawMessage = $rawMessage.Substring(0, 150) + "..." } 
                } catch { $rawMessage = "[No disponible]" }
                
                $tipoEvento = switch ($evento.Id) {
                    4625 { "Login Fallido" }
                    4648 { "Login Explícito" }
                    4771 { "Kerberos Fallido" }
                    4776 { "Validación" }
                    4740 { "Cuenta Bloqueada" }
                    4767 { "Desbloqueada" }
                    4624 { "Login" }
                    4634 { "Logout" }
                    4647 { "Logout Usuario" }
                    4616 { "Cambio hora" }
                    4720 { "Cuenta creada" }
                    4722 { "Cuenta habilitada" }
                    4724 { "Cambio pwd" }
                    4738 { "Cambio cuenta" }
                    default { "ID:$($evento.Id)" }
                }
                
[PSCustomObject]@{TimeCreated=$evento.TimeCreated;Id=$evento.Id;Account=$cuenta;SourceIP=$ip;EventType=$tipoEvento;RawMessage=$rawMessage}
            }
        } catch {
            Write-Host "   Error al obtener logs de seguridad: $_" -ForegroundColor Red
            $logsSeguridad = @()
        }
        
        Write-Host "   Total eventos encontrados - Sistema: $($logsSistema.Count), Aplicación: $($logsAplicacion.Count), Seguridad: $($logsSeguridad.Count)" -ForegroundColor Yellow
        
        return @{
            LogsSistema = $logsSistema
            LogsAplicacion = $logsAplicacion
            LogsSeguridad = $logsSeguridad
        }
    } catch {
        Write-Warning "Error al obtener logs de eventos: $_"
        return @{ Error = $_.Exception.Message }
    }
}

function Get-AnalisisConfiabilidad {
    try {
        Write-Progress -Activity "Analizando confiabilidad del sistema" -PercentComplete 35
        Write-Host "   Analizando registros de confiabilidad..." -ForegroundColor Yellow
        
        $datos = @{}
        
        try {
            $reliabilityRecords = Get-CimInstance -ClassName Win32_ReliabilityRecords -ErrorAction SilentlyContinue | 
                                 Where-Object { $_.TimeGenerated -gt (Get-Date).AddDays(-30) } |
                                 Sort-Object TimeGenerated -Descending |
                                 Select-Object -First 100
            
            if ($reliabilityRecords) {
                Write-Host "   Se encontraron $($reliabilityRecords.Count) registros de confiabilidad" -ForegroundColor Yellow
                
                $eventosConfiabilidad = $reliabilityRecords | ForEach-Object {
                    $tipoEvento = switch ($_.EventIdentifier) {
                        1001 { "Inicio de aplicación" }
                        1002 { "Fallo de aplicación" }
                        1003 { "Cuelgue de aplicación" }
                        1004 { "Instalación exitosa" }
                        1005 { "Fallo de instalación" }
                        1006 { "Inicio del sistema" }
                        1007 { "Apagado del sistema" }
                        1008 { "Fallo del sistema" }
                        default { "Evento ID: $($_.EventIdentifier)" }
                    }
                    
                    [PSCustomObject]@{
                        Fecha = $_.TimeGenerated
                        TipoEvento = $tipoEvento
                        Fuente = $_.SourceName
                        Descripcion = if ($_.Message) { $_.Message.Substring(0, [Math]::Min(200, $_.Message.Length)) } else { "N/A" }
                        EventID = $_.EventIdentifier
                        Criticidad = switch ($_.EventIdentifier) {
                            { $_ -in @(1002, 1003, 1005, 1008) } { "Alto" }
                            { $_ -in @(1001, 1004, 1006, 1007) } { "Normal" }
                            default { "Medio" }
                        }
                    }
                }
                
                $datos.EventosConfiabilidad = $eventosConfiabilidad
                
                $fallosAplicacion = ($eventosConfiabilidad | Where-Object { $_.EventID -in @(1002, 1003) }).Count
                $fallosSistema = ($eventosConfiabilidad | Where-Object { $_.EventID -eq 1008 }).Count
                $reinicios = ($eventosConfiabilidad | Where-Object { $_.EventID -in @(1006, 1007) }).Count
                
                $datos.EstadisticasEstabilidad = @{
                    FallosAplicacion = $fallosAplicacion
                    FallosSistema = $fallosSistema
                    ReiniciosDetectados = $reinicios
                    TotalEventos = $eventosConfiabilidad.Count
                    PeriodoAnalisis = "Últimos 30 días"
                    IndiceEstabilidad = switch ($true) {
                        { $fallosSistema -gt 5 -or $fallosAplicacion -gt 20 } { "Baja" }
                        { $fallosSistema -gt 2 -or $fallosAplicacion -gt 10 } { "Media" }
                        default { "Alta" }
                    }
                }
                
                $tendenciasSemana = $eventosConfiabilidad | 
                    Group-Object { (Get-Date $_.Fecha).ToString("yyyy-MM-dd") } |
                    Sort-Object Name -Descending |
                    Select-Object -First 7 |
                    ForEach-Object {
                        $fallosDia = ($_.Group | Where-Object { $_.Criticidad -eq "Alto" }).Count
                        [PSCustomObject]@{
                            Fecha = $_.Name
                            TotalEventos = $_.Count
                            EventosCriticos = $fallosDia
                            Estabilidad = if ($fallosDia -eq 0) { "Estable" } elseif ($fallosDia -lt 3) { "Moderada" } else { "Inestable" }
                        }
                    }
                
                $datos.TendenciasSemanales = $tendenciasSemana
                
            } else {
                Write-Host "   No se encontraron registros de confiabilidad o no están disponibles" -ForegroundColor Yellow
                $datos.EventosConfiabilidad = @()
                $datos.EstadisticasEstabilidad = @{ Error = "No hay datos de confiabilidad disponibles" }
                $datos.TendenciasSemanales = @()
            }
            
        } catch {
            Write-Host "   Error al acceder a registros de confiabilidad: $_" -ForegroundColor Red
            
            try {
                Write-Host "   Intentando método alternativo con Event Log..." -ForegroundColor Yellow
                $eventosAlternativos = Get-WinEvent -FilterHashtable @{
                    LogName = 'System'
                    Level = 1,2  # Critical, Error
                    StartTime = (Get-Date).AddDays(-7)
                } -MaxEvents 50 -ErrorAction SilentlyContinue |
                Where-Object { $_.Id -in @(1001, 1074, 6005, 6006, 6008, 6009, 6013) } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Fecha = $_.TimeCreated
                        TipoEvento = switch ($_.Id) {
                            1074 { "Apagado del sistema" }
                            6005 { "Inicio del servicio Event Log" }
                            6006 { "Parada del servicio Event Log" }
                            6008 { "Apagado inesperado" }
                            6009 { "Información de versión del procesador" }
                            6013 { "Tiempo de actividad del sistema" }
                            default { "Evento del sistema" }
                        }
                        EventID = $_.Id
                        Fuente = $_.ProviderName
                        Criticidad = if ($_.Id -eq 6008) { "Alto" } else { "Normal" }
                    }
                }
                
                $datos.EventosConfiabilidad = $eventosAlternativos
                $datos.EstadisticasEstabilidad = @{
                    Metodo = "Event Log alternativo"
                    ApagadosInesperados = ($eventosAlternativos | Where-Object { $_.EventID -eq 6008 }).Count
                    TotalEventos = $eventosAlternativos.Count
                    PeriodoAnalisis = "Últimos 7 días"
                }
                
            } catch {
                $datos.EventosConfiabilidad = @()
                $datos.EstadisticasEstabilidad = @{ Error = "No se pudo obtener información de confiabilidad: $_" }
            }
        }
        
        return $datos
        
    } catch {
        Write-Warning "Error en análisis de confiabilidad: $_"
        return @{ Error = $_.Exception.Message }
    }
}

function Get-DiagnosticoHardwareAvanzado {
    try {
        Write-Progress -Activity "Realizando diagnóstico avanzado de hardware" -PercentComplete 45
        Write-Host "   Analizando hardware avanzado..." -ForegroundColor Yellow
        
        $datos = @{}
        
        Write-Host "   Verificando estado SMART de discos..." -ForegroundColor Yellow
        try {
            $discosSMART = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue | ForEach-Object {
                $disco = $_
                $smartData = $null
                
                try {
                    # Intentar obtener datos SMART
                    $smartData = Get-CimInstance -ClassName MSStorageDriver_FailurePredictStatus -Namespace "root\wmi" -ErrorAction SilentlyContinue |
                                Where-Object { $_.InstanceName -like "*$($disco.PNPDeviceID)*" }
                    
                    if (-not $smartData) {
                        # Método alternativo
                        $smartData = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue |
                                    Where-Object { $_.DeviceID -eq $disco.DeviceID } |
                                    ForEach-Object {
                                        [PSCustomObject]@{
                                            PredictFailure = $false
                                            Reason = "Datos SMART no disponibles"
                                        }
                                    }
                    }
                } catch {
                    $smartData = [PSCustomObject]@{
                        PredictFailure = $null
                        Reason = "Error al acceder a SMART: $_"
                    }
                }
                
                [PSCustomObject]@{
                    Modelo = $disco.Model
                    NumeroSerie = $disco.SerialNumber
                    TamanoGB = [math]::Round($disco.Size / 1GB, 2)
                    Interfaz = $disco.InterfaceType
                    EstadoSMART = if ($smartData.PredictFailure -eq $true) { "FALLO PREDICHO" } 
                                 elseif ($smartData.PredictFailure -eq $false) { "Saludable" } 
                                 else { "No disponible" }
                    DetallesSMART = $smartData.Reason
                    Particiones = (Get-CimInstance -ClassName Win32_DiskPartition -ErrorAction SilentlyContinue | 
                                  Where-Object { $_.DiskIndex -eq $disco.Index }).Count
                    Estado = switch ($smartData.PredictFailure) {
                        $true { "Crítico" }
                        $false { "Normal" }
                        default { "Desconocido" }
                    }
                }
            }
            
            $datos.DiscosSMART = $discosSMART
            
        } catch {
            Write-Host "   Error al obtener datos SMART: $_" -ForegroundColor Red
            $datos.DiscosSMART = @{ Error = "No se pudo obtener información SMART: $_" }
        }
        
        Write-Host "   Verificando temperaturas de componentes..." -ForegroundColor Yellow
        try {
            $temperaturas = @()
            
            $tempCPU = Get-CimInstance -ClassName Win32_PerfRawData_Counters_ThermalZoneInformation -ErrorAction SilentlyContinue |
                      Where-Object { $_.Name -like "*CPU*" -or $_.Name -like "*Processor*" } |
                      ForEach-Object {
                          $tempKelvin = $_.Temperature
                          $tempCelsius = if ($tempKelvin -and $tempKelvin -gt 0) { 
                              [math]::Round(($tempKelvin / 10) - 273.15, 1) 
                          } else { $null }
                          
                          [PSCustomObject]@{
                              Componente = "CPU - $($_.Name)"
                              TemperaturaC = $tempCelsius
                              Estado = if ($tempCelsius -gt 80) { "Crítico" } 
                                      elseif ($tempCelsius -gt 70) { "Alto" } 
                                      elseif ($tempCelsius -gt 0) { "Normal" } 
                                      else { "No disponible" }
                          }
                      }
            
            if ($tempCPU) { $temperaturas += $tempCPU }
            
            try {
                $wmiTemp = Get-CimInstance -ClassName MSAcpi_ThermalZoneTemperature -Namespace "root\wmi" -ErrorAction SilentlyContinue |
                          ForEach-Object {
                              $tempKelvin = $_.CurrentTemperature
                              $tempCelsius = if ($tempKelvin -and $tempKelvin -gt 0) { 
                                  [math]::Round(($tempKelvin / 10) - 273.15, 1) 
                              } else { $null }
                              
                              [PSCustomObject]@{
                                  Componente = "Zona Térmica - $($_.InstanceName)"
                                  TemperaturaC = $tempCelsius
                                  Estado = if ($tempCelsius -gt 80) { "Crítico" } 
                                          elseif ($tempCelsius -gt 70) { "Alto" } 
                                          elseif ($tempCelsius -gt 0) { "Normal" } 
                                          else { "No disponible" }
                              }
                          }
                
                if ($wmiTemp) { $temperaturas += $wmiTemp }
                
            } catch {
                Write-Host "   Método WMI para temperaturas no disponible" -ForegroundColor Yellow
            }
            
            if ($temperaturas.Count -eq 0) {
                $temperaturas = @([PSCustomObject]@{
                    Componente = "Sistema"
                    TemperaturaC = $null
                    Estado = "Sensores de temperatura no disponibles"
                })
            }
            
            $datos.Temperaturas = $temperaturas
            
        } catch {
            Write-Host "   Error al obtener temperaturas: $_" -ForegroundColor Red
            $datos.Temperaturas = @{ Error = "No se pudo obtener información de temperatura: $_" }
        }
        
        Write-Host "   Verificando estado de la batería..." -ForegroundColor Yellow
        try {
            $baterias = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
            
            if ($baterias) {
                $estadoBaterias = $baterias | ForEach-Object {
                    $bateria = $_
                    
                    $batteryStatus = Get-CimInstance -ClassName BatteryStatus -Namespace "root\wmi" -ErrorAction SilentlyContinue |
                                    Where-Object { $_.InstanceName -like "*$($bateria.DeviceID)*" }
                    
                    $estadoCarga = switch ($bateria.BatteryStatus) {
                        1 { "Desconocido" }
                        2 { "Cargando" }
                        3 { "Descargando" }
                        4 { "Crítico" }
                        5 { "Bajo" }
                        6 { "Cargando y Alto" }
                        7 { "Cargando y Bajo" }
                        8 { "Cargando y Crítico" }
                        9 { "Indefinido" }
                        10 { "Parcialmente Cargado" }
                        11 { "Completamente Cargado" }
                        default { "Estado $($bateria.BatteryStatus)" }
                    }
                    
                    [PSCustomObject]@{
                        Nombre = $bateria.Name
                        Fabricante = $bateria.Manufacturer
                        EstadoCarga = $estadoCarga
                        PorcentajeCarga = $bateria.EstimatedChargeRemaining
                        TiempoRestante = if ($bateria.EstimatedRunTime -and $bateria.EstimatedRunTime -ne 71582788) { 
                            "$([math]::Round($bateria.EstimatedRunTime / 60, 1)) horas" 
                        } else { "Calculando..." }
                        CapacidadDiseño = if ($bateria.DesignCapacity) { "$($bateria.DesignCapacity) mWh" } else { "N/A" }
                        CapacidadCompleta = if ($bateria.FullChargeCapacity) { "$($bateria.FullChargeCapacity) mWh" } else { "N/A" }
                        SaludBateria = if ($bateria.DesignCapacity -and $bateria.FullChargeCapacity) {
                            $salud = [math]::Round(($bateria.FullChargeCapacity / $bateria.DesignCapacity) * 100, 1)
                            "$salud%"
                        } else { "No disponible" }
                        Estado = switch ($true) {
                            { $bateria.BatteryStatus -in @(4, 8) } { "Crítico" }
                            { $bateria.BatteryStatus -in @(5, 7) } { "Bajo" }
                            { $bateria.EstimatedChargeRemaining -lt 20 } { "Advertencia" }
                            default { "Normal" }
                        }
                    }
                }
                
                $datos.Baterias = $estadoBaterias
                
            } else {
                $datos.Baterias = @([PSCustomObject]@{
                    Estado = "No se detectaron baterías (sistema de escritorio)"
                })
            }
            
        } catch {
            Write-Host "   Error al obtener estado de batería: $_" -ForegroundColor Red
            $datos.Baterias = @{ Error = "No se pudo obtener información de batería: $_" }
        }
        
        Write-Host "   Recopilando información adicional de hardware..." -ForegroundColor Yellow
        try {
            $memoriaFisica = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue | ForEach-Object {
                [PSCustomObject]@{
                    Ubicacion = $_.DeviceLocator
                    Capacidad = "$([math]::Round($_.Capacity / 1GB, 2)) GB"
                    Velocidad = "$($_.Speed) MHz"
                    Fabricante = $_.Manufacturer
                    NumeroSerie = $_.SerialNumber
                    TipoMemoria = switch ($_.MemoryType) {
                        20 { "DDR" }
                        21 { "DDR2" }
                        22 { "DDR2 FB-DIMM" }
                        24 { "DDR3" }
                        26 { "DDR4" }
                        default { "Tipo $($_.MemoryType)" }
                    }
                }
            }
            
            $datos.MemoriaFisica = $memoriaFisica
            
            $ventiladores = Get-CimInstance -ClassName Win32_Fan -ErrorAction SilentlyContinue | ForEach-Object {
                [PSCustomObject]@{
                    Nombre = $_.Name
                    Descripcion = $_.Description
                    Estado = switch ($_.Status) {
                        "OK" { "Normal" }
                        "Error" { "Error" }
                        "Degraded" { "Degradado" }
                        default { $_.Status }
                    }
                    Activo = $_.ActiveCooling
                }
            }
            
            if ($ventiladores.Count -eq 0) {
                $ventiladores = @([PSCustomObject]@{
                    Estado = "No se detectaron ventiladores o información no disponible"
                })
            }
            
            $datos.Ventiladores = $ventiladores
            
        } catch {
            Write-Host "   Error al obtener información adicional de hardware: $_" -ForegroundColor Red
            $datos.MemoriaFisica = @{ Error = "No disponible" }
            $datos.Ventiladores = @{ Error = "No disponible" }
        }
        
        return $datos
        
    } catch {
        Write-Warning "Error en diagnóstico de hardware avanzado: $_"
        return @{ Error = $_.Exception.Message }
    }
}

function Get-AnalisisRolesServidor {
    try {
        Write-Progress -Activity "Analizando roles y características del servidor" -PercentComplete 55
        Write-Host "   Analizando roles de Windows Server..." -ForegroundColor Yellow
        
        $datos = @{}
        
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $esServidor = $os.ProductType -eq 2 -or $os.ProductType -eq 3 -or $os.Caption -like "*Server*"
        
        if (-not $esServidor) {
            Write-Host "   Sistema detectado como cliente Windows, no servidor" -ForegroundColor Yellow
            return @{
                TipoSistema = "Cliente Windows"
                EsServidor = $false
                Mensaje = "Este sistema no es Windows Server. Análisis de roles no aplicable."
            }
        }
        
        Write-Host "   Sistema Windows Server detectado, analizando roles..." -ForegroundColor Yellow
        $datos.EsServidor = $true
        $datos.TipoSistema = "Windows Server"
        
        # 1. ROLES Y CARACTERÍSTICAS USANDO DISM (más compatible)
        try {
            Write-Host "   Obteniendo características mediante DISM..." -ForegroundColor Yellow
            $dismFeatures = dism /online /get-features /format:table | Out-String
            
            # Parsear salida de DISM para características habilitadas
            $caracteristicasHabilitadas = @()
            $lineas = $dismFeatures -split "`n" | Where-Object { $_ -match "Enabled" }
            
            foreach ($linea in $lineas) {
                if ($linea -match "^([^\|]+)\|.*Enabled") {
                    $nombreCaracteristica = $matches[1].Trim()
                    if ($nombreCaracteristica -and $nombreCaracteristica -ne "Feature Name") {
                        $caracteristicasHabilitadas += $nombreCaracteristica
                    }
                }
            }
            
            $datos.CaracteristicasDISM = $caracteristicasHabilitadas
            
        } catch {
            Write-Host "   Error al usar DISM: $_" -ForegroundColor Red
            $datos.CaracteristicasDISM = @("Error al obtener características via DISM")
        }
        
        # 2. SERVICIOS DE ROLES ESPECÍFICOS
        Write-Host "   Analizando servicios de roles específicos..." -ForegroundColor Yellow
        $rolesDetectados = @()
        
        # Active Directory Domain Services
        $addsService = Get-Service -Name "NTDS" -ErrorAction SilentlyContinue
        if ($addsService) {
            $rolesDetectados += [PSCustomObject]@{
                Rol = "Active Directory Domain Services"
                Servicio = "NTDS"
                Estado = $addsService.Status
                TipoInicio = $addsService.StartType
                Descripcion = "Controlador de dominio Active Directory"
                Critico = $true
            }
        }
        
        # DNS Server
        $dnsService = Get-Service -Name "DNS" -ErrorAction SilentlyContinue
        if ($dnsService) {
            $rolesDetectados += [PSCustomObject]@{
                Rol = "DNS Server"
                Servicio = "DNS"
                Estado = $dnsService.Status
                TipoInicio = $dnsService.StartType
                Descripcion = "Servidor DNS"
                Critico = $true
            }
        }
        
        # DHCP Server
        $dhcpService = Get-Service -Name "DHCPServer" -ErrorAction SilentlyContinue
        if ($dhcpService) {
            $rolesDetectados += [PSCustomObject]@{
                Rol = "DHCP Server"
                Servicio = "DHCPServer"
                Estado = $dhcpService.Status
                TipoInicio = $dhcpService.StartType
                Descripcion = "Servidor DHCP"
                Critico = $false
            }
        }
        
        # IIS (Internet Information Services)
        $iisService = Get-Service -Name "W3SVC" -ErrorAction SilentlyContinue
        if ($iisService) {
            $rolesDetectados += [PSCustomObject]@{
                Rol = "Web Server (IIS)"
                Servicio = "W3SVC"
                Estado = $iisService.Status
                TipoInicio = $iisService.StartType
                Descripcion = "Servidor web IIS"
                Critico = $false
            }
            
            # Información adicional de IIS
            try {
                $iisInfo = Get-CimInstance -ClassName Win32_Service -Filter "Name='W3SVC'" -ErrorAction SilentlyContinue
                $sitiosIIS = @()
                
                # Intentar obtener sitios web si IIS está instalado
                if (Get-Command "Get-IISSite" -ErrorAction SilentlyContinue) {
                    $sitiosIIS = Get-IISSite | ForEach-Object {
                        [PSCustomObject]@{
                            Nombre = $_.Name
                            Estado = $_.State
                            Puerto = ($_.Bindings | ForEach-Object { $_.EndPoint.Port }) -join ", "
                            RutaFisica = $_.PhysicalPath
                        }
                    }
                } else {
                    # Método alternativo usando appcmd si está disponible
                    try {
                        $appcmdPath = "$env:SystemRoot\System32\inetsrv\appcmd.exe"
                        if (Test-Path $appcmdPath) {
                            $sitiosOutput = & $appcmdPath list sites
                            $sitiosIIS = $sitiosOutput | ForEach-Object {
                                if ($_ -match 'SITE "([^"]+)" $$id:(\d+),bindings:([^,]+),state:(\w+)$$') {
                                    [PSCustomObject]@{
                                        Nombre = $matches[1]
                                        ID = $matches[2]
                                        Bindings = $matches[3]
                                        Estado = $matches[4]
                                    }
                                }
                            }
                        }
                    } catch {
                        $sitiosIIS = @([PSCustomObject]@{ Info = "No se pudo obtener información de sitios IIS" })
                    }
                }
                
                $datos.SitiosIIS = $sitiosIIS
                
            } catch {
                $datos.SitiosIIS = @{ Error = "Error al obtener información de IIS: $_" }
            }
        }
        
        # File and Storage Services
        $lanmanService = Get-Service -Name "LanmanServer" -ErrorAction SilentlyContinue
        if ($lanmanService -and $lanmanService.Status -eq "Running") {
            $rolesDetectados += [PSCustomObject]@{
                Rol = "File and Storage Services"
                Servicio = "LanmanServer"
                Estado = $lanmanService.Status
                TipoInicio = $lanmanService.StartType
                Descripcion = "Servicios de archivos y almacenamiento"
                Critico = $false
            }
        }
        
        # Print and Document Services
        $spoolerService = Get-Service -Name "Spooler" -ErrorAction SilentlyContinue
        if ($spoolerService -and $spoolerService.Status -eq "Running") {
            # Verificar si hay impresoras compartidas
            $impresorasCompartidas = Get-CimInstance -ClassName Win32_Printer -ErrorAction SilentlyContinue | 
                                    Where-Object { $_.Shared -eq $true }
            
            if ($impresorasCompartidas) {
                $rolesDetectados += [PSCustomObject]@{
                    Rol = "Print and Document Services"
                    Servicio = "Spooler"
                    Estado = $spoolerService.Status
                    TipoInicio = $spoolerService.StartType
                    Descripcion = "Servicios de impresión y documentos"
                    Critico = $false
                }
            }
        }
        
        # Remote Desktop Services
        $termService = Get-Service -Name "TermService" -ErrorAction SilentlyContinue
        if ($termService) {
            $rolesDetectados += [PSCustomObject]@{
                Rol = "Remote Desktop Services"
                Servicio = "TermService"
                Estado = $termService.Status
                TipoInicio = $termService.StartType
                Descripcion = "Servicios de escritorio remoto"
                Critico = $false
            }
        }
        
        # Windows Server Update Services (WSUS)
        $wsusService = Get-Service -Name "WsusService" -ErrorAction SilentlyContinue
        if ($wsusService) {
            $rolesDetectados += [PSCustomObject]@{
                Rol = "Windows Server Update Services"
                Servicio = "WsusService"
                Estado = $wsusService.Status
                TipoInicio = $wsusService.StartType
                Descripcion = "Servidor WSUS"
                Critico = $false
            }
        }
        
        # Hyper-V
        $hypervService = Get-Service -Name "vmms" -ErrorAction SilentlyContinue
        if ($hypervService) {
            $rolesDetectados += [PSCustomObject]@{
                Rol = "Hyper-V"
                Servicio = "vmms"
                Estado = $hypervService.Status
                TipoInicio = $hypervService.StartType
                Descripcion = "Plataforma de virtualización Hyper-V"
                Critico = $false
            }
            
            # Información adicional de Hyper-V
            try {
                if (Get-Command "Get-VM" -ErrorAction SilentlyContinue) {
                    $vms = Get-VM -ErrorAction SilentlyContinue | ForEach-Object {
                        [PSCustomObject]@{
                            Nombre = $_.Name
                            Estado = $_.State
                            CPUs = $_.ProcessorCount
                            MemoriaGB = [math]::Round($_.MemoryAssigned / 1GB, 2)
                            TiempoActividad = if ($_.Uptime) { $_.Uptime.ToString() } else { "N/A" }
                        }
                    }
                    $datos.MaquinasVirtuales = $vms
                } else {
                    $datos.MaquinasVirtuales = @{ Info = "Cmdlets de Hyper-V no disponibles" }
                }
            } catch {
                $datos.MaquinasVirtuales = @{ Error = "Error al obtener VMs: $_" }
            }
        }
        
        $datos.RolesDetectados = $rolesDetectados
        
        # 3. INFORMACIÓN ADICIONAL DEL SERVIDOR
        Write-Host "   Recopilando información adicional del servidor..." -ForegroundColor Yellow
        
        # Información de dominio
        try {
            $dominioInfo = Get-CimInstance -ClassName Win32_ComputerSystem
            $datos.InformacionDominio = @{
                ParteDominio = $dominioInfo.PartOfDomain
                Dominio = $dominioInfo.Domain
                Workgroup = $dominioInfo.Workgroup
                Rol = switch ($dominioInfo.DomainRole) {
                    0 { "Standalone Workstation" }
                    1 { "Member Workstation" }
                    2 { "Standalone Server" }
                    3 { "Member Server" }
                    4 { "Backup Domain Controller" }
                    5 { "Primary Domain Controller" }
                    default { "Desconocido ($($dominioInfo.DomainRole))" }
                }
            }
        } catch {
            $datos.InformacionDominio = @{ Error = "No se pudo obtener información de dominio" }
        }
        
        # Características de Windows instaladas
        try {
            $caracteristicasWindows = Get-WindowsFeature -ErrorAction SilentlyContinue | 
                                     Where-Object { $_.InstallState -eq "Installed" } |
                                     Select-Object Name, DisplayName, InstallState |
                                     Sort-Object DisplayName
            
            if ($caracteristicasWindows) {
                $datos.CaracteristicasWindows = $caracteristicasWindows
            } else {
                # Método alternativo si Get-WindowsFeature no está disponible
                $datos.CaracteristicasWindows = @{ Info = "Get-WindowsFeature no disponible en esta versión" }
            }
        } catch {
            $datos.CaracteristicasWindows = @{ Error = "Error al obtener características de Windows: $_" }
        }
        
        # Resumen de estado de roles
        $rolesCriticos = $rolesDetectados | Where-Object { $_.Critico -eq $true }
        $rolesDetenidos = $rolesDetectados | Where-Object { $_.Estado -ne "Running" }
        
        $datos.ResumenRoles = @{
            TotalRoles = $rolesDetectados.Count
            RolesCriticos = $rolesCriticos.Count
            RolesDetenidos = $rolesDetenidos.Count
            EstadoGeneral = if ($rolesDetenidos.Count -eq 0) { "Todos los roles funcionando" } 
                           elseif ($rolesDetenidos.Count -eq 1) { "1 rol detenido" } 
                           else { "$($rolesDetenidos.Count) roles detenidos" }
        }
        
        return $datos
        
    } catch {
        Write-Warning "Error en análisis de roles del servidor: $_"
        return @{ Error = $_.Exception.Message }
    }
}

# NUEVA FUNCIÓN: Análisis de Políticas de Grupo (GPO)
function Get-AnalisisPoliticasGrupo {
    try {
        Write-Progress -Activity "Analizando políticas de grupo aplicadas" -PercentComplete 60
        Write-Host "   Analizando políticas de grupo (GPO)..." -ForegroundColor Yellow
        
        $datos = @{}
        
        # 1. OBTENER INFORMACIÓN BÁSICA DE GPO
        try {
            # Verificar si el sistema está en un dominio
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
            if (-not $computerSystem.PartOfDomain) {
                Write-Host "   Sistema no está en dominio, análisis GPO limitado" -ForegroundColor Yellow
                return @{
                    EnDominio = $false
                    Mensaje = "El sistema no está unido a un dominio. Análisis de GPO no aplicable."
                    PoliticasLocales = Get-AnalisisPoliticasLocales
                }
            }
            
            Write-Host "   Sistema en dominio detectado, analizando GPOs..." -ForegroundColor Yellow
            $datos.EnDominio = $true
            
            # Ejecutar gpresult para obtener información de GPO
            $gpresultOutput = gpresult /r /scope:computer 2>$null
            $gpresultUser = gpresult /r /scope:user 2>$null
            
            # Parsear información de GPO del equipo
            $gpoComputer = @()
            $gpoUser = @()
            
            if ($gpresultOutput) {
                $inGPOSection = $false
                foreach ($line in $gpresultOutput) {
                    if ($line -match "Applied Group Policy Objects") {
                        $inGPOSection = $true
                        continue
                    }
                    if ($inGPOSection -and $line.Trim() -ne "" -and $line -notmatch "^-+$") {
                        if ($line -match "The following GPOs were not applied") {
                            break
                        }
                        $gpoName = $line.Trim()
                        if ($gpoName -and $gpoName -ne "None") {
                            $gpoComputer += [PSCustomObject]@{
                                Nombre = $gpoName
                                Tipo = "Equipo"
                                Estado = "Aplicada"
                            }
                        }
                    }
                }
            }
            
            # Parsear información de GPO del usuario
            if ($gpresultUser) {
                $inGPOSection = $false
                foreach ($line in $gpresultUser) {
                    if ($line -match "Applied Group Policy Objects") {
                        $inGPOSection = $true
                        continue
                    }
                    if ($inGPOSection -and $line.Trim() -ne "" -and $line -notmatch "^-+$") {
                        if ($line -match "The following GPOs were not applied") {
                            break
                        }
                        $gpoName = $line.Trim()
                        if ($gpoName -and $gpoName -ne "None") {
                            $gpoUser += [PSCustomObject]@{
                                Nombre = $gpoName
                                Tipo = "Usuario"
                                Estado = "Aplicada"
                            }
                        }
                    }
                }
            }
            
            $datos.GPOsEquipo = $gpoComputer
            $datos.GPOsUsuario = $gpoUser
            
            # Obtener información detallada con gpresult /v
            try {
                Write-Host "   Obteniendo detalles de configuración GPO..." -ForegroundColor Yellow
                $gpresultDetailed = gpresult /v /scope:computer 2>$null
                
                # Parsear configuraciones específicas
                $configuracionesSeguridad = @()
                $configuracionesRed = @()
                $configuracionesAuditoria = @()
                
                if ($gpresultDetailed) {
                    $currentSection = ""
                    foreach ($line in $gpresultDetailed) {
                        # Identificar secciones importantes
                        if ($line -match "Security Settings") {
                            $currentSection = "Security"
                        } elseif ($line -match "Network") {
                            $currentSection = "Network"
                        } elseif ($line -match "Audit") {
                            $currentSection = "Audit"
                        }
                        
                        # Extraer configuraciones relevantes
                        if ($line -match "^\s+(.+):\s+(.+)$") {
                            $setting = $matches[1].Trim()
                            $value = $matches[2].Trim()
                            
                            $configObj = [PSCustomObject]@{
                                Configuracion = $setting
                                Valor = $value
                                Seccion = $currentSection
                            }
                            
                            switch ($currentSection) {
                                "Security" { $configuracionesSeguridad += $configObj }
                                "Network" { $configuracionesRed += $configObj }
                                "Audit" { $configuracionesAuditoria += $configObj }
                            }
                        }
                    }
                }
                
                $datos.ConfiguracionesSeguridad = $configuracionesSeguridad
                $datos.ConfiguracionesRed = $configuracionesRed
                $datos.ConfiguracionesAuditoria = $configuracionesAuditoria
                
            } catch {
                Write-Host "   Error al obtener detalles de GPO: $_" -ForegroundColor Red
                $datos.ConfiguracionesSeguridad = @()
                $datos.ConfiguracionesRed = @()
                $datos.ConfiguracionesAuditoria = @()
            }
            
            # Información de última actualización de GPO
            try {
                $gpoUpdateInfo = gpresult /r | Select-String "Last time Group Policy was applied"
                if ($gpoUpdateInfo) {
                    $datos.UltimaActualizacionGPO = $gpoUpdateInfo.ToString().Split(":")[1].Trim()
                } else {
                    $datos.UltimaActualizacionGPO = "No disponible"
                }
            } catch {
                $datos.UltimaActualizacionGPO = "Error al obtener fecha"
            }
            
        } catch {
            Write-Host "   Error al analizar GPOs: $_" -ForegroundColor Red
            $datos.Error = "Error al obtener información de GPO: $_"
        }
        
        # 2. ANÁLISIS DE POLÍTICAS LOCALES (siempre disponible)
        $datos.PoliticasLocales = Get-AnalisisPoliticasLocales
        
        return $datos
        
    } catch {
        Write-Warning "Error en análisis de políticas de grupo: $_"
        return @{ Error = $_.Exception.Message }
    }
}

function Get-AnalisisPoliticasLocales {
    try {
        Write-Host "   Analizando políticas de seguridad locales..." -ForegroundColor Yellow
        
        $politicasLocales = @()
        
        # Políticas de contraseñas
        try {
            $secpol = secedit /export /cfg "$env:TEMP\secpol.cfg" 2>$null
            if (Test-Path "$env:TEMP\secpol.cfg") {
                $secpolContent = Get-Content "$env:TEMP\secpol.cfg"
                
                foreach ($line in $secpolContent) {
                    if ($line -match "^(.+)\s*=\s*(.+)$") {
                        $setting = $matches[1].Trim()
                        $value = $matches[2].Trim()
                        
                        # Mapear configuraciones importantes
                        $descripcion = switch ($setting) {
                            "MinimumPasswordAge" { "Edad mínima de contraseña (días)" }
                            "MaximumPasswordAge" { "Edad máxima de contraseña (días)" }
                            "MinimumPasswordLength" { "Longitud mínima de contraseña" }
                            "PasswordComplexity" { "Complejidad de contraseña requerida" }
                            "PasswordHistorySize" { "Historial de contraseñas" }
                            "LockoutBadCount" { "Umbral de bloqueo de cuenta" }
                            "LockoutDuration" { "Duración de bloqueo (minutos)" }
                            "ResetLockoutCount" { "Restablecer contador después de (minutos)" }
                            "RequireLogonToChangePassword" { "Requerir logon para cambiar contraseña" }
                            "ClearTextPassword" { "Almacenar contraseñas con cifrado reversible" }
                            default { $setting }
                        }
                        
                        if ($descripcion -ne $setting) {
                            $politicasLocales += [PSCustomObject]@{
                                Politica = $descripcion
                                Valor = $value
                                Configuracion = $setting
                                Categoria = "Políticas de Contraseña"
                            }
                        }
                    }
                }
                
                Remove-Item "$env:TEMP\secpol.cfg" -Force -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Host "   Error al obtener políticas de seguridad locales: $_" -ForegroundColor Red
        }
        
        # Políticas de auditoría
        try {
            $auditPolicies = auditpol /get /category:* 2>$null
            if ($auditPolicies) {
                $currentCategory = ""
                foreach ($line in $auditPolicies) {
                    if ($line -match "^\s*(.+):$") {
                        $currentCategory = $matches[1].Trim()
                    } elseif ($line -match "^\s+(.+?)\s+(Success and Failure|Success|Failure|No Auditing)$") {
                        $auditType = $matches[1].Trim()
                        $auditSetting = $matches[2].Trim()
                        
                        $politicasLocales += [PSCustomObject]@{
                            Politica = $auditType
                            Valor = $auditSetting
                            Configuracion = $auditType
                            Categoria = "Auditoría - $currentCategory"
                        }
                    }
                }
            }
        } catch {
            Write-Host "   Error al obtener políticas de auditoría: $_" -ForegroundColor Red
        }
        
        return $politicasLocales
        
    } catch {
        Write-Warning "Error al obtener políticas locales: $_"
        return @()
    }
}

# NUEVA FUNCIÓN: Verificación de Cumplimiento (CIS Benchmarks básicos)
function Get-VerificacionCumplimiento {
    try {
        Write-Progress -Activity "Verificando cumplimiento con estándares de seguridad" -PercentComplete 65
        Write-Host "   Verificando cumplimiento con CIS Benchmarks..." -ForegroundColor Yellow
        
        $verificaciones = @()
        
        # 1. VERIFICACIONES DE POLÍTICAS DE CONTRASEÑA
        Write-Host "   Verificando políticas de contraseña..." -ForegroundColor Yellow
        
        # Obtener políticas de contraseña actuales
        $passwordPolicy = net accounts 2>$null
        $minPasswordLength = 0
        $maxPasswordAge = 0
        $minPasswordAge = 0
        $passwordComplexity = $false
        
        if ($passwordPolicy) {
            foreach ($line in $passwordPolicy) {
                if ($line -match "Minimum password length:\s*(\d+)") {
                    $minPasswordLength = [int]$matches[1]
                } elseif ($line -match "Maximum password age $$days$$:\s*(\d+)") {
                    $maxPasswordAge = [int]$matches[1]
                } elseif ($line -match "Minimum password age $$days$$:\s*(\d+)") {
                    $minPasswordAge = [int]$matches[1]
                }
            }
        }
        
        # Verificar complejidad de contraseña
        try {
            $secpol = secedit /export /cfg "$env:TEMP\secpol_check.cfg" 2>$null
            if (Test-Path "$env:TEMP\secpol_check.cfg") {
                $secpolContent = Get-Content "$env:TEMP\secpol_check.cfg"
                $complexityLine = $secpolContent | Where-Object { $_ -match "PasswordComplexity" }
                if ($complexityLine -and $complexityLine -match "=\s*1") {
                    $passwordComplexity = $true
                }
                Remove-Item "$env:TEMP\secpol_check.cfg" -Force -ErrorAction SilentlyContinue
            }
        } catch {}
        
        # CIS 1.1.1 - Enforce password history
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.1.1"
            Descripcion = "Enforce password history: 24 or more passwords remembered"
            Categoria = "Políticas de Contraseña"
            EstadoActual = "Verificar manualmente"
            Recomendacion = "24 o más contraseñas"
            Cumple = "Pendiente"
            Criticidad = "Media"
        }
        
        # CIS 1.1.2 - Maximum password age
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.1.2"
            Descripcion = "Maximum password age: 365 or fewer days"
            Categoria = "Políticas de Contraseña"
            EstadoActual = "$maxPasswordAge días"
            Recomendacion = "365 días o menos"
            Cumple = if ($maxPasswordAge -le 365 -and $maxPasswordAge -gt 0) { "Sí" } else { "No" }
            Criticidad = "Media"
        }
        
        # CIS 1.1.3 - Minimum password age
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.1.3"
            Descripcion = "Minimum password age: 1 or more days"
            Categoria = "Políticas de Contraseña"
            EstadoActual = "$minPasswordAge días"
            Recomendacion = "1 día o más"
            Cumple = if ($minPasswordAge -ge 1) { "Sí" } else { "No" }
            Criticidad = "Baja"
        }
        
        # CIS 1.1.4 - Minimum password length
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.1.4"
            Descripcion = "Minimum password length: 14 or more characters"
            Categoria = "Políticas de Contraseña"
            EstadoActual = "$minPasswordLength caracteres"
            Recomendacion = "14 caracteres o más"
            Cumple = if ($minPasswordLength -ge 14) { "Sí" } else { "No" }
            Criticidad = "Alta"
        }
        
        # CIS 1.1.5 - Password complexity
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.1.5"
            Descripcion = "Password must meet complexity requirements"
            Categoria = "Políticas de Contraseña"
            EstadoActual = if ($passwordComplexity) { "Habilitado" } else { "Deshabilitado" }
            Recomendacion = "Habilitado"
            Cumple = if ($passwordComplexity) { "Sí" } else { "No" }
            Criticidad = "Alta"
        }
        
        # 2. VERIFICACIONES DE POLÍTICAS DE BLOQUEO DE CUENTA
        Write-Host "   Verificando políticas de bloqueo de cuenta..." -ForegroundColor Yellow
        
        $lockoutThreshold = 0
        $lockoutDuration = 0
        
        if ($passwordPolicy) {
            foreach ($line in $passwordPolicy) {
                if ($line -match "Lockout threshold:\s*(\d+)") {
                    $lockoutThreshold = [int]$matches[1]
                } elseif ($line -match "Lockout duration $$minutes$$:\s*(\d+)") {
                    $lockoutDuration = [int]$matches[1]
                }
            }
        }
        
        # CIS 1.2.1 - Account lockout threshold
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.2.1"
            Descripcion = "Account lockout threshold: 5 or fewer invalid attempts"
            Categoria = "Políticas de Bloqueo"
            EstadoActual = if ($lockoutThreshold -eq 0) { "Sin bloqueo" } else { "$lockoutThreshold intentos" }
            Recomendacion = "5 intentos o menos"
            Cumple = if ($lockoutThreshold -gt 0 -and $lockoutThreshold -le 5) { "Sí" } else { "No" }
            Criticidad = "Media"
        }
        
        # CIS 1.2.2 - Account lockout duration
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.2.2"
            Descripcion = "Account lockout duration: 15 or more minutes"
            Categoria = "Políticas de Bloqueo"
            EstadoActual = "$lockoutDuration minutos"
            Recomendacion = "15 minutos o más"
            Cumple = if ($lockoutDuration -ge 15) { "Sí" } else { "No" }
            Criticidad = "Media"
        }
        
        # 3. VERIFICACIONES DE SERVICIOS Y CARACTERÍSTICAS
        Write-Host "   Verificando servicios y características de seguridad..." -ForegroundColor Yellow
        
        # Windows Defender
        $defenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-18.9.39.1"
            Descripcion = "Windows Defender Antivirus: Real-time protection enabled"
            Categoria = "Antivirus"
            EstadoActual = if ($defenderStatus -and $defenderStatus.RealTimeProtectionEnabled) { "Habilitado" } else { "Deshabilitado" }
            Recomendacion = "Habilitado"
            Cumple = if ($defenderStatus -and $defenderStatus.RealTimeProtectionEnabled) { "Sí" } else { "No" }
            Criticidad = "Alta"
        }
        
        # Windows Firewall
        $firewallProfiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue
        foreach ($profile in $firewallProfiles) {
            $verificaciones += [PSCustomObject]@{
                ID = "CIS-9.1.$($profile.Name)"
                Descripcion = "Windows Firewall: $($profile.Name) profile enabled"
                Categoria = "Firewall"
                EstadoActual = if ($profile.Enabled) { "Habilitado" } else { "Deshabilitado" }
                Recomendacion = "Habilitado"
                Cumple = if ($profile.Enabled) { "Sí" } else { "No" }
                Criticidad = "Alta"
            }
        }
        
        # UAC
        $uacSettings = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -ErrorAction SilentlyContinue
        $uacEnabled = $uacSettings.EnableLUA -eq 1
        
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-2.3.17.1"
            Descripcion = "User Account Control: Admin Approval Mode for Built-in Administrator"
            Categoria = "Control de Acceso"
            EstadoActual = if ($uacEnabled) { "Habilitado" } else { "Deshabilitado" }
            Recomendacion = "Habilitado"
            Cumple = if ($uacEnabled) { "Sí" } else { "No" }
            Criticidad = "Alta"
        }
        
        # 4. VERIFICACIONES DE AUDITORÍA
        Write-Host "   Verificando configuración de auditoría..." -ForegroundColor Yellow
        
        $auditCategories = @(
            @{ Name = "Logon/Logoff"; ID = "CIS-17.1" },
            @{ Name = "Account Management"; ID = "CIS-17.2" },
            @{ Name = "Privilege Use"; ID = "CIS-17.3" },
            @{ Name = "System"; ID = "CIS-17.4" }
        )
        
        try {
            $auditPol = auditpol /get /category:* 2>$null
            foreach ($category in $auditCategories) {
                $auditEnabled = $false
                if ($auditPol) {
                    $categorySection = $auditPol | Select-String -Pattern $category.Name -Context 0,10
                    if ($categorySection -and $categorySection.ToString() -match "(Success and Failure|Success|Failure)") {
                        $auditEnabled = $true
                    }
                }
                
                $verificaciones += [PSCustomObject]@{
                    ID = $category.ID
                    Descripcion = "Audit $($category.Name) events"
                    Categoria = "Auditoría"
                    EstadoActual = if ($auditEnabled) { "Configurado" } else { "No configurado" }
                    Recomendacion = "Success and Failure"
                    Cumple = if ($auditEnabled) { "Sí" } else { "No" }
                    Criticidad = "Media"
                }
            }
        } catch {
            Write-Host "   Error al verificar auditoría: $_" -ForegroundColor Red
        }
        
        # 5. VERIFICACIONES DE REGISTRO
        Write-Host "   Verificando configuraciones del registro..." -ForegroundColor Yellow
        
        # SMB v1
        $smbv1Enabled = $false
        try {
            $smbv1Feature = Get-WindowsOptionalFeature -Online -FeatureName "SMB1Protocol" -ErrorAction SilentlyContinue
            $smbv1Enabled = $smbv1Feature -and $smbv1Feature.State -eq "Enabled"
        } catch {
            # Método alternativo para versiones más antiguas
            $smbv1Reg = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\mrxsmb10" -Name "Start" -ErrorAction SilentlyContinue
            $smbv1Enabled = $smbv1Reg -and $smbv1Reg.Start -ne 4
        }
        
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-18.3.1"
            Descripcion = "SMB v1 protocol disabled"
            Categoria = "Protocolos de Red"
            EstadoActual = if ($smbv1Enabled) { "Habilitado" } else { "Deshabilitado" }
            Recomendacion = "Deshabilitado"
            Cumple = if (-not $smbv1Enabled) { "Sí" } else { "No" }
            Criticidad = "Alta"
        }
        
        # Remote Desktop
        $rdpEnabled = $false
        try {
            $rdpSetting = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -ErrorAction SilentlyContinue
            $rdpEnabled = $rdpSetting -and $rdpSetting.fDenyTSConnections -eq 0
        } catch {}
        
        $verificaciones += [PSCustomObject]@{
            ID = "CIS-18.9.48.3.1"
            Descripcion = "Remote Desktop connections security"
            Categoria = "Acceso Remoto"
            EstadoActual = if ($rdpEnabled) { "Habilitado" } else { "Deshabilitado" }
            Recomendacion = "Configurado según necesidad"
            Cumple = "Revisar"
            Criticidad = "Media"
        }
        
        # 6. RESUMEN DE CUMPLIMIENTO
        $totalVerificaciones = $verificaciones.Count
        $cumpleCompleto = ($verificaciones | Where-Object { $_.Cumple -eq "Sí" }).Count
        $noCumple = ($verificaciones | Where-Object { $_.Cumple -eq "No" }).Count
        $pendienteRevision = ($verificaciones | Where-Object { $_.Cumple -in @("Pendiente", "Revisar") }).Count
        
        $porcentajeCumplimiento = if ($totalVerificaciones -gt 0) {
            [math]::Round(($cumpleCompleto / $totalVerificaciones) * 100, 1)
        } else { 0 }
        
        $resumenCumplimiento = @{
            TotalVerificaciones = $totalVerificaciones
            Cumple = $cumpleCompleto
            NoCumple = $noCumple
            PendienteRevision = $pendienteRevision
            PorcentajeCumplimiento = $porcentajeCumplimiento
            NivelCumplimiento = switch ($porcentajeCumplimiento) {
                { $_ -ge 90 } { "Excelente" }
                { $_ -ge 80 } { "Bueno" }
                { $_ -ge 70 } { "Aceptable" }
                { $_ -ge 60 } { "Mejorable" }
                default { "Deficiente" }
            }
        }
        
        return @{
            Verificaciones = $verificaciones
            ResumenCumplimiento = $resumenCumplimiento
        }
        
    } catch {
        Write-Warning "Error en verificación de cumplimiento: $_"
        return @{ Error = $_.Exception.Message }
    }
}

# NUEVA FUNCIÓN: Análisis de Permisos de Carpetas Sensibles
function Get-AnalisisPermisos {
    try {
        Write-Progress -Activity "Analizando permisos de carpetas sensibles" -PercentComplete 70
        Write-Host "   Analizando permisos de carpetas sensibles..." -ForegroundColor Yellow
        
        $analisisPermisos = @()
        
        # Definir carpetas sensibles del sistema
        $carpetasSensibles = @(
            @{ Ruta = "C:\Windows\System32"; Descripcion = "Archivos del sistema Windows" },
            @{ Ruta = "C:\Windows\SysWOW64"; Descripcion = "Archivos del sistema Windows (32-bit)" },
            @{ Ruta = "C:\Program Files"; Descripcion = "Programas instalados" },
            @{ Ruta = "C:\Program Files (x86)"; Descripcion = "Programas instalados (32-bit)" },
            @{ Ruta = "C:\Windows\Temp"; Descripcion = "Archivos temporales del sistema" },
            @{ Ruta = "C:\ProgramData"; Descripcion = "Datos de aplicaciones" },
            @{ Ruta = "C:\Users\Public"; Descripcion = "Carpeta pública de usuarios" },
            @{ Ruta = "C:\inetpub"; Descripcion = "Sitios web IIS" }
        )
        
        # Agregar carpetas compartidas
        try {
            $carpetasCompartidas = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "IPC$" -and $_.Name -ne "ADMIN$" -and $_.Name -notlike "*$" }
            foreach ($share in $carpetasCompartidas) {
                $carpetasSensibles += @{ Ruta = $share.Path; Descripcion = "Carpeta compartida: $($share.Name)" }
            }
        } catch {
            Write-Host "   No se pudieron obtener carpetas compartidas" -ForegroundColor Yellow
        }
        
        foreach ($carpeta in $carpetasSensibles) {
            try {
                if (Test-Path $carpeta.Ruta) {
                    Write-Host "   Analizando: $($carpeta.Ruta)" -ForegroundColor Yellow
                    
                    # Obtener ACL de la carpeta
                    $acl = Get-Acl $carpeta.Ruta -ErrorAction SilentlyContinue
                    
                    if ($acl) {
                        $permisosProblematicos = @()
                        $permisosNormales = @()
                        
                        foreach ($access in $acl.Access) {
                            $usuario = $access.IdentityReference.Value
                            $permisos = $access.FileSystemRights.ToString()
                            $tipo = $access.AccessControlType.ToString()
                            $herencia = $access.IsInherited
                            
                            # Identificar permisos problemáticos
                            $esProblematico = $false
                            $razon = ""
                            
                            # Verificar si Everyone o Users tienen permisos excesivos
                            if ($usuario -match "(Everyone|Users|Authenticated Users)" -and $tipo -eq "Allow") {
                                if ($permisos -match "(FullControl|Modify|Write)" -and $carpeta.Ruta -match "(System32|Program Files|Windows)") {
                                    $esProblematico = $true
                                    $razon = "Permisos excesivos para grupo amplio en carpeta del sistema"
                                }
                            }
                            
                            # Verificar permisos de escritura en carpetas del sistema
                            if ($permisos -match "(Write|Modify|FullControl)" -and $tipo -eq "Allow" -and $carpeta.Ruta -match "(System32|SysWOW64)") {
                                if ($usuario -notmatch "(SYSTEM|Administrators|TrustedInstaller)") {
                                    $esProblematico = $true
                                    $razon = "Permisos de escritura en carpeta crítica del sistema"
                                }
                            }
                            
                            # Verificar permisos en carpetas temporales
                            if ($carpeta.Ruta -match "Temp" -and $permisos -match "FullControl" -and $usuario -match "Everyone") {
                                $esProblematico = $true
                                $razon = "Control total para Everyone en carpeta temporal"
                            }
                            
                            $permisoObj = [PSCustomObject]@{
                                Usuario = $usuario
                                Permisos = $permisos
                                Tipo = $tipo
                                Heredado = $herencia
                                Problematico = $esProblematico
                                Razon = $razon
                            }
                            
                            if ($esProblematico) {
                                $permisosProblematicos += $permisoObj
                            } else {
                                $permisosNormales += $permisoObj
                            }
                        }
                        
                        $analisisPermisos += [PSCustomObject]@{
                            Ruta = $carpeta.Ruta
                            Descripcion = $carpeta.Descripcion
                            Propietario = $acl.Owner
                            PermisosProblematicos = $permisosProblematicos
                            PermisosNormales = $permisosNormales
                            TotalPermisos = $acl.Access.Count
                            PermisosProblematicosCount = $permisosProblematicos.Count
                            Estado = if ($permisosProblematicos.Count -eq 0) { "Normal" } 
                                    elseif ($permisosProblematicos.Count -le 2) { "Advertencia" } 
                                    else { "Crítico" }
                        }
                    }
                } else {
                    Write-Host "   Carpeta no existe: $($carpeta.Ruta)" -ForegroundColor Gray
                }
            } catch {
                Write-Host "   Error al analizar $($carpeta.Ruta): $_" -ForegroundColor Red
                $analisisPermisos += [PSCustomObject]@{
                    Ruta = $carpeta.Ruta
                    Descripcion = $carpeta.Descripcion
                    Error = $_.Exception.Message
                    Estado = "Error"
                }
            }
        }
        
        # Resumen del análisis de permisos
        $totalCarpetas = $analisisPermisos.Count
        $carpetasConProblemas = ($analisisPermisos | Where-Object { $_.Estado -in @("Advertencia", "Crítico") }).Count
        $carpetasCriticas = ($analisisPermisos | Where-Object { $_.Estado -eq "Crítico" }).Count
        
        $resumenPermisos = @{
            TotalCarpetasAnalizadas = $totalCarpetas
            CarpetasConProblemas = $carpetasConProblemas
            CarpetasCriticas = $carpetasCriticas
            CarpetasNormales = $totalCarpetas - $carpetasConProblemas
            PorcentajeSeguras = if ($totalCarpetas -gt 0) { 
                [math]::Round((($totalCarpetas - $carpetasConProblemas) / $totalCarpetas) * 100, 1) 
            } else { 0 }
            NivelSeguridad = switch ($carpetasCriticas) {
                0 { if ($carpetasConProblemas -eq 0) { "Excelente" } else { "Bueno" } }
                { $_ -le 2 } { "Aceptable" }
                default { "Deficiente" }
            }
        }
        
        return @{
            AnalisisPermisos = $analisisPermisos
            ResumenPermisos = $resumenPermisos
        }
        
    } catch {
        Write-Warning "Error en análisis de permisos: $_"
        return @{ Error = $_.Exception.Message }
    }
}

# NUEVA FUNCIÓN: Auditoría de Software Instalado
function Get-AuditoriaSoftware {
    try {
        Write-Progress -Activity "Realizando auditoría de software instalado" -PercentComplete 75
        Write-Host "   Auditando software instalado..." -ForegroundColor Yellow
        
        $softwareInstalado = @()
        $softwareProblematico = @()
        
        # 1. OBTENER SOFTWARE INSTALADO DESDE EL REGISTRO
        Write-Host "   Obteniendo lista de software instalado..." -ForegroundColor Yellow
        
        $registryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($path in $registryPaths) {
            try {
                $installedSoftware = Get-ItemProperty $path -ErrorAction SilentlyContinue | 
                                   Where-Object { $_.DisplayName -and $_.DisplayName -notmatch "^(KB|Update for)" }
                
                foreach ($software in $installedSoftware) {
                    $nombre = $software.DisplayName
                    $version = $software.DisplayVersion
                    $fabricante = $software.Publisher
                    $fechaInstalacion = $software.InstallDate
                    $tamaño = $software.EstimatedSize
                    
                    # Convertir fecha de instalación
                    $fechaInstalacionFormateada = "No disponible"
                    if ($fechaInstalacion -and $fechaInstalacion -match "^\d{8}$") {
                        try {
                            $fechaInstalacionFormateada = [DateTime]::ParseExact($fechaInstalacion, "yyyyMMdd", $null).ToString("dd/MM/yyyy")
                        } catch {}
                    }
                    
                    # Convertir tamaño
                    $tamañoMB = if ($tamaño) { [math]::Round($tamaño / 1024, 2) } else { 0 }
                    
                    $softwareInstalado += [PSCustomObject]@{
                        Nombre = $nombre
                        Version = $version
                        Fabricante = $fabricante
                        FechaInstalacion = $fechaInstalacionFormateada
                        TamañoMB = $tamañoMB
                        RegistryPath = $software.PSPath
                    }
                }
            } catch {
                Write-Host "   Error al leer registro: $_" -ForegroundColor Red
            }
        }
        
        # Eliminar duplicados
        $softwareInstalado = $softwareInstalado | Sort-Object Nombre, Version | Get-Unique -AsString
        
        Write-Host "   Se encontraron $($softwareInstalado.Count) programas instalados" -ForegroundColor Yellow
        
        # 2. IDENTIFICAR SOFTWARE PROBLEMÁTICO
        Write-Host "   Identificando software problemático..." -ForegroundColor Yellow
        
        # Lista de software conocido como problemático o desactualizado
        $softwareProblemas = @(
            @{ Nombre = "Adobe Flash Player"; Razon = "Software descontinuado y vulnerable"; Criticidad = "Alta" },
            @{ Nombre = "Java"; VersionMinima = "8.0.300"; Razon = "Versiones antiguas de Java son vulnerables"; Criticidad = "Alta" },
            @{ Nombre = "Adobe Reader"; VersionMinima = "2021.0"; Razon = "Versiones antiguas tienen vulnerabilidades"; Criticidad = "Media" },
            @{ Nombre = "VLC media player"; VersionMinima = "3.0.16"; Razon = "Versiones antiguas pueden tener vulnerabilidades"; Criticidad = "Baja" },
            @{ Nombre = "WinRAR"; VersionMinima = "6.0"; Razon = "Versiones antiguas tienen vulnerabilidades conocidas"; Criticidad = "Media" },
            @{ Nombre = "7-Zip"; VersionMinima = "21.0"; Razon = "Versiones antiguas pueden ser vulnerables"; Criticidad = "Baja" },
            @{ Nombre = "Google Chrome"; Razon = "Verificar que esté actualizado"; Criticidad = "Media" },
            @{ Nombre = "Mozilla Firefox"; Razon = "Verificar que esté actualizado"; Criticidad = "Media" },
            @{ Nombre = "Internet Explorer"; Razon = "Navegador descontinuado"; Criticidad = "Alta" }
        )
        
        # Software sin soporte conocido
        $softwareSinSoporte = @(
            "Windows XP", "Windows Vista", "Windows 7", "Office 2010", "Office 2013", 
            "Adobe Flash", "Internet Explorer", "Silverlight"
        )
        
        foreach ($software in $softwareInstalado) {
            $problemas = @()
            
            # Verificar si está en la lista de software problemático
            foreach ($problema in $softwareProblemas) {
                if ($software.Nombre -like "*$($problema.Nombre)*") {
                    if ($problema.VersionMinima) {
                        # Comparar versiones (simplificado)
                        $versionActual = $software.Version
                        if ($versionActual) {
                            try {
                                $versionActualNum = [Version]$versionActual.Split(' ')[0]
                                $versionMinimaNum = [Version]$problema.VersionMinima
                                
                                if ($versionActualNum -lt $versionMinimaNum) {
                                    $problemas += @{
                                        Tipo = "Versión desactualizada"
                                        Descripcion = $problema.Razon
                                        Criticidad = $problema.Criticidad
                                        Recomendacion = "Actualizar a versión $($problema.VersionMinima) o superior"
                                    }
                                }
                            } catch {
                                $problemas += @{
                                    Tipo = "Verificación de versión"
                                    Descripcion = "No se pudo verificar la versión automáticamente"
                                    Criticidad = "Baja"
                                    Recomendacion = "Verificar manualmente la versión"
                                }
                            }
                        }
                    } else {
                        $problemas += @{
                            Tipo = "Software problemático"
                            Descripcion = $problema.Razon
                            Criticidad = $problema.Criticidad
                            Recomendacion = "Considerar desinstalar o reemplazar"
                        }
                    }
                }
            }
            
            # Verificar software sin soporte
            foreach ($sinSoporte in $softwareSinSoporte) {
                if ($software.Nombre -like "*$sinSoporte*") {
                    $problemas += @{
                        Tipo = "Sin soporte"
                        Descripcion = "Software sin soporte del fabricante"
                        Criticidad = "Alta"
                        Recomendacion = "Migrar a versión soportada"
                    }
                }
            }
            
            # Verificar software muy antiguo (más de 5 años)
            if ($software.FechaInstalacion -ne "No disponible") {
                try {
                    $fechaInstalacion = [DateTime]::ParseExact($software.FechaInstalacion, "dd/MM/yyyy", $null)
                    $añosAntiguedad = ((Get-Date) - $fechaInstalacion).Days / 365
                    
                    if ($añosAntiguedad -gt 5) {
                        $problemas += @{
                            Tipo = "Software muy antiguo"
                            Descripcion = "Instalado hace más de 5 años"
                            Criticidad = "Baja"
                            Recomendacion = "Verificar si hay actualizaciones disponibles"
                        }
                    }
                } catch {}
            }
            
            # Verificar software sin fabricante conocido
            if (-not $software.Fabricante -or $software.Fabricante -eq "") {
                $problemas += @{
                    Tipo = "Fabricante desconocido"
                    Descripcion = "Software sin información del fabricante"
                    Criticidad = "Media"
                    Recomendacion = "Verificar la legitimidad del software"
                }
            }
            
            if ($problemas.Count -gt 0) {
                $softwareProblematico += [PSCustomObject]@{
                    Nombre = $software.Nombre
                    Version = $software.Version
                    Fabricante = $software.Fabricante
                    FechaInstalacion = $software.FechaInstalacion
                    Problemas = $problemas
                    CriticidadMaxima = ($problemas | ForEach-Object { $_.Criticidad } | Sort-Object { 
                        switch ($_) { "Alta" { 3 } "Media" { 2 } "Baja" { 1 } default { 0 } } 
                    } -Descending)[0]
                }
            }
        }
        
        # 3. ANÁLISIS DE NAVEGADORES Y PLUGINS
        Write-Host "   Analizando navegadores y plugins..." -ForegroundColor Yellow
        
        $navegadores = $softwareInstalado | Where-Object { 
            $_.Nombre -match "(Chrome|Firefox|Edge|Internet Explorer|Opera|Safari)" 
        }
        
        # 4. RESUMEN DE AUDITORÍA
        $totalSoftware = $softwareInstalado.Count
        $softwareConProblemas = $softwareProblematico.Count
        $softwareCritico = ($softwareProblematico | Where-Object { $_.CriticidadMaxima -eq "Alta" }).Count
        $softwareMedia = ($softwareProblematico | Where-Object { $_.CriticidadMaxima -eq "Media" }).Count
        $softwareBaja = ($softwareProblematico | Where-Object { $_.CriticidadMaxima -eq "Baja" }).Count
        
        $resumenAuditoria = @{
            TotalSoftwareInstalado = $totalSoftware
            SoftwareConProblemas = $softwareConProblemas
            SoftwareCritico = $softwareCritico
            SoftwareRiesgoMedio = $softwareMedia
            SoftwareRiesgoBajo = $softwareBaja
            SoftwareSeguro = $totalSoftware - $softwareConProblemas
            PorcentajeSeguro = if ($totalSoftware -gt 0) { 
                [math]::Round((($totalSoftware - $softwareConProblemas) / $totalSoftware) * 100, 1) 
            } else { 0 }
            NivelRiesgo = switch ($softwareCritico) {
                0 { if ($softwareMedia -eq 0) { "Bajo" } else { "Medio" } }
                { $_ -le 2 } { "Medio" }
                default { "Alto" }
            }
        }
        
        return @{
            SoftwareInstalado = $softwareInstalado
            SoftwareProblematico = $softwareProblematico
            Navegadores = $navegadores
            ResumenAuditoria = $resumenAuditoria
        }
        
    } catch {
        Write-Warning "Error en auditoría de software: $_"
        return @{ Error = $_.Exception.Message }
    }
}

function Get-DatosExtendidos {
    param(
        [bool]$ParchesFaltantes = $false,
        [bool]$RevisarServicioTerceros = $false
    )
    
    try {
        Write-Progress -Activity "Recopilando datos extendidos completos" -PercentComplete 40
        
        $datos = @{}
        
        # 1. PARCHES - Últimos instalados + Faltantes (si se solicita)
        Write-Host "   Analizando parches del sistema..." -ForegroundColor Yellow
        $datos.UltimosParches = Get-HotFix | Sort-Object InstalledOn -Descending | 
                               Select-Object -First 15 HotFixID, Description, InstalledOn, InstalledBy
        
        if ($ParchesFaltantes) {
            try {
                Write-Host "   Verificando parches faltantes..." -ForegroundColor Yellow
                $updateSession = New-Object -ComObject Microsoft.Update.Session
                $updateSearcher = $updateSession.CreateUpdateSearcher()
                $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software'")
                
                $datos.ParchesFaltantes = @()
                for ($i = 0; $i -lt $searchResult.Updates.Count; $i++) {
                    $update = $searchResult.Updates.Item($i)
                    $datos.ParchesFaltantes += @{
                        Titulo = $update.Title
                        Descripcion = $update.Description.Substring(0,[Math]::Min(200,$update.Description.Length))
                        Tamaño = [math]::Round($update.MaxDownloadSize / 1MB, 2)
                    }
                }
            } catch {
                $datos.ParchesFaltantes = @{ Error = "No se pudo verificar parches faltantes: $_" }
            }
        }
        
        # 2. SERVICIOS - Automáticos detenidos + Terceros (si se solicita)
        Write-Host "   Analizando servicios..." -ForegroundColor Yellow
        $serviciosDetenidos = Get-Service | Where-Object { $_.StartType -eq "Automatic" -and $_.Status -ne "Running" }
        
        if ($RevisarServicioTerceros) {
            # Lista de servicios Microsoft conocidos (como respaldo)
            $serviciosMicrosoft = @('Spooler', 'BITS', 'Themes', 'AudioSrv', 'Dnscache', 'eventlog', 'PlugPlay', 
                                    'RpcSs', 'lanmanserver', 'W32Time', 'Winmgmt', 'Schedule', 'LanmanWorkstation',
                                    'DHCP', 'Netlogon', 'PolicyAgent', 'TermService', 'UmRdpService', 'SessionEnv',
                                    'RemoteRegistry')
            
            # Usar análisis de fabricante del ejecutable para clasificación más precisa
            $datos.ServiciosDetenidos = foreach ($servicio in $serviciosDetenidos) {
                # Obtener detalles del servicio mediante CIM
                $servicioCIM = Get-CimInstance Win32_Service -Filter "Name='$($servicio.Name)'" -ErrorAction SilentlyContinue
                $compania = "Desconocido"
                $esMicrosoft = $false
                
                # Intentar determinar la compañía por el ejecutable del servicio
                if ($servicioCIM -and $servicioCIM.PathName) {
                    if ($servicioCIM.PathName -match '"([^"]+)"' -or $servicioCIM.PathName -match '([^\s]+\.exe)') {
                        $exePath = $matches[1]
                        
                        if (Test-Path $exePath -ErrorAction SilentlyContinue) {
                            $fileInfo = (Get-Item $exePath -ErrorAction SilentlyContinue).VersionInfo
                            if ($fileInfo.CompanyName) {
                                $compania = $fileInfo.CompanyName
                                $esMicrosoft = $compania -like "*Microsoft*"
                            }
                        }
                    }
                }
                
                # Si no se pudo determinar, usar la lista de servicios conocidos
                if ($compania -eq "Desconocido") {
                    $esMicrosoft = $serviciosMicrosoft -contains $servicio.Name
                }
                
                [PSCustomObject]@{
                    DisplayName = $servicio.DisplayName
                    Name = $servicio.Name
                    Status = $servicio.Status
                    StartType = $servicio.StartType
                    Compania = $compania
                    EsMicrosoft = $esMicrosoft
                    Tipo = if ($esMicrosoft) { "Microsoft" } else { "Tercero" }
                }
            }
        } else {
            $datos.ServiciosDetenidos = $serviciosDetenidos | Select-Object DisplayName, Name, Status, StartType
        }
        
        # 3. TOP 5 PROCESOS por CPU y Memoria (ORIGINAL)
        Write-Host "   Analizando procesos con alto consumo..." -ForegroundColor Yellow
        $datos.ProcesosCPU = Get-Process | Where-Object { $_.CPU -gt 0 } | 
                            Sort-Object CPU -Descending | Select-Object -First 5 |
                            Select-Object Name, @{Name="CPU";Expression={[math]::Round($_.CPU,2)}}, Id, 
                                        @{Name="Memoria_MB";Expression={[math]::Round($_.WS/1MB,2)}}
        
        $datos.ProcesosMemoria = Get-Process | Sort-Object WS -Descending | Select-Object -First 5 |
                                Select-Object Name, @{Name="Memoria_MB";Expression={[math]::Round($_.WS/1MB,2)}}, 
                                            Id, @{Name="CPU";Expression={[math]::Round($_.CPU,2)}}
        
        # 4. PUERTOS ABIERTOS (ORIGINAL)
        Write-Host "   Analizando puertos de red..." -ForegroundColor Yellow
        $datos.PuertosAbiertos = Get-NetTCPConnection | Where-Object { $_.State -eq "Listen" } |
                                Select-Object LocalAddress, LocalPort, State, OwningProcess |
                                Sort-Object LocalPort
        
        # 5. CONEXIONES ACTIVAS (ORIGINAL)
        $datos.ConexionesActivas = Get-NetTCPConnection | Where-Object { $_.State -eq "Established" } |
                                  Group-Object RemoteAddress | Select-Object Count, Name |
                                  Sort-Object Count -Descending | Select-Object -First 10
        
        # 6. USUARIOS LOCALES (ORIGINAL)
        Write-Host "   Analizando usuarios locales..." -ForegroundColor Yellow
        $datos.UsuariosLocales = Get-LocalUser | Select-Object Name, Enabled, 
                                @{Name="UltimoLogon";Expression={
                                    if($_.LastLogon -eq $null){"Nunca"}else{$_.LastLogon.ToString("dd/MM/yyyy HH:mm")}
                                }},
                                @{Name="PasswordExpires";Expression={
                                    if($_.PasswordExpires -eq $null){"Nunca"}else{$_.PasswordExpires.ToString("dd/MM/yyyy")}
                                }}
        
        # 7. LOGINS FALLIDOS RECIENTES - TOP 5 (ORIGINAL)
        Write-Host "   Analizando intentos de login fallidos..." -ForegroundColor Yellow
        $datos.LoginsFallidos = Get-WinEvent -LogName "Security" -FilterXPath "*[System[EventID=4625]]" -MaxEvents 5 -ErrorAction SilentlyContinue | 
                               Select-Object TimeCreated, 
                                           @{Name="Cuenta";Expression={$_.Properties[5].Value}}, 
                                           @{Name="IPOrigen";Expression={$_.Properties[19].Value}},
                                           @{Name="TipoFallo";Expression={$_.Properties[10].Value}}
        
        # 8. TAREAS PROGRAMADAS (ORIGINAL)
        Write-Host "   Analizando tareas programadas..." -ForegroundColor Yellow
        $datos.TareasProgramadas = Get-ScheduledTask | Where-Object { $_.State -eq "Ready" } | 
                                  Get-ScheduledTaskInfo | Select-Object -First 20 |
                                  Select-Object TaskName, LastRunTime, LastTaskResult,
                                              @{Name="Estado";Expression={
                                                  switch ($_.LastTaskResult) {
                                                      0 { "Exitoso" }
                                                      1 { "Falso/Incorrecto" }
                                                      2 { "Acceso Denegado" }
                                                      default { "Error ($($_.LastTaskResult))" }
                                                  }
                                              }}
        
        # 9. INFORMACIÓN SNMP COMPLETA (ORIGINAL)
        Write-Host "   Verificando configuración SNMP..." -ForegroundColor Yellow
        $datos.ServicioSNMP = Get-Service -Name "SNMP" -ErrorAction SilentlyContinue | 
                             Select-Object Name, Status, StartType
        
        $datos.ConfigSNMP = try {
            $snmpConfig = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities" -ErrorAction SilentlyContinue
            if ($snmpConfig) {
                $snmpConfig.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } | 
                Select-Object Name, Value
            } else {
                @{ Info = "No hay comunidades SNMP configuradas" }
            }
        } catch {
            @{ Error = "No se pudo obtener configuración SNMP" }
        }
        
        Write-Host "   Analizando rendimiento de discos físicos..." -ForegroundColor Yellow
        $datos.RendimientoDisco = Get-Counter "\PhysicalDisk(*)\Avg. Disk sec/Read", "\PhysicalDisk(*)\Avg. Disk sec/Write" -ErrorAction SilentlyContinue
        $datos.DiscosFisicos = Get-PhysicalDisk -ErrorAction SilentlyContinue | Select-Object FriendlyName, HealthStatus, OperationalStatus, @{Name="SizeGB";Expression={[math]::Round($_.Size / 1GB, 2)}}, MediaType, BusType
        
        Write-Host "   Verificando configuración de seguridad..." -ForegroundColor Yellow
        
        # Información detallada Windows Defender
        try {
            $defenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
            if ($defenderStatus) {
                $datos.WindowsDefender = @{
                    Habilitado = $defenderStatus.AntivirusEnabled
                    TiempoRealActivo = $defenderStatus.RealTimeProtectionEnabled
                    ProteccionSpyware = $defenderStatus.AntispywareEnabled
                    ServicioActivo = $defenderStatus.AMServiceEnabled
                    UltimoAnalisisRapido = if ($defenderStatus.QuickScanEndTime) { $defenderStatus.QuickScanEndTime.ToString("dd/MM/yyyy HH:mm") } else { "No disponible" }
                    UltimoAnalisisCompleto = if ($defenderStatus.FullScanEndTime) { $defenderStatus.FullScanEndTime.ToString("dd/MM/yyyy HH:mm") } else { "No disponible" }
                    EdadDefinicionesAV = "$($defenderStatus.AntivirusSignatureAge) días"
                    VersionEngine = $defenderStatus.AMEngineVersion
                    VersionSignature = $defenderStatus.AntispywareSignatureVersion
                    TamperProtection = if (Get-Member -InputObject $defenderStatus -Name "IsTamperProtected" -MemberType Properties) { 
                        $defenderStatus.IsTamperProtected 
                    } else { "No disponible" }
                    ProteccionBasadaNube = $defenderStatus.CloudBlockLevel
                    ProteccionRed = if (Get-Member -InputObject $defenderStatus -Name "NISEnabled" -MemberType Properties) {
                        $defenderStatus.NISEnabled
                    } else { "No disponible" }
                }
            } else {
                # Obtener información de Defender mediante WMI si los cmdlets no están disponibles
                $wmiBuscador = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntivirusProduct -ErrorAction SilentlyContinue
                if ($wmiBuscador) {
                    $defenderWMI = $wmiBuscador | Where-Object { $_.displayName -like "*Defender*" }
                    
                    if ($defenderWMI) {
                        # Convertir el código de estado a valores legibles
                        $habilitado = [bool]($defenderWMI.productState -band 0x1000)
                        $actualizado = [bool]($defenderWMI.productState -band 0x10)
                        
                        $datos.WindowsDefender = @{
                            Habilitado = $habilitado
                            TiempoRealActivo = $habilitado  # Asumiendo que si está habilitado, la protección en tiempo real también
                            DefinicionesActualizadas = $actualizado
                            FechaUltimaActualizacion = "No disponible"
                            UltimoAnalisisCompleto = "No disponible"
                        }
                    } else {
                        $datos.WindowsDefender = @{ Estado = "No detectado" }
                    }
                } else {
                    $datos.WindowsDefender = @{ Estado = "No disponible" }
                }
            }
        } catch {
            $datos.WindowsDefender = @{ Error = "No se pudo obtener estado de Windows Defender: $_" }
        }
        
        # Otros antivirus instalados
        try {
            $otrosAntivirus = @()
            $antivirusList = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntivirusProduct -ErrorAction SilentlyContinue
            
            if ($antivirusList) {
                foreach ($av in $antivirusList) {
                    if ($av.displayName -notlike "*Defender*") {
                        # Determinar estado basado en el productState
                        $habilitado = [bool]($av.productState -band 0x1000)
                        $actualizado = [bool]($av.productState -band 0x10)
                        
                        $otrosAntivirus += @{
                            Nombre = $av.displayName
                            Habilitado = $habilitado
                            DefinicionesActualizadas = $actualizado
                            Fabricante = $av.companyName
                            Path = $av.pathToSignedProductExe
                        }
                    }
                }
            }
            
            $datos.OtrosAntivirus = $otrosAntivirus
        } catch {
            $datos.OtrosAntivirus = @{ Error = "No se pudo obtener información de otros antivirus: $_" }
        }
        
        # Estado detallado del Firewall por perfil
        try {
            $firewallPerfiles = @()
            $perfilesFirewall = Get-NetFirewallProfile -ErrorAction SilentlyContinue
            
            if ($perfilesFirewall) {
                foreach ($perfil in $perfilesFirewall) {
                    $firewallPerfiles += @{
                        Nombre = $perfil.Name
                        Habilitado = $perfil.Enabled
                        ConexionesEntrantes = $perfil.DefaultInboundAction
                        ConexionesSalientes = $perfil.DefaultOutboundAction
                        NotificacionesActivas = $perfil.NotifyOnListen
                        LoggingAllowed = $perfil.LogAllowed
                        LoggingBlocked = $perfil.LogBlocked
                        RutaLogs = $perfil.LogFileName
                    }
                }
                $datos.Firewall = $firewallPerfiles
            } else {
                # Alternativa usando netsh
                $netshOutput = netsh advfirewall show allprofiles | Out-String
                
                # Procesar la salida para extraer el estado
                $perfilesNombres = @("Domain", "Private", "Public")
                $firewallNetsh = @()
                foreach ($nombre in $perfilesNombres) {
                    if ($netshOutput -match "$nombre Profile Settings:") {
                        $habilitadoMatch = $netshOutput -match "$nombre Profile Settings:\s*\n.*?State\s*(ON|OFF)"
                        $habilitado = if ($matches -and $matches[1] -eq "ON") { $true } else { $false }
                        
                        $firewallNetsh += @{
                            Nombre = $nombre
                            Habilitado = $habilitado
                            Detalles = "Información limitada disponible a través de netsh"
                        }
                    }
                }
                $datos.Firewall = $firewallNetsh
            }
        } catch {
            $datos.Firewall = @{ Error = "No se pudo obtener información del Firewall: $_" }
        }
        
        # Control de Acceso de Usuario (UAC)
        try {
            $uacPolicies = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -ErrorAction SilentlyContinue
            $uacEnabled = $uacPolicies.EnableLUA -eq 1
            
            $nivelUAC = switch ($uacPolicies.ConsentPromptBehaviorAdmin) {
                0 { "Nunca notificar" }
                1 { "Notificar solo aplicaciones (sin oscurecer escritorio)" }
                2 { "Notificar siempre" }
                5 { "Notificar siempre y oscurecer escritorio (predeterminado)" }
                default { "Desconocido ($($uacPolicies.ConsentPromptBehaviorAdmin))" }
            }
            
            $datos.UAC = @{
                Habilitado = $uacEnabled
                Nivel = $nivelUAC
                ElevacionNoSegura = $uacPolicies.EnableInstallerDetection -eq 0
                AdminSinPrompt = $uacPolicies.PromptOnSecureDesktop -eq 0
                AprobacionModo = $uacPolicies.ValidateAdminCodeSignatures -eq 1
            }
        } catch {
            $datos.UAC = @{ Error = "No se pudo obtener información de UAC: $_" }
        }
        
        # SmartScreen y otras protecciones de seguridad
        try {
            $smartScreenPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
            $smartScreenEnabled = Get-ItemProperty -Path $smartScreenPath -Name "SmartScreenEnabled" -ErrorAction SilentlyContinue
            
            $datos.SmartScreen = @{
                Habilitado = if ($smartScreenEnabled) {
                    switch ($smartScreenEnabled.SmartScreenEnabled) {
                        "RequireAdmin" { "Activo (requiere aprobación)" }
                        "Warn" { "Activo (solo advertencias)" }
                        "Off" { "Desactivado" }
                        default { $smartScreenEnabled.SmartScreenEnabled }
                    }
                } else { "No disponible" }
            }
            
            # Información sobre Secure Boot si está disponible
            if (Get-Command Get-SecureBootUEFI -ErrorAction SilentlyContinue) {
                try {
                    $secureBootStatus = Get-SecureBootUEFI -Name SetupMode -ErrorAction SilentlyContinue
                    $datos.SecureBoot = @{
                        Habilitado = $secureBootStatus -eq 0 -or $secureBootStatus.Bytes[0] -eq 0
                    }
                } catch {
                    $datos.SecureBoot = @{ Estado = "No disponible o no compatible" }
                }
            }
            
            # Verificar BitLocker
            if (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue) {
                try {
                    $bitlockerVolumes = Get-BitLockerVolume -ErrorAction SilentlyContinue
                    $datos.BitLocker = $bitlockerVolumes | ForEach-Object {
                        @{
                            Unidad = $_.MountPoint
                            ProteccionActiva = $_.ProtectionStatus
                            MetodoEncriptacion = $_.EncryptionMethod
                            EstadoEncriptacion = $_.VolumeStatus
                            PorcentajeEncriptado = $_.EncryptionPercentage
                        }
                    }
                } catch {
                    $datos.BitLocker = @{ Error = "No se pudo obtener información de BitLocker" }
                }
            }
            
            # Verificar Device Guard y Credential Guard
            try {
                $deviceGuard = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard -ErrorAction SilentlyContinue
                if ($deviceGuard) {
                    $datos.DeviceGuard = @{
                        VirtualizationBasedSecurityStatus = $deviceGuard.VirtualizationBasedSecurityStatus
                        CredentialGuardStatus = if ($deviceGuard.SecurityServicesRunning -contains 1) { "Activo" } else { "Inactivo" }
                        HypervisorEnforcedCodeIntegrity = if ($deviceGuard.SecurityServicesRunning -contains 2) { "Activo" } else { "Inactivo" }
                    }
                }
            } catch {
                $datos.DeviceGuard = @{ Estado = "No disponible" }
            }
            
        } catch {
            $datos.SmartScreen = @{ Error = "No se pudo obtener información de SmartScreen: $_" }
        }
        
        return $datos
    } catch {
        Write-Warning "Error al obtener datos extendidos: $_"
        return @{ Error = $_.Exception.Message }
    }
}

function Generate-CompleteHTML {
    param($InfoSistema, $MetricasRendimiento, $LogsEventos, $DatosExtendidos, $AnalisisConfiabilidad, $DiagnosticoHardware, $AnalisisRoles, $AnalisisPoliticas, $VerificacionCumplimiento, $AnalisisPermisos, $AuditoriaSoftware)
    
    # Extraer el nombre del servidor y SO directamente
    $nombreServidor = $InfoSistema.NombreServidor
    $sistemaOperativo = $InfoSistema.NombreSO
    
    # Convertir la imagen a Base64 para usar en el HTML
    $ms = New-Object System.IO.MemoryStream
    $imagenl.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $imagenBytes = $ms.ToArray()
    $ms.Close()
    $imagenBase64 = [Convert]::ToBase64String($imagenBytes)
    
    # Crear el CSS como string separado para evitar problemas de parsing
    $cssStyles = @'
        * { box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; padding: 20px; 
            background: linear-gradient(135deg, rgb(102,126,234) 0%, rgb(118,75,162) 100%);
            min-height: 100vh;
        }
        .container { 
            max-width: 1400px; margin: 0 auto; 
            background: white; padding: 30px; 
            border-radius: 15px; 
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        }
        h1 { 
            color: white; text-align: center; margin-bottom: 30px; 
            border-bottom: 4px solid rgb(255,255,255); padding-bottom: 20px;
            font-size: 2.5em; text-shadow: 2px 2px 6px rgba(0,0,0,0.8);
        }
        h2 { 
            color: white; margin: 40px 0 20px 0; padding: 20px; 
            background: linear-gradient(135deg, rgb(52,152,219), rgb(41,128,185)); 
            border-radius: 10px; font-size: 1.5em;
            box-shadow: 0 4px 15px rgba(52, 152, 219, 0.3);
        }
        h3 { 
            color: rgb(44,62,80); margin: 30px 0 15px 0; 
            border-left: 6px solid rgb(52,152,219); padding-left: 20px;
            font-size: 1.3em;
        }
        .metrics-grid { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); 
            gap: 25px; margin: 30px 0; 
        }
        .metric-card { 
            background: linear-gradient(135deg, rgb(248,249,250), rgb(233,236,239)); 
            padding: 25px; border-radius: 12px; 
            border-left: 6px solid rgb(52,152,219);
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s ease;
        }
        .metric-card:hover { transform: translateY(-5px); }
        .metric-card h4 { margin-top: 0; color: rgb(44,62,80); font-size: 1.2em; }
        .metric-value { font-size: 2em; font-weight: bold; color: rgb(52,152,219); margin: 10px 0; }
        .status-indicator { 
            display: inline-block; padding: 5px 15px; 
            border-radius: 20px; color: white; font-weight: bold; 
        }
        .status-normal { background-color: rgb(39,174,96); }
        .status-warning { background-color: rgb(243,156,18); }
        .status-critical { background-color: rgb(231,76,60); }
        .status-high { background-color: rgb(230,126,34); }
        .status-medium { background-color: rgb(241,196,15); color: rgb(44,62,80); }
        
        table { 
            border-collapse: collapse; width: 100%; margin: 20px 0; 
            background: white; border-radius: 10px; overflow: hidden;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        th { 
            background: linear-gradient(135deg, rgb(52,152,219), rgb(41,128,185)); 
            color: white; padding: 15px; text-align: left; font-weight: 600;
        }
        td { padding: 12px 15px; border-bottom: 1px solid rgb(236,240,241); }
        tr:nth-child(even) { background-color: rgb(248,249,250); }
        tr:hover { background-color: rgb(214,234,248); }
        
        .summary-box { 
            background: linear-gradient(135deg, rgb(46,204,113), rgb(39,174,96)); 
            color: white; padding: 25px; margin: 30px 0; 
            border-radius: 12px; text-align: center;
            box-shadow: 0 6px 20px rgba(46, 204, 113, 0.3);
        }
        .summary-box h3 { color: white; border: none; padding: 0; margin-bottom: 15px; }
        
        .warning-box { 
            background: linear-gradient(135deg, rgb(231,76,60), rgb(192,57,43)); 
            color: white; padding: 25px; margin: 30px 0; 
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(231, 76, 60, 0.3);
        }
        .warning-box h3 { color: white; border: none; padding: 0; margin-bottom: 15px; }
        
        .info-box { 
            background: linear-gradient(135deg, rgb(52,152,219), rgb(41,128,185)); 
            color: white; padding: 25px; margin: 30px 0; 
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(52, 152, 219, 0.3);
        }
        .info-box h3 { color: white; border: none; padding: 0; margin-bottom: 15px; }
        
        .header-info { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); 
            gap: 20px; margin: 30px 0; 
        }
        .header-card { 
            background: white; padding: 20px; border-radius: 10px; 
            border-left: 5px solid rgb(52,152,219);
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        .header-card strong { color: rgb(44,62,80); }
        
        .logo { 
            text-align: center; margin-bottom: 30px; 
        }
        .logo img { 
            max-width: 200px; height: auto; 
            filter: drop-shadow(0 4px 8px rgba(0,0,0,0.3));
        }
        
        .footer { 
            text-align: center; margin-top: 50px; padding: 30px; 
            background: linear-gradient(135deg, rgb(44,62,80), rgb(52,73,94)); 
            color: white; border-radius: 10px;
        }
        
        .progress-bar { 
            background-color: rgb(236,240,241); 
            border-radius: 10px; height: 20px; 
            overflow: hidden; margin: 10px 0;
        }
        .progress-fill { 
            height: 100%; 
            background: linear-gradient(90deg, rgb(46,204,113), rgb(39,174,96)); 
            transition: width 0.3s ease;
        }
        
        .tabs { 
            display: flex; 
            background: rgb(236,240,241); 
            border-radius: 10px 10px 0 0; 
            overflow: hidden;
        }
        .tab { 
            flex: 1; padding: 15px; text-align: center; 
            background: rgb(189,195,199); color: rgb(44,62,80); 
            cursor: pointer; transition: background 0.3s;
        }
        .tab.active { 
            background: rgb(52,152,219); color: white; 
        }
        .tab-content { 
            background: white; padding: 30px; 
            border-radius: 0 0 10px 10px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        
        .compliance-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        
        .compliance-card {
            background: white;
            border-radius: 10px;
            padding: 20px;
            border-left: 5px solid;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        
        .compliance-pass { border-left-color: rgb(39,174,96); }
        .compliance-fail { border-left-color: rgb(231,76,60); }
        .compliance-pending { border-left-color: rgb(243,156,18); }
        
        .security-summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }
        
        .security-metric {
            text-align: center;
            padding: 20px;
            background: linear-gradient(135deg, rgb(248,249,250), rgb(233,236,239));
            border-radius: 10px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        
        .security-metric .number {
            font-size: 2.5em;
            font-weight: bold;
            margin: 10px 0;
        }
        
        .risk-high { color: rgb(231,76,60); }
        .risk-medium { color: rgb(243,156,18); }
        .risk-low { color: rgb(39,174,96); }
        
        @media (max-width: 768px) {
            .container { padding: 15px; }
            .metrics-grid { grid-template-columns: 1fr; }
            .header-info { grid-template-columns: 1fr; }
            h1 { font-size: 2em; }
            h2 { font-size: 1.3em; }
        }
'@

    $htmlContent = @"
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Informe de Salud del Sistema - $nombreServidor</title>
    <style>$cssStyles</style>
</head>
<body>
    <div class="container">
        <div class="logo">
            <img src="data:image/png;base64,$imagenBase64" alt="Logo">
        </div>
        
        <h1>🖥️ INFORME COMPLETO DE SALUD DEL SISTEMA</h1>
        
        <div class="summary-box">
            <h3>📊 Resumen Ejecutivo</h3>
            <p><strong>Servidor:</strong> $nombreServidor | <strong>Sistema:</strong> $sistemaOperativo</p>
            <p><strong>Fecha del Informe:</strong> $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")</p>
            <p><strong>Tiempo de Actividad:</strong> $($InfoSistema.TiempoActividad.Days) días, $($InfoSistema.TiempoActividad.Hours) horas</p>
        </div>

        <div class="header-info">
            <div class="header-card">
                <strong>🏢 Información del Sistema</strong><br>
                Fabricante: $($InfoSistema.Fabricante)<br>
                Modelo: $($InfoSistema.Modelo)<br>
                Procesador: $($InfoSistema.Procesador)<br>
                Memoria Total: $($InfoSistema.MemoriaTotal) GB
            </div>
            <div class="header-card">
                <strong>🌐 Configuración de Red</strong><br>
                IP Principal: $($InfoSistema.DireccionIP)<br>
                $($InfoSistema.DominioWorkgroup)<br>
                Zona Horaria: $($InfoSistema.TimeZone)
            </div>
            <div class="header-card">
                <strong>🔧 Detalles Técnicos</strong><br>
                Versión SO: $($InfoSistema.VersionSO)<br>
                Build: $($InfoSistema.BuildNumber)<br>
                Arquitectura: $($InfoSistema.Arquitectura)<br>
                Último Reinicio: $($InfoSistema.UltimoReinicio.ToString("dd/MM/yyyy HH:mm"))
            </div>
        </div>

        <h2>📈 MÉTRICAS DE RENDIMIENTO DE LOS 4 SUBSISTEMAS</h2>
        
        <div class="metrics-grid">
            <div class="metric-card">
                <h4>🔥 Subsistema CPU</h4>
                <div class="metric-value">$($MetricasRendimiento.CPU.UsoCPU)%</div>
                <span class="status-indicator status-$($MetricasRendimiento.CPU.Estado.ToLower())">$($MetricasRendimiento.CPU.Estado)</span>
                <p><strong>Cola del Procesador:</strong> $($MetricasRendimiento.CPU.ColaProcesor)</p>
                <p><strong>Interrupciones/seg:</strong> $($MetricasRendimiento.CPU.Interrupciones)</p>
            </div>
            
            <div class="metric-card">
                <h4>🧠 Subsistema Memoria</h4>
                <div class="metric-value">$($MetricasRendimiento.Memoria.PorcentajeUsado)%</div>
                <span class="status-indicator status-$($MetricasRendimiento.Memoria.Estado.ToLower())">$($MetricasRendimiento.Memoria.Estado)</span>
                <p><strong>Usada:</strong> $($MetricasRendimiento.Memoria.UsadaMB) MB / $($MetricasRendimiento.Memoria.TotalMB) MB</p>
                <p><strong>Páginas/seg:</strong> $($MetricasRendimiento.Memoria.PaginasPorSeg)</p>
                <p><strong>Caché:</strong> $($MetricasRendimiento.Memoria.CacheMB) MB</p>
            </div>
        </div>

        <h3>💾 Subsistema Disco</h3>
        <table>
            <tr>
                <th>Unidad</th>
                <th>Total (GB)</th>
                <th>Libre (GB)</th>
                <th>% Usado</th>
                <th>Sistema Archivos</th>
                <th>Tiempo Lectura (ms)</th>
                <th>Tiempo Escritura (ms)</th>
                <th>Estado</th>
            </tr>
"@

    foreach ($disco in $MetricasRendimiento.Discos) {
        $statusClass = switch ($disco.Estado) {
            "Normal" { "status-normal" }
            "Advertencia" { "status-warning" }
            "Crítico" { "status-critical" }
            default { "status-normal" }
        }
        
        $htmlContent += @"
            <tr>
                <td><strong>$($disco.Unidad)</strong></td>
                <td>$($disco.TotalGB)</td>
                <td>$($disco.LibreGB)</td>
                <td>$($disco.PorcentajeUsado)%</td>
                <td>$($disco.SistemaArchivos)</td>
                <td>$($disco.TiempoLectura)</td>
                <td>$($disco.TiempoEscritura)</td>
                <td><span class="status-indicator $statusClass">$($disco.Estado)</span></td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🌐 Subsistema Red</h3>
        <table>
            <tr>
                <th>Interfaz</th>
                <th>Bytes Enviados (KB)</th>
                <th>Bytes Recibidos (KB)</th>
                <th>Paquetes Enviados</th>
                <th>Paquetes Recibidos</th>
                <th>Errores Envío</th>
                <th>Errores Recepción</th>
            </tr>
"@

    foreach ($interfaz in $MetricasRendimiento.Red) {
        $htmlContent += @"
            <tr>
                <td>$($interfaz.Interfaz)</td>
                <td>$($interfaz.BytesEnviados)</td>
                <td>$($interfaz.BytesRecibidos)</td>
                <td>$($interfaz.PaquetesEnviados)</td>
                <td>$($interfaz.PaquetesRecibidos)</td>
                <td>$($interfaz.ErroresEnvio)</td>
                <td>$($interfaz.ErroresRecepcion)</td>
            </tr>
"@
    }

    # NUEVA SECCIÓN: ANÁLISIS DE CONFIABILIDAD
    if ($AnalisisConfiabilidad -and $AnalisisConfiabilidad.EstadisticasEstabilidad) {
        $htmlContent += @"
        </table>

        <h2>🔍 ANÁLISIS DE CONFIABILIDAD DEL SISTEMA</h2>
        
        <div class="metrics-grid">
            <div class="metric-card">
                <h4>📊 Estadísticas de Estabilidad</h4>
                <p><strong>Índice de Estabilidad:</strong> <span class="status-indicator status-$(if($AnalisisConfiabilidad.EstadisticasEstabilidad.IndiceEstabilidad -eq 'Alta'){'normal'}elseif($AnalisisConfiabilidad.EstadisticasEstabilidad.IndiceEstabilidad -eq 'Media'){'warning'}else{'critical'})">$($AnalisisConfiabilidad.EstadisticasEstabilidad.IndiceEstabilidad)</span></p>
                <p><strong>Fallos de Aplicación:</strong> $($AnalisisConfiabilidad.EstadisticasEstabilidad.FallosAplicacion)</p>
                <p><strong>Fallos del Sistema:</strong> $($AnalisisConfiabilidad.EstadisticasEstabilidad.FallosSistema)</p>
                <p><strong>Reinicios Detectados:</strong> $($AnalisisConfiabilidad.EstadisticasEstabilidad.ReiniciosDetectados)</p>
                <p><strong>Período de Análisis:</strong> $($AnalisisConfiabilidad.EstadisticasEstabilidad.PeriodoAnalisis)</p>
            </div>
        </div>
"@

        if ($AnalisisConfiabilidad.TendenciasSemanales -and $AnalisisConfiabilidad.TendenciasSemanales.Count -gt 0) {
            $htmlContent += @"
        <h3>📈 Tendencias de Estabilidad (Últimos 7 días)</h3>
        <table>
            <tr>
                <th>Fecha</th>
                <th>Total Eventos</th>
                <th>Eventos Críticos</th>
                <th>Estado de Estabilidad</th>
            </tr>
"@
            foreach ($tendencia in $AnalisisConfiabilidad.TendenciasSemanales) {
                $estabilidadClass = switch ($tendencia.Estabilidad) {
                    "Estable" { "status-normal" }
                    "Moderada" { "status-warning" }
                    "Inestable" { "status-critical" }
                    default { "status-normal" }
                }
                
                $htmlContent += @"
            <tr>
                <td>$($tendencia.Fecha)</td>
                <td>$($tendencia.TotalEventos)</td>
                <td>$($tendencia.EventosCriticos)</td>
                <td><span class="status-indicator $estabilidadClass">$($tendencia.Estabilidad)</span></td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }
    }

    # NUEVA SECCIÓN: DIAGNÓSTICO DE HARDWARE AVANZADO
    if ($DiagnosticoHardware) {
        $htmlContent += @"
        <h2>🔧 DIAGNÓSTICO AVANZADO DE HARDWARE</h2>
"@

        # Estado SMART de discos
        if ($DiagnosticoHardware.DiscosSMART -and $DiagnosticoHardware.DiscosSMART.Count -gt 0) {
            $htmlContent += @"
        <h3>💾 Estado SMART de Discos Duros</h3>
        <table>
            <tr>
                <th>Modelo</th>
                <th>Número de Serie</th>
                <th>Tamaño (GB)</th>
                <th>Interfaz</th>
                <th>Estado SMART</th>
                <th>Particiones</th>
                <th>Estado General</th>
            </tr>
"@
            foreach ($disco in $DiagnosticoHardware.DiscosSMART) {
                $estadoClass = switch ($disco.Estado) {
                    "Normal" { "status-normal" }
                    "Crítico" { "status-critical" }
                    default { "status-warning" }
                }
                
                $htmlContent += @"
            <tr>
                <td>$($disco.Modelo)</td>
                <td>$($disco.NumeroSerie)</td>
                <td>$($disco.TamanoGB)</td>
                <td>$($disco.Interfaz)</td>
                <td>$($disco.EstadoSMART)</td>
                <td>$($disco.Particiones)</td>
                <td><span class="status-indicator $estadoClass">$($disco.Estado)</span></td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }

        # Temperaturas
        if ($DiagnosticoHardware.Temperaturas -and $DiagnosticoHardware.Temperaturas.Count -gt 0) {
            $htmlContent += @"
        <h3>🌡️ Temperaturas de Componentes</h3>
        <table>
            <tr>
                <th>Componente</th>
                <th>Temperatura (°C)</th>
                <th>Estado</th>
            </tr>
"@
            foreach ($temp in $DiagnosticoHardware.Temperaturas) {
                $tempClass = switch ($temp.Estado) {
                    "Normal" { "status-normal" }
                    "Alto" { "status-warning" }
                    "Crítico" { "status-critical" }
                    default { "status-warning" }
                }
                
                $tempValue = if ($temp.TemperaturaC) { "$($temp.TemperaturaC)°C" } else { "No disponible" }
                
                $htmlContent += @"
            <tr>
                <td>$($temp.Componente)</td>
                <td>$tempValue</td>
                <td><span class="status-indicator $tempClass">$($temp.Estado)</span></td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }

        # Estado de baterías
        if ($DiagnosticoHardware.Baterias -and $DiagnosticoHardware.Baterias.Count -gt 0) {
            $htmlContent += @"
        <h3>🔋 Estado de Baterías</h3>
        <table>
            <tr>
                <th>Nombre</th>
                <th>Fabricante</th>
                <th>Estado de Carga</th>
                <th>% Carga</th>
                <th>Tiempo Restante</th>
                <th>Salud de Batería</th>
                <th>Estado</th>
            </tr>
"@
            foreach ($bateria in $DiagnosticoHardware.Baterias) {
                if ($bateria.Estado -ne "No se detectaron baterías (sistema de escritorio)") {
                    $bateriaClass = switch ($bateria.Estado) {
                        "Normal" { "status-normal" }
                        "Advertencia" { "status-warning" }
                        "Bajo" { "status-warning" }
                        "Crítico" { "status-critical" }
                        default { "status-normal" }
                    }
                    
                    $htmlContent += @"
            <tr>
                <td>$($bateria.Nombre)</td>
                <td>$($bateria.Fabricante)</td>
                <td>$($bateria.EstadoCarga)</td>
                <td>$($bateria.PorcentajeCarga)%</td>
                <td>$($bateria.TiempoRestante)</td>
                <td>$($bateria.SaludBateria)</td>
                <td><span class="status-indicator $bateriaClass">$($bateria.Estado)</span></td>
            </tr>
"@
                } else {
                    $htmlContent += @"
            <tr>
                <td colspan="7" style="text-align: center; font-style: italic;">$($bateria.Estado)</td>
            </tr>
"@
                }
            }
            $htmlContent += "</table>"
        }
    }

    # NUEVA SECCIÓN: ANÁLISIS DE ROLES DE SERVIDOR
    if ($AnalisisRoles -and $AnalisisRoles.EsServidor) {
        $htmlContent += @"
        <h2>🖥️ ANÁLISIS DE ROLES DE WINDOWS SERVER</h2>
        
        <div class="info-box">
            <h3>📊 Resumen de Roles</h3>
            <p><strong>Tipo de Sistema:</strong> $($AnalisisRoles.TipoSistema)</p>
            <p><strong>Total de Roles Detectados:</strong> $($AnalisisRoles.ResumenRoles.TotalRoles)</p>
            <p><strong>Roles Críticos:</strong> $($AnalisisRoles.ResumenRoles.RolesCriticos)</p>
            <p><strong>Roles Detenidos:</strong> $($AnalisisRoles.ResumenRoles.RolesDetenidos)</p>
            <p><strong>Estado General:</strong> $($AnalisisRoles.ResumenRoles.EstadoGeneral)</p>
        </div>
"@

        if ($AnalisisRoles.RolesDetectados -and $AnalisisRoles.RolesDetectados.Count -gt 0) {
            $htmlContent += @"
        <h3>🔧 Roles y Servicios Detectados</h3>
        <table>
            <tr>
                <th>Rol</th>
                <th>Servicio</th>
                <th>Estado</th>
                <th>Tipo de Inicio</th>
                <th>Descripción</th>
                <th>Crítico</th>
            </tr>
"@
            foreach ($rol in $AnalisisRoles.RolesDetectados) {
                $estadoClass = if ($rol.Estado -eq "Running") { "status-normal" } else { "status-critical" }
                $criticoIcon = if ($rol.Critico) { "⚠️" } else { "ℹ️" }
                
                $htmlContent += @"
            <tr>
                <td><strong>$($rol.Rol)</strong></td>
                <td>$($rol.Servicio)</td>
                <td><span class="status-indicator $estadoClass">$($rol.Estado)</span></td>
                <td>$($rol.TipoInicio)</td>
                <td>$($rol.Descripcion)</td>
                <td>$criticoIcon</td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }

        # Información de dominio
        if ($AnalisisRoles.InformacionDominio) {
            $htmlContent += @"
        <h3>🌐 Información de Dominio</h3>
        <div class="metric-card">
            <p><strong>Parte de Dominio:</strong> $(if($AnalisisRoles.InformacionDominio.ParteDominio){'Sí'}else{'No'})</p>
            <p><strong>Dominio:</strong> $($AnalisisRoles.InformacionDominio.Dominio)</p>
            <p><strong>Rol del Servidor:</strong> $($AnalisisRoles.InformacionDominio.Rol)</p>
        </div>
"@
        }
    } elseif ($AnalisisRoles -and -not $AnalisisRoles.EsServidor) {
        $htmlContent += @"
        <div class="info-box">
            <h3>ℹ️ Información del Sistema</h3>
            <p>Este sistema es un <strong>$($AnalisisRoles.TipoSistema)</strong>, no un Windows Server.</p>
            <p>El análisis de roles de servidor no es aplicable.</p>
        </div>
"@
    }

    # NUEVA SECCIÓN: ANÁLISIS DE POLÍTICAS DE GRUPO
    if ($AnalisisPoliticas) {
        $htmlContent += @"
        <h2>🛡️ ANÁLISIS DE POLÍTICAS DE GRUPO Y SEGURIDAD</h2>
"@

        if ($AnalisisPoliticas.EnDominio) {
            $htmlContent += @"
        <div class="info-box">
            <h3>📋 Información de GPO</h3>
            <p><strong>Sistema en Dominio:</strong> Sí</p>
            <p><strong>Última Actualización GPO:</strong> $($AnalisisPoliticas.UltimaActualizacionGPO)</p>
        </div>
"@

            # GPOs aplicadas al equipo
            if ($AnalisisPoliticas.GPOsEquipo -and $AnalisisPoliticas.GPOsEquipo.Count -gt 0) {
                $htmlContent += @"
        <h3>🖥️ GPOs Aplicadas al Equipo</h3>
        <table>
            <tr>
                <th>Nombre de la GPO</th>
                <th>Tipo</th>
                <th>Estado</th>
            </tr>
"@
                foreach ($gpo in $AnalisisPoliticas.GPOsEquipo) {
                    $htmlContent += @"
            <tr>
                <td>$($gpo.Nombre)</td>
                <td>$($gpo.Tipo)</td>
                <td><span class="status-indicator status-normal">$($gpo.Estado)</span></td>
            </tr>
"@
                }
                $htmlContent += "</table>"
            }

            # GPOs aplicadas al usuario
            if ($AnalisisPoliticas.GPOsUsuario -and $AnalisisPoliticas.GPOsUsuario.Count -gt 0) {
                $htmlContent += @"
        <h3>👤 GPOs Aplicadas al Usuario</h3>
        <table>
            <tr>
                <th>Nombre de la GPO</th>
                <th>Tipo</th>
                <th>Estado</th>
            </tr>
"@
                foreach ($gpo in $AnalisisPoliticas.GPOsUsuario) {
                    $htmlContent += @"
            <tr>
                <td>$($gpo.Nombre)</td>
                <td>$($gpo.Tipo)</td>
                <td><span class="status-indicator status-normal">$($gpo.Estado)</span></td>
            </tr>
"@
                }
                $htmlContent += "</table>"
            }
        } else {
            $htmlContent += @"
        <div class="warning-box">
            <h3>⚠️ Sistema No Unido a Dominio</h3>
            <p>Este sistema no está unido a un dominio Active Directory.</p>
            <p>Las políticas de grupo de dominio no son aplicables.</p>
        </div>
"@
        }

        # Políticas locales
        if ($AnalisisPoliticas.PoliticasLocales -and $AnalisisPoliticas.PoliticasLocales.Count -gt 0) {
            $htmlContent += @"
        <h3>🔒 Políticas de Seguridad Locales</h3>
        <table>
            <tr>
                <th>Política</th>
                <th>Valor</th>
                <th>Categoría</th>
            </tr>
"@
            foreach ($politica in $AnalisisPoliticas.PoliticasLocales | Select-Object -First 20) {
                $htmlContent += @"
            <tr>
                <td>$($politica.Politica)</td>
                <td>$($politica.Valor)</td>
                <td>$($politica.Categoria)</td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }
    }

    # NUEVA SECCIÓN: VERIFICACIÓN DE CUMPLIMIENTO
    if ($VerificacionCumplimiento -and $VerificacionCumplimiento.ResumenCumplimiento) {
        $htmlContent += @"
        <h2>✅ VERIFICACIÓN DE CUMPLIMIENTO (CIS BENCHMARKS)</h2>
        
        <div class="security-summary">
            <div class="security-metric">
                <div class="number risk-low">$($VerificacionCumplimiento.ResumenCumplimiento.PorcentajeCumplimiento)%</div>
                <div>Cumplimiento General</div>
            </div>
            <div class="security-metric">
                <div class="number risk-low">$($VerificacionCumplimiento.ResumenCumplimiento.Cumple)</div>
                <div>Verificaciones Exitosas</div>
            </div>
            <div class="security-metric">
                <div class="number risk-high">$($VerificacionCumplimiento.ResumenCumplimiento.NoCumple)</div>
                <div>No Cumple</div>
            </div>
            <div class="security-metric">
                <div class="number risk-medium">$($VerificacionCumplimiento.ResumenCumplimiento.PendienteRevision)</div>
                <div>Pendiente Revisión</div>
            </div>
        </div>
        
        <div class="$(if($VerificacionCumplimiento.ResumenCumplimiento.NivelCumplimiento -eq 'Excelente' -or $VerificacionCumplimiento.ResumenCumplimiento.NivelCumplimiento -eq 'Bueno'){'summary-box'}elseif($VerificacionCumplimiento.ResumenCumplimiento.NivelCumplimiento -eq 'Aceptable'){'info-box'}else{'warning-box'})">
            <h3>📊 Nivel de Cumplimiento: $($VerificacionCumplimiento.ResumenCumplimiento.NivelCumplimiento)</h3>
            <p>Total de verificaciones realizadas: $($VerificacionCumplimiento.ResumenCumplimiento.TotalVerificaciones)</p>
        </div>
"@

        if ($VerificacionCumplimiento.Verificaciones -and $VerificacionCumplimiento.Verificaciones.Count -gt 0) {
            # Mostrar solo las verificaciones que no cumplen o están pendientes
            $verificacionesProblematicas = $VerificacionCumplimiento.Verificaciones | Where-Object { $_.Cumple -ne "Sí" }
            
            if ($verificacionesProblematicas.Count -gt 0) {
                $htmlContent += @"
        <h3>⚠️ Verificaciones que Requieren Atención</h3>
        <div class="compliance-grid">
"@
                foreach ($verificacion in $verificacionesProblematicas) {
                    $cardClass = switch ($verificacion.Cumple) {
                        "No" { "compliance-fail" }
                        "Pendiente" { "compliance-pending" }
                        "Revisar" { "compliance-pending" }
                        default { "compliance-pending" }
                    }
                    
                    $criticidadClass = switch ($verificacion.Criticidad) {
                        "Alta" { "risk-high" }
                        "Media" { "risk-medium" }
                        "Baja" { "risk-low" }
                        default { "risk-medium" }
                    }
                    
                    $htmlContent += @"
            <div class="compliance-card $cardClass">
                <h4>$($verificacion.ID)</h4>
                <p><strong>$($verificacion.Descripcion)</strong></p>
                <p><strong>Estado Actual:</strong> $($verificacion.EstadoActual)</p>
                <p><strong>Recomendación:</strong> $($verificacion.Recomendacion)</p>
                <p><strong>Criticidad:</strong> <span class="$criticidadClass">$($verificacion.Criticidad)</span></p>
                <p><strong>Categoría:</strong> $($verificacion.Categoria)</p>
            </div>
"@
                }
                $htmlContent += "</div>"
            }
        }
    }

    # NUEVA SECCIÓN: ANÁLISIS DE PERMISOS
    if ($AnalisisPermisos -and $AnalisisPermisos.ResumenPermisos) {
        $htmlContent += @"
        <h2>🔐 ANÁLISIS DE PERMISOS DE CARPETAS SENSIBLES</h2>
        
        <div class="security-summary">
            <div class="security-metric">
                <div class="number risk-low">$($AnalisisPermisos.ResumenPermisos.PorcentajeSeguras)%</div>
                <div>Carpetas Seguras</div>
            </div>
            <div class="security-metric">
                <div class="number risk-medium">$($AnalisisPermisos.ResumenPermisos.CarpetasConProblemas)</div>
                <div>Con Problemas</div>
            </div>
            <div class="security-metric">
                <div class="number risk-high">$($AnalisisPermisos.ResumenPermisos.CarpetasCriticas)</div>
                <div>Críticas</div>
            </div>
            <div class="security-metric">
                <div class="number">$($AnalisisPermisos.ResumenPermisos.TotalCarpetasAnalizadas)</div>
                <div>Total Analizadas</div>
            </div>
        </div>
        
        <div class="$(if($AnalisisPermisos.ResumenPermisos.NivelSeguridad -eq 'Excelente'){'summary-box'}elseif($AnalisisPermisos.ResumenPermisos.NivelSeguridad -eq 'Bueno' -or $AnalisisPermisos.ResumenPermisos.NivelSeguridad -eq 'Aceptable'){'info-box'}else{'warning-box'})">
            <h3>🛡️ Nivel de Seguridad: $($AnalisisPermisos.ResumenPermisos.NivelSeguridad)</h3>
        </div>
"@

        if ($AnalisisPermisos.AnalisisPermisos) {
            $carpetasProblematicas = $AnalisisPermisos.AnalisisPermisos | Where-Object { $_.Estado -in @("Advertencia", "Crítico") }
            
            if ($carpetasProblematicas.Count -gt 0) {
                $htmlContent += @"
        <h3>⚠️ Carpetas con Permisos Problemáticos</h3>
        <table>
            <tr>
                <th>Ruta</th>
                <th>Descripción</th>
                <th>Propietario</th>
                <th>Permisos Problemáticos</th>
                <th>Estado</th>
            </tr>
"@
                foreach ($carpeta in $carpetasProblematicas) {
                    $estadoClass = switch ($carpeta.Estado) {
                        "Crítico" { "status-critical" }
                        "Advertencia" { "status-warning" }
                        default { "status-normal" }
                    }
                    
                    $htmlContent += @"
            <tr>
                <td><strong>$($carpeta.Ruta)</strong></td>
                <td>$($carpeta.Descripcion)</td>
                <td>$($carpeta.Propietario)</td>
                <td>$($carpeta.PermisosProblematicosCount)</td>
                <td><span class="status-indicator $estadoClass">$($carpeta.Estado)</span></td>
            </tr>
"@
                }
                $htmlContent += "</table>"
            }
        }
    }

    # NUEVA SECCIÓN: AUDITORÍA DE SOFTWARE
    if ($AuditoriaSoftware -and $AuditoriaSoftware.ResumenAuditoria) {
        $htmlContent += @"
        <h2>💿 AUDITORÍA DE SOFTWARE INSTALADO</h2>
        
        <div class="security-summary">
            <div class="security-metric">
                <div class="number">$($AuditoriaSoftware.ResumenAuditoria.TotalSoftwareInstalado)</div>
                <div>Total Programas</div>
            </div>
            <div class="security-metric">
                <div class="number risk-high">$($AuditoriaSoftware.ResumenAuditoria.SoftwareCritico)</div>
                <div>Riesgo Alto</div>
            </div>
            <div class="security-metric">
                <div class="number risk-medium">$($AuditoriaSoftware.ResumenAuditoria.SoftwareRiesgoMedio)</div>
                <div>Riesgo Medio</div>
            </div>
            <div class="security-metric">
                <div class="number risk-low">$($AuditoriaSoftware.ResumenAuditoria.PorcentajeSeguro)%</div>
                <div>Software Seguro</div>
            </div>
        </div>
        
        <div class="$(if($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo -eq 'Bajo'){'summary-box'}elseif($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo -eq 'Medio'){'info-box'}else{'warning-box'})">
            <h3>⚠️ Nivel de Riesgo: $($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo)</h3>
        </div>
        
"@
# NUEVA SECCIÓN: SOFTWARE INSTALADO
if ($AuditoriaSoftware -and $AuditoriaSoftware.SoftwareInstalado) {
    $htmlContent += @"
    <h2>📦 INVENTARIO DE SOFTWARE INSTALADO</h2>
    
    <div class="metrics-grid">
        <div class="metric-card">
            <h4>📊 Resumen de Auditoría de Software</h4>
            <p><strong>Total Software:</strong> $($AuditoriaSoftware.ResumenAuditoria.TotalSoftwareInstalado)</p>
            <p><strong>Software Problemático:</strong> $($AuditoriaSoftware.ResumenAuditoria.SoftwareConProblemas)</p>
            <p><strong>Nivel de Riesgo:</strong> <span class="status-indicator status-$(if($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo -eq 'Bajo'){'normal'}elseif($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo -eq 'Medio'){'warning'}else{'critical'})">$($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo)</span></p>
            <p><strong>Software Crítico:</strong> $($AuditoriaSoftware.ResumenAuditoria.SoftwareCritico)</p>
            <p><strong>Software Riesgo Medio:</strong> $($AuditoriaSoftware.ResumenAuditoria.SoftwareRiesgoMedio)</p>
        </div>
    </div>
"@

    # Software problemático
    if ($AuditoriaSoftware.SoftwareProblematico -and $AuditoriaSoftware.SoftwareProblematico.Count -gt 0) {
        $htmlContent += @"
    <h3>⚠️ Software Problemático Detectado</h3>
    <table>
        <tr>
            <th>Software</th>
            <th>Versión</th>
            <th>Fabricante</th>
            <th>Instalado</th>
            <th>Criticidad</th>
            <th>Problema</th>
        </tr>
"@

        foreach ($software in $AuditoriaSoftware.SoftwareProblematico) {
            $criticidadClass = switch ($software.CriticidadMaxima) {
                "Alta" { "status-critical" }
                "Media" { "status-warning" }
                "Baja" { "status-normal" }
                default { "" }
            }
            
            $descripcionProblemas = ""
            foreach ($problema in $software.Problemas) {
                $descripcionProblemas += "$($problema.Descripcion)<br>"
            }
            
            $htmlContent += @"
        <tr>
            <td><strong>$($software.Nombre)</strong></td>
            <td>$($software.Version)</td>
            <td>$($software.Fabricante)</td>
            <td>$($software.FechaInstalacion)</td>
            <td><span class="status-indicator $criticidadClass">$($software.CriticidadMaxima)</span></td>
            <td>$descripcionProblemas</td>
        </tr>
"@
        }
        $htmlContent += "</table>"
    }

    # Lista completa de software
    $htmlContent += @"
    <h3>📋 Lista Completa de Software Instalado</h3>
    <table>
        <tr>
            <th>Nombre</th>
            <th>Versión</th>
            <th>Fabricante</th>
            <th>Fecha Instalación</th>
            <th>Tamaño (MB)</th>
        </tr>
"@

    foreach ($software in ($AuditoriaSoftware.SoftwareInstalado | Sort-Object Nombre)) {
        $htmlContent += @"
        <tr>
            <td>$($software.Nombre)</td>
            <td>$($software.Version)</td>
            <td>$($software.Fabricante)</td>
            <td>$($software.FechaInstalacion)</td>
            <td>$($software.TamañoMB)</td>
        </tr>
"@
    }
    $htmlContent += "</table>"
}
        if ($AuditoriaSoftware.SoftwareProblematico -and $AuditoriaSoftware.SoftwareProblematico.Count -gt 0) {
            $htmlContent += @"
        <h3>⚠️ Software que Requiere Atención</h3>
        <table>
            <tr>
                <th>Nombre</th>
                <th>Versión</th>
                <th>Fabricante</th>
                <th>Fecha Instalación</th>
                <th>Criticidad</th>
                <th>Problemas Detectados</th>
            </tr>
"@
            foreach ($software in $AuditoriaSoftware.SoftwareProblematico | Select-Object -First 20) {
                $criticidadClass = switch ($software.CriticidadMaxima) {
                    "Alta" { "status-critical" }
                    "Media" { "status-warning" }
                    "Baja" { "status-normal" }
                    default { "status-normal" }
                }
                
                $problemasTexto = ($software.Problemas | ForEach-Object { $_.Tipo }) -join ", "
                
                $htmlContent += @"
            <tr>
                <td><strong>$($software.Nombre)</strong></td>
                <td>$($software.Version)</td>
                <td>$($software.Fabricante)</td>
                <td>$($software.FechaInstalacion)</td>
                <td><span class="status-indicator $criticidadClass">$($software.CriticidadMaxima)</span></td>
                <td>$problemasTexto</td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }
    }

    $htmlContent += @"
        <h2>📊 LOGS DE EVENTOS DE LOS 3 TIPOS PRINCIPALES</h2>
        
        <h3>🔴 Logs del Sistema (Últimos eventos críticos)</h3>
        <table>
            <tr>
                <th>Fecha/Hora</th>
                <th>Nivel</th>
                <th>Proveedor</th>
                <th>ID Evento</th>
                <th>Mensaje</th>
            </tr>
"@

    foreach ($log in $LogsEventos.LogsSistema | Select-Object -First 10) {
        $nivelClass = switch ($log.LevelDisplayName) {
            "Critical" { "status-critical" }
            "Error" { "status-critical" }
            "Warning" { "status-warning" }
            default { "status-normal" }
        }
        
        $htmlContent += @"
            <tr>
                <td>$($log.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss"))</td>
                <td><span class="status-indicator $nivelClass">$($log.LevelDisplayName)</span></td>
                <td>$($log.ProviderName)</td>
                <td>$($log.Id)</td>
                <td>$($log.Message)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>📱 Logs de Aplicación (Últimos eventos críticos)</h3>
        <table>
            <tr>
                <th>Fecha/Hora</th>
                <th>Nivel</th>
                <th>Proveedor</th>
                <th>ID Evento</th>
                <th>Mensaje</th>
            </tr>
"@

    foreach ($log in $LogsEventos.LogsAplicacion | Select-Object -First 10) {
        $nivelClass = switch ($log.LevelDisplayName) {
            "Critical" { "status-critical" }
            "Error" { "status-critical" }
            "Warning" { "status-warning" }
            default { "status-normal" }
        }
        
        $htmlContent += @"
            <tr>
                <td>$($log.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss"))</td>
                <td><span class="status-indicator $nivelClass">$($log.LevelDisplayName)</span></td>
                <td>$($log.ProviderName)</td>
                <td>$($log.Id)</td>
                <td>$($log.Message)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🔒 Logs de Seguridad (Eventos de autenticación)</h3>
        <table>
            <tr>
                <th>Fecha/Hora</th>
                <th>ID Evento</th>
                <th>Cuenta</th>
                <th>IP Origen</th>
                <th>Tipo de Evento</th>
                <th>Mensaje</th>
            </tr>
"@

    foreach ($log in $LogsEventos.LogsSeguridad | Select-Object -First 15) {
        $tipoClass = switch ($log.EventType) {
            "Login Fallido" { "status-critical" }
            "Cuenta Bloqueada" { "status-critical" }
            "Login" { "status-normal" }
            "Logout" { "status-normal" }
            default { "status-warning" }
        }
        
        $htmlContent += @"
            <tr>
                <td>$($log.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss"))</td>
                <td>$($log.Id)</td>
                <td>$($log.Account)</td>
                <td>$($log.SourceIP)</td>
                <td><span class="status-indicator $tipoClass">$($log.EventType)</span></td>
                <td>$($log.RawMessage)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h2>🔍 DATOS EXTENDIDOS Y ANÁLISIS AVANZADO</h2>
        
        <h3>🔄 Últimos Parches Instalados</h3>
        <table>
            <tr>
                <th>HotFix ID</th>
                <th>Descripción</th>
                <th>Fecha Instalación</th>
                <th>Instalado Por</th>
            </tr>
"@

    foreach ($parche in $DatosExtendidos.UltimosParches) {
        $fechaInstalacion = if ($parche.InstalledOn) { $parche.InstalledOn.ToString("dd/MM/yyyy") } else { "No disponible" }
        $htmlContent += @"
            <tr>
                <td><strong>$($parche.HotFixID)</strong></td>
                <td>$($parche.Description)</td>
                <td>$fechaInstalacion</td>
                <td>$($parche.InstalledBy)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>⚠️ Servicios Automáticos Detenidos</h3>
        <table>
            <tr>
                <th>Nombre del Servicio</th>
                <th>Nombre Interno</th>
                <th>Estado</th>
                <th>Tipo de Inicio</th>
"@

    if ($RevisarServicioTerceros) {
        $htmlContent += "<th>Compañía</th><th>Tipo</th>"
    }

    $htmlContent += "</tr>"

    foreach ($servicio in $DatosExtendidos.ServiciosDetenidos) {
        $htmlContent += @"
            <tr>
                <td><strong>$($servicio.DisplayName)</strong></td>
                <td>$($servicio.Name)</td>
                <td><span class="status-indicator status-critical">$($servicio.Status)</span></td>
                <td>$($servicio.StartType)</td>
"@
        
        if ($RevisarServicioTerceros) {
            $tipoClass = if ($servicio.EsMicrosoft) { "status-normal" } else { "status-warning" }
            $htmlContent += @"
                <td>$($servicio.Compania)</td>
                <td><span class="status-indicator $tipoClass">$($servicio.Tipo)</span></td>
"@
        }
        
        $htmlContent += "</tr>"
    }

    $htmlContent += @"
        </table>

        <h3>🔥 TOP 5 Procesos por Uso de CPU</h3>
        <table>
            <tr>
                <th>Proceso</th>
                <th>CPU (segundos)</th>
                <th>ID Proceso</th>
                <th>Memoria (MB)</th>
            </tr>
"@

    foreach ($proceso in $DatosExtendidos.ProcesosCPU) {
        $htmlContent += @"
            <tr>
                <td><strong>$($proceso.Name)</strong></td>
                <td>$($proceso.CPU)</td>
                <td>$($proceso.Id)</td>
                <td>$($proceso.Memoria_MB)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🧠 TOP 5 Procesos por Uso de Memoria</h3>
        <table>
            <tr>
                <th>Proceso</th>
                <th>Memoria (MB)</th>
                <th>ID Proceso</th>
                <th>CPU (segundos)</th>
            </tr>
"@

    foreach ($proceso in $DatosExtendidos.ProcesosMemoria) {
        $htmlContent += @"
            <tr>
                <td><strong>$($proceso.Name)</strong></td>
                <td>$($proceso.Memoria_MB)</td>
                <td>$($proceso.Id)</td>
                <td>$($proceso.CPU)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🌐 Puertos Abiertos (Listening)</h3>
        <table>
            <tr>
                <th>Dirección Local</th>
                <th>Puerto</th>
                <th>Estado</th>
                <th>Proceso ID</th>
            </tr>
"@

    foreach ($puerto in $DatosExtendidos.PuertosAbiertos | Select-Object -First 20) {
        $htmlContent += @"
            <tr>
                <td>$($puerto.LocalAddress)</td>
                <td><strong>$($puerto.LocalPort)</strong></td>
                <td><span class="status-indicator status-normal">$($puerto.State)</span></td>
                <td>$($puerto.OwningProcess)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🔗 TOP 10 Conexiones Activas por IP</h3>
        <table>
            <tr>
                <th>IP Remota</th>
                <th>Número de Conexiones</th>
            </tr>
"@

    foreach ($conexion in $DatosExtendidos.ConexionesActivas) {
        $htmlContent += @"
            <tr>
                <td><strong>$($conexion.Name)</strong></td>
                <td>$($conexion.Count)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>👥 Usuarios Locales</h3>
        <table>
            <tr>
                <th>Nombre</th>
                <th>Habilitado</th>
                <th>Último Logon</th>
                <th>Contraseña Expira</th>
            </tr>
"@

    foreach ($usuario in $DatosExtendidos.UsuariosLocales) {
        $habilitadoClass = if ($usuario.Enabled) { "status-normal" } else { "status-warning" }
        $htmlContent += @"
            <tr>
                <td><strong>$($usuario.Name)</strong></td>
                <td><span class="status-indicator $habilitadoClass">$(if($usuario.Enabled){'Sí'}else{'No'})</span></td>
                <td>$($usuario.UltimoLogon)</td>
                <td>$($usuario.PasswordExpires)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🚫 Últimos 5 Intentos de Login Fallidos</h3>
        <table>
            <tr>
                <th>Fecha/Hora</th>
                <th>Cuenta</th>
                <th>IP Origen</th>
                <th>Tipo de Fallo</th>
            </tr>
"@

    if ($DatosExtendidos.LoginsFallidos -and $DatosExtendidos.LoginsFallidos.Count -gt 0) {
        foreach ($login in $DatosExtendidos.LoginsFallidos) {
            $htmlContent += @"
            <tr>
                <td>$($login.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss"))</td>
                <td><strong>$($login.Cuenta)</strong></td>
                <td>$($login.IPOrigen)</td>
                <td>$($login.TipoFallo)</td>
            </tr>
"@
        }
    } else {
        $htmlContent += @"
            <tr>
                <td colspan="4" style="text-align: center; font-style: italic;">No se encontraron intentos de login fallidos recientes</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>⏰ Tareas Programadas (Estado)</h3>
        <table>
            <tr>
                <th>Nombre de Tarea</th>
                <th>Última Ejecución</th>
                <th>Resultado</th>
                <th>Estado</th>
            </tr>
"@

    foreach ($tarea in $DatosExtendidos.TareasProgramadas | Select-Object -First 15) {
        $estadoClass = switch ($tarea.Estado) {
            "Exitoso" { "status-normal" }
            "Error*" { "status-critical" }
            default { "status-warning" }
        }
        
        $ultimaEjecucion = if ($tarea.LastRunTime) { $tarea.LastRunTime.ToString("dd/MM/yyyy HH:mm") } else { "Nunca" }
        
        $htmlContent += @"
            <tr>
                <td><strong>$($tarea.TaskName)</strong></td>
                <td>$ultimaEjecucion</td>
                <td>$($tarea.LastTaskResult)</td>
                <td><span class="status-indicator $estadoClass">$($tarea.Estado)</span></td>
            </tr>
"@
    }

    # SECCIÓN DE SEGURIDAD EXTENDIDA
    $htmlContent += @"
        </table>

        <h2>🛡️ ANÁLISIS DE SEGURIDAD EXTENDIDO</h2>
        
        <h3>🔍 Windows Defender</h3>
        <div class="metrics-grid">
            <div class="metric-card">
                <h4>Estado de Windows Defender</h4>
"@

    if ($DatosExtendidos.WindowsDefender -and -not $DatosExtendidos.WindowsDefender.Error) {
        $defenderStatus = if ($DatosExtendidos.WindowsDefender.Habilitado) { "status-normal" } else { "status-critical" }
        $realtimeStatus = if ($DatosExtendidos.WindowsDefender.TiempoRealActivo) { "status-normal" } else { "status-critical" }
        
        $htmlContent += @"
                <p><strong>Antivirus Habilitado:</strong> <span class="status-indicator $defenderStatus">$(if($DatosExtendidos.WindowsDefender.Habilitado){'Sí'}else{'No'})</span></p>
                <p><strong>Protección en Tiempo Real:</strong> <span class="status-indicator $realtimeStatus">$(if($DatosExtendidos.WindowsDefender.TiempoRealActivo){'Activa'}else{'Inactiva'})</span></p>
                <p><strong>Último Análisis Rápido:</strong> $($DatosExtendidos.WindowsDefender.UltimoAnalisisRapido)</p>
                <p><strong>Último Análisis Completo:</strong> $($DatosExtendidos.WindowsDefender.UltimoAnalisisCompleto)</p>
                <p><strong>Edad de Definiciones:</strong> $($DatosExtendidos.WindowsDefender.EdadDefinicionesAV)</p>
"@
    } else {
        $htmlContent += @"
                <p><span class="status-indicator status-warning">Estado no disponible</span></p>
                <p>$($DatosExtendidos.WindowsDefender.Error)</p>
"@
    }

    $htmlContent += @"
            </div>
        </div>

        <h3>🔥 Estado del Firewall por Perfil</h3>
        <table>
            <tr>
                <th>Perfil</th>
                <th>Habilitado</th>
                <th>Conexiones Entrantes</th>
                <th>Conexiones Salientes</th>
                <th>Notificaciones</th>
            </tr>
"@

    if ($DatosExtendidos.Firewall -and $DatosExtendidos.Firewall.Count -gt 0) {
        foreach ($perfil in $DatosExtendidos.Firewall) {
            $habilitadoClass = if ($perfil.Habilitado) { "status-normal" } else { "status-critical" }
            
            $htmlContent += @"
            <tr>
                <td><strong>$($perfil.Nombre)</strong></td>
                <td><span class="status-indicator $habilitadoClass">$(if($perfil.Habilitado){'Sí'}else{'No'})</span></td>
                <td>$($perfil.ConexionesEntrantes)</td>
                <td>$($perfil.ConexionesSalientes)</td>
                <td>$(if($perfil.NotificacionesActivas){'Activas'}else{'Inactivas'})</td>
            </tr>
"@
        }
    } else {
        $htmlContent += @"
            <tr>
                <td colspan="5" style="text-align: center; font-style: italic;">Información del firewall no disponible</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🔒 Control de Acceso de Usuario (UAC)</h3>
        <div class="metric-card">
"@

    if ($DatosExtendidos.UAC -and -not $DatosExtendidos.UAC.Error) {
        $uacStatus = if ($DatosExtendidos.UAC.Habilitado) { "status-normal" } else { "status-critical" }
        
        $htmlContent += @"
            <p><strong>UAC Habilitado:</strong> <span class="status-indicator $uacStatus">$(if($DatosExtendidos.UAC.Habilitado){'Sí'}else{'No'})</span></p>
            <p><strong>Nivel de UAC:</strong> $($DatosExtendidos.UAC.Nivel)</p>
            <p><strong>Elevación No Segura:</strong> $(if($DatosExtendidos.UAC.ElevacionNoSegura){'Sí'}else{'No'})</p>
"@
    } else {
        $htmlContent += @"
            <p><span class="status-indicator status-warning">Información de UAC no disponible</span></p>
"@
    }

    $htmlContent += @"
        </div>

        <h3>📡 Configuración SNMP</h3>
        <div class="metric-card">
"@

    if ($DatosExtendidos.ServicioSNMP) {
        $snmpStatus = if ($DatosExtendidos.ServicioSNMP.Status -eq "Running") { "status-normal" } else { "status-warning" }
        
        $htmlContent += @"
            <p><strong>Servicio SNMP:</strong> <span class="status-indicator $snmpStatus">$($DatosExtendidos.ServicioSNMP.Status)</span></p>
            <p><strong>Tipo de Inicio:</strong> $($DatosExtendidos.ServicioSNMP.StartType)</p>
"@
    } else {
        $htmlContent += @"
            <p><span class="status-indicator status-normal">Servicio SNMP no instalado</span></p>
"@
    }

    $htmlContent += @"
        </div>

        <div class="footer">
            <h3>📋 RESUMEN DEL INFORME</h3>
            <p><strong>Servidor Analizado:</strong> $nombreServidor</p>
            <p><strong>Sistema Operativo:</strong> $sistemaOperativo</p>
            <p><strong>Fecha y Hora del Informe:</strong> $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")</p>
            <p><strong>Generado por:</strong> Script de Salud del Sistema v3.0 - Análisis Completo</p>
            <hr style="margin: 20px 0; border: 1px solid rgba(255,255,255,0.3);">
            <p style="font-size: 0.9em; opacity: 0.8;">
                Este informe incluye análisis de los 4 subsistemas principales (CPU, Memoria, Disco, Red), 
                logs de eventos de seguridad, análisis de confiabilidad, diagnóstico avanzado de hardware,
                verificación de cumplimiento con CIS Benchmarks, análisis de permisos de carpetas sensibles,
                auditoría de software instalado, y análisis completo de seguridad del sistema.
            </p>
        </div>
    </div>

    <script>
        // Agregar interactividad básica
        document.addEventListener('DOMContentLoaded', function() {
            // Efecto hover en las tarjetas
            const cards = document.querySelectorAll('.metric-card, .header-card, .compliance-card');
            cards.forEach(card => {
                card.addEventListener('mouseenter', function() {
                    this.style.transform = 'translateY(-5px)';
                    this.style.transition = 'transform 0.3s ease';
                });
                card.addEventListener('mouseleave', function() {
                    this.style.transform = 'translateY(0)';
                });
            });

            // Resaltar filas de tabla al hacer hover
            const rows = document.querySelectorAll('tr');
            rows.forEach(row => {
                row.addEventListener('mouseenter', function() {
                    if (this.parentElement.tagName === 'TBODY' || this.parentElement.parentElement.tagName === 'TABLE') {
                        this.style.backgroundColor = '#e3f2fd';
                        this.style.transition = 'background-color 0.2s ease';
                    }
                });
                row.addEventListener('mouseleave', function() {
                    if (this.parentElement.tagName === 'TBODY' || this.parentElement.parentElement.tagName === 'TABLE') {
                        this.style.backgroundColor = '';
                    }
                });
            });

            // Animación de aparición para las secciones
            const sections = document.querySelectorAll('h2, h3');
            const observer = new IntersectionObserver((entries) => {
                entries.forEach(entry => {
                    if (entry.isIntersecting) {
                        entry.target.style.opacity = '1';
                        entry.target.style.transform = 'translateY(0)';
                    }
                });
            });

            sections.forEach(section => {
                section.style.opacity = '0';
                section.style.transform = 'translateY(20px)';
                section.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
                observer.observe(section);
            });
        });
    </script>
</body>
</html>
"@

    return $htmlContent
}

# FUNCIÓN PRINCIPAL MEJORADA
function Main {
    try {
        Write-Host "`n🚀 INICIANDO ANÁLISIS COMPLETO DE SALUD DEL SISTEMA" -ForegroundColor Green
        Write-Host "=" * 80 -ForegroundColor Green
        Write-Host "Servidor: $NombreServidor" -ForegroundColor Cyan
        Write-Host "Fecha: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Cyan
        Write-Host "=" * 80 -ForegroundColor Green
        
        # 1. Información básica del sistema
        Write-Host "`n📋 FASE 1: Recopilando información del sistema..." -ForegroundColor Yellow
        $infoSistema = Get-InformacionSistema
        
        # 2. Métricas de rendimiento de los 4 subsistemas
        Write-Host "`n📊 FASE 2: Analizando métricas de rendimiento..." -ForegroundColor Yellow
        $metricasRendimiento = Get-MetricasRendimiento
        
        # 3. Logs de eventos
        Write-Host "`n📝 FASE 3: Recopilando logs de eventos..." -ForegroundColor Yellow
        $logsEventos = Get-LogsEventos -Dias $DiasLogs
        
        # 4. Datos extendidos
        Write-Host "`n🔍 FASE 4: Recopilando datos extendidos..." -ForegroundColor Yellow
        $datosExtendidos = Get-DatosExtendidos -ParchesFaltantes $ParchesFaltantes -RevisarServicioTerceros $RevisarServicioTerceros
        
        # 5. NUEVAS FUNCIONALIDADES DE SEGURIDAD
        $analisisConfiabilidad = $null
        $diagnosticoHardware = $null
        $analisisRoles = $null
        $analisisPoliticas = $null
        $verificacionCumplimiento = $null
        $analisisPermisos = $null
        $auditoriaSoftware = $null
        
        if ($AnalisisSeguridad) {
            Write-Host "`n🛡️ FASE 5: Ejecutando análisis de seguridad avanzado..." -ForegroundColor Yellow
            
            # 5.1 Análisis de confiabilidad
            Write-Host "   5.1 Análisis de confiabilidad..." -ForegroundColor Cyan
            $analisisConfiabilidad = Get-AnalisisConfiabilidad
            
            # 5.2 Diagnóstico de hardware avanzado
            Write-Host "   5.2 Diagnóstico de hardware avanzado..." -ForegroundColor Cyan
            $diagnosticoHardware = Get-DiagnosticoHardwareAvanzado
            
            # 5.3 Análisis de roles de servidor
            Write-Host "   5.3 Análisis de roles de servidor..." -ForegroundColor Cyan
            $analisisRoles = Get-AnalisisRolesServidor
            
            # 5.4 Análisis de políticas de grupo
            Write-Host "   5.4 Análisis de políticas de grupo..." -ForegroundColor Cyan
            $analisisPoliticas = Get-AnalisisPoliticasGrupo
            
            # 5.5 Análisis de permisos
            Write-Host "   5.5 Análisis de permisos de carpetas..." -ForegroundColor Cyan
            $analisisPermisos = Get-AnalisisPermisos
            
            # 5.6 Auditoría de software
            Write-Host "   5.6 Auditoría de software instalado..." -ForegroundColor Cyan
            $auditoriaSoftware = Get-AuditoriaSoftware
        }
        
        if ($VerificarCumplimiento) {
            Write-Host "`n✅ FASE 6: Verificando cumplimiento con estándares..." -ForegroundColor Yellow
            $verificacionCumplimiento = Get-VerificacionCumplimiento
        }
        
        # Generar informe
        Write-Host "`n📄 FASE FINAL: Generando informe completo..." -ForegroundColor Yellow
        Write-Progress -Activity "Generando informe HTML" -PercentComplete 90
        
        $htmlContent = Generate-CompleteHTML -InfoSistema $infoSistema -MetricasRendimiento $metricasRendimiento -LogsEventos $logsEventos -DatosExtendidos $datosExtendidos -AnalisisConfiabilidad $analisisConfiabilidad -DiagnosticoHardware $diagnosticoHardware -AnalisisRoles $analisisRoles -AnalisisPoliticas $analisisPoliticas -VerificacionCumplimiento $verificacionCumplimiento -AnalisisPermisos $analisisPermisos -AuditoriaSoftware $auditoriaSoftware
        
        # Guardar archivo
        $archivoHTML = "$ArchivoSalida.html"
        $htmlContent | Out-File -FilePath $archivoHTML -Encoding UTF8
        
        Write-Progress -Activity "Completado" -PercentComplete 100
        
        # Resumen final
        Write-Host "`n" + "=" * 80 -ForegroundColor Green
        Write-Host "✅ ANÁLISIS COMPLETADO EXITOSAMENTE" -ForegroundColor Green
        Write-Host "=" * 80 -ForegroundColor Green
        Write-Host "📁 Archivo generado: $archivoHTML" -ForegroundColor Cyan
        Write-Host "📊 Servidor analizado: $($infoSistema.NombreServidor)" -ForegroundColor Cyan
        Write-Host "🖥️ Sistema operativo: $($infoSistema.NombreSO)" -ForegroundColor Cyan
        Write-Host "⏱️ Tiempo de actividad: $($infoSistema.TiempoActividad.Days) días, $($infoSistema.TiempoActividad.Hours) horas" -ForegroundColor Cyan
        
        # Estadísticas del análisis
        Write-Host "`n📈 ESTADÍSTICAS DEL ANÁLISIS:" -ForegroundColor Yellow
        Write-Host "   • Logs del sistema analizados: $($logsEventos.LogsSistema.Count)" -ForegroundColor White
        Write-Host "   • Logs de aplicación analizados: $($logsEventos.LogsAplicacion.Count)" -ForegroundColor White
        Write-Host "   • Logs de seguridad analizados: $($logsEventos.LogsSeguridad.Count)" -ForegroundColor White
        Write-Host "   • Servicios automáticos detenidos: $($datosExtendidos.ServiciosDetenidos.Count)" -ForegroundColor White
        Write-Host "   • Parches recientes instalados: $($datosExtendidos.UltimosParches.Count)" -ForegroundColor White
        
        if ($AnalisisSeguridad) {
            Write-Host "`n🛡️ ANÁLISIS DE SEGURIDAD:" -ForegroundColor Yellow
            if ($verificacionCumplimiento -and $verificacionCumplimiento.ResumenCumplimiento) {
                Write-Host "   • Cumplimiento general: $($verificacionCumplimiento.ResumenCumplimiento.PorcentajeCumplimiento)%" -ForegroundColor White
                Write-Host "   • Verificaciones exitosas: $($verificacionCumplimiento.ResumenCumplimiento.Cumple)" -ForegroundColor White
                Write-Host "   • Verificaciones fallidas: $($verificacionCumplimiento.ResumenCumplimiento.NoCumple)" -ForegroundColor White
            }
            if ($analisisPermisos -and $analisisPermisos.ResumenPermisos) {
                Write-Host "   • Carpetas analizadas: $($analisisPermisos.ResumenPermisos.TotalCarpetasAnalizadas)" -ForegroundColor White
                Write-Host "   • Carpetas con problemas: $($analisisPermisos.ResumenPermisos.CarpetasConProblemas)" -ForegroundColor White
            }
            if ($auditoriaSoftware -and $auditoriaSoftware.ResumenAuditoria) {
                Write-Host "   • Software instalado: $($auditoriaSoftware.ResumenAuditoria.TotalSoftwareInstalado)" -ForegroundColor White
                Write-Host "   • Software problemático: $($auditoriaSoftware.ResumenAuditoria.SoftwareConProblemas)" -ForegroundColor White
            }
        }
        
        Write-Host "`n🌐 Para ver el informe completo, abra el archivo HTML en su navegador." -ForegroundColor Green
        Write-Host "=" * 80 -ForegroundColor Green
        
        # Abrir automáticamente el archivo si es posible
        try {
            Start-Process $archivoHTML
            Write-Host "🚀 Abriendo informe en el navegador predeterminado..." -ForegroundColor Green
        } catch {
            Write-Host "ℹ️ Abra manualmente el archivo: $archivoHTML" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Error "❌ Error durante la ejecución del script: $_"
        Write-Host "📧 Si el problema persiste, contacte al administrador del sistema." -ForegroundColor Red
        exit 1
    }
}

# Ejecutar función principal
Main
