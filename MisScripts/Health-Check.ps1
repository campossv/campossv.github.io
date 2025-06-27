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
.PARAMETER FormatoExportar
    Formato de exportación (HTML, JSON, CSV, EXCEL)
.PARAMETER ParchesFaltantes
    Revisa parches faltantes
.PARAMETER RevisarServicioTerceros
    Revisa servicios de terceros
.PARAMETER AnalisisSeguridad
    Realiza análisis de seguridad
.PARAMETER VerificarCumplimiento
    Verifica cumplimiento
.EXAMPLE
    .\Health-Checkps1 -SalidaArchivo "C:\Informes" -DiasLogs 7 -FormatoExportar HTML -ParchesFaltantes -RevisarServicioTerceros -AnalisisSeguridad -VerificarCumplimiento
 .AUTHOR
    Vladimir Campos
.NOTES
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
Install-Module -Name ImportExcel -Scope CurrentUser -Force
Add-Type -AssemblyName System.Drawing
$LPNG = "iVBORw0KGgoAAAANSUhEUgAAAMgAAAAuCAYAAABtRVYBAAAACXBIWXMAAAsTAAALEwEAmpwYAAAGq2lUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNy4xLWMwMDAgNzkuYTg3MzFiOSwgMjAyMS8wOS8wOS0wMDozNzozOCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6cGhvdG9zaG9wPSJodHRwOi8vbnMuYWRvYmUuY29tL3Bob3Rvc2hvcC8xLjAvIiB4bXBNTTpPcmlnaW5hbERvY3VtZW50SUQ9InhtcC5kaWQ6ODIxNzExM2UtNmZjNS0zNjQyLWIzMDctOGEzNDM3ZmY2NWRkIiB4bXBNTTpEb2N1bWVudElEPSJhZG9iZTpkb2NpZDpwaG90b3Nob3A6Mzk4YTY5ZDMtYzljYS0zYzRhLWE4YTctZjhmYmM2MmYxOWU0IiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjEzNDk3YzZlLWVjNTgtMzM0YS1hZWY2LWFhMWFlODRjNGE0YiIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgMjMuMCAoV2luZG93cykiIHhtcDpDcmVhdGVEYXRlPSIyMDI0LTExLTIwVDEzOjU2OjExLTA2OjAwIiB4bXA6TW9kaWZ5RGF0ZT0iMjAyNC0xMS0yMFQxNzo1MTo0My0wNjowMCIgeG1wOk1ldGFkYXRhRGF0ZT0iMjAyNC0xMS0yMFQxNzo1MTo0My0wNjowMCIgZGM6Zm9ybWF0PSJpbWFnZS9wbmciIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDo4MjE3MTEzZS02ZmM1LTM2NDItYjMwNy04YTM0MzdmZjY1ZGQiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6ODIxNzExM2UtNmZjNS0zNjQyLWIzMDctOGEzNDM3ZmY2NWRkIi8+IDx4bXBNTTpIaXN0b3J5PiA8cmRmOlNlcT4gPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjQ4OTAyZGY4LTNjNzQtNzc0MC05YjM1LTBjYjkyODRjYTgyMCIgc3RFdnQ6d2hlbj0iMjAyNC0xMS0yMFQxNzo0NzoyMi0wNjowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDIzLjAgKFdpbmRvd3MpIiBzdEV2dDpjaGFuZ2VkPSIvIi8+IDxyZGY6bGkgc3RFdnQ6YWN0aW9uPSJzYXZlZCIgc3RFdnQ6aW5zdGFuY2VJRD0ieG1wLmlpZDoxMzQ5N2M2ZS1lYzU4LTMzNGEtYWVmNi1hYTFhZTg0YzRhNGIiIHN0RXZ0OndoZW49IjIwMjQtMTEtMjBUMTc6NTE6NDMtMDY6MDAiIHN0RXZ0OnNvZnR3YXJlQWdlbnQ9IkFkb2JlIFBob3Rvc2hvcCAyMy4wIChXaW5kb3dzKSIgc3RFdnQ6Y2hhbmdlZD0iLyIvPiA8L3JkZjpTZXE+IDwveG1wTU06SGlzdG9yeT4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz55i9rdAAAuNUlEQVR4nO2deXyU1fnovzOTmclkZ0IC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"
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
                    Level = 1,2
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

                    $smartData = Get-CimInstance -ClassName MSStorageDriver_FailurePredictStatus -Namespace "root\wmi" -ErrorAction SilentlyContinue |
                                Where-Object { $_.InstanceName -like "*$($disco.PNPDeviceID)*" }

                    if (-not $smartData) {

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


        try {
            Write-Host "   Obteniendo características mediante DISM..." -ForegroundColor Yellow
            $dismFeatures = dism /online /get-features /format:table | Out-String


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


        Write-Host "   Analizando servicios de roles específicos..." -ForegroundColor Yellow
        $rolesDetectados = @()


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


            try {
                $iisInfo = Get-CimInstance -ClassName Win32_Service -Filter "Name='W3SVC'" -ErrorAction SilentlyContinue
                $sitiosIIS = @()


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


        $spoolerService = Get-Service -Name "Spooler" -ErrorAction SilentlyContinue
        if ($spoolerService -and $spoolerService.Status -eq "Running") {

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


        Write-Host "   Recopilando información adicional del servidor..." -ForegroundColor Yellow


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


        try {
            $caracteristicasWindows = Get-WindowsFeature -ErrorAction SilentlyContinue |
                                     Where-Object { $_.InstallState -eq "Installed" } |
                                     Select-Object Name, DisplayName, InstallState |
                                     Sort-Object DisplayName

            if ($caracteristicasWindows) {
                $datos.CaracteristicasWindows = $caracteristicasWindows
            } else {

                $datos.CaracteristicasWindows = @{ Info = "Get-WindowsFeature no disponible en esta versión" }
            }
        } catch {
            $datos.CaracteristicasWindows = @{ Error = "Error al obtener características de Windows: $_" }
        }


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


function Get-AnalisisPoliticasGrupo {
    try {
        Write-Progress -Activity "Analizando políticas de grupo aplicadas" -PercentComplete 60
        Write-Host "   Analizando políticas de grupo (GPO)..." -ForegroundColor Yellow

        $datos = @{}


        try {

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


            $gpresultOutput = gpresult /r /scope:computer 2>$null
            $gpresultUser = gpresult /r /scope:user 2>$null


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


            try {
                Write-Host "   Obteniendo detalles de configuración GPO..." -ForegroundColor Yellow
                $gpresultDetailed = gpresult /v /scope:computer 2>$null


                $configuracionesSeguridad = @()
                $configuracionesRed = @()
                $configuracionesAuditoria = @()

                if ($gpresultDetailed) {
                    $currentSection = ""
                    foreach ($line in $gpresultDetailed) {

                        if ($line -match "Security Settings") {
                            $currentSection = "Security"
                        } elseif ($line -match "Network") {
                            $currentSection = "Network"
                        } elseif ($line -match "Audit") {
                            $currentSection = "Audit"
                        }


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


        try {
            $secpol = secedit /export /cfg "$env:TEMP\secpol.cfg" 2>$null
            if (Test-Path "$env:TEMP\secpol.cfg") {
                $secpolContent = Get-Content "$env:TEMP\secpol.cfg"

                foreach ($line in $secpolContent) {
                    if ($line -match "^(.+)\s*=\s*(.+)$") {
                        $setting = $matches[1].Trim()
                        $value = $matches[2].Trim()


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


function Get-VerificacionCumplimiento {
    try {
        Write-Progress -Activity "Verificando cumplimiento con estándares de seguridad" -PercentComplete 65
        Write-Host "   Verificando cumplimiento con CIS Benchmarks..." -ForegroundColor Yellow

        $verificaciones = @()


        Write-Host "   Verificando políticas de contraseña..." -ForegroundColor Yellow


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


        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.1.1"
            Descripcion = "Enforce password history: 24 or more passwords remembered"
            Categoria = "Políticas de Contraseña"
            EstadoActual = "Verificar manualmente"
            Recomendacion = "24 o más contraseñas"
            Cumple = "Pendiente"
            Criticidad = "Media"
        }


        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.1.2"
            Descripcion = "Maximum password age: 365 or fewer days"
            Categoria = "Políticas de Contraseña"
            EstadoActual = "$maxPasswordAge días"
            Recomendacion = "365 días o menos"
            Cumple = if ($maxPasswordAge -le 365 -and $maxPasswordAge -gt 0) { "Sí" } else { "No" }
            Criticidad = "Media"
        }


        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.1.3"
            Descripcion = "Minimum password age: 1 or more days"
            Categoria = "Políticas de Contraseña"
            EstadoActual = "$minPasswordAge días"
            Recomendacion = "1 día o más"
            Cumple = if ($minPasswordAge -ge 1) { "Sí" } else { "No" }
            Criticidad = "Baja"
        }


        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.1.4"
            Descripcion = "Minimum password length: 14 or more characters"
            Categoria = "Políticas de Contraseña"
            EstadoActual = "$minPasswordLength caracteres"
            Recomendacion = "14 caracteres o más"
            Cumple = if ($minPasswordLength -ge 14) { "Sí" } else { "No" }
            Criticidad = "Alta"
        }


        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.1.5"
            Descripcion = "Password must meet complexity requirements"
            Categoria = "Políticas de Contraseña"
            EstadoActual = if ($passwordComplexity) { "Habilitado" } else { "Deshabilitado" }
            Recomendacion = "Habilitado"
            Cumple = if ($passwordComplexity) { "Sí" } else { "No" }
            Criticidad = "Alta"
        }


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


        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.2.1"
            Descripcion = "Account lockout threshold: 5 or fewer invalid attempts"
            Categoria = "Políticas de Bloqueo"
            EstadoActual = if ($lockoutThreshold -eq 0) { "Sin bloqueo" } else { "$lockoutThreshold intentos" }
            Recomendacion = "5 intentos o menos"
            Cumple = if ($lockoutThreshold -gt 0 -and $lockoutThreshold -le 5) { "Sí" } else { "No" }
            Criticidad = "Media"
        }


        $verificaciones += [PSCustomObject]@{
            ID = "CIS-1.2.2"
            Descripcion = "Account lockout duration: 15 or more minutes"
            Categoria = "Políticas de Bloqueo"
            EstadoActual = "$lockoutDuration minutos"
            Recomendacion = "15 minutos o más"
            Cumple = if ($lockoutDuration -ge 15) { "Sí" } else { "No" }
            Criticidad = "Media"
        }


        Write-Host "   Verificando servicios y características de seguridad..." -ForegroundColor Yellow


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


        Write-Host "   Verificando configuraciones del registro..." -ForegroundColor Yellow


        $smbv1Enabled = $false
        try {
            $smbv1Feature = Get-WindowsOptionalFeature -Online -FeatureName "SMB1Protocol" -ErrorAction SilentlyContinue
            $smbv1Enabled = $smbv1Feature -and $smbv1Feature.State -eq "Enabled"
        } catch {

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


function Get-AnalisisPermisos {
    try {
        Write-Progress -Activity "Analizando permisos de carpetas sensibles" -PercentComplete 70
        Write-Host "   Analizando permisos de carpetas sensibles..." -ForegroundColor Yellow

        $analisisPermisos = @()


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


                    $acl = Get-Acl $carpeta.Ruta -ErrorAction SilentlyContinue

                    if ($acl) {
                        $permisosProblematicos = @()
                        $permisosNormales = @()

                        foreach ($access in $acl.Access) {
                            $usuario = $access.IdentityReference.Value
                            $permisos = $access.FileSystemRights.ToString()
                            $tipo = $access.AccessControlType.ToString()
                            $herencia = $access.IsInherited


                            $esProblematico = $false
                            $razon = ""


                            if ($usuario -match "(Everyone|Users|Authenticated Users)" -and $tipo -eq "Allow") {
                                if ($permisos -match "(FullControl|Modify|Write)" -and $carpeta.Ruta -match "(System32|Program Files|Windows)") {
                                    $esProblematico = $true
                                    $razon = "Permisos excesivos para grupo amplio en carpeta del sistema"
                                }
                            }


                            if ($permisos -match "(Write|Modify|FullControl)" -and $tipo -eq "Allow" -and $carpeta.Ruta -match "(System32|SysWOW64)") {
                                if ($usuario -notmatch "(SYSTEM|Administrators|TrustedInstaller)") {
                                    $esProblematico = $true
                                    $razon = "Permisos de escritura en carpeta crítica del sistema"
                                }
                            }


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


function Get-AuditoriaSoftware {
    try {
        Write-Progress -Activity "Realizando auditoría de software instalado" -PercentComplete 75
        Write-Host "   Auditando software instalado..." -ForegroundColor Yellow

        $softwareInstalado = @()
        $softwareProblematico = @()


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


                    $fechaInstalacionFormateada = "No disponible"
                    if ($fechaInstalacion -and $fechaInstalacion -match "^\d{8}$") {
                        try {
                            $fechaInstalacionFormateada = [DateTime]::ParseExact($fechaInstalacion, "yyyyMMdd", $null).ToString("dd/MM/yyyy")
                        } catch {}
                    }


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


        $softwareInstalado = $softwareInstalado | Sort-Object Nombre, Version | Get-Unique -AsString

        Write-Host "   Se encontraron $($softwareInstalado.Count) programas instalados" -ForegroundColor Yellow


        Write-Host "   Identificando software problemático..." -ForegroundColor Yellow


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


        $softwareSinSoporte = @(
            "Windows XP", "Windows Vista", "Windows 7", "Office 2010", "Office 2013",
            "Adobe Flash", "Internet Explorer", "Silverlight"
        )

        foreach ($software in $softwareInstalado) {
            $problemas = @()


            foreach ($problema in $softwareProblemas) {
                if ($software.Nombre -like "*$($problema.Nombre)*") {
                    if ($problema.VersionMinima) {

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


        Write-Host "   Analizando navegadores y plugins..." -ForegroundColor Yellow

        $navegadores = $softwareInstalado | Where-Object {
            $_.Nombre -match "(Chrome|Firefox|Edge|Internet Explorer|Opera|Safari)"
        }


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


        Write-Host "   Analizando servicios..." -ForegroundColor Yellow
        $serviciosDetenidos = Get-Service | Where-Object { $_.StartType -eq "Automatic" -and $_.Status -ne "Running" }

        if ($RevisarServicioTerceros) {

            $serviciosMicrosoft = @('Spooler', 'BITS', 'Themes', 'AudioSrv', 'Dnscache', 'eventlog', 'PlugPlay',
                                    'RpcSs', 'lanmanserver', 'W32Time', 'Winmgmt', 'Schedule', 'LanmanWorkstation',
                                    'DHCP', 'Netlogon', 'PolicyAgent', 'TermService', 'UmRdpService', 'SessionEnv',
                                    'RemoteRegistry')


            $datos.ServiciosDetenidos = foreach ($servicio in $serviciosDetenidos) {

                $servicioCIM = Get-CimInstance Win32_Service -Filter "Name='$($servicio.Name)'" -ErrorAction SilentlyContinue
                $compania = "Desconocido"
                $esMicrosoft = $false


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


        Write-Host "   Analizando procesos con alto consumo..." -ForegroundColor Yellow
        $datos.ProcesosCPU = Get-Process | Where-Object { $_.CPU -gt 0 } |
                            Sort-Object CPU -Descending | Select-Object -First 5 |
                            Select-Object Name, @{Name="CPU";Expression={[math]::Round($_.CPU,2)}}, Id,
                                        @{Name="Memoria_MB";Expression={[math]::Round($_.WS/1MB,2)}}

        $datos.ProcesosMemoria = Get-Process | Sort-Object WS -Descending | Select-Object -First 5 |
                                Select-Object Name, @{Name="Memoria_MB";Expression={[math]::Round($_.WS/1MB,2)}},
                                            Id, @{Name="CPU";Expression={[math]::Round($_.CPU,2)}}


        Write-Host "   Analizando puertos de red..." -ForegroundColor Yellow
        $datos.PuertosAbiertos = Get-NetTCPConnection | Where-Object { $_.State -eq "Listen" } |
                                Select-Object LocalAddress, LocalPort, State, OwningProcess |
                                Sort-Object LocalPort


        $datos.ConexionesActivas = Get-NetTCPConnection | Where-Object { $_.State -eq "Established" } |
                                  Group-Object RemoteAddress | Select-Object Count, Name |
                                  Sort-Object Count -Descending | Select-Object -First 10


        Write-Host "   Analizando usuarios locales..." -ForegroundColor Yellow
        $datos.UsuariosLocales = Get-LocalUser | Select-Object Name, Enabled,
                                @{Name="UltimoLogon";Expression={
                                    if($_.LastLogon -eq $null){"Nunca"}else{$_.LastLogon.ToString("dd/MM/yyyy HH:mm")}
                                }},
                                @{Name="PasswordExpires";Expression={
                                    if($_.PasswordExpires -eq $null){"Nunca"}else{$_.PasswordExpires.ToString("dd/MM/yyyy")}
                                }}


        Write-Host "   Analizando intentos de login fallidos..." -ForegroundColor Yellow
        $datos.LoginsFallidos = Get-WinEvent -LogName "Security" -FilterXPath "*[System[EventID=4625]]" -MaxEvents 5 -ErrorAction SilentlyContinue |
                               Select-Object TimeCreated,
                                           @{Name="Cuenta";Expression={$_.Properties[5].Value}},
                                           @{Name="IPOrigen";Expression={$_.Properties[19].Value}},
                                           @{Name="TipoFallo";Expression={$_.Properties[10].Value}}


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

                $wmiBuscador = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntivirusProduct -ErrorAction SilentlyContinue
                if ($wmiBuscador) {
                    $defenderWMI = $wmiBuscador | Where-Object { $_.displayName -like "*Defender*" }

                    if ($defenderWMI) {

                        $habilitado = [bool]($defenderWMI.productState -band 0x1000)
                        $actualizado = [bool]($defenderWMI.productState -band 0x10)

                        $datos.WindowsDefender = @{
                            Habilitado = $habilitado
                            TiempoRealActivo = $habilitado
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


        try {
            $otrosAntivirus = @()
            $antivirusList = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntivirusProduct -ErrorAction SilentlyContinue

            if ($antivirusList) {
                foreach ($av in $antivirusList) {
                    if ($av.displayName -notlike "*Defender*") {

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

                $netshOutput = netsh advfirewall show allprofiles | Out-String


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


    $nombreServidor = $InfoSistema.NombreServidor
    $sistemaOperativo = $InfoSistema.NombreSO


    $ms = New-Object System.IO.MemoryStream
    $imagenl.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $imagenBytes = $ms.ToArray()
    $ms.Close()
    $imagenBase64 = [Convert]::ToBase64String($imagenBytes)


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


    if ($DiagnosticoHardware) {
        $htmlContent += @"
        <h2>🔧 DIAGNÓSTICO AVANZADO DE HARDWARE</h2>
"@


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
                        this.style.backgroundColor = '
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


function Main {
    try {
        Write-Host "`n🚀 INICIANDO ANÁLISIS COMPLETO DE SALUD DEL SISTEMA" -ForegroundColor Green
        Write-Host "=" * 80 -ForegroundColor Green
        Write-Host "Servidor: $NombreServidor" -ForegroundColor Cyan
        Write-Host "Fecha: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Cyan
        Write-Host "=" * 80 -ForegroundColor Green


        Write-Host "`n📋 FASE 1: Recopilando información del sistema..." -ForegroundColor Yellow
        $infoSistema = Get-InformacionSistema


        Write-Host "`n📊 FASE 2: Analizando métricas de rendimiento..." -ForegroundColor Yellow
        $metricasRendimiento = Get-MetricasRendimiento


        Write-Host "`n📝 FASE 3: Recopilando logs de eventos..." -ForegroundColor Yellow
        $logsEventos = Get-LogsEventos -Dias $DiasLogs


        Write-Host "`n🔍 FASE 4: Recopilando datos extendidos..." -ForegroundColor Yellow
        $datosExtendidos = Get-DatosExtendidos -ParchesFaltantes $ParchesFaltantes -RevisarServicioTerceros $RevisarServicioTerceros


        $analisisConfiabilidad = $null
        $diagnosticoHardware = $null
        $analisisRoles = $null
        $analisisPoliticas = $null
        $verificacionCumplimiento = $null
        $analisisPermisos = $null
        $auditoriaSoftware = $null

        if ($AnalisisSeguridad) {
            Write-Host "`n🛡️ FASE 5: Ejecutando análisis de seguridad avanzado..." -ForegroundColor Yellow


            Write-Host "   5.1 Análisis de confiabilidad..." -ForegroundColor Cyan
            $analisisConfiabilidad = Get-AnalisisConfiabilidad


            Write-Host "   5.2 Diagnóstico de hardware avanzado..." -ForegroundColor Cyan
            $diagnosticoHardware = Get-DiagnosticoHardwareAvanzado


            Write-Host "   5.3 Análisis de roles de servidor..." -ForegroundColor Cyan
            $analisisRoles = Get-AnalisisRolesServidor


            Write-Host "   5.4 Análisis de políticas de grupo..." -ForegroundColor Cyan
            $analisisPoliticas = Get-AnalisisPoliticasGrupo


            Write-Host "   5.5 Análisis de permisos de carpetas..." -ForegroundColor Cyan
            $analisisPermisos = Get-AnalisisPermisos


            Write-Host "   5.6 Auditoría de software instalado..." -ForegroundColor Cyan
            $auditoriaSoftware = Get-AuditoriaSoftware
        }

        if ($VerificarCumplimiento) {
            Write-Host "`n✅ FASE 6: Verificando cumplimiento con estándares..." -ForegroundColor Yellow
            $verificacionCumplimiento = Get-VerificacionCumplimiento
        }


        Write-Host "`n📄 FASE FINAL: Generando informe completo..." -ForegroundColor Yellow
        Write-Progress -Activity "Generando informe HTML" -PercentComplete 90

        $htmlContent = Generate-CompleteHTML -InfoSistema $infoSistema -MetricasRendimiento $metricasRendimiento -LogsEventos $logsEventos -DatosExtendidos $datosExtendidos -AnalisisConfiabilidad $analisisConfiabilidad -DiagnosticoHardware $diagnosticoHardware -AnalisisRoles $analisisRoles -AnalisisPoliticas $analisisPoliticas -VerificacionCumplimiento $verificacionCumplimiento -AnalisisPermisos $analisisPermisos -AuditoriaSoftware $auditoriaSoftware


        $archivoHTML = "$ArchivoSalida.html"
        $htmlContent | Out-File -FilePath $archivoHTML -Encoding UTF8

        Write-Progress -Activity "Completado" -PercentComplete 100


        Write-Host "`n" + "=" * 80 -ForegroundColor Green
        Write-Host "✅ ANÁLISIS COMPLETADO EXITOSAMENTE" -ForegroundColor Green
        Write-Host "=" * 80 -ForegroundColor Green
        Write-Host "📁 Archivo generado: $archivoHTML" -ForegroundColor Cyan
        Write-Host "📊 Servidor analizado: $($infoSistema.NombreServidor)" -ForegroundColor Cyan
        Write-Host "🖥️ Sistema operativo: $($infoSistema.NombreSO)" -ForegroundColor Cyan
        Write-Host "⏱️ Tiempo de actividad: $($infoSistema.TiempoActividad.Days) días, $($infoSistema.TiempoActividad.Hours) horas" -ForegroundColor Cyan


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


Main
