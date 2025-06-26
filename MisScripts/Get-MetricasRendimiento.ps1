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
