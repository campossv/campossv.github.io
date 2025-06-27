function Get-MetricasRendimiento {
    try {
        # CPU
        $cpuSamples = @()
        for ($i = 0; $i -lt 3; $i++) {
            $cpuSamples += (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
            Start-Sleep -Milliseconds 500
        }
        $usoCPU = [math]::Round(($cpuSamples | Measure-Object -Average).Average, 2)
        
        # Memoria
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $memTotal = [math]::Round($os.TotalVisibleMemorySize / 1KB, 2)
        $memLibre = [math]::Round($os.FreePhysicalMemory / 1KB, 2)
        $porcentajeMemoria = [math]::Round(($memTotal - $memLibre) / $memTotal * 100, 2)
        
        # Discos
        $discos = Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
        $metricasDisco = @()
        foreach ($disco in $discos) {
            $metricasDisco += @{
                Unidad = $disco.DeviceID
                TotalGB = [math]::Round($disco.Size / 1GB, 2)
                LibreGB = [math]::Round($disco.FreeSpace / 1GB, 2)
                PorcentajeUsado = [math]::Round(($disco.Size - $disco.FreeSpace) / $disco.Size * 100, 2)
                SistemaArchivos = $disco.FileSystem
            }
        }
        
        # Red
        $interfacesRed = Get-CimInstance -ClassName Win32_PerfRawData_Tcpip_NetworkInterface
        $metricasRed = @()
        foreach ($interfaz in $interfacesRed) {
            $metricasRed += @{
                Interfaz = $interfaz.Name
                BytesEnviados = [math]::Round($interfaz.BytesSentPerSec / 1KB, 2)
                BytesRecibidos = [math]::Round($interfaz.BytesReceivedPerSec / 1KB, 2)
            }
        }
        
        return @{
            CPU = @{ UsoCPU = $usoCPU }
            Memoria = @{ PorcentajeUsado = $porcentajeMemoria; TotalMB = $memTotal }
            Discos = $metricasDisco
            Red = $metricasRed
        }
    } catch {
        Write-Warning "Error al obtener métricas de rendimiento: $_"
        return @{ Error = $_.Exception.Message }
    }
}