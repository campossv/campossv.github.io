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