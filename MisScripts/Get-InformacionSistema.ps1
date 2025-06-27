function Get-InformacionSistema {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $cpu = Get-CimInstance -ClassName Win32_Processor
        $system = Get-CimInstance -ClassName Win32_ComputerSystem
        $bios = Get-CimInstance -ClassName Win32_BIOS
        
        return @{
            NombreServidor = $env:COMPUTERNAME
            DireccionIP = (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1).IPv4Address.IPAddressToString
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
            MemoriaTotal = [math]::Round($system.TotalPhysicalMemory / 1GB, 2)
            DominioWorkgroup = if ($system.PartOfDomain) {"Dominio: $($system.Domain)"} else {"Grupo de trabajo: $($system.Workgroup)"}
            TimeZone = (Get-TimeZone).DisplayName
        }
    } catch {
        Write-Warning "Error al obtener información del sistema: $_"
        return @{ Error = $_.Exception.Message }
    }
}