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