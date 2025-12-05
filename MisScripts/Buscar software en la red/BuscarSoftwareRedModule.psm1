#requires -Version 5.1

function Get-DomainComputers {
    param(
        [ValidateSet('All', 'Servers', 'Workstations')]
        [string]$Type = 'All'
    )

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        throw "No se pudo importar el módulo ActiveDirectory. Asegúrate de tener RSAT/AD instalado. Detalle: $_"
    }

    try {
        $ldapFilter = '(objectClass=computer)'

        switch ($Type) {
            'Servers' {
                $ldapFilter = '(operatingSystem=*Server*)'
            }
            'Workstations' {
                $ldapFilter = '(!(operatingSystem=*Server*))'
            }
        }

        $computers = Get-ADComputer -LDAPFilter $ldapFilter -Properties OperatingSystem |
        Select-Object Name, OperatingSystem

        return $computers | Sort-Object Name
    }
    catch {
        throw "Error al obtener equipos del dominio: $_"
    }
}

function Get-RemoteSoftwareFromRegistry {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$SearchText
    )

    $scriptBlock = {
        param($SearchTextInner)
        $paths = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )

        $apps = foreach ($path in $paths) {
            Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and ($_.DisplayName -like "*" + $SearchTextInner + "*") } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
        }

        return $apps
    }

    try {
        $sessionOptions = New-PSSessionOption -OperationTimeout 120000
        $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $SearchText -SessionOption $sessionOptions -ErrorAction Stop
        foreach ($app in $result) {
            [PSCustomObject]@{
                ComputerName   = $ComputerName
                DisplayName    = $app.DisplayName
                DisplayVersion = $app.DisplayVersion
                Publisher      = $app.Publisher
                InstallDate    = $app.InstallDate
            }
        }
    }
    catch {
        @()
    }
}

function Get-RemoteSoftwareFromPackage {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$SearchText
    )

    $scriptBlock = {
        param($SearchTextInner)
        try {
            $packages = Get-Package -ErrorAction Stop |
            Where-Object { $_.Name -and ($_.Name -like "*" + $SearchTextInner + "*") }

            foreach ($pkg in $packages) {
                [PSCustomObject]@{
                    DisplayName    = $pkg.Name
                    DisplayVersion = $pkg.Version
                    Publisher      = $pkg.ProviderName
                    InstallDate    = $null
                }
            }
        }
        catch {
            @()
        }
    }

    try {
        $sessionOptions = New-PSSessionOption -OperationTimeout 120000
        $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $SearchText -SessionOption $sessionOptions -ErrorAction Stop
        foreach ($app in $result) {
            [PSCustomObject]@{
                ComputerName   = $ComputerName
                DisplayName    = $app.DisplayName
                DisplayVersion = $app.DisplayVersion
                Publisher      = $app.Publisher
                InstallDate    = $app.InstallDate
            }
        }
    }
    catch {
        @()
    }
}

function Get-RemoteSoftwareFromWmi {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$SearchText
    )

    try {
        $instances = Get-CimInstance -ClassName Win32_Product -ComputerName $ComputerName -OperationTimeoutSec 120 -ErrorAction Stop |
        Where-Object { $_.Name -and ($_.Name -like "*" + $SearchText + "*") }

        foreach ($inst in $instances) {
            [PSCustomObject]@{
                ComputerName   = $ComputerName
                DisplayName    = $inst.Name
                DisplayVersion = $inst.Version
                Publisher      = $inst.Vendor
                InstallDate    = $inst.InstallDate
            }
        }
    }
    catch {
        @()
    }
}

function Get-InstalledSoftwareRemote {
    param(
        [Parameter(Mandatory)] [string]$ComputerName,
        [Parameter(Mandatory)] [string]$SearchText,
        [switch]$UseWin32Product = $true
    )

    if (-not (Test-ComputerOnline -ComputerName $ComputerName)) {
        return @()
    }

    $results = @()

    # 1) Registro (método principal)
    $results += Get-RemoteSoftwareFromRegistry -ComputerName $ComputerName -SearchText $SearchText

    # 2) Get-Package como fallback si no hubo resultados
    if (-not $results -or $results.Count -eq 0) {
        $results += Get-RemoteSoftwareFromPackage -ComputerName $ComputerName -SearchText $SearchText
    }

    # 3) WMI (Win32_Product) como último recurso automático si no hay resultados
    if ($UseWin32Product -and (-not $results -or $results.Count -eq 0)) {
        $results += Get-RemoteSoftwareFromWmi -ComputerName $ComputerName -SearchText $SearchText
    }

    if (-not $results -or $results.Count -eq 0) {
        return @()
    }

    return $results
}

function Test-ComputerOnline {
    param(
        [Parameter(Mandatory)][string]$ComputerName
    )

    try {
        # ICMP + resolución DNS; -Quiet devuelve True/False
        return Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue
    }
    catch {
        return $false
    }
}