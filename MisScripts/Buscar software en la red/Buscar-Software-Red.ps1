#requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

function Write-BuscarSoftwareLog {
    param(
        [Parameter(Mandatory)][string]$Message
    )

    try {
        $logDirectory = 'C:\Logs'
        $logPath = Join-Path -Path $logDirectory -ChildPath 'BuscarSoftwareRed.log'

        if (-not (Test-Path -Path $logDirectory)) {
            New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
        }

        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $entry = "[$timestamp] $Message"
        Add-Content -Path $logPath -Value $entry -Encoding UTF8
    }
    catch {
        # No romper la ejecución del módulo si falla el log
    }
}

function Get-ImageFromBase64 {
    param(
        [Parameter(Mandatory)][string]$Base64
    )

    if ([string]::IsNullOrWhiteSpace($Base64)) {
        return $null
    }

    try {
        $bytes = [Convert]::FromBase64String($Base64)
        $ms = New-Object System.IO.MemoryStream(, $bytes)
        return [System.Drawing.Image]::FromStream($ms)
    }
    catch {
        return $null
    }
}

function Get-DomainComputers {
    param(
        [ValidateSet('All', 'Servers', 'Workstations')]
        [string]$Type = 'All'
    )

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("No se pudo importar el módulo ActiveDirectory.\nAsegúrate de tener RSAT/AD instalado.", "Error", 'OK', 'Error') | Out-Null
        return @()
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
        $errorMessage = $_.Exception.Message
        $fullError = $_ | Out-String
        Write-BuscarSoftwareLog -Message "Get-DomainComputers error. Type=$Type; Message=$errorMessage; Details=$fullError"
        [System.Windows.Forms.MessageBox]::Show("Error al obtener equipos del dominio. Detalle: $errorMessage", "Error", 'OK', 'Error') | Out-Null
        return @()
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
        $errorMessage = $_.Exception.Message
        $fullError = $_ | Out-String
        Write-BuscarSoftwareLog -Message "Get-RemoteSoftwareFromRegistry error. Computer=$ComputerName; Message=$errorMessage; Details=$fullError"
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
        $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $SearchText -ErrorAction Stop
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
        $errorMessage = $_.Exception.Message
        $fullError = $_ | Out-String
        Write-BuscarSoftwareLog -Message "Get-RemoteSoftwareFromPackage error. Computer=$ComputerName; Message=$errorMessage; Details=$fullError"
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
        $errorMessage = $_.Exception.Message
        $fullError = $_ | Out-String
        Write-BuscarSoftwareLog -Message "Get-RemoteSoftwareFromWmi error. Computer=$ComputerName; Message=$errorMessage; Details=$fullError"
        @()
    }
}

function Get-InstalledSoftwareRemote {
    param(
        [Parameter(Mandatory)] [string]$ComputerName,
        [Parameter(Mandatory)] [string]$SearchText,
        [switch]$UseWin32Product = $true
    )

    # 0) Comprobar que el equipo responde a ICMP
    if (-not (Test-ComputerOnline -ComputerName $ComputerName)) {
        Write-BuscarSoftwareLog -Message "Equipo no accesible por ICMP (posible problema de DNS, apagado o firewall ICMP): $ComputerName"
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

function Get-DefaultSilentArguments {
    param(
        [Parameter(Mandatory)][string]$InstallerPath
    )

    $ext = [System.IO.Path]::GetExtension($InstallerPath).ToLowerInvariant()

    switch ($ext) {
        '.msi' { '/i `"{0}`" /qn /norestart' -f $InstallerPath }
        '.exe' { '/silent /norestart' }
        default { '' }
    }
}

function Invoke-RemoteInstall {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$InstallerPath,
        [string]$SilentArgs,
        [string]$CustomCommand,
        [string]$ExpectedDisplayName
    )

    if (-not (Test-ComputerOnline -ComputerName $ComputerName)) {
        Write-BuscarSoftwareLog -Message "Instalacion omitida. Equipo sin respuesta ICMP: $ComputerName"
        return $false
    }

    if (-not (Test-Path -Path $InstallerPath)) {
        Write-BuscarSoftwareLog -Message "Ruta de instalador no valida: $InstallerPath"
        return $false
    }

    $fileName = [System.IO.Path]::GetFileName($InstallerPath)
    $ext = [System.IO.Path]::GetExtension($InstallerPath).ToLowerInvariant()

    if ([string]::IsNullOrWhiteSpace($SilentArgs)) {
        $SilentArgs = Get-DefaultSilentArguments -InstallerPath $InstallerPath
    }

    $session = $null
    try {
        $session = New-PSSession -ComputerName $ComputerName -ErrorAction Stop

        $remoteFolder = 'C:\Temp\SoftwareDeploy'

        Invoke-Command -Session $session -ScriptBlock {
            param($folder)
            if (-not (Test-Path -Path $folder)) {
                New-Item -Path $folder -ItemType Directory -Force | Out-Null
            }
        } -ArgumentList $remoteFolder -ErrorAction Stop

        $uniqueName = ([guid]::NewGuid().ToString('N').Substring(0, 8) + '_' + $fileName)
        $remotePath = Join-Path -Path $remoteFolder -ChildPath $uniqueName

        Copy-Item -Path $InstallerPath -Destination $remotePath -ToSession $session -Force

        $commandLine = $null

        if (-not [string]::IsNullOrWhiteSpace($CustomCommand)) {
            $commandLine = $CustomCommand.Replace('{InstallerPath}', '"{0}"' -f $remotePath)
        }
        else {
            if ($ext -eq '.msi') {
                if (-not $SilentArgs) { $SilentArgs = '/qn /norestart' }
                $commandLine = "msiexec.exe /i `"$remotePath`" $SilentArgs"
            }
            else {
                $commandLine = '"' + $remotePath + '" ' + $SilentArgs
            }
        }

        $scriptBlock = {
            param($cmd, $expectedName)

            Start-Process -FilePath 'cmd.exe' -ArgumentList "/c $cmd" -WindowStyle Hidden -Wait -ErrorAction Stop

            $installedOk = $true

            if ($expectedName -and -not [string]::IsNullOrWhiteSpace($expectedName)) {
                $paths = @(
                    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
                    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                )

                $found = $false
                foreach ($p in $paths) {
                    $item = Get-ItemProperty -Path $p -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -and ($_.DisplayName -like "*" + $expectedName + "*") } |
                    Select-Object -First 1 DisplayName

                    if ($item) {
                        $found = $true
                        break
                    }
                }

                if (-not $found) {
                    $installedOk = $false
                }
            }

            return $installedOk
        }

        $remoteResult = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $commandLine, $ExpectedDisplayName -ErrorAction Stop

        if ($remoteResult) {
            Write-BuscarSoftwareLog -Message "Instalacion verificada via PSSession en $ComputerName. Comando: $commandLine"
            return $true
        }
        else {
            Write-BuscarSoftwareLog -Message "Instalador se ejecuto pero no se encontro el software esperado en el registro en $ComputerName. ExpectedDisplayName: $ExpectedDisplayName"
            return $false
        }
    }
    catch {
        $msg = $_.Exception.Message
        Write-BuscarSoftwareLog -Message "Error en instalacion remota para $ComputerName. Detalle: $msg"
        return $false
    }
    finally {
        if ($session) {
            try { Remove-PSSession -Session $session -ErrorAction SilentlyContinue } catch { }
        }
    }
}

function Invoke-RemoteUninstall {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$DisplayName
    )

    if (-not (Test-ComputerOnline -ComputerName $ComputerName)) {
        Write-BuscarSoftwareLog -Message "Desinstalacion omitida. Equipo sin respuesta ICMP: $ComputerName"
        return $false
    }

    $methods = @('RegistryQuiet', 'RegistryNormal', 'Win32Product')

    foreach ($method in $methods) {
        try {
            switch ($method) {
                'RegistryQuiet' {
                    $scriptBlock = {
                        param($name)
                        $paths = @(
                            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
                            'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                        )

                        foreach ($path in $paths) {
                            Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
                            Where-Object { $_.DisplayName -and ($_.DisplayName -like "*" + $name + "*") -and $_.QuietUninstallString } |
                            Select-Object -First 1 @{ Name = 'Command'; Expression = { $_.QuietUninstallString } }
                        }
                    }

                    $cmdObj = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $DisplayName -ErrorAction Stop
                    if ($cmdObj -and $cmdObj.Command) {
                        $cmd = [string]$cmdObj.Command
                        $sb = {
                            param($uCmd)
                            Start-Process -FilePath 'cmd.exe' -ArgumentList "/c $uCmd" -WindowStyle Hidden -Wait -ErrorAction Stop
                        }
                        Invoke-Command -ComputerName $ComputerName -ScriptBlock $sb -ArgumentList $cmd -ErrorAction Stop
                        Write-BuscarSoftwareLog -Message "Desinstalacion exitosa (QuietUninstallString) en $ComputerName. DisplayName: $DisplayName Comando: $cmd"
                        return $true
                    }
                }
                'RegistryNormal' {
                    $scriptBlock2 = {
                        param($name)
                        $paths = @(
                            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
                            'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                        )

                        foreach ($path in $paths) {
                            Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
                            Where-Object { $_.DisplayName -and ($_.DisplayName -like "*" + $name + "*") -and $_.UninstallString } |
                            Select-Object -First 1 DisplayName, UninstallString
                        }
                    }

                    $info = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock2 -ArgumentList $DisplayName -ErrorAction Stop
                    if ($info -and $info.UninstallString) {
                        $u = [string]$info.UninstallString

                        if ($u -match 'msiexec(.+?)/I\s*\{') {
                            $u = $u -replace '/I', '/x'
                            if ($u -notmatch '/qn') {
                                $u += ' /qn /norestart'
                            }
                        }

                        $sb2 = {
                            param($uCmd)
                            Start-Process -FilePath 'cmd.exe' -ArgumentList "/c $uCmd" -WindowStyle Hidden -Wait -ErrorAction Stop
                        }
                        Invoke-Command -ComputerName $ComputerName -ScriptBlock $sb2 -ArgumentList $u -ErrorAction Stop
                        Write-BuscarSoftwareLog -Message "Desinstalacion iniciada (UninstallString) en $ComputerName. DisplayName: $DisplayName Comando: $u"
                        return $true
                    }
                }
                'Win32Product' {
                    $sb3 = {
                        param($name)
                        $product = Get-CimInstance -ClassName Win32_Product -ErrorAction Stop |
                        Where-Object { $_.Name -and ($_.Name -like "*" + $name + "*") } |
                        Select-Object -First 1

                        if ($product) {
                            $result = $product.Uninstall()
                            return $result.ReturnValue
                        }
                        else {
                            return $null
                        }
                    }

                    $rv = Invoke-Command -ComputerName $ComputerName -ScriptBlock $sb3 -ArgumentList $DisplayName -ErrorAction Stop
                    if ($rv -eq 0) {
                        Write-BuscarSoftwareLog -Message "Desinstalacion exitosa via Win32_Product en $ComputerName. DisplayName: $DisplayName"
                        return $true
                    }
                }
            }
        }
        catch {
            $msg = $_.Exception.Message
            Write-BuscarSoftwareLog -Message "Metodo de desinstalacion '$method' fallo para $ComputerName. DisplayName: $DisplayName Detalle: $msg"
        }
    }

    Write-BuscarSoftwareLog -Message "Todos los metodos de desinstalacion fallaron para $ComputerName. DisplayName: $DisplayName"
    return $false
}

# Crear formulario
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Buscador de software en la red'
$form.Size = New-Object System.Drawing.Size(900, 780)
$form.StartPosition = 'CenterScreen'
$form.TopMost = $false
$(
    # Estilo general del formulario
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
)

# Etiqueta y textbox de búsqueda
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = 'Nombre (o parte) del software:'
$lblSearch.AutoSize = $true
$lblSearch.Location = New-Object System.Drawing.Point(10, 90)
$lblSearch.ForeColor = [System.Drawing.Color]::FromArgb(45, 52, 63)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(190, 90)
$txtSearch.Width = 300
$txtSearch.BorderStyle = 'FixedSingle'

# Filtro de tipo de equipo
$lblType = New-Object System.Windows.Forms.Label
$lblType.Text = 'Tipo de equipo:'
$lblType.AutoSize = $true
$lblType.Location = New-Object System.Drawing.Point(10, 60)
$lblType.ForeColor = [System.Drawing.Color]::FromArgb(45, 52, 63)

$cmbType = New-Object System.Windows.Forms.ComboBox
$cmbType.Location = New-Object System.Drawing.Point(100, 55)
$cmbType.Width = 150
$cmbType.DropDownStyle = 'DropDownList'
$cmbType.FlatStyle = 'Flat'
[void]$cmbType.Items.Add('Todos')
[void]$cmbType.Items.Add('Servidores')
[void]$cmbType.Items.Add('Workstations')
$cmbType.SelectedIndex = 0

# Checkbox para uso de Win32_Product
$chkUseWin32Product = New-Object System.Windows.Forms.CheckBox
$chkUseWin32Product.Text = 'Usar Win32_Product (lento)'
$chkUseWin32Product.AutoSize = $true
$chkUseWin32Product.Location = New-Object System.Drawing.Point(500, 90)
$chkUseWin32Product.Checked = $true
$chkUseWin32Product.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)

# Controles para instalacion desatendida
$lblInstaller = New-Object System.Windows.Forms.Label
$lblInstaller.Text = 'Instalador:'
$lblInstaller.AutoSize = $true
$lblInstaller.Location = New-Object System.Drawing.Point(60, 600)
$lblInstaller.ForeColor = [System.Drawing.Color]::FromArgb(45, 52, 63)

$txtInstallerPath = New-Object System.Windows.Forms.TextBox
$txtInstallerPath.Location = New-Object System.Drawing.Point(125, 600)
$txtInstallerPath.Width = 360
$txtInstallerPath.BorderStyle = 'FixedSingle'

$btnBrowseInstaller = New-Object System.Windows.Forms.Button
$btnBrowseInstaller.Text = '...'
$btnBrowseInstaller.Location = New-Object System.Drawing.Point(488, 600)
$btnBrowseInstaller.Width = 40
$btnBrowseInstaller.FlatStyle = 'Flat'
$btnBrowseInstaller.BackColor = [System.Drawing.Color]::FromArgb(127, 140, 141)
$btnBrowseInstaller.ForeColor = [System.Drawing.Color]::White

$lblInstallArgs = New-Object System.Windows.Forms.Label
$lblInstallArgs.Text = 'Parametros silenciosos (opcional):'
$lblInstallArgs.AutoSize = $true
$lblInstallArgs.Location = New-Object System.Drawing.Point(60, 630)
$lblInstallArgs.ForeColor = [System.Drawing.Color]::FromArgb(45, 52, 63)

$txtInstallArgs = New-Object System.Windows.Forms.TextBox
$txtInstallArgs.Location = New-Object System.Drawing.Point(250, 630)
$txtInstallArgs.Width = 280
$txtInstallArgs.BorderStyle = 'FixedSingle'

$lblExpectedName = New-Object System.Windows.Forms.Label
$lblExpectedName.Text = 'Nombre esperado (registro):'
$lblExpectedName.AutoSize = $true
$lblExpectedName.Location = New-Object System.Drawing.Point(60, 660)
$lblExpectedName.ForeColor = [System.Drawing.Color]::FromArgb(45, 52, 63)

$txtExpectedName = New-Object System.Windows.Forms.TextBox
$txtExpectedName.Location = New-Object System.Drawing.Point(250, 658)
$txtExpectedName.Width = 280
$txtExpectedName.BorderStyle = 'FixedSingle'

# Botón cargar equipos
$btnLoadComputers = New-Object System.Windows.Forms.Button
$btnLoadComputers.Text = 'Cargar equipos'
$btnLoadComputers.Location = New-Object System.Drawing.Point(270, 55)
$btnLoadComputers.Width = 150
$btnLoadComputers.FlatStyle = 'Flat'
$btnLoadComputers.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
$btnLoadComputers.ForeColor = [System.Drawing.Color]::White

# Lista de equipos
$lblComputers = New-Object System.Windows.Forms.Label
$lblComputers.Text = 'Equipos del dominio:'
$lblComputers.AutoSize = $true
$lblComputers.Location = New-Object System.Drawing.Point(10, 130)
$lblComputers.ForeColor = [System.Drawing.Color]::FromArgb(45, 52, 63)

$lstComputers = New-Object System.Windows.Forms.ListBox
$lstComputers.Location = New-Object System.Drawing.Point(10, 160)
$lstComputers.Size = New-Object System.Drawing.Size(250, 400)
$lstComputers.SelectionMode = 'MultiExtended'
$lstComputers.BorderStyle = 'FixedSingle'

# Botón buscar
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = 'Buscar en seleccionados'
$btnSearch.Location = New-Object System.Drawing.Point(680, 90)
$btnSearch.Width = 180
$btnSearch.FlatStyle = 'Flat'
$btnSearch.BackColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
$btnSearch.ForeColor = [System.Drawing.Color]::White

# Botón exportar a CSV
$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = 'Exportar a CSV'
$btnExportCsv.Location = New-Object System.Drawing.Point(600, 570)
$btnExportCsv.Width = 130
$btnExportCsv.FlatStyle = 'Flat'
$btnExportCsv.BackColor = [System.Drawing.Color]::FromArgb(52, 73, 94)
$btnExportCsv.ForeColor = [System.Drawing.Color]::White

# Botón copiar fila
$btnCopySelected = New-Object System.Windows.Forms.Button
$btnCopySelected.Text = 'Copiar selección'
$btnCopySelected.Location = New-Object System.Drawing.Point(740, 570)
$btnCopySelected.Width = 130
$btnCopySelected.FlatStyle = 'Flat'
$btnCopySelected.BackColor = [System.Drawing.Color]::FromArgb(127, 140, 141)
$btnCopySelected.ForeColor = [System.Drawing.Color]::White

$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Text = 'Instalar en seleccionados'
$btnInstall.Location = New-Object System.Drawing.Point(600, 600)
$btnInstall.Width = 190
$btnInstall.Height = 60
$btnInstall.FlatStyle = 'Flat'
$btnInstall.BackColor = [System.Drawing.Color]::FromArgb(231, 76, 60)
$btnInstall.ForeColor = [System.Drawing.Color]::White

$btnUninstall = New-Object System.Windows.Forms.Button
$btnUninstall.Text = 'Desinstalar seleccionados'
$btnUninstall.Location = New-Object System.Drawing.Point(600, 670)
$btnUninstall.Width = 190
$btnUninstall.Height = 40
$btnUninstall.FlatStyle = 'Flat'
$btnUninstall.BackColor = [System.Drawing.Color]::FromArgb(192, 57, 43)
$btnUninstall.ForeColor = [System.Drawing.Color]::White

# Logo centrado arriba del nombre del software
$logoBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAJwAAABJCAYAAADIS0/RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAHzcAAB83AeXxie8AAAgCaVRYdFhNTDpjb20uYWRvYmUueG1wAAAAAAA8P3hwYWNrZXQgYmVnaW49Iu+7vyIgaWQ9Ilc1TTBNcENlaGlIenJlU3pOVGN6a2M5ZCI/Pg0KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgOS4xLWMwMDIgNzkuYTFjZDEyZiwgMjAyNC8xMS8xMS0xOTowODo0NiAgICAgICAgIj4NCgk8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPg0KCQk8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczpBdHRyaWI9Imh0dHA6Ly9ucy5hdHRyaWJ1dGlvbi5jb20vYWRzLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczpwaG90b3Nob3A9Imh0dHA6Ly9ucy5hZG9iZS5jb20vcGhvdG9zaG9wLzEuMC8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiBkYzpmb3JtYXQ9ImltYWdlL3BuZyIgeG1wOkNyZWF0b3JUb29sPSJDYW52YSBkb2M9REFHbVNQOVF5VDQgdXNlcj1VQUZzbDYwbHp0ZyBicmFuZD1DQU5WQSBQUk8gMyB0ZW1wbGF0ZT1Db2xsYWdlIGRlIEZlbGljaXRhY2lvbiBDdW1wbGVhw7FvcyBGb3RvcyBNb2Rlcm5vIFBhc3RlbCIgeG1wOkNyZWF0ZURhdGU9IjIwMjUtMDUtMTlUMTE6NTg6MTctMDY6MDAiIHhtcDpNb2RpZnlEYXRlPSIyMDI1LTA1LTE5VDEyOjAyOjQyLTA2OjAwIiB4bXA6TWV0YWRhdGFEYXRlPSIyMDI1LTA1LTE5VDEyOjAyOjQyLTA2OjAwIiBwaG90b3Nob3A6Q29sb3JNb2RlPSIzIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOmQ0YjlhMjRhLWZiM2MtYTg0NS04MTUzLWNlYjBjZWI1ZTNiMiIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDpkNGI5YTI0YS1mYjNjLWE4NDUtODE1My1jZWIwY2ViNWUzYjIiIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDpkNGI5YTI0YS1mYjNjLWE4NDUtODE1My1jZWIwY2ViNWUzYjIiPg0KCQkJPEF0dHJpYjpBZHM+DQoJCQkJPHJkZjpTZXE+DQoJCQkJCTxyZGY6bGkgQXR0cmliOkNyZWF0ZWQ9IjIwMjUtMDUtMTkiIEF0dHJpYjpFeHRJZD0iOTIyNGJiODItZTg3Zi00N2Q3LTg2N2MtYjhkYzYzOTM4NzIzIiBBdHRyaWI6RmJJZD0iNTI1MjY1OTE0MTc5NTgwIiBBdHRyaWI6VG91Y2hUeXBlPSIyIi8+DQoJCQkJPC9yZGY6U2VxPg0KCQkJPC9BdHRyaWI6QWRzPg0KCQkJPGRjOnRpdGxlPg0KCQkJCTxyZGY6QWx0Pg0KCQkJCQk8cmRmOmxpIHhtbDpsYW5nPSJ4LWRlZmF1bHQiPsKhQmllbnZlbmlkb3MgYWwgRXF1aXBvIERldmVsISAtIEx1bmVzIE1vdGl2YWNpb25hbDwvcmRmOmxpPg0KCQkJCTwvcmRmOkFsdD4NCgkJCTwvZGM6dGl0bGU+DQoJCQk8ZGM6Y3JlYXRvcj4NCgkJCQk8cmRmOlNlcT4NCgkJCQkJPHJkZjpsaT5XZWJtYXN0ZXIgRGV2ZWw8L3JkZjpsaT4NCgkJCQk8L3JkZjpTZXE+DQoJCQk8L2RjOmNyZWF0b3I+DQoJCQk8eG1wTU06SGlzdG9yeT4NCgkJCQk8cmRmOlNlcT4NCgkJCQkJPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOmQ0YjlhMjRhLWZiM2MtYTg0NS04MTUzLWNlYjBjZWI1ZTNiMiIgc3RFdnQ6d2hlbj0iMjAyNS0wNS0xOVQxMjowMjo0Mi0wNjowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDI2LjQgKFdpbmRvd3MpIiBzdEV2dDpjaGFuZ2VkPSIvIi8+DQoJCQkJPC9yZGY6U2VxPg0KCQkJPC94bXBNTTpIaXN0b3J5Pg0KCQk8L3JkZjpEZXNjcmlwdGlvbj4NCgkJPHJkZjpEZXNjcmlwdGlvbiB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyI+PHRpZmY6T3JpZW50YXRpb24+MTwvdGlmZjpPcmllbnRhdGlvbj48L3JkZjpEZXNjcmlwdGlvbj48L3JkZjpSREY+DQo8L3g6eG1wbWV0YT4NCjw/eHBhY2tldCBlbmQ9J3cnPz4C+79wAAAgy0lEQVR4Xu2dd1xUV/r/P1OZxtBBRcWCIqCAgFjWtRKjhohGEo0aY4lrTOJaIjHZrDExWTXWtcZOYkRBjSurEVGiGEA6DAjCgEMvQxvqFKbd3x+5szu5YWA0+E02v3m/XufFzHnOee557j1zynPOuQAWLFiwYMGCBQsWLFiwYMGCBQsWLFiwYOEPBY0a0RP3IyI4GDLEUafTEVRZbzQ0NGh3797dmZeXJ6fKLPy+qa6uduBwOFxqPABwuVwCAFFYWNgYGBioocqfFToA28MHD4arVCqFUqmUKxUK84JSKVcqlXK5XN7W3t5e3dLSkiiVSo8+EokWrly50ol6IQu/K7gABhcWFt7V6XQKU0GtVitEItFoauZfAwvAjHVr194g+pCWlpYGiURyIioqypd6QQu/C5wBrI6Pj6+kPjsq2dnZZj1DOjXCBDQAAqVKxacKfg22trZOw4YNWxsSEpJZWlp6fPv27c7UNBZ+U5gAbNVqNYMqMEar1YIgCLOGWeZWOADQ02g0PTWyL+Dz+cyhQ4eu27hxY2ZiYuLLVLmF3wwCgI5Go5lVmczhaSrcc8fW1nZQUFDQv7OysrZSZRb+GPyuKhwAsNls+Pv7787Ozt5OlVn43+d3V+EMjB079tOsrKy11HgL/9s8lwrX0NCgr6mp0dfW1upkMhlVbDajRo06duXKlfHUeAv/u/R5hdPpdFi3du0xfz+/Lb6+vjvGjx9/IDg4+OLmzZuTYmNja9vb26lZTMLj8RhTpkw5A4BNlVn4Y8MGELJyxYp4qv+FilarJT777LPppNPQBcAwAOMALATwNx8fn8tXr17t1a9jTGZm5kZqgZ4BOjnN7/Mf2e+YX2tzfwAbY2Njq6nPxBiNRkNkZWX5UDN3h7lLW2wAs1auWLHxXETETKrQGJ1Oh9TU1LGTJ08WUUR8AAMB+AEI2Lp168ydO3f60+m934u2trbazz//fOT+/fvNXhY7ceJEYFBQ0FQHB4cAHo83nMFgWDOZTJZer9fpdDq5RqOp7+joyGxsbHz42muvPaiurlZSdZAPinHjxo2FkydPXqrRaLTUBAbodDpNpVIVPXz4cNdrr73WRpWb4rq/y0uTnLiLoadZgyC6dTvRGDR0afXy20rOO289FHdQ5SS0EydOBEyYMGGKg4NDAIfDcWcwGNYMBoNFEITBZqlcLs+sr69P6cFmY/oDWBQbG7tl9uzZrlShAa1Wi7y8PN+AgIA8quxZeaoWLi4uLoCqwAg7AHMBfLFt27Y8an5TZGRkvEVV1A3c6OjoVaWlpSlyuZyqwiQtLS2SkpKSj0NDQ20p+pgAAsaNG7dHr9dTs3VLdXX1eYoOU1iPs+Uskc0aoideGUEQoe6mw2seRPMLQ66aaCC4UVFRK5/WZplMJikpKflbcHCwDVWhEX3ewpkLG0DIir6pcAAgBPAKgMP37t1rpOroDqlUep+qxJh58+ZNEYlEydR8T4NMJnuSkJAw10gtDUAAgM/OnTtXQk3fHUqlUhMTE+NupKM7GAAmRPg4ZRALRhDE3GGmQ8hwomvuMM3N8W6eVCVz586dkpub+2ttLklISJhD1U3S5xWu9/7s+dAOIB2AeMeOHTk6nY4q/wUCgSDo6NGj/ajxANhvvfXWmqNHj9709fWdRBU+DXZ2dsMnTpz4fUZGxjoyigBQCqB03759jxQKBSXHL+FwOEw/P7/expz2bjzm9BAXvj+03fai/4VJR3OX7lpIWkWhUSx71apVa06cOHHTx8fn19rsPmHChFvp6el/ocqeBr1eb9ZqxHOpcAwGw5yL1wEoSkhIEOfn55sal/wHPp/Pmzhx4jhKNC84OHjVjh07jg0aNMiaInsmSMfz8aSkpEVkVCsA8ePHj4vu3LlTTUneLU5OTsvPnz9val2YDmBo+DDbEEc+i46ednrRaNCodUSxQrPbKJY3Y8aMVZ9//nmf2WxlZQV/f/+TqampYVRZX/NcKpyZ6ACUA6hKTk6uowq7QyAQjDT6ymGz2cH79+//on///iyj+G5pb29HUVGRvLCwsLOlpYUq/hl0Oh2+vr6nzpw540a2chIApXv37n2k1/fSIv20R8x60qRJb1PjSeyc2czJC13446DrRReLDplK9/20lNocMobDYDCCDxw48MWAAQPMslksFneaYzODwYC3t/fZo0ePulFl5kCn07sbX/6C37LCAUAzgIaSkpJWqqA7uFzuQPIjDYD7jh071vv4+DhQkv2M+vp69YkTJ6JnzZr1vqen5xYvL69d48ePj9i+fXt2U1OTyb5cIBAIZ8yYsYv82gyg+OHDh48TExMbKUm7xcXFZd327dsFlGg6ALfw4bYh/azZrJ5bN0Cv1UOi1BjKQAfg/vnnn6/39fXt0WapVKo+ceJE1KxZs94fNWpUuJHNOc3NzT3a/NJLL/2DGv9b8FSThvj4eH+qAhPwAIStWbPmAVVPd8hksrNkPgcbG5uN5eXlGmoaYxoaGlqWLFkSSl6HS05WBgOYCeCDoKCgf9fX16up+QyoVCrN1atXR5DXHADgndDQ0HvUdKYoLi6mtnJ2Qib93coZg5XEy8N/OUEwDqHuRFPwkHtGee1Jm02Wl/jJZtnixYvndWNzMICtQUFBN3qyWaFQaKOjow2THrMnDRkZGWOMymqS37qF0wFQAzDp3zKGxfpPLzJ4yZIlM93c3Jg/T/Ff9Ho9cf/+/fUXL16MAaAAoCQnK5UAfgDwr/T09Dvh4eGp1LwGrKysmH5+fovJrw0ASmJiYgpEIpFZfjZnZ+dN5OZVkK3y4PeH24YMElpxemvdoCdQotQZWjcaALelS5fOcHNzM9mV6nQ64ocffngvKirq393YHE/aHBceHp5qavsal8tl+Pv7G8avfc5vXeGAn2Z2Zi1dKZVKg+PXPTg42Isi/hkqlUptY2PDzs7ODk1PT3/FEDIzMxfk5uaGxsfHB4SHh7uxWCyis7Oz+7sPwMbG5kXyoxZAMQDJ8ePHH1OSdYuNjc3IgoKC+eRXIYuOoKX9+H9GbxM6Jh0ylS51YnLVXTKGB8B9+vTp3pSUP6Orq0vt6OjIys3NnZeWlrbAOOTm5s6Li4vz37JlyxAGg9Gjzba2trPJjybTPG+ea5e6c+fOLKqe7igrK/uYdMaGi0SiDqrcGHMdtUQvaTs6Opp37dplR5aXDSCEwWCcLC4uVlDTdkdLS0sKmXdU+DCb60So+y+7T2qY507k/nlQqNF96k/a3E7V/6z0YnPTmTNnrAE4/NG6VCsAtuPGjTO5bGKMWq2WALARCoXOjo6OHKrcGBrNrEkT0EtaHo9n7+/vP4z8qgZQpNPpJBEREUWUpN1ia2s74ZtLl0IBeKweJJwOGtlBmgpWDLQqtPm+iVU3jNRYC4VCZwcHB55R3K+iJ5u5XK6Dr6/vENLePuW3rnDCwYMHjwoICDDls/oPXV1dOrFYLAJg5+DgYM3n802O3/oSOp0OoVBo7O+qAfDk+PHjj6RSqVkPxN3HL3wJH6EeQrZQr9ZDT6DbANDQodJ2SLq0WwAY+0x4zs7ONjY2Nj2eLegrGAwGrK2theQYu0/5LSscA8Dgjz/+eKadnZ3pnxuJXC4Xz5s3rxiADZPJZJuz6N9XCIVC4zOZSgDitra24osXLz4xijfJWPdhkwR/mkZ7JUO6NyRDevgXIV165OUM6aElovqP14lk0wITq+MoKqzodDqrp1aprxEIBJznMYYz1wI2gFkrVqzYGGHGbpGEhISA4ODgbKqMguPcuXP/fvny5Q18fu+HwUpKSv4xcuTIvwMY7+rquiYvL2+Vvb29yfJ3dHRoq6urOwHQn+UQCEEQNDabrba1tW0tLy9fGRgYmGQktgYQMnDgwNBHjx69amtr22vtfyQWf+czatRyALYm7rseQBs5u6Qywc3N7S2RSLTa1pa6v+C/GNlMe5bKabDZzs6upbOzc4mbm5sYwGpzdouIRCKfcePGPaLKnpW+njQwZ8+evam8vNysgbdcLu86e/asYRwVwGazD5eWlnZR0xkjFotLASwD8CaAFc8QVgEIBdDd+i0AjAXw2ddffy2hXrs71Gq1IjExcTBViZkEstnsI73ZXFhY+ATA0l9p8zzyPCqehx/OXJ6qwt29e3csVYERTocPH97bk/ORSkFBwddG+b0B7ExNTW2hpjOmra2tXiAQOJITE14PgUV278xuQk/NhB2AN0ePHn1dpVJRL98t1dXVB6hKzMQbwD9SUlJkVJ3GUGzmdxOe1uY+r3C9dgXPQn5+vooSJTh27Ni4+Pj4j0tKShLXr1+/xdnZ2aQD05i2trbOy5cvf2oU1QlAlpmZ2WAU9wsEAoEz2dJ2kd2UqUC7ffu2zf379wXUkJiYaJuYmGhwiVBpBVCcn59fdOfOnVqqsDvs7e1XnTx50pEabwYdAGRZWVk92mxtbe384MEDX9JmeTfBLJtzcnJM99v/R5jdwun1eqKuri6lrq7udm1t7R2pVJrc2NhYrlCY1Xv+gjt37rxHKYsNgDemTJkS25MviSAIoqmpKdfHx8fkAHHYsGGBNTU1j+RyeatCoWjuJshqamp+oOYzwgnAmilTpsTqdDrq5bulrKzsI6oSM3gam3N6snno0KHjerO5trY2nkzu0tctnLmYXeH6kpycnCvUgpCt8jQABxMTE5upeag0NzdnJCcnvxgQEGDsw7LduHHjW3l5eVXU9FSKi4t7WsymA5gK4EBiYqJZG0k7Ojqq/vKXvzytP81g8z8fPHjQY7dK/FTpMpKSkmZRbd68ebNZNkskkh1knj7vUs3l/7zCFRQUJJMLz90xEMB706ZNu2Nuy9LW1lbZ2Nj4UCqVJkil0gqNpsd1f4IgCKKzs7Nj9+7dhh0qphgA4N2wsLD71PymEIvFa6hKzMBg811zbW5vb69oampKlkqlCfX19eXm2NzR0dH+5ZdfDiCv+f9HhcvLy/uRz+f35AxmAJgCYNeBAwfE1Px9RXJy8hbqhbuBCWAWgKO5ubltVB3d0draWkDa8DQwDTbv37//edr8vtE1/9gVjtztcIH0c/WGLYDXABy7ePFir93E0yISiaKpF+yBIQA2rVu3LoWqxxQ5OTkLqErMwBbAqwCOPg+bc3JyoijX++NWuJKSkpp9+/a9S71wL/Qn/U5HDh8+3Ge/+pSUlIuka8FcrAC8zGazT0skEiVVX3fIZLJkqhIz6Wdsc2+TCHNJTk6OJJ+zMX+8CicWi+sjIiKOAOhxu1EPuJAO2h0LFy68m5aW1kq9hrmUl5fXnzlzJpx6ATMZCWDr1q1bc6h6TfH48eMQqhIzMdj82YIFC36tzdJz586ZGjr0eYXrybFpDBvAi6tXrw4/c+bMn6nCp6GtrQ1SqVRWUlJSeP/+/bsHDhy4CSCHslj9tAgAjCIdpF6vv/66X1hY2KjAwMD+gwYN6nENsqmpCRUVFcVpaWnff/DBB1/L5fJnPczLI3cSz9y7d+8Uf39/BxqNhu42OrJYLPWAAQMaNRrNd56envupcjPhA/A02Lx48WK/V1991TMwMLBfbzY3NjYSlZWVJSkpKd9/+OGHEXK53NSS1AAAi+Pj4z+eOXOmPVVoTHp6+tjx48dTD7//AtOl+jksANO9vLwWhIWFDXmaykGj0QiCILStra1ttbW1TWVlZeVZWVmFAIrIk1t9tSOBRnr/B5PB1draevC4ceNcx44d6+zq6mptb2/PYjKZRFdXV5dMJmt+8uRJaVpaWq5IJEoHUNIHZXEiK4F1D/eWTm4AyANQTxU+JTQA9gAGkTYPJG0eQLFZ39XVpW5paWkSi8WlGRkZeWba7AwgZPHixTO8vb251ANEBEHQHB0dGydPnixqaWm5NH369F7Pppi6KVToZPPqSBbwlz9b0+jJfVUqI093FzVRH8MhH7o12fqxyVkeyJ27CnLrdTu5cmH2D8gM6GbOQPW9POynhUO6kazJ1s+UzW3kqoM5NlsBcCX/6rt57jRSp1mn7ixYsGDBggULFixYsGChV8ydpf6M7777zn/MmDHjGQxG4/nz528BYAYFBY186aWXMg1pIiIihlhbW7OOHDlSt2/fvokCgYBLp9O1FRUVolmzZtUCYFy/fn2sh4dHf51OR9Pr9QyFQlE5YcKErJs3bwYMHz58gE6nowGQjR49+j/bu2/cuOHr7u4+XqPRlJ4+fTrxyJEjXbdu3Ro+dOhQd71ez2AymTSlUtnm5+eXhJ9eLCNISEiYw+Vy7VtaWlICAgLyTp48acPj8fq98cYbYlIt7cqVK14//vjjE1dXV86cOXMCWSwWh8PhaBQKRd7o0aOlAOjffvvtqDfeeONJaWlpf4FAMESn0yl1Oh2TIAirnJycR3Q63e2bb755dOXKFTUADBw4kHvq1KnRc+fOzTCU/8MPP7SbM2fOmCFDhrA1Gg1RUFBQGhoaWgYAFy9e9Pb09HQWCoUMADS5XN7p4+OTEh0d7eXl5eVibW1Na21trfbz8ys2lPvixYtjvL29nbhcLkOn09E6Oztbx40bl3bt2jVPb29vJx6Px9DpdDSFQtHq5eWV3dzcPEir1ba4uLgobt26FbRjx47c1NRUJQDcunXLp6Ojo3nRokU1sbGx/u7u7jYsFouuVqvpMpmsZsKECY9v3rwZ4OHhIeTz+br8/Pwn5LM0m6fZgEkDYB0TE7PHycnpSnZ29qSKioo358yZcysyMtLL1dX1pEgkmoGf9lyNDAgIuCoSiXynTJmyjMvlnissLFxQVVUVNmTIkH9FRkYGAxju7Oz8jUQieVssFoe1tLS80NTUNAaAp6Oj4zclJSVrnzx5sqCrq2tbUVHROQCCPXv2rBEKhRcfP34cqFKpPly4cOEGALZCofBQfX3934qKihZWVVW9RKfTvQDQly5dOuH7779PKi4u/kteXt44DocTceXKlTUVFRWTfH19PyPtsgbg5uHhcTgvL28Mn89fQRDEieLi4vmVlZVhAoHgcnx8/BwANl5eXgeXLl3qHxkZOSU5OXlxQ0PDqebm5o+kUum069evj2AwGOt27969x6A3IiLiVP/+/V8zun9OSqVy5cSJEx90dHRcJgjiwrRp00QpKSkbAAzy8vK65OTkFCeXyyMBnGMwGNsBeHl4eFyys7O7rVarz7u6uqY2NzdHDh482A6A5+jRoy/b2trGqlSqCywW6xyTyfwIwEgPD4/LPB7vdmtr6wW9Xn9Wr9d/QC6LxUkkksUAvAMDA1NiYmIukm6c/qNGjYp0cXF5G4Cnt7f3HTqd/u+Ojo4LLBbrtFarXcNisfz8/PziaTRajEqluhAUFJQnlUo/Ie3rcxyWL1++ISMjo9zLy2sygDEAPF9//fUXAQSHh4evqqqqSgPgcenSpdjU1NR/AnD75JNPjiQkJJwjNxEOj4uLO3fz5s0jAGbeu3cvBcBoo3VLGwCLExISkgH4GlYPHj9+nDFmzJiwK1euxMbExJwl/U0sLy+vwQCmxMfH3/3rX/8aQmmxvW/cuPHw2rVrhwEMBTBm4MCBo2fMmPHCunXr1hcXFxveVDkYQPCjR49+8PLymrdr1679cXFxB8mK6P7VV19tEIlEdwD4ZGVl3fnss8/+BMADwIRLly59GxMTs4zUM8HFxWW6WCxO/+KLLxa8+eabG0UiUcamTZsMJ74YACYvW7bs68bGxk4AiwHMi4yMvNLU1NQOYGlWVlZddHT0aQD+ZLkGAFiZnZ1dHx0d/Q2AhS+88MJfm5ub2x8+fHgWwILs7OzGY8eOHSJfnDiQ9JutzM3NbTx16tRe8v46k7Kw2tra6sjIyI8BbKioqOggCIK4ffv2fgCTiouLxf/617/+CeDNsrIy+cqVKz8gz244AHDncDjhlZWV8rVr1+4A8OqGDRs+USqVRFlZ2Sukjb1ibgvHBNA/ICAgiMfjXXz8+HES+aK+2kuXLt0FINu7d684ISEhNzU19Za7uzsRFhb2EQBbuVze4u3tHZiZmfnlDz/88NXAgQOnnTlz5ioAhp2dHVsikfytuLj4kEQiOblu3bqZAOrt7e21lZWVnqdOnQrYuXPnWq1Wq6iurpafP3/+1ogRI3yKi4u/LiwsXNfe3q4DoObxeKpt27ataGho2N3Y2Lj/9OnTrzAYjGGurq78xMTETwGUAaisrq6W3Lt3r0av13P1/3WbMwAw9Xo9aDQawWAwZIGBgfYKhcL3k08+menp6RlSXV2dBICr0+mwfPnyGgBVANT29vbqSZMmdZJ6pPX19dizZ0/EjBkz9q5du3ZtTU3NuoMHDxreo0sDwCQIQsnj8XRRUVFOGzZsCHR3d/dtbm4WA+jSarUdc+bMCa6rq/uyoaHhm1OnTr0GoIlGo2mDg4PzAGTcvXu3JiYm5vrIkSNfJsvUtmTJkpdqa2u/rK+vv3D8+PH5ANr0en17WFhYcE1Nzcc1NTXHdu/e/SIAPUEQenKowuvs7NRHRERETZo0aVNYWNiLcrm8U6vVMgDo6XS6as+ePcvq6ur2NjQ0RL/99tuBarW6mUajaRYtWpQGoOTQoUOPsrOzs52dnc1+r5y5FY4FgEun01uFQqHhFVRy0mutB1AIgPvGG2+k2tracmg02k7yhcU8tVrtUF5ero6MjJTGxsamVVdXx73zzjtLAFi3tLSwr1271hAdHf2ksLDwB4lEUg3AQaFQ9CsoKNjg4eHx4auvvrrw5MmTr7e0tOTfuHEj3cvL65MLFy5ktba2Tr9+/fppALy2tjb+/fv3286ePVv58OHD1Lt377YA0FtZWalXr15t2AHRRi4pqfV6Peh0uiFeB0BDEASrs7OT1dTUxJPJZJOTk5M3LV++fC+fz38QEhJy0LAZ1MbGxrBCoKUcP6wC0Hr27Nm6mpqaun79+t02HtMavPQKhYLFYDCsPTw8Ptq5c+e2kSNHds2dO3cFAAFBELyioqKG6OjonPz8/Ot5eXl5P53FphP29vbN5Etp2ng83hCNRtNO/lC4eXl50qioKNGjR4+u5+TkPAYg1Gq1VqWlpaykpCRIpdLq9vb2VkMPQJa7ic/nM7766quCb7/99tKhQ4feFwgE/dRqdRcArlarZSUmJlbGxMSkVlZWRotEoioWiwU6nU5Mnz69BoAIP60L99fr9T2/fM4IcyucDgD722+/TVKpVMElJSUvAsD777/vnJGRsf7+/ftW5MsFW5ubm8tGjBhhWCPUODg4aFksVtHBgweP7Nu3745IJKpxdHQMAqC0srKSx8bGXty2bdvpkJCQy3fu3GkAIFAoFMrXX3/9q6lTp37U1NQkXr9+/VwAVVFRUWP+/ve/O+3YsSP+nXfe+Y7L5Q4DQOPz+Wq5XH77o48+OhMaGnrl8uXLEp1OpxCLxbkODg5Hr1+/bg0AP/74Y8inn34ampaWVkin073FYrEvAO3mzZtfYDAYrIqKigo7Ozv7/Pz8tBdeeGHPkSNHPhk8ePD4sLCwAQC0DAaD3dXVRScfHI1Go7FpNJphGUsHoAJAW2traxmTySz5z90zgs/nc9vb24mxY8f+c+vWrRFsNnvQkiVLBgKQW1lZsTgczpPy8vLbEokkifznt3wGgwGCIMZkZ2dPOnPmzKqXX375z3FxcScAdHK5XBaNRhNLJJK4J0+eJAkEAgCgW1lZCRoaGq4sWrRoXUBAwKadO3emAeDT6XQO2YB0cblcmp+fH/3dd99NzMjIaHB3dx8gl8t1AFRWVlZMjUaTVVFREdfY2PjQ09OTS6fTWSwWi9nW1jb+6tWrL3733XcbRo0a5VRUVHSCaqcpzK1wagCN6enpsj179pyQyWS7kpOTYxcvXnxdrVaPbGhoMJyM6lSr1VKVSmXorugSiaRZKBT65ebmfp2UlLRr6tSpc6Oior4EoFOr1drDhw/vfPToUWRRUdG/r169ugVANYPBKJw/f34+gPZNmzYd7OrqWjZ//vypZWVl1vPmzVt37969nUePHt2Sm5t7HEBbR0eHbObMmW9JJJKvKyoqLhYUFGwAQLzyyitX0tLSdEOGDLmZkpJyUygUhnO53Nq8vLzWCxcuXOvs7Pw6ISHh/OLFi1/Mycn5BEAbjUar9PDwyAXQdOjQoSSxWJy/atWqdQBUHR0ddRqNRkNWOIZSqawjCMJ4wboLQLter2+Uy+XUA800/PRWpyatVlvg4OBQfPTo0djMzMyUhQsX/g2AVXl5eZWTk9PszZs3Xw4NDf1+xYoVewDwSkpK6urq6lbb2trGTZky5U+ZmZlvr1ix4hsA9hUVFXXDhw9/ZevWrdHz58+/tWzZst0AOrq6uh7PmDGjgDxdBnKtlZDL5U/kcnkrALpMJqv09vbOBNC8aNGiyyUlJcUEQSgBMOvq6uqmTp267r333rs6YcKE26tXr35XqVSiqqqqqbOzc7e/v/9lHx8fu/v37899mgPQT+MWYQMYQe465Y8fP95Zo9GUZGdnp5FyRwCDeDxenUKhqCd1ewEY1r9/f0a/fv3YLBarMz09PZV8MH/icrlcd3d3FoPBwIABAyq1Wm3bnTt3WNbW1rUdHR3NANx/ercKt9nFxcWpvLycA0A4adIk+/b29tz8/PwSAOMBWHt6ejL4fD6jX79+0vfee08ye/ZsOllemru7u6urq6vmwYMH98mudSwAPpvNtgsKCuIkJSUlA6gl02vJ1ppL7nGTTp06lfPgwQMbPp9fJ5fLG8mH5wZARr43zgCH/EcoSrKLNX7vHZ0cwNsLhcK89vZ2gixHw9SpU1UPHjzwAODAYrF0bDYbbDa7TCgU1lZUVHgCcGaxWBorKyt5Z2dnHjkmtQMwGYAjh8PRsVgsgslkSqytraWVlZX9AVSTLa4BF9K+UnJM7s7hcJ6oVKoqsszuAEo4HI5GpVL5A7DmcDg6JpOp5fF4YgCdDQ0NPgAEHA5HS6fTWxQKRRaAJqNr9MrTVDiQN01AFtiw+8MAg2yqjc+kcgy/LHKspwKgIdPySX0EGZTkA2Ib7SZhkDNYhcEtQ15bSQYaWTEMuyIIsjU25DccCNaRY05DBWCSdoCMN/yvdhZlF4dhnKchdRlso5E6tJQdFDQyj9bEThArMo1Bj2EGqyTLyTTS10XawiPjje8fyHvDpexMUZF5rMi/xjtCDAef1eRnttHzo5N5usjrG54NyO8q0h4eGa8zirNgwYIFCxYsWLBgwYIFCxYsWLBgwYIFCxYsWLBg4Y/I/wPI0LhNafBifwAAAABJRU5ErkJggg=='

$pictureLogo = New-Object System.Windows.Forms.PictureBox
$pictureLogo.Size = New-Object System.Drawing.Size(150, 60)
$pictureLogo.Location = New-Object System.Drawing.Point(720, 5)
$pictureLogo.SizeMode = 'Zoom'
$pictureLogo.BorderStyle = 'None'

$logoImage = Get-ImageFromBase64 -Base64 $logoBase64
if ($logoImage) {
    $pictureLogo.Image = $logoImage
}

# DataGridView para resultados
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(270, 160)
$grid.Size = New-Object System.Drawing.Size(600, 400)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.BorderStyle = 'FixedSingle'
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.EnableHeadersVisualStyles = $false

$headerStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
$headerStyle.BackColor = [System.Drawing.Color]::FromArgb(52, 73, 94)
$headerStyle.ForeColor = [System.Drawing.Color]::White
$headerStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersDefaultCellStyle = $headerStyle

$rowStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
$rowStyle.BackColor = [System.Drawing.Color]::White
$rowStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
$rowStyle.SelectionForeColor = [System.Drawing.Color]::White
$grid.DefaultCellStyle = $rowStyle

$altRowStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
$altRowStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$grid.AlternatingRowsDefaultCellStyle = $altRowStyle

# Barra de progreso para carga/busqueda
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(60, 685)
$progressBar.Size = New-Object System.Drawing.Size(500, 15)
$progressBar.Style = 'Continuous'
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Visible = $false

# Barra de progreso para instalacion/desinstalacion (parte inferior)
$progressBarInstall = New-Object System.Windows.Forms.ProgressBar
$progressBarInstall.Location = New-Object System.Drawing.Point(60, 685)
$progressBarInstall.Size = New-Object System.Drawing.Size(470, 15)
$progressBarInstall.Style = 'Continuous'
$progressBarInstall.Minimum = 0
$progressBarInstall.Maximum = 100
$progressBarInstall.Value = 0
$progressBarInstall.Visible = $false

# Label de estado
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Listo.'
$lblStatus.AutoSize = $true
$lblStatus.Location = New-Object System.Drawing.Point(10, 560)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(90, 98, 110)

# Agregar controles al formulario
$form.Controls.AddRange(@(
        $lblSearch,
        $txtSearch,
        $lblType,
        $cmbType,
        $btnLoadComputers,
        $lblComputers,
        $lstComputers,
        $btnSearch,
        $btnExportCsv,
        $btnCopySelected,
        $btnInstall,
        $btnUninstall,
        $pictureLogo,
        $chkUseWin32Product,
        $lblInstaller,
        $txtInstallerPath,
        $btnBrowseInstaller,
        $lblInstallArgs,
        $txtInstallArgs,
        $lblExpectedName,
        $txtExpectedName,
        $grid,
        $lblStatus,
        $progressBar,
        $progressBarInstall
    ))

# Evento: Cargar equipos
$btnLoadComputers.Add_Click({
        $lblStatus.Text = 'Cargando equipos del dominio...'
        $form.Refresh()

        $btnLoadComputers.Enabled = $false
        $btnSearch.Enabled = $false
        $lstComputers.Enabled = $false

        $progressBar.Style = 'Marquee'
        $progressBar.MarqueeAnimationSpeed = 30
        $progressBar.Visible = $true

        switch ($cmbType.SelectedItem) {
            'Servidores' { $type = 'Servers' }
            'Workstations' { $type = 'Workstations' }
            default { $type = 'All' }
        }

        $lstComputers.Items.Clear()
        $computers = Get-DomainComputers -Type $type

        if (-not $computers -or $computers.Count -eq 0) {
            $lblStatus.Text = 'No se encontraron equipos o hubo un error.'
        }
        else {
            foreach ($c in $computers) {
                [void]$lstComputers.Items.Add($c.Name)
            }

            $lblStatus.Text = "Equipos cargados: $($computers.Count)"
        }

        $progressBar.Visible = $false
        $progressBar.Style = 'Continuous'
        $progressBar.Value = 0

        $btnLoadComputers.Enabled = $true
        $btnSearch.Enabled = $true
        $lstComputers.Enabled = $true
    })

# Convertir lista de resultados a DataTable para el grid
function ConvertTo-DataTable {
    param([Parameter(Mandatory, ValueFromPipeline)][PSObject]$InputObject)
    begin {
        $dt = New-Object System.Data.DataTable
        $first = $true
    }
    process {
        $obj = $_
        if ($first) {
            foreach ($prop in $obj.PSObject.Properties.Name) {
                [void]$dt.Columns.Add($prop)            
            }
            $first = $false
        }
        $row = $dt.NewRow()
        foreach ($prop in $obj.PSObject.Properties.Name) {
            $row[$prop] = $obj.$prop
        }
        [void]$dt.Rows.Add($row)
    }
    end {
        return $dt
    }
}

function Test-ComputerOnline {
    param(
        [Parameter(Mandatory)][string]$ComputerName
    )

    try {
        return Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue
    }
    catch {
        return $false
    }
}

# Evento: Buscar software en equipos seleccionados
$btnSearch.Add_Click({
        $searchText = $txtSearch.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($searchText)) {
            [System.Windows.Forms.MessageBox]::Show('Escribe el nombre (o parte) del software a buscar.', 'Aviso', 'OK', 'Information') | Out-Null
            return
        }

        if ($lstComputers.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Selecciona al menos un equipo de la lista.', 'Aviso', 'OK', 'Information') | Out-Null
            return
        }

        $selectedCount = $lstComputers.SelectedItems.Count
        if ($selectedCount -gt 20) {
            $answer = [System.Windows.Forms.MessageBox]::Show(
                "Has seleccionado $selectedCount equipos. La búsqueda puede tardar varios minutos. ¿Deseas continuar?",
                'Confirmación',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }
        }

        $lblStatus.Text = 'Buscando en equipos seleccionados...'
        $form.Refresh()

        $btnSearch.Enabled = $false
        $btnLoadComputers.Enabled = $false
        $lstComputers.Enabled = $false

        $progressBar.Style = 'Continuous'
        $progressBar.Minimum = 0
        $progressBar.Maximum = $lstComputers.SelectedItems.Count
        $progressBar.Value = 0
        $progressBar.Visible = $true

        $useWin32 = $chkUseWin32Product.Checked

        $modulePath = Join-Path -Path $PSScriptRoot -ChildPath 'BuscarSoftwareRedModule.psm1'
        if (-not (Test-Path $modulePath)) {
            [System.Windows.Forms.MessageBox]::Show("No se encontró el módulo de lógica: $modulePath", 'Error', 'OK', 'Error') | Out-Null
            $btnSearch.Enabled = $true
            $btnLoadComputers.Enabled = $true
            $lstComputers.Enabled = $true
            return
        }

        $jobs = @()
        $offlineComputers = @()

        foreach ($item in $lstComputers.SelectedItems) {
            $computerName = [string]$item

            if (-not (Test-ComputerOnline -ComputerName $computerName)) {
                Write-BuscarSoftwareLog -Message "Equipo no accesible por ICMP (posible problema de DNS, apagado o firewall ICMP): $computerName"
                $offlineComputers += $computerName
                continue
            }

            $lblStatus.Text = "Creando tarea para $computerName..."
            $form.Refresh()

            $job = Start-Job -ScriptBlock {
                param($ComputerNameInner, $SearchTextInner, $UseWin32Inner, $ModulePathInner)

                Import-Module $ModulePathInner -ErrorAction Stop

                if ($UseWin32Inner) {
                    Get-InstalledSoftwareRemote -ComputerName $ComputerNameInner -SearchText $SearchTextInner -UseWin32Product
                }
                else {
                    Get-InstalledSoftwareRemote -ComputerName $ComputerNameInner -SearchText $SearchTextInner -UseWin32Product:$false
                }
            } -ArgumentList $computerName, $searchText, $useWin32, $modulePath

            $jobs += $job
        }

        $results = @()

        foreach ($job in $jobs) {
            $lblStatus.Text = "Esperando resultados de tareas en segundo plano..."
            $form.Refresh()

            $null = Wait-Job -Job $job -Timeout 300

            $res = Receive-Job -Job $job -ErrorAction SilentlyContinue
            if ($res) {
                $results += $res
            }

            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null

            if ($progressBar.Value -lt $progressBar.Maximum) {
                $progressBar.Value += 1
            }
        }

        if ($results.Count -eq 0) {
            $lblStatus.Text = 'Búsqueda finalizada. No se encontraron coincidencias.'
            $grid.DataSource = $null
        }
        else {
            $lblStatus.Text = "Búsqueda finalizada. Resultados: $($results.Count)"

            # Construir DataTable explícito con normalización de InstallDate
            $dt = New-Object System.Data.DataTable
            [void]$dt.Columns.Add('ComputerName')
            [void]$dt.Columns.Add('DisplayName')
            [void]$dt.Columns.Add('DisplayVersion')
            [void]$dt.Columns.Add('Publisher')
            [void]$dt.Columns.Add('InstallDate')

            foreach ($r in $results) {
                $installDateText = ''

                if ($null -ne $r.InstallDate -and $r.InstallDate -ne '') {
                    try {
                        $dtParsed = [datetime]$r.InstallDate
                        $installDateText = $dtParsed.ToString('yyyy-MM-dd')
                    }
                    catch {
                        $installDateText = [string]$r.InstallDate
                    }
                }

                $row = $dt.NewRow()
                $row['ComputerName'] = [string]$r.ComputerName
                $row['DisplayName'] = [string]$r.DisplayName
                $row['DisplayVersion'] = [string]$r.DisplayVersion
                $row['Publisher'] = [string]$r.Publisher
                $row['InstallDate'] = $installDateText
                [void]$dt.Rows.Add($row)
            }

            $grid.DataSource = $null
            $grid.DataSource = $dt
        }

        $progressBar.Visible = $false
        $progressBar.Value = 0

        if ($offlineComputers.Count -gt 0) {
            $mensajeOffline = "Los siguientes equipos no están accesibles (offline o sin respuesta ICMP):`r`n`r`n" + ($offlineComputers -join "`r`n")
            [System.Windows.Forms.MessageBox]::Show($mensajeOffline, 'Equipos no accesibles', 'OK', 'Warning') | Out-Null
        }

        $btnSearch.Enabled = $true
        $btnLoadComputers.Enabled = $true
        $lstComputers.Enabled = $true
    })

function Copy-CurrentGridRowToClipboard {
    if (-not $grid.CurrentRow) { return }

    $values = @()
    foreach ($cell in $grid.CurrentRow.Cells) {
        $values += [string]$cell.Value
    }

    $text = [string]::Join("`t", $values)
    [System.Windows.Forms.Clipboard]::SetText($text)
}

$btnExportCsv.Add_Click({
        if (-not $grid.DataSource -or -not $grid.Rows.Count) {
            [System.Windows.Forms.MessageBox]::Show('No hay resultados para exportar.', 'Aviso', 'OK', 'Information') | Out-Null
            return
        }

        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = 'CSV (*.csv)|*.csv|Todos los archivos (*.*)|*.*'
        $dialog.Title = 'Guardar resultados como CSV'
        $dialog.FileName = 'Resultados_BuscarSoftware.csv'

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $path = $dialog.FileName
            $dataTable = [System.Data.DataTable]$grid.DataSource
            $dataTable | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Resultados exportados a: $path", 'Información', 'OK', 'Information') | Out-Null
        }
    })

$btnCopySelected.Add_Click({
        if (-not $grid.CurrentRow) {
            [System.Windows.Forms.MessageBox]::Show('No hay ninguna fila seleccionada.', 'Aviso', 'OK', 'Information') | Out-Null
            return
        }

        Copy-CurrentGridRowToClipboard
        [System.Windows.Forms.MessageBox]::Show('Datos de la fila copiados al portapapeles.', 'Información', 'OK', 'Information') | Out-Null
    })

$btnBrowseInstaller.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Instaladores (*.exe;*.msi)|*.exe;*.msi|Todos los archivos (*.*)|*.*'
        $dialog.Title = 'Seleccionar instalador'

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtInstallerPath.Text = $dialog.FileName
        }
    })

$btnInstall.Add_Click({
        if ($lstComputers.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Selecciona al menos un equipo para instalar el software.', 'Aviso', 'OK', 'Information') | Out-Null
            return
        }

        $installerPath = $txtInstallerPath.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($installerPath)) {
            [System.Windows.Forms.MessageBox]::Show('Selecciona o escribe la ruta del instalador.', 'Aviso', 'OK', 'Information') | Out-Null
            return
        }

        $silentArgs = $txtInstallArgs.Text.Trim()

        $customCommand = $null

        $expectedDisplayName = $txtExpectedName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($expectedDisplayName)) {
            $expectedDisplayName = $txtSearch.Text.Trim()
        }
        $txtExpectedName.Text = $expectedDisplayName

        $msg = "Se intentara instalar el software de forma desatendida en los equipos seleccionados.\n\n" +
        "Instalador: $installerPath\n" +
        "Parametros: $silentArgs\n\n" +
        "Nota: Se probaran varios metodos (Invoke-Command, copia remota, WMI)." 

        $answer = [System.Windows.Forms.MessageBox]::Show($msg, 'Confirmar instalación', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $btnInstall.Enabled = $false
        $btnSearch.Enabled = $false
        $btnLoadComputers.Enabled = $false
        $lstComputers.Enabled = $false

        $lblStatus.Text = 'Iniciando instalación desatendida en equipos seleccionados...'
        $form.Refresh()

        $progressBar.Style = 'Continuous'
        $progressBar.Minimum = 0
        $progressBar.Maximum = $lstComputers.SelectedItems.Count
        $progressBar.Value = 0
        $progressBar.Visible = $true

        $ok = 0
        $fail = 0

        foreach ($item in $lstComputers.SelectedItems) {
            $computerName = [string]$item
            $lblStatus.Text = "Instalando en $computerName..."
            $form.Refresh()

            $result = Invoke-RemoteInstall -ComputerName $computerName -InstallerPath $installerPath -SilentArgs $silentArgs -CustomCommand $customCommand -ExpectedDisplayName $expectedDisplayName
            if ($result) { $ok++ } else { $fail++ }

            if ($progressBarInstall.Value -lt $progressBarInstall.Maximum) {
                $progressBarInstall.Value += 1
            }
        }

        $progressBarInstall.Visible = $false
        $progressBarInstall.Value = 0

        $lblStatus.Text = "Instalacion finalizada. Exitosas: $ok. Fallidas: $fail."

        $btnInstall.Enabled = $true
        $btnSearch.Enabled = $true
        $btnLoadComputers.Enabled = $true
        $lstComputers.Enabled = $true

        [System.Windows.Forms.MessageBox]::Show("Instalacion finalizada. Exitosas: $ok. Fallidas: $fail.", 'Resultado de instalación', 'OK', 'Information') | Out-Null
    })

$btnUninstall.Add_Click({
        if (-not $grid.DataSource -or -not $grid.Rows.Count) {
            [System.Windows.Forms.MessageBox]::Show('No hay resultados en la tabla de software para desinstalar.', 'Aviso', 'OK', 'Information') | Out-Null
            return
        }

        if ($grid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Selecciona al menos una fila en la tabla de resultados para desinstalar.', 'Aviso', 'OK', 'Information') | Out-Null
            return
        }

        $count = $grid.SelectedRows.Count
        $msg = "Se intentara desinstalar de forma desatendida el software de las filas seleccionadas.\n\n" +
        "Filas seleccionadas: $count\n\n" +
        "Se usaran metodos en cascada (QuietUninstallString, UninstallString, Win32_Product)." 

        $answer = [System.Windows.Forms.MessageBox]::Show($msg, 'Confirmar desinstalación', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $btnUninstall.Enabled = $false
        $btnInstall.Enabled = $false
        $btnSearch.Enabled = $false
        $btnLoadComputers.Enabled = $false
        $lstComputers.Enabled = $false

        $lblStatus.Text = 'Iniciando desinstalación en equipos seleccionados...'
        $form.Refresh()

        $progressBarInstall.Style = 'Continuous'
        $progressBarInstall.Minimum = 0
        $progressBarInstall.Maximum = $grid.SelectedRows.Count
        $progressBarInstall.Value = 0
        $progressBarInstall.Visible = $true

        $ok = 0
        $fail = 0

        foreach ($row in $grid.SelectedRows) {
            $computerName = [string]$row.Cells['ComputerName'].Value
            $displayName = [string]$row.Cells['DisplayName'].Value

            if ([string]::IsNullOrWhiteSpace($computerName) -or [string]::IsNullOrWhiteSpace($displayName)) {
                continue
            }

            $lblStatus.Text = "Desinstalando '$displayName' en $computerName..."
            $form.Refresh()

            $result = Invoke-RemoteUninstall -ComputerName $computerName -DisplayName $displayName
            if ($result) { $ok++ } else { $fail++ }

            if ($progressBar.Value -lt $progressBar.Maximum) {
                $progressBar.Value += 1
            }
        }

        $progressBar.Visible = $false
        $progressBar.Value = 0

        $lblStatus.Text = "Desinstalacion finalizada. Exitosas: $ok. Fallidas: $fail."

        $btnUninstall.Enabled = $true
        $btnInstall.Enabled = $true
        $btnSearch.Enabled = $true
        $btnLoadComputers.Enabled = $true
        $lstComputers.Enabled = $true

        [System.Windows.Forms.MessageBox]::Show("Desinstalacion finalizada. Exitosas: $ok. Fallidas: $fail.", 'Resultado de desinstalación', 'OK', 'Information') | Out-Null
    })

$grid.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0) {
            Copy-CurrentGridRowToClipboard
        }
    })

[void]$form.ShowDialog()
