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