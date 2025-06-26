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