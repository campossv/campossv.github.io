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