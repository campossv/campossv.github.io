function Get-LogsEventos {
    param([int]$Dias)

    try {
        Write-Progress -Activity "Recopilando logs de los 3 tipos principales" -PercentComplete 25

        $fechaInicio = (Get-Date).AddDays(-$Dias)
        Write-Host "   Período de logs: desde $($fechaInicio) hasta $(Get-Date)" -ForegroundColor Yellow

        Write-Host "   Recopilando logs del Sistema..." -ForegroundColor Yellow
        $logsSistema = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Level = 1,2,3
            StartTime = $fechaInicio
        } -MaxEvents 200 -ErrorAction SilentlyContinue |
        Select-Object TimeCreated, LevelDisplayName, ProviderName, Id,
                     @{Name="Message";Expression={$_.Message.Substring(0,[Math]::Min(300,$_.Message.Length))}}

        Write-Host "   Recopilando logs de Aplicación..." -ForegroundColor Yellow
        $logsAplicacion = Get-WinEvent -FilterHashtable @{
            LogName = 'Application'
            Level = 1,2,3
            StartTime = $fechaInicio
        } -MaxEvents 200 -ErrorAction SilentlyContinue |
        Select-Object TimeCreated, LevelDisplayName, ProviderName, Id,
                     @{Name="Message";Expression={$_.Message.Substring(0,[Math]::Min(300,$_.Message.Length))}}

        Write-Host "   Recopilando logs de Seguridad..." -ForegroundColor Yellow
        try {
            $eventosSeguridad = Get-WinEvent -FilterHashtable @{
                LogName = 'Security'
                ID = 4624, 4625, 4634, 4647, 4648, 4740, 4767, 4771, 4776, 4616, 4720, 4722, 4724, 4738
                StartTime = $fechaInicio
            } -MaxEvents 150 -ErrorAction Stop

            Write-Host "   Se encontraron $($eventosSeguridad.Count) eventos de seguridad" -ForegroundColor Yellow


            $logsSeguridad = $eventosSeguridad | ForEach-Object {
                $evento = $_
                $cuenta = "N/A"
                $ip = "N/A"

                try { $cuenta = if ($evento.Properties.Count -gt 5) { $evento.Properties[5].Value } elseif ($evento.Properties.Count -gt 1) { $evento.Properties[1].Value }
                    if ([string]::IsNullOrEmpty($cuenta)) {
                        for ($i = 0; $i -lt [Math]::Min(10, $evento.Properties.Count); $i++) {
                            if (![string]::IsNullOrEmpty($evento.Properties[$i].Value)) { $cuenta = $evento.Properties[$i].Value; break }
                        }
                    }
                } catch { $cuenta = "No disponible" }

                $ip = if ($evento.Properties.Count -gt 19) { $evento.Properties[19].Value } elseif ($evento.Properties.Count -gt 9) { $evento.Properties[9].Value }

                try { $rawMessage = $evento.Message; if ($rawMessage.Length -gt 150) { $rawMessage = $rawMessage.Substring(0, 150) + "..." }
                } catch { $rawMessage = "[No disponible]" }

                $tipoEvento = switch ($evento.Id) {
                    4625 { "Login Fallido" }
                    4648 { "Login Explícito" }
                    4771 { "Kerberos Fallido" }
                    4776 { "Validación" }
                    4740 { "Cuenta Bloqueada" }
                    4767 { "Desbloqueada" }
                    4624 { "Login" }
                    4634 { "Logout" }
                    4647 { "Logout Usuario" }
                    4616 { "Cambio hora" }
                    4720 { "Cuenta creada" }
                    4722 { "Cuenta habilitada" }
                    4724 { "Cambio pwd" }
                    4738 { "Cambio cuenta" }
                    default { "ID:$($evento.Id)" }
                }

[PSCustomObject]@{TimeCreated=$evento.TimeCreated;Id=$evento.Id;Account=$cuenta;SourceIP=$ip;EventType=$tipoEvento;RawMessage=$rawMessage}
            }
        } catch {
            Write-Host "   Error al obtener logs de seguridad: $_" -ForegroundColor Red
            $logsSeguridad = @()
        }

        Write-Host "   Total eventos encontrados - Sistema: $($logsSistema.Count), Aplicación: $($logsAplicacion.Count), Seguridad: $($logsSeguridad.Count)" -ForegroundColor Yellow

        return @{
            LogsSistema = $logsSistema
            LogsAplicacion = $logsAplicacion
            LogsSeguridad = $logsSeguridad
        }
    } catch {
        Write-Warning "Error al obtener logs de eventos: $_"
        return @{ Error = $_.Exception.Message }
    }
}
