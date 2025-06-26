function Get-AnalisisConfiabilidad {
    try {
        Write-Progress -Activity "Analizando confiabilidad del sistema" -PercentComplete 35
        Write-Host "   Analizando registros de confiabilidad..." -ForegroundColor Yellow

        $datos = @{}

        try {
            $reliabilityRecords = Get-CimInstance -ClassName Win32_ReliabilityRecords -ErrorAction SilentlyContinue |
                                 Where-Object { $_.TimeGenerated -gt (Get-Date).AddDays(-30) } |
                                 Sort-Object TimeGenerated -Descending |
                                 Select-Object -First 100

            if ($reliabilityRecords) {
                Write-Host "   Se encontraron $($reliabilityRecords.Count) registros de confiabilidad" -ForegroundColor Yellow

                $eventosConfiabilidad = $reliabilityRecords | ForEach-Object {
                    $tipoEvento = switch ($_.EventIdentifier) {
                        1001 { "Inicio de aplicación" }
                        1002 { "Fallo de aplicación" }
                        1003 { "Cuelgue de aplicación" }
                        1004 { "Instalación exitosa" }
                        1005 { "Fallo de instalación" }
                        1006 { "Inicio del sistema" }
                        1007 { "Apagado del sistema" }
                        1008 { "Fallo del sistema" }
                        default { "Evento ID: $($_.EventIdentifier)" }
                    }

                    [PSCustomObject]@{
                        Fecha = $_.TimeGenerated
                        TipoEvento = $tipoEvento
                        Fuente = $_.SourceName
                        Descripcion = if ($_.Message) { $_.Message.Substring(0, [Math]::Min(200, $_.Message.Length)) } else { "N/A" }
                        EventID = $_.EventIdentifier
                        Criticidad = switch ($_.EventIdentifier) {
                            { $_ -in @(1002, 1003, 1005, 1008) } { "Alto" }
                            { $_ -in @(1001, 1004, 1006, 1007) } { "Normal" }
                            default { "Medio" }
                        }
                    }
                }

                $datos.EventosConfiabilidad = $eventosConfiabilidad

                $fallosAplicacion = ($eventosConfiabilidad | Where-Object { $_.EventID -in @(1002, 1003) }).Count
                $fallosSistema = ($eventosConfiabilidad | Where-Object { $_.EventID -eq 1008 }).Count
                $reinicios = ($eventosConfiabilidad | Where-Object { $_.EventID -in @(1006, 1007) }).Count

                $datos.EstadisticasEstabilidad = @{
                    FallosAplicacion = $fallosAplicacion
                    FallosSistema = $fallosSistema
                    ReiniciosDetectados = $reinicios
                    TotalEventos = $eventosConfiabilidad.Count
                    PeriodoAnalisis = "Últimos 30 días"
                    IndiceEstabilidad = switch ($true) {
                        { $fallosSistema -gt 5 -or $fallosAplicacion -gt 20 } { "Baja" }
                        { $fallosSistema -gt 2 -or $fallosAplicacion -gt 10 } { "Media" }
                        default { "Alta" }
                    }
                }

                $tendenciasSemana = $eventosConfiabilidad |
                    Group-Object { (Get-Date $_.Fecha).ToString("yyyy-MM-dd") } |
                    Sort-Object Name -Descending |
                    Select-Object -First 7 |
                    ForEach-Object {
                        $fallosDia = ($_.Group | Where-Object { $_.Criticidad -eq "Alto" }).Count
                        [PSCustomObject]@{
                            Fecha = $_.Name
                            TotalEventos = $_.Count
                            EventosCriticos = $fallosDia
                            Estabilidad = if ($fallosDia -eq 0) { "Estable" } elseif ($fallosDia -lt 3) { "Moderada" } else { "Inestable" }
                        }
                    }

                $datos.TendenciasSemanales = $tendenciasSemana

            } else {
                Write-Host "   No se encontraron registros de confiabilidad o no están disponibles" -ForegroundColor Yellow
                $datos.EventosConfiabilidad = @()
                $datos.EstadisticasEstabilidad = @{ Error = "No hay datos de confiabilidad disponibles" }
                $datos.TendenciasSemanales = @()
            }

        } catch {
            Write-Host "   Error al acceder a registros de confiabilidad: $_" -ForegroundColor Red

            try {
                Write-Host "   Intentando método alternativo con Event Log..." -ForegroundColor Yellow
                $eventosAlternativos = Get-WinEvent -FilterHashtable @{
                    LogName = 'System'
                    Level = 1,2
                    StartTime = (Get-Date).AddDays(-7)
                } -MaxEvents 50 -ErrorAction SilentlyContinue |
                Where-Object { $_.Id -in @(1001, 1074, 6005, 6006, 6008, 6009, 6013) } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Fecha = $_.TimeCreated
                        TipoEvento = switch ($_.Id) {
                            1074 { "Apagado del sistema" }
                            6005 { "Inicio del servicio Event Log" }
                            6006 { "Parada del servicio Event Log" }
                            6008 { "Apagado inesperado" }
                            6009 { "Información de versión del procesador" }
                            6013 { "Tiempo de actividad del sistema" }
                            default { "Evento del sistema" }
                        }
                        EventID = $_.Id
                        Fuente = $_.ProviderName
                        Criticidad = if ($_.Id -eq 6008) { "Alto" } else { "Normal" }
                    }
                }

                $datos.EventosConfiabilidad = $eventosAlternativos
                $datos.EstadisticasEstabilidad = @{
                    Metodo = "Event Log alternativo"
                    ApagadosInesperados = ($eventosAlternativos | Where-Object { $_.EventID -eq 6008 }).Count
                    TotalEventos = $eventosAlternativos.Count
                    PeriodoAnalisis = "Últimos 7 días"
                }

            } catch {
                $datos.EventosConfiabilidad = @()
                $datos.EstadisticasEstabilidad = @{ Error = "No se pudo obtener información de confiabilidad: $_" }
            }
        }

        return $datos

    } catch {
        Write-Warning "Error en análisis de confiabilidad: $_"
        return @{ Error = $_.Exception.Message }
    }
}
