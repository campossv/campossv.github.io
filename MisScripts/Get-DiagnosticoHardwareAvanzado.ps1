function Get-DiagnosticoHardwareAvanzado {
    try {
        Write-Progress -Activity "Realizando diagnóstico avanzado de hardware" -PercentComplete 45
        Write-Host "   Analizando hardware avanzado..." -ForegroundColor Yellow

        $datos = @{}

        Write-Host "   Verificando estado SMART de discos..." -ForegroundColor Yellow
        try {
            $discosSMART = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue | ForEach-Object {
                $disco = $_
                $smartData = $null

                try {

                    $smartData = Get-CimInstance -ClassName MSStorageDriver_FailurePredictStatus -Namespace "root\wmi" -ErrorAction SilentlyContinue |
                                Where-Object { $_.InstanceName -like "*$($disco.PNPDeviceID)*" }

                    if (-not $smartData) {

                        $smartData = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue |
                                    Where-Object { $_.DeviceID -eq $disco.DeviceID } |
                                    ForEach-Object {
                                        [PSCustomObject]@{
                                            PredictFailure = $false
                                            Reason = "Datos SMART no disponibles"
                                        }
                                    }
                    }
                } catch {
                    $smartData = [PSCustomObject]@{
                        PredictFailure = $null
                        Reason = "Error al acceder a SMART: $_"
                    }
                }

                [PSCustomObject]@{
                    Modelo = $disco.Model
                    NumeroSerie = $disco.SerialNumber
                    TamanoGB = [math]::Round($disco.Size / 1GB, 2)
                    Interfaz = $disco.InterfaceType
                    EstadoSMART = if ($smartData.PredictFailure -eq $true) { "FALLO PREDICHO" }
                                 elseif ($smartData.PredictFailure -eq $false) { "Saludable" }
                                 else { "No disponible" }
                    DetallesSMART = $smartData.Reason
                    Particiones = (Get-CimInstance -ClassName Win32_DiskPartition -ErrorAction SilentlyContinue |
                                  Where-Object { $_.DiskIndex -eq $disco.Index }).Count
                    Estado = switch ($smartData.PredictFailure) {
                        $true { "Crítico" }
                        $false { "Normal" }
                        default { "Desconocido" }
                    }
                }
            }

            $datos.DiscosSMART = $discosSMART

        } catch {
            Write-Host "   Error al obtener datos SMART: $_" -ForegroundColor Red
            $datos.DiscosSMART = @{ Error = "No se pudo obtener información SMART: $_" }
        }

        Write-Host "   Verificando temperaturas de componentes..." -ForegroundColor Yellow
        try {
            $temperaturas = @()

            $tempCPU = Get-CimInstance -ClassName Win32_PerfRawData_Counters_ThermalZoneInformation -ErrorAction SilentlyContinue |
                      Where-Object { $_.Name -like "*CPU*" -or $_.Name -like "*Processor*" } |
                      ForEach-Object {
                          $tempKelvin = $_.Temperature
                          $tempCelsius = if ($tempKelvin -and $tempKelvin -gt 0) {
                              [math]::Round(($tempKelvin / 10) - 273.15, 1)
                          } else { $null }

                          [PSCustomObject]@{
                              Componente = "CPU - $($_.Name)"
                              TemperaturaC = $tempCelsius
                              Estado = if ($tempCelsius -gt 80) { "Crítico" }
                                      elseif ($tempCelsius -gt 70) { "Alto" }
                                      elseif ($tempCelsius -gt 0) { "Normal" }
                                      else { "No disponible" }
                          }
                      }

            if ($tempCPU) { $temperaturas += $tempCPU }

            try {
                $wmiTemp = Get-CimInstance -ClassName MSAcpi_ThermalZoneTemperature -Namespace "root\wmi" -ErrorAction SilentlyContinue |
                          ForEach-Object {
                              $tempKelvin = $_.CurrentTemperature
                              $tempCelsius = if ($tempKelvin -and $tempKelvin -gt 0) {
                                  [math]::Round(($tempKelvin / 10) - 273.15, 1)
                              } else { $null }

                              [PSCustomObject]@{
                                  Componente = "Zona Térmica - $($_.InstanceName)"
                                  TemperaturaC = $tempCelsius
                                  Estado = if ($tempCelsius -gt 80) { "Crítico" }
                                          elseif ($tempCelsius -gt 70) { "Alto" }
                                          elseif ($tempCelsius -gt 0) { "Normal" }
                                          else { "No disponible" }
                              }
                          }

                if ($wmiTemp) { $temperaturas += $wmiTemp }

            } catch {
                Write-Host "   Método WMI para temperaturas no disponible" -ForegroundColor Yellow
            }

            if ($temperaturas.Count -eq 0) {
                $temperaturas = @([PSCustomObject]@{
                    Componente = "Sistema"
                    TemperaturaC = $null
                    Estado = "Sensores de temperatura no disponibles"
                })
            }

            $datos.Temperaturas = $temperaturas

        } catch {
            Write-Host "   Error al obtener temperaturas: $_" -ForegroundColor Red
            $datos.Temperaturas = @{ Error = "No se pudo obtener información de temperatura: $_" }
        }

        Write-Host "   Verificando estado de la batería..." -ForegroundColor Yellow
        try {
            $baterias = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue

            if ($baterias) {
                $estadoBaterias = $baterias | ForEach-Object {
                    $bateria = $_

                    $batteryStatus = Get-CimInstance -ClassName BatteryStatus -Namespace "root\wmi" -ErrorAction SilentlyContinue |
                                    Where-Object { $_.InstanceName -like "*$($bateria.DeviceID)*" }

                    $estadoCarga = switch ($bateria.BatteryStatus) {
                        1 { "Desconocido" }
                        2 { "Cargando" }
                        3 { "Descargando" }
                        4 { "Crítico" }
                        5 { "Bajo" }
                        6 { "Cargando y Alto" }
                        7 { "Cargando y Bajo" }
                        8 { "Cargando y Crítico" }
                        9 { "Indefinido" }
                        10 { "Parcialmente Cargado" }
                        11 { "Completamente Cargado" }
                        default { "Estado $($bateria.BatteryStatus)" }
                    }

                    [PSCustomObject]@{
                        Nombre = $bateria.Name
                        Fabricante = $bateria.Manufacturer
                        EstadoCarga = $estadoCarga
                        PorcentajeCarga = $bateria.EstimatedChargeRemaining
                        TiempoRestante = if ($bateria.EstimatedRunTime -and $bateria.EstimatedRunTime -ne 71582788) {
                            "$([math]::Round($bateria.EstimatedRunTime / 60, 1)) horas"
                        } else { "Calculando..." }
                        CapacidadDiseño = if ($bateria.DesignCapacity) { "$($bateria.DesignCapacity) mWh" } else { "N/A" }
                        CapacidadCompleta = if ($bateria.FullChargeCapacity) { "$($bateria.FullChargeCapacity) mWh" } else { "N/A" }
                        SaludBateria = if ($bateria.DesignCapacity -and $bateria.FullChargeCapacity) {
                            $salud = [math]::Round(($bateria.FullChargeCapacity / $bateria.DesignCapacity) * 100, 1)
                            "$salud%"
                        } else { "No disponible" }
                        Estado = switch ($true) {
                            { $bateria.BatteryStatus -in @(4, 8) } { "Crítico" }
                            { $bateria.BatteryStatus -in @(5, 7) } { "Bajo" }
                            { $bateria.EstimatedChargeRemaining -lt 20 } { "Advertencia" }
                            default { "Normal" }
                        }
                    }
                }

                $datos.Baterias = $estadoBaterias

            } else {
                $datos.Baterias = @([PSCustomObject]@{
                    Estado = "No se detectaron baterías (sistema de escritorio)"
                })
            }

        } catch {
            Write-Host "   Error al obtener estado de batería: $_" -ForegroundColor Red
            $datos.Baterias = @{ Error = "No se pudo obtener información de batería: $_" }
        }

        Write-Host "   Recopilando información adicional de hardware..." -ForegroundColor Yellow
        try {
            $memoriaFisica = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue | ForEach-Object {
                [PSCustomObject]@{
                    Ubicacion = $_.DeviceLocator
                    Capacidad = "$([math]::Round($_.Capacity / 1GB, 2)) GB"
                    Velocidad = "$($_.Speed) MHz"
                    Fabricante = $_.Manufacturer
                    NumeroSerie = $_.SerialNumber
                    TipoMemoria = switch ($_.MemoryType) {
                        20 { "DDR" }
                        21 { "DDR2" }
                        22 { "DDR2 FB-DIMM" }
                        24 { "DDR3" }
                        26 { "DDR4" }
                        default { "Tipo $($_.MemoryType)" }
                    }
                }
            }

            $datos.MemoriaFisica = $memoriaFisica

            $ventiladores = Get-CimInstance -ClassName Win32_Fan -ErrorAction SilentlyContinue | ForEach-Object {
                [PSCustomObject]@{
                    Nombre = $_.Name
                    Descripcion = $_.Description
                    Estado = switch ($_.Status) {
                        "OK" { "Normal" }
                        "Error" { "Error" }
                        "Degraded" { "Degradado" }
                        default { $_.Status }
                    }
                    Activo = $_.ActiveCooling
                }
            }

            if ($ventiladores.Count -eq 0) {
                $ventiladores = @([PSCustomObject]@{
                    Estado = "No se detectaron ventiladores o información no disponible"
                })
            }

            $datos.Ventiladores = $ventiladores

        } catch {
            Write-Host "   Error al obtener información adicional de hardware: $_" -ForegroundColor Red
            $datos.MemoriaFisica = @{ Error = "No disponible" }
            $datos.Ventiladores = @{ Error = "No disponible" }
        }

        return $datos

    } catch {
        Write-Warning "Error en diagnóstico de hardware avanzado: $_"
        return @{ Error = $_.Exception.Message }
    }
}