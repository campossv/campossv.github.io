function Generate-CompleteHTML {
    param($InfoSistema, $MetricasRendimiento, $LogsEventos, $DatosExtendidos, $AnalisisConfiabilidad, $DiagnosticoHardware, $AnalisisRoles, $AnalisisPoliticas, $VerificacionCumplimiento, $AnalisisPermisos, $AuditoriaSoftware)


    $nombreServidor = $InfoSistema.NombreServidor
    $sistemaOperativo = $InfoSistema.NombreSO


    $ms = New-Object System.IO.MemoryStream
    $imagenl.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $imagenBytes = $ms.ToArray()
    $ms.Close()
    $imagenBase64 = [Convert]::ToBase64String($imagenBytes)


    $cssStyles = @'
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0; padding: 20px;
            background: linear-gradient(135deg, rgb(102,126,234) 0%, rgb(118,75,162) 100%);
            min-height: 100vh;
        }
        .container {
            max-width: 1400px; margin: 0 auto;
            background: white; padding: 30px;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        }
        h1 {
            color: white; text-align: center; margin-bottom: 30px;
            border-bottom: 4px solid rgb(255,255,255); padding-bottom: 20px;
            font-size: 2.5em; text-shadow: 2px 2px 6px rgba(0,0,0,0.8);
        }
        h2 {
            color: white; margin: 40px 0 20px 0; padding: 20px;
            background: linear-gradient(135deg, rgb(52,152,219), rgb(41,128,185));
            border-radius: 10px; font-size: 1.5em;
            box-shadow: 0 4px 15px rgba(52, 152, 219, 0.3);
        }
        h3 {
            color: rgb(44,62,80); margin: 30px 0 15px 0;
            border-left: 6px solid rgb(52,152,219); padding-left: 20px;
            font-size: 1.3em;
        }
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 25px; margin: 30px 0;
        }
        .metric-card {
            background: linear-gradient(135deg, rgb(248,249,250), rgb(233,236,239));
            padding: 25px; border-radius: 12px;
            border-left: 6px solid rgb(52,152,219);
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s ease;
        }
        .metric-card:hover { transform: translateY(-5px); }
        .metric-card h4 { margin-top: 0; color: rgb(44,62,80); font-size: 1.2em; }
        .metric-value { font-size: 2em; font-weight: bold; color: rgb(52,152,219); margin: 10px 0; }
        .status-indicator {
            display: inline-block; padding: 5px 15px;
            border-radius: 20px; color: white; font-weight: bold;
        }
        .status-normal { background-color: rgb(39,174,96); }
        .status-warning { background-color: rgb(243,156,18); }
        .status-critical { background-color: rgb(231,76,60); }
        .status-high { background-color: rgb(230,126,34); }
        .status-medium { background-color: rgb(241,196,15); color: rgb(44,62,80); }

        table {
            border-collapse: collapse; width: 100%; margin: 20px 0;
            background: white; border-radius: 10px; overflow: hidden;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        th {
            background: linear-gradient(135deg, rgb(52,152,219), rgb(41,128,185));
            color: white; padding: 15px; text-align: left; font-weight: 600;
        }
        td { padding: 12px 15px; border-bottom: 1px solid rgb(236,240,241); }
        tr:nth-child(even) { background-color: rgb(248,249,250); }
        tr:hover { background-color: rgb(214,234,248); }

        .summary-box {
            background: linear-gradient(135deg, rgb(46,204,113), rgb(39,174,96));
            color: white; padding: 25px; margin: 30px 0;
            border-radius: 12px; text-align: center;
            box-shadow: 0 6px 20px rgba(46, 204, 113, 0.3);
        }
        .summary-box h3 { color: white; border: none; padding: 0; margin-bottom: 15px; }

        .warning-box {
            background: linear-gradient(135deg, rgb(231,76,60), rgb(192,57,43));
            color: white; padding: 25px; margin: 30px 0;
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(231, 76, 60, 0.3);
        }
        .warning-box h3 { color: white; border: none; padding: 0; margin-bottom: 15px; }

        .info-box {
            background: linear-gradient(135deg, rgb(52,152,219), rgb(41,128,185));
            color: white; padding: 25px; margin: 30px 0;
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(52, 152, 219, 0.3);
        }
        .info-box h3 { color: white; border: none; padding: 0; margin-bottom: 15px; }

        .header-info {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px; margin: 30px 0;
        }
        .header-card {
            background: white; padding: 20px; border-radius: 10px;
            border-left: 5px solid rgb(52,152,219);
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        .header-card strong { color: rgb(44,62,80); }

        .logo {
            text-align: center; margin-bottom: 30px;
        }
        .logo img {
            max-width: 200px; height: auto;
            filter: drop-shadow(0 4px 8px rgba(0,0,0,0.3));
        }

        .footer {
            text-align: center; margin-top: 50px; padding: 30px;
            background: linear-gradient(135deg, rgb(44,62,80), rgb(52,73,94));
            color: white; border-radius: 10px;
        }

        .progress-bar {
            background-color: rgb(236,240,241);
            border-radius: 10px; height: 20px;
            overflow: hidden; margin: 10px 0;
        }
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, rgb(46,204,113), rgb(39,174,96));
            transition: width 0.3s ease;
        }

        .tabs {
            display: flex;
            background: rgb(236,240,241);
            border-radius: 10px 10px 0 0;
            overflow: hidden;
        }
        .tab {
            flex: 1; padding: 15px; text-align: center;
            background: rgb(189,195,199); color: rgb(44,62,80);
            cursor: pointer; transition: background 0.3s;
        }
        .tab.active {
            background: rgb(52,152,219); color: white;
        }
        .tab-content {
            background: white; padding: 30px;
            border-radius: 0 0 10px 10px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }

        .compliance-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }

        .compliance-card {
            background: white;
            border-radius: 10px;
            padding: 20px;
            border-left: 5px solid;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }

        .compliance-pass { border-left-color: rgb(39,174,96); }
        .compliance-fail { border-left-color: rgb(231,76,60); }
        .compliance-pending { border-left-color: rgb(243,156,18); }

        .security-summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }

        .security-metric {
            text-align: center;
            padding: 20px;
            background: linear-gradient(135deg, rgb(248,249,250), rgb(233,236,239));
            border-radius: 10px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }

        .security-metric .number {
            font-size: 2.5em;
            font-weight: bold;
            margin: 10px 0;
        }

        .risk-high { color: rgb(231,76,60); }
        .risk-medium { color: rgb(243,156,18); }
        .risk-low { color: rgb(39,174,96); }

        @media (max-width: 768px) {
            .container { padding: 15px; }
            .metrics-grid { grid-template-columns: 1fr; }
            .header-info { grid-template-columns: 1fr; }
            h1 { font-size: 2em; }
            h2 { font-size: 1.3em; }
        }
'@

    $htmlContent = @"
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Informe de Salud del Sistema - $nombreServidor</title>
    <style>$cssStyles</style>
</head>
<body>
    <div class="container">
        <div class="logo">
            <img src="data:image/png;base64,$imagenBase64" alt="Logo">
        </div>

        <h1>🖥️ INFORME COMPLETO DE SALUD DEL SISTEMA</h1>

        <div class="summary-box">
            <h3>📊 Resumen Ejecutivo</h3>
            <p><strong>Servidor:</strong> $nombreServidor | <strong>Sistema:</strong> $sistemaOperativo</p>
            <p><strong>Fecha del Informe:</strong> $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")</p>
            <p><strong>Tiempo de Actividad:</strong> $($InfoSistema.TiempoActividad.Days) días, $($InfoSistema.TiempoActividad.Hours) horas</p>
        </div>

        <div class="header-info">
            <div class="header-card">
                <strong>🏢 Información del Sistema</strong><br>
                Fabricante: $($InfoSistema.Fabricante)<br>
                Modelo: $($InfoSistema.Modelo)<br>
                Procesador: $($InfoSistema.Procesador)<br>
                Memoria Total: $($InfoSistema.MemoriaTotal) GB
            </div>
            <div class="header-card">
                <strong>🌐 Configuración de Red</strong><br>
                IP Principal: $($InfoSistema.DireccionIP)<br>
                $($InfoSistema.DominioWorkgroup)<br>
                Zona Horaria: $($InfoSistema.TimeZone)
            </div>
            <div class="header-card">
                <strong>🔧 Detalles Técnicos</strong><br>
                Versión SO: $($InfoSistema.VersionSO)<br>
                Build: $($InfoSistema.BuildNumber)<br>
                Arquitectura: $($InfoSistema.Arquitectura)<br>
                Último Reinicio: $($InfoSistema.UltimoReinicio.ToString("dd/MM/yyyy HH:mm"))
            </div>
        </div>

        <h2>📈 MÉTRICAS DE RENDIMIENTO DE LOS 4 SUBSISTEMAS</h2>

        <div class="metrics-grid">
            <div class="metric-card">
                <h4>🔥 Subsistema CPU</h4>
                <div class="metric-value">$($MetricasRendimiento.CPU.UsoCPU)%</div>
                <span class="status-indicator status-$($MetricasRendimiento.CPU.Estado.ToLower())">$($MetricasRendimiento.CPU.Estado)</span>
                <p><strong>Cola del Procesador:</strong> $($MetricasRendimiento.CPU.ColaProcesor)</p>
                <p><strong>Interrupciones/seg:</strong> $($MetricasRendimiento.CPU.Interrupciones)</p>
            </div>

            <div class="metric-card">
                <h4>🧠 Subsistema Memoria</h4>
                <div class="metric-value">$($MetricasRendimiento.Memoria.PorcentajeUsado)%</div>
                <span class="status-indicator status-$($MetricasRendimiento.Memoria.Estado.ToLower())">$($MetricasRendimiento.Memoria.Estado)</span>
                <p><strong>Usada:</strong> $($MetricasRendimiento.Memoria.UsadaMB) MB / $($MetricasRendimiento.Memoria.TotalMB) MB</p>
                <p><strong>Páginas/seg:</strong> $($MetricasRendimiento.Memoria.PaginasPorSeg)</p>
                <p><strong>Caché:</strong> $($MetricasRendimiento.Memoria.CacheMB) MB</p>
            </div>
        </div>

        <h3>💾 Subsistema Disco</h3>
        <table>
            <tr>
                <th>Unidad</th>
                <th>Total (GB)</th>
                <th>Libre (GB)</th>
                <th>% Usado</th>
                <th>Sistema Archivos</th>
                <th>Tiempo Lectura (ms)</th>
                <th>Tiempo Escritura (ms)</th>
                <th>Estado</th>
            </tr>
"@

    foreach ($disco in $MetricasRendimiento.Discos) {
        $statusClass = switch ($disco.Estado) {
            "Normal" { "status-normal" }
            "Advertencia" { "status-warning" }
            "Crítico" { "status-critical" }
            default { "status-normal" }
        }

        $htmlContent += @"
            <tr>
                <td><strong>$($disco.Unidad)</strong></td>
                <td>$($disco.TotalGB)</td>
                <td>$($disco.LibreGB)</td>
                <td>$($disco.PorcentajeUsado)%</td>
                <td>$($disco.SistemaArchivos)</td>
                <td>$($disco.TiempoLectura)</td>
                <td>$($disco.TiempoEscritura)</td>
                <td><span class="status-indicator $statusClass">$($disco.Estado)</span></td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🌐 Subsistema Red</h3>
        <table>
            <tr>
                <th>Interfaz</th>
                <th>Bytes Enviados (KB)</th>
                <th>Bytes Recibidos (KB)</th>
                <th>Paquetes Enviados</th>
                <th>Paquetes Recibidos</th>
                <th>Errores Envío</th>
                <th>Errores Recepción</th>
            </tr>
"@

    foreach ($interfaz in $MetricasRendimiento.Red) {
        $htmlContent += @"
            <tr>
                <td>$($interfaz.Interfaz)</td>
                <td>$($interfaz.BytesEnviados)</td>
                <td>$($interfaz.BytesRecibidos)</td>
                <td>$($interfaz.PaquetesEnviados)</td>
                <td>$($interfaz.PaquetesRecibidos)</td>
                <td>$($interfaz.ErroresEnvio)</td>
                <td>$($interfaz.ErroresRecepcion)</td>
            </tr>
"@
    }


    if ($AnalisisConfiabilidad -and $AnalisisConfiabilidad.EstadisticasEstabilidad) {
        $htmlContent += @"
        </table>

        <h2>🔍 ANÁLISIS DE CONFIABILIDAD DEL SISTEMA</h2>

        <div class="metrics-grid">
            <div class="metric-card">
                <h4>📊 Estadísticas de Estabilidad</h4>
                <p><strong>Índice de Estabilidad:</strong> <span class="status-indicator status-$(if($AnalisisConfiabilidad.EstadisticasEstabilidad.IndiceEstabilidad -eq 'Alta'){'normal'}elseif($AnalisisConfiabilidad.EstadisticasEstabilidad.IndiceEstabilidad -eq 'Media'){'warning'}else{'critical'})">$($AnalisisConfiabilidad.EstadisticasEstabilidad.IndiceEstabilidad)</span></p>
                <p><strong>Fallos de Aplicación:</strong> $($AnalisisConfiabilidad.EstadisticasEstabilidad.FallosAplicacion)</p>
                <p><strong>Fallos del Sistema:</strong> $($AnalisisConfiabilidad.EstadisticasEstabilidad.FallosSistema)</p>
                <p><strong>Reinicios Detectados:</strong> $($AnalisisConfiabilidad.EstadisticasEstabilidad.ReiniciosDetectados)</p>
                <p><strong>Período de Análisis:</strong> $($AnalisisConfiabilidad.EstadisticasEstabilidad.PeriodoAnalisis)</p>
            </div>
        </div>
"@

        if ($AnalisisConfiabilidad.TendenciasSemanales -and $AnalisisConfiabilidad.TendenciasSemanales.Count -gt 0) {
            $htmlContent += @"
        <h3>📈 Tendencias de Estabilidad (Últimos 7 días)</h3>
        <table>
            <tr>
                <th>Fecha</th>
                <th>Total Eventos</th>
                <th>Eventos Críticos</th>
                <th>Estado de Estabilidad</th>
            </tr>
"@
            foreach ($tendencia in $AnalisisConfiabilidad.TendenciasSemanales) {
                $estabilidadClass = switch ($tendencia.Estabilidad) {
                    "Estable" { "status-normal" }
                    "Moderada" { "status-warning" }
                    "Inestable" { "status-critical" }
                    default { "status-normal" }
                }

                $htmlContent += @"
            <tr>
                <td>$($tendencia.Fecha)</td>
                <td>$($tendencia.TotalEventos)</td>
                <td>$($tendencia.EventosCriticos)</td>
                <td><span class="status-indicator $estabilidadClass">$($tendencia.Estabilidad)</span></td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }
    }


    if ($DiagnosticoHardware) {
        $htmlContent += @"
        <h2>🔧 DIAGNÓSTICO AVANZADO DE HARDWARE</h2>
"@


        if ($DiagnosticoHardware.DiscosSMART -and $DiagnosticoHardware.DiscosSMART.Count -gt 0) {
            $htmlContent += @"
        <h3>💾 Estado SMART de Discos Duros</h3>
        <table>
            <tr>
                <th>Modelo</th>
                <th>Número de Serie</th>
                <th>Tamaño (GB)</th>
                <th>Interfaz</th>
                <th>Estado SMART</th>
                <th>Particiones</th>
                <th>Estado General</th>
            </tr>
"@
            foreach ($disco in $DiagnosticoHardware.DiscosSMART) {
                $estadoClass = switch ($disco.Estado) {
                    "Normal" { "status-normal" }
                    "Crítico" { "status-critical" }
                    default { "status-warning" }
                }

                $htmlContent += @"
            <tr>
                <td>$($disco.Modelo)</td>
                <td>$($disco.NumeroSerie)</td>
                <td>$($disco.TamanoGB)</td>
                <td>$($disco.Interfaz)</td>
                <td>$($disco.EstadoSMART)</td>
                <td>$($disco.Particiones)</td>
                <td><span class="status-indicator $estadoClass">$($disco.Estado)</span></td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }


        if ($DiagnosticoHardware.Temperaturas -and $DiagnosticoHardware.Temperaturas.Count -gt 0) {
            $htmlContent += @"
        <h3>🌡️ Temperaturas de Componentes</h3>
        <table>
            <tr>
                <th>Componente</th>
                <th>Temperatura (°C)</th>
                <th>Estado</th>
            </tr>
"@
            foreach ($temp in $DiagnosticoHardware.Temperaturas) {
                $tempClass = switch ($temp.Estado) {
                    "Normal" { "status-normal" }
                    "Alto" { "status-warning" }
                    "Crítico" { "status-critical" }
                    default { "status-warning" }
                }

                $tempValue = if ($temp.TemperaturaC) { "$($temp.TemperaturaC)°C" } else { "No disponible" }

                $htmlContent += @"
            <tr>
                <td>$($temp.Componente)</td>
                <td>$tempValue</td>
                <td><span class="status-indicator $tempClass">$($temp.Estado)</span></td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }


        if ($DiagnosticoHardware.Baterias -and $DiagnosticoHardware.Baterias.Count -gt 0) {
            $htmlContent += @"
        <h3>🔋 Estado de Baterías</h3>
        <table>
            <tr>
                <th>Nombre</th>
                <th>Fabricante</th>
                <th>Estado de Carga</th>
                <th>% Carga</th>
                <th>Tiempo Restante</th>
                <th>Salud de Batería</th>
                <th>Estado</th>
            </tr>
"@
            foreach ($bateria in $DiagnosticoHardware.Baterias) {
                if ($bateria.Estado -ne "No se detectaron baterías (sistema de escritorio)") {
                    $bateriaClass = switch ($bateria.Estado) {
                        "Normal" { "status-normal" }
                        "Advertencia" { "status-warning" }
                        "Bajo" { "status-warning" }
                        "Crítico" { "status-critical" }
                        default { "status-normal" }
                    }

                    $htmlContent += @"
            <tr>
                <td>$($bateria.Nombre)</td>
                <td>$($bateria.Fabricante)</td>
                <td>$($bateria.EstadoCarga)</td>
                <td>$($bateria.PorcentajeCarga)%</td>
                <td>$($bateria.TiempoRestante)</td>
                <td>$($bateria.SaludBateria)</td>
                <td><span class="status-indicator $bateriaClass">$($bateria.Estado)</span></td>
            </tr>
"@
                } else {
                    $htmlContent += @"
            <tr>
                <td colspan="7" style="text-align: center; font-style: italic;">$($bateria.Estado)</td>
            </tr>
"@
                }
            }
            $htmlContent += "</table>"
        }
    }


    if ($AnalisisRoles -and $AnalisisRoles.EsServidor) {
        $htmlContent += @"
        <h2>🖥️ ANÁLISIS DE ROLES DE WINDOWS SERVER</h2>

        <div class="info-box">
            <h3>📊 Resumen de Roles</h3>
            <p><strong>Tipo de Sistema:</strong> $($AnalisisRoles.TipoSistema)</p>
            <p><strong>Total de Roles Detectados:</strong> $($AnalisisRoles.ResumenRoles.TotalRoles)</p>
            <p><strong>Roles Críticos:</strong> $($AnalisisRoles.ResumenRoles.RolesCriticos)</p>
            <p><strong>Roles Detenidos:</strong> $($AnalisisRoles.ResumenRoles.RolesDetenidos)</p>
            <p><strong>Estado General:</strong> $($AnalisisRoles.ResumenRoles.EstadoGeneral)</p>
        </div>
"@

        if ($AnalisisRoles.RolesDetectados -and $AnalisisRoles.RolesDetectados.Count -gt 0) {
            $htmlContent += @"
        <h3>🔧 Roles y Servicios Detectados</h3>
        <table>
            <tr>
                <th>Rol</th>
                <th>Servicio</th>
                <th>Estado</th>
                <th>Tipo de Inicio</th>
                <th>Descripción</th>
                <th>Crítico</th>
            </tr>
"@
            foreach ($rol in $AnalisisRoles.RolesDetectados) {
                $estadoClass = if ($rol.Estado -eq "Running") { "status-normal" } else { "status-critical" }
                $criticoIcon = if ($rol.Critico) { "⚠️" } else { "ℹ️" }

                $htmlContent += @"
            <tr>
                <td><strong>$($rol.Rol)</strong></td>
                <td>$($rol.Servicio)</td>
                <td><span class="status-indicator $estadoClass">$($rol.Estado)</span></td>
                <td>$($rol.TipoInicio)</td>
                <td>$($rol.Descripcion)</td>
                <td>$criticoIcon</td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }


        if ($AnalisisRoles.InformacionDominio) {
            $htmlContent += @"
        <h3>🌐 Información de Dominio</h3>
        <div class="metric-card">
            <p><strong>Parte de Dominio:</strong> $(if($AnalisisRoles.InformacionDominio.ParteDominio){'Sí'}else{'No'})</p>
            <p><strong>Dominio:</strong> $($AnalisisRoles.InformacionDominio.Dominio)</p>
            <p><strong>Rol del Servidor:</strong> $($AnalisisRoles.InformacionDominio.Rol)</p>
        </div>
"@
        }
    } elseif ($AnalisisRoles -and -not $AnalisisRoles.EsServidor) {
        $htmlContent += @"
        <div class="info-box">
            <h3>ℹ️ Información del Sistema</h3>
            <p>Este sistema es un <strong>$($AnalisisRoles.TipoSistema)</strong>, no un Windows Server.</p>
            <p>El análisis de roles de servidor no es aplicable.</p>
        </div>
"@
    }


    if ($AnalisisPoliticas) {
        $htmlContent += @"
        <h2>🛡️ ANÁLISIS DE POLÍTICAS DE GRUPO Y SEGURIDAD</h2>
"@

        if ($AnalisisPoliticas.EnDominio) {
            $htmlContent += @"
        <div class="info-box">
            <h3>📋 Información de GPO</h3>
            <p><strong>Sistema en Dominio:</strong> Sí</p>
            <p><strong>Última Actualización GPO:</strong> $($AnalisisPoliticas.UltimaActualizacionGPO)</p>
        </div>
"@


            if ($AnalisisPoliticas.GPOsEquipo -and $AnalisisPoliticas.GPOsEquipo.Count -gt 0) {
                $htmlContent += @"
        <h3>🖥️ GPOs Aplicadas al Equipo</h3>
        <table>
            <tr>
                <th>Nombre de la GPO</th>
                <th>Tipo</th>
                <th>Estado</th>
            </tr>
"@
                foreach ($gpo in $AnalisisPoliticas.GPOsEquipo) {
                    $htmlContent += @"
            <tr>
                <td>$($gpo.Nombre)</td>
                <td>$($gpo.Tipo)</td>
                <td><span class="status-indicator status-normal">$($gpo.Estado)</span></td>
            </tr>
"@
                }
                $htmlContent += "</table>"
            }


            if ($AnalisisPoliticas.GPOsUsuario -and $AnalisisPoliticas.GPOsUsuario.Count -gt 0) {
                $htmlContent += @"
        <h3>👤 GPOs Aplicadas al Usuario</h3>
        <table>
            <tr>
                <th>Nombre de la GPO</th>
                <th>Tipo</th>
                <th>Estado</th>
            </tr>
"@
                foreach ($gpo in $AnalisisPoliticas.GPOsUsuario) {
                    $htmlContent += @"
            <tr>
                <td>$($gpo.Nombre)</td>
                <td>$($gpo.Tipo)</td>
                <td><span class="status-indicator status-normal">$($gpo.Estado)</span></td>
            </tr>
"@
                }
                $htmlContent += "</table>"
            }
        } else {
            $htmlContent += @"
        <div class="warning-box">
            <h3>⚠️ Sistema No Unido a Dominio</h3>
            <p>Este sistema no está unido a un dominio Active Directory.</p>
            <p>Las políticas de grupo de dominio no son aplicables.</p>
        </div>
"@
        }


        if ($AnalisisPoliticas.PoliticasLocales -and $AnalisisPoliticas.PoliticasLocales.Count -gt 0) {
            $htmlContent += @"
        <h3>🔒 Políticas de Seguridad Locales</h3>
        <table>
            <tr>
                <th>Política</th>
                <th>Valor</th>
                <th>Categoría</th>
            </tr>
"@
            foreach ($politica in $AnalisisPoliticas.PoliticasLocales | Select-Object -First 20) {
                $htmlContent += @"
            <tr>
                <td>$($politica.Politica)</td>
                <td>$($politica.Valor)</td>
                <td>$($politica.Categoria)</td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }
    }


    if ($VerificacionCumplimiento -and $VerificacionCumplimiento.ResumenCumplimiento) {
        $htmlContent += @"
        <h2>✅ VERIFICACIÓN DE CUMPLIMIENTO (CIS BENCHMARKS)</h2>

        <div class="security-summary">
            <div class="security-metric">
                <div class="number risk-low">$($VerificacionCumplimiento.ResumenCumplimiento.PorcentajeCumplimiento)%</div>
                <div>Cumplimiento General</div>
            </div>
            <div class="security-metric">
                <div class="number risk-low">$($VerificacionCumplimiento.ResumenCumplimiento.Cumple)</div>
                <div>Verificaciones Exitosas</div>
            </div>
            <div class="security-metric">
                <div class="number risk-high">$($VerificacionCumplimiento.ResumenCumplimiento.NoCumple)</div>
                <div>No Cumple</div>
            </div>
            <div class="security-metric">
                <div class="number risk-medium">$($VerificacionCumplimiento.ResumenCumplimiento.PendienteRevision)</div>
                <div>Pendiente Revisión</div>
            </div>
        </div>

        <div class="$(if($VerificacionCumplimiento.ResumenCumplimiento.NivelCumplimiento -eq 'Excelente' -or $VerificacionCumplimiento.ResumenCumplimiento.NivelCumplimiento -eq 'Bueno'){'summary-box'}elseif($VerificacionCumplimiento.ResumenCumplimiento.NivelCumplimiento -eq 'Aceptable'){'info-box'}else{'warning-box'})">
            <h3>📊 Nivel de Cumplimiento: $($VerificacionCumplimiento.ResumenCumplimiento.NivelCumplimiento)</h3>
            <p>Total de verificaciones realizadas: $($VerificacionCumplimiento.ResumenCumplimiento.TotalVerificaciones)</p>
        </div>
"@

        if ($VerificacionCumplimiento.Verificaciones -and $VerificacionCumplimiento.Verificaciones.Count -gt 0) {

            $verificacionesProblematicas = $VerificacionCumplimiento.Verificaciones | Where-Object { $_.Cumple -ne "Sí" }

            if ($verificacionesProblematicas.Count -gt 0) {
                $htmlContent += @"
        <h3>⚠️ Verificaciones que Requieren Atención</h3>
        <div class="compliance-grid">
"@
                foreach ($verificacion in $verificacionesProblematicas) {
                    $cardClass = switch ($verificacion.Cumple) {
                        "No" { "compliance-fail" }
                        "Pendiente" { "compliance-pending" }
                        "Revisar" { "compliance-pending" }
                        default { "compliance-pending" }
                    }

                    $criticidadClass = switch ($verificacion.Criticidad) {
                        "Alta" { "risk-high" }
                        "Media" { "risk-medium" }
                        "Baja" { "risk-low" }
                        default { "risk-medium" }
                    }

                    $htmlContent += @"
            <div class="compliance-card $cardClass">
                <h4>$($verificacion.ID)</h4>
                <p><strong>$($verificacion.Descripcion)</strong></p>
                <p><strong>Estado Actual:</strong> $($verificacion.EstadoActual)</p>
                <p><strong>Recomendación:</strong> $($verificacion.Recomendacion)</p>
                <p><strong>Criticidad:</strong> <span class="$criticidadClass">$($verificacion.Criticidad)</span></p>
                <p><strong>Categoría:</strong> $($verificacion.Categoria)</p>
            </div>
"@
                }
                $htmlContent += "</div>"
            }
        }
    }


    if ($AnalisisPermisos -and $AnalisisPermisos.ResumenPermisos) {
        $htmlContent += @"
        <h2>🔐 ANÁLISIS DE PERMISOS DE CARPETAS SENSIBLES</h2>

        <div class="security-summary">
            <div class="security-metric">
                <div class="number risk-low">$($AnalisisPermisos.ResumenPermisos.PorcentajeSeguras)%</div>
                <div>Carpetas Seguras</div>
            </div>
            <div class="security-metric">
                <div class="number risk-medium">$($AnalisisPermisos.ResumenPermisos.CarpetasConProblemas)</div>
                <div>Con Problemas</div>
            </div>
            <div class="security-metric">
                <div class="number risk-high">$($AnalisisPermisos.ResumenPermisos.CarpetasCriticas)</div>
                <div>Críticas</div>
            </div>
            <div class="security-metric">
                <div class="number">$($AnalisisPermisos.ResumenPermisos.TotalCarpetasAnalizadas)</div>
                <div>Total Analizadas</div>
            </div>
        </div>

        <div class="$(if($AnalisisPermisos.ResumenPermisos.NivelSeguridad -eq 'Excelente'){'summary-box'}elseif($AnalisisPermisos.ResumenPermisos.NivelSeguridad -eq 'Bueno' -or $AnalisisPermisos.ResumenPermisos.NivelSeguridad -eq 'Aceptable'){'info-box'}else{'warning-box'})">
            <h3>🛡️ Nivel de Seguridad: $($AnalisisPermisos.ResumenPermisos.NivelSeguridad)</h3>
        </div>
"@

        if ($AnalisisPermisos.AnalisisPermisos) {
            $carpetasProblematicas = $AnalisisPermisos.AnalisisPermisos | Where-Object { $_.Estado -in @("Advertencia", "Crítico") }

            if ($carpetasProblematicas.Count -gt 0) {
                $htmlContent += @"
        <h3>⚠️ Carpetas con Permisos Problemáticos</h3>
        <table>
            <tr>
                <th>Ruta</th>
                <th>Descripción</th>
                <th>Propietario</th>
                <th>Permisos Problemáticos</th>
                <th>Estado</th>
            </tr>
"@
                foreach ($carpeta in $carpetasProblematicas) {
                    $estadoClass = switch ($carpeta.Estado) {
                        "Crítico" { "status-critical" }
                        "Advertencia" { "status-warning" }
                        default { "status-normal" }
                    }

                    $htmlContent += @"
            <tr>
                <td><strong>$($carpeta.Ruta)</strong></td>
                <td>$($carpeta.Descripcion)</td>
                <td>$($carpeta.Propietario)</td>
                <td>$($carpeta.PermisosProblematicosCount)</td>
                <td><span class="status-indicator $estadoClass">$($carpeta.Estado)</span></td>
            </tr>
"@
                }
                $htmlContent += "</table>"
            }
        }
    }


    if ($AuditoriaSoftware -and $AuditoriaSoftware.ResumenAuditoria) {
        $htmlContent += @"
        <h2>💿 AUDITORÍA DE SOFTWARE INSTALADO</h2>

        <div class="security-summary">
            <div class="security-metric">
                <div class="number">$($AuditoriaSoftware.ResumenAuditoria.TotalSoftwareInstalado)</div>
                <div>Total Programas</div>
            </div>
            <div class="security-metric">
                <div class="number risk-high">$($AuditoriaSoftware.ResumenAuditoria.SoftwareCritico)</div>
                <div>Riesgo Alto</div>
            </div>
            <div class="security-metric">
                <div class="number risk-medium">$($AuditoriaSoftware.ResumenAuditoria.SoftwareRiesgoMedio)</div>
                <div>Riesgo Medio</div>
            </div>
            <div class="security-metric">
                <div class="number risk-low">$($AuditoriaSoftware.ResumenAuditoria.PorcentajeSeguro)%</div>
                <div>Software Seguro</div>
            </div>
        </div>

        <div class="$(if($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo -eq 'Bajo'){'summary-box'}elseif($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo -eq 'Medio'){'info-box'}else{'warning-box'})">
            <h3>⚠️ Nivel de Riesgo: $($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo)</h3>
        </div>

"@

if ($AuditoriaSoftware -and $AuditoriaSoftware.SoftwareInstalado) {
    $htmlContent += @"
    <h2>📦 INVENTARIO DE SOFTWARE INSTALADO</h2>

    <div class="metrics-grid">
        <div class="metric-card">
            <h4>📊 Resumen de Auditoría de Software</h4>
            <p><strong>Total Software:</strong> $($AuditoriaSoftware.ResumenAuditoria.TotalSoftwareInstalado)</p>
            <p><strong>Software Problemático:</strong> $($AuditoriaSoftware.ResumenAuditoria.SoftwareConProblemas)</p>
            <p><strong>Nivel de Riesgo:</strong> <span class="status-indicator status-$(if($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo -eq 'Bajo'){'normal'}elseif($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo -eq 'Medio'){'warning'}else{'critical'})">$($AuditoriaSoftware.ResumenAuditoria.NivelRiesgo)</span></p>
            <p><strong>Software Crítico:</strong> $($AuditoriaSoftware.ResumenAuditoria.SoftwareCritico)</p>
            <p><strong>Software Riesgo Medio:</strong> $($AuditoriaSoftware.ResumenAuditoria.SoftwareRiesgoMedio)</p>
        </div>
    </div>
"@


    if ($AuditoriaSoftware.SoftwareProblematico -and $AuditoriaSoftware.SoftwareProblematico.Count -gt 0) {
        $htmlContent += @"
    <h3>⚠️ Software Problemático Detectado</h3>
    <table>
        <tr>
            <th>Software</th>
            <th>Versión</th>
            <th>Fabricante</th>
            <th>Instalado</th>
            <th>Criticidad</th>
            <th>Problema</th>
        </tr>
"@

        foreach ($software in $AuditoriaSoftware.SoftwareProblematico) {
            $criticidadClass = switch ($software.CriticidadMaxima) {
                "Alta" { "status-critical" }
                "Media" { "status-warning" }
                "Baja" { "status-normal" }
                default { "" }
            }

            $descripcionProblemas = ""
            foreach ($problema in $software.Problemas) {
                $descripcionProblemas += "$($problema.Descripcion)<br>"
            }

            $htmlContent += @"
        <tr>
            <td><strong>$($software.Nombre)</strong></td>
            <td>$($software.Version)</td>
            <td>$($software.Fabricante)</td>
            <td>$($software.FechaInstalacion)</td>
            <td><span class="status-indicator $criticidadClass">$($software.CriticidadMaxima)</span></td>
            <td>$descripcionProblemas</td>
        </tr>
"@
        }
        $htmlContent += "</table>"
    }


    $htmlContent += @"
    <h3>📋 Lista Completa de Software Instalado</h3>
    <table>
        <tr>
            <th>Nombre</th>
            <th>Versión</th>
            <th>Fabricante</th>
            <th>Fecha Instalación</th>
            <th>Tamaño (MB)</th>
        </tr>
"@

    foreach ($software in ($AuditoriaSoftware.SoftwareInstalado | Sort-Object Nombre)) {
        $htmlContent += @"
        <tr>
            <td>$($software.Nombre)</td>
            <td>$($software.Version)</td>
            <td>$($software.Fabricante)</td>
            <td>$($software.FechaInstalacion)</td>
            <td>$($software.TamañoMB)</td>
        </tr>
"@
    }
    $htmlContent += "</table>"
}
        if ($AuditoriaSoftware.SoftwareProblematico -and $AuditoriaSoftware.SoftwareProblematico.Count -gt 0) {
            $htmlContent += @"
        <h3>⚠️ Software que Requiere Atención</h3>
        <table>
            <tr>
                <th>Nombre</th>
                <th>Versión</th>
                <th>Fabricante</th>
                <th>Fecha Instalación</th>
                <th>Criticidad</th>
                <th>Problemas Detectados</th>
            </tr>
"@
            foreach ($software in $AuditoriaSoftware.SoftwareProblematico | Select-Object -First 20) {
                $criticidadClass = switch ($software.CriticidadMaxima) {
                    "Alta" { "status-critical" }
                    "Media" { "status-warning" }
                    "Baja" { "status-normal" }
                    default { "status-normal" }
                }

                $problemasTexto = ($software.Problemas | ForEach-Object { $_.Tipo }) -join ", "

                $htmlContent += @"
            <tr>
                <td><strong>$($software.Nombre)</strong></td>
                <td>$($software.Version)</td>
                <td>$($software.Fabricante)</td>
                <td>$($software.FechaInstalacion)</td>
                <td><span class="status-indicator $criticidadClass">$($software.CriticidadMaxima)</span></td>
                <td>$problemasTexto</td>
            </tr>
"@
            }
            $htmlContent += "</table>"
        }
    }

    $htmlContent += @"
        <h2>📊 LOGS DE EVENTOS DE LOS 3 TIPOS PRINCIPALES</h2>

        <h3>🔴 Logs del Sistema (Últimos eventos críticos)</h3>
        <table>
            <tr>
                <th>Fecha/Hora</th>
                <th>Nivel</th>
                <th>Proveedor</th>
                <th>ID Evento</th>
                <th>Mensaje</th>
            </tr>
"@

    foreach ($log in $LogsEventos.LogsSistema | Select-Object -First 10) {
        $nivelClass = switch ($log.LevelDisplayName) {
            "Critical" { "status-critical" }
            "Error" { "status-critical" }
            "Warning" { "status-warning" }
            default { "status-normal" }
        }

        $htmlContent += @"
            <tr>
                <td>$($log.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss"))</td>
                <td><span class="status-indicator $nivelClass">$($log.LevelDisplayName)</span></td>
                <td>$($log.ProviderName)</td>
                <td>$($log.Id)</td>
                <td>$($log.Message)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>📱 Logs de Aplicación (Últimos eventos críticos)</h3>
        <table>
            <tr>
                <th>Fecha/Hora</th>
                <th>Nivel</th>
                <th>Proveedor</th>
                <th>ID Evento</th>
                <th>Mensaje</th>
            </tr>
"@

    foreach ($log in $LogsEventos.LogsAplicacion | Select-Object -First 10) {
        $nivelClass = switch ($log.LevelDisplayName) {
            "Critical" { "status-critical" }
            "Error" { "status-critical" }
            "Warning" { "status-warning" }
            default { "status-normal" }
        }

        $htmlContent += @"
            <tr>
                <td>$($log.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss"))</td>
                <td><span class="status-indicator $nivelClass">$($log.LevelDisplayName)</span></td>
                <td>$($log.ProviderName)</td>
                <td>$($log.Id)</td>
                <td>$($log.Message)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🔒 Logs de Seguridad (Eventos de autenticación)</h3>
        <table>
            <tr>
                <th>Fecha/Hora</th>
                <th>ID Evento</th>
                <th>Cuenta</th>
                <th>IP Origen</th>
                <th>Tipo de Evento</th>
                <th>Mensaje</th>
            </tr>
"@

    foreach ($log in $LogsEventos.LogsSeguridad | Select-Object -First 15) {
        $tipoClass = switch ($log.EventType) {
            "Login Fallido" { "status-critical" }
            "Cuenta Bloqueada" { "status-critical" }
            "Login" { "status-normal" }
            "Logout" { "status-normal" }
            default { "status-warning" }
        }

        $htmlContent += @"
            <tr>
                <td>$($log.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss"))</td>
                <td>$($log.Id)</td>
                <td>$($log.Account)</td>
                <td>$($log.SourceIP)</td>
                <td><span class="status-indicator $tipoClass">$($log.EventType)</span></td>
                <td>$($log.RawMessage)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h2>🔍 DATOS EXTENDIDOS Y ANÁLISIS AVANZADO</h2>

        <h3>🔄 Últimos Parches Instalados</h3>
        <table>
            <tr>
                <th>HotFix ID</th>
                <th>Descripción</th>
                <th>Fecha Instalación</th>
                <th>Instalado Por</th>
            </tr>
"@

    foreach ($parche in $DatosExtendidos.UltimosParches) {
        $fechaInstalacion = if ($parche.InstalledOn) { $parche.InstalledOn.ToString("dd/MM/yyyy") } else { "No disponible" }
        $htmlContent += @"
            <tr>
                <td><strong>$($parche.HotFixID)</strong></td>
                <td>$($parche.Description)</td>
                <td>$fechaInstalacion</td>
                <td>$($parche.InstalledBy)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>⚠️ Servicios Automáticos Detenidos</h3>
        <table>
            <tr>
                <th>Nombre del Servicio</th>
                <th>Nombre Interno</th>
                <th>Estado</th>
                <th>Tipo de Inicio</th>
"@

    if ($RevisarServicioTerceros) {
        $htmlContent += "<th>Compañía</th><th>Tipo</th>"
    }

    $htmlContent += "</tr>"

    foreach ($servicio in $DatosExtendidos.ServiciosDetenidos) {
        $htmlContent += @"
            <tr>
                <td><strong>$($servicio.DisplayName)</strong></td>
                <td>$($servicio.Name)</td>
                <td><span class="status-indicator status-critical">$($servicio.Status)</span></td>
                <td>$($servicio.StartType)</td>
"@

        if ($RevisarServicioTerceros) {
            $tipoClass = if ($servicio.EsMicrosoft) { "status-normal" } else { "status-warning" }
            $htmlContent += @"
                <td>$($servicio.Compania)</td>
                <td><span class="status-indicator $tipoClass">$($servicio.Tipo)</span></td>
"@
        }

        $htmlContent += "</tr>"
    }

    $htmlContent += @"
        </table>

        <h3>🔥 TOP 5 Procesos por Uso de CPU</h3>
        <table>
            <tr>
                <th>Proceso</th>
                <th>CPU (segundos)</th>
                <th>ID Proceso</th>
                <th>Memoria (MB)</th>
            </tr>
"@

    foreach ($proceso in $DatosExtendidos.ProcesosCPU) {
        $htmlContent += @"
            <tr>
                <td><strong>$($proceso.Name)</strong></td>
                <td>$($proceso.CPU)</td>
                <td>$($proceso.Id)</td>
                <td>$($proceso.Memoria_MB)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🧠 TOP 5 Procesos por Uso de Memoria</h3>
        <table>
            <tr>
                <th>Proceso</th>
                <th>Memoria (MB)</th>
                <th>ID Proceso</th>
                <th>CPU (segundos)</th>
            </tr>
"@

    foreach ($proceso in $DatosExtendidos.ProcesosMemoria) {
        $htmlContent += @"
            <tr>
                <td><strong>$($proceso.Name)</strong></td>
                <td>$($proceso.Memoria_MB)</td>
                <td>$($proceso.Id)</td>
                <td>$($proceso.CPU)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🌐 Puertos Abiertos (Listening)</h3>
        <table>
            <tr>
                <th>Dirección Local</th>
                <th>Puerto</th>
                <th>Estado</th>
                <th>Proceso ID</th>
            </tr>
"@

    foreach ($puerto in $DatosExtendidos.PuertosAbiertos | Select-Object -First 20) {
        $htmlContent += @"
            <tr>
                <td>$($puerto.LocalAddress)</td>
                <td><strong>$($puerto.LocalPort)</strong></td>
                <td><span class="status-indicator status-normal">$($puerto.State)</span></td>
                <td>$($puerto.OwningProcess)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🔗 TOP 10 Conexiones Activas por IP</h3>
        <table>
            <tr>
                <th>IP Remota</th>
                <th>Número de Conexiones</th>
            </tr>
"@

    foreach ($conexion in $DatosExtendidos.ConexionesActivas) {
        $htmlContent += @"
            <tr>
                <td><strong>$($conexion.Name)</strong></td>
                <td>$($conexion.Count)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>👥 Usuarios Locales</h3>
        <table>
            <tr>
                <th>Nombre</th>
                <th>Habilitado</th>
                <th>Último Logon</th>
                <th>Contraseña Expira</th>
            </tr>
"@

    foreach ($usuario in $DatosExtendidos.UsuariosLocales) {
        $habilitadoClass = if ($usuario.Enabled) { "status-normal" } else { "status-warning" }
        $htmlContent += @"
            <tr>
                <td><strong>$($usuario.Name)</strong></td>
                <td><span class="status-indicator $habilitadoClass">$(if($usuario.Enabled){'Sí'}else{'No'})</span></td>
                <td>$($usuario.UltimoLogon)</td>
                <td>$($usuario.PasswordExpires)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🚫 Últimos 5 Intentos de Login Fallidos</h3>
        <table>
            <tr>
                <th>Fecha/Hora</th>
                <th>Cuenta</th>
                <th>IP Origen</th>
                <th>Tipo de Fallo</th>
            </tr>
"@

    if ($DatosExtendidos.LoginsFallidos -and $DatosExtendidos.LoginsFallidos.Count -gt 0) {
        foreach ($login in $DatosExtendidos.LoginsFallidos) {
            $htmlContent += @"
            <tr>
                <td>$($login.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss"))</td>
                <td><strong>$($login.Cuenta)</strong></td>
                <td>$($login.IPOrigen)</td>
                <td>$($login.TipoFallo)</td>
            </tr>
"@
        }
    } else {
        $htmlContent += @"
            <tr>
                <td colspan="4" style="text-align: center; font-style: italic;">No se encontraron intentos de login fallidos recientes</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>⏰ Tareas Programadas (Estado)</h3>
        <table>
            <tr>
                <th>Nombre de Tarea</th>
                <th>Última Ejecución</th>
                <th>Resultado</th>
                <th>Estado</th>
            </tr>
"@

    foreach ($tarea in $DatosExtendidos.TareasProgramadas | Select-Object -First 15) {
        $estadoClass = switch ($tarea.Estado) {
            "Exitoso" { "status-normal" }
            "Error*" { "status-critical" }
            default { "status-warning" }
        }

        $ultimaEjecucion = if ($tarea.LastRunTime) { $tarea.LastRunTime.ToString("dd/MM/yyyy HH:mm") } else { "Nunca" }

        $htmlContent += @"
            <tr>
                <td><strong>$($tarea.TaskName)</strong></td>
                <td>$ultimaEjecucion</td>
                <td>$($tarea.LastTaskResult)</td>
                <td><span class="status-indicator $estadoClass">$($tarea.Estado)</span></td>
            </tr>
"@
    }


    $htmlContent += @"
        </table>

        <h2>🛡️ ANÁLISIS DE SEGURIDAD EXTENDIDO</h2>

        <h3>🔍 Windows Defender</h3>
        <div class="metrics-grid">
            <div class="metric-card">
                <h4>Estado de Windows Defender</h4>
"@

    if ($DatosExtendidos.WindowsDefender -and -not $DatosExtendidos.WindowsDefender.Error) {
        $defenderStatus = if ($DatosExtendidos.WindowsDefender.Habilitado) { "status-normal" } else { "status-critical" }
        $realtimeStatus = if ($DatosExtendidos.WindowsDefender.TiempoRealActivo) { "status-normal" } else { "status-critical" }

        $htmlContent += @"
                <p><strong>Antivirus Habilitado:</strong> <span class="status-indicator $defenderStatus">$(if($DatosExtendidos.WindowsDefender.Habilitado){'Sí'}else{'No'})</span></p>
                <p><strong>Protección en Tiempo Real:</strong> <span class="status-indicator $realtimeStatus">$(if($DatosExtendidos.WindowsDefender.TiempoRealActivo){'Activa'}else{'Inactiva'})</span></p>
                <p><strong>Último Análisis Rápido:</strong> $($DatosExtendidos.WindowsDefender.UltimoAnalisisRapido)</p>
                <p><strong>Último Análisis Completo:</strong> $($DatosExtendidos.WindowsDefender.UltimoAnalisisCompleto)</p>
                <p><strong>Edad de Definiciones:</strong> $($DatosExtendidos.WindowsDefender.EdadDefinicionesAV)</p>
"@
    } else {
        $htmlContent += @"
                <p><span class="status-indicator status-warning">Estado no disponible</span></p>
                <p>$($DatosExtendidos.WindowsDefender.Error)</p>
"@
    }

    $htmlContent += @"
            </div>
        </div>

        <h3>🔥 Estado del Firewall por Perfil</h3>
        <table>
            <tr>
                <th>Perfil</th>
                <th>Habilitado</th>
                <th>Conexiones Entrantes</th>
                <th>Conexiones Salientes</th>
                <th>Notificaciones</th>
            </tr>
"@

    if ($DatosExtendidos.Firewall -and $DatosExtendidos.Firewall.Count -gt 0) {
        foreach ($perfil in $DatosExtendidos.Firewall) {
            $habilitadoClass = if ($perfil.Habilitado) { "status-normal" } else { "status-critical" }

            $htmlContent += @"
            <tr>
                <td><strong>$($perfil.Nombre)</strong></td>
                <td><span class="status-indicator $habilitadoClass">$(if($perfil.Habilitado){'Sí'}else{'No'})</span></td>
                <td>$($perfil.ConexionesEntrantes)</td>
                <td>$($perfil.ConexionesSalientes)</td>
                <td>$(if($perfil.NotificacionesActivas){'Activas'}else{'Inactivas'})</td>
            </tr>
"@
        }
    } else {
        $htmlContent += @"
            <tr>
                <td colspan="5" style="text-align: center; font-style: italic;">Información del firewall no disponible</td>
            </tr>
"@
    }

    $htmlContent += @"
        </table>

        <h3>🔒 Control de Acceso de Usuario (UAC)</h3>
        <div class="metric-card">
"@

    if ($DatosExtendidos.UAC -and -not $DatosExtendidos.UAC.Error) {
        $uacStatus = if ($DatosExtendidos.UAC.Habilitado) { "status-normal" } else { "status-critical" }

        $htmlContent += @"
            <p><strong>UAC Habilitado:</strong> <span class="status-indicator $uacStatus">$(if($DatosExtendidos.UAC.Habilitado){'Sí'}else{'No'})</span></p>
            <p><strong>Nivel de UAC:</strong> $($DatosExtendidos.UAC.Nivel)</p>
            <p><strong>Elevación No Segura:</strong> $(if($DatosExtendidos.UAC.ElevacionNoSegura){'Sí'}else{'No'})</p>
"@
    } else {
        $htmlContent += @"
            <p><span class="status-indicator status-warning">Información de UAC no disponible</span></p>
"@
    }

    $htmlContent += @"
        </div>

        <h3>📡 Configuración SNMP</h3>
        <div class="metric-card">
"@

    if ($DatosExtendidos.ServicioSNMP) {
        $snmpStatus = if ($DatosExtendidos.ServicioSNMP.Status -eq "Running") { "status-normal" } else { "status-warning" }

        $htmlContent += @"
            <p><strong>Servicio SNMP:</strong> <span class="status-indicator $snmpStatus">$($DatosExtendidos.ServicioSNMP.Status)</span></p>
            <p><strong>Tipo de Inicio:</strong> $($DatosExtendidos.ServicioSNMP.StartType)</p>
"@
    } else {
        $htmlContent += @"
            <p><span class="status-indicator status-normal">Servicio SNMP no instalado</span></p>
"@
    }

    $htmlContent += @"
        </div>

        <div class="footer">
            <h3>📋 RESUMEN DEL INFORME</h3>
            <p><strong>Servidor Analizado:</strong> $nombreServidor</p>
            <p><strong>Sistema Operativo:</strong> $sistemaOperativo</p>
            <p><strong>Fecha y Hora del Informe:</strong> $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")</p>
            <p><strong>Generado por:</strong> Script de Salud del Sistema v3.0 - Análisis Completo</p>
            <hr style="margin: 20px 0; border: 1px solid rgba(255,255,255,0.3);">
            <p style="font-size: 0.9em; opacity: 0.8;">
                Este informe incluye análisis de los 4 subsistemas principales (CPU, Memoria, Disco, Red),
                logs de eventos de seguridad, análisis de confiabilidad, diagnóstico avanzado de hardware,
                verificación de cumplimiento con CIS Benchmarks, análisis de permisos de carpetas sensibles,
                auditoría de software instalado, y análisis completo de seguridad del sistema.
            </p>
        </div>
    </div>

    <script>
        // Agregar interactividad básica
        document.addEventListener('DOMContentLoaded', function() {
            // Efecto hover en las tarjetas
            const cards = document.querySelectorAll('.metric-card, .header-card, .compliance-card');
            cards.forEach(card => {
                card.addEventListener('mouseenter', function() {
                    this.style.transform = 'translateY(-5px)';
                    this.style.transition = 'transform 0.3s ease';
                });
                card.addEventListener('mouseleave', function() {
                    this.style.transform = 'translateY(0)';
                });
            });

            // Resaltar filas de tabla al hacer hover
            const rows = document.querySelectorAll('tr');
            rows.forEach(row => {
                row.addEventListener('mouseenter', function() {
                    if (this.parentElement.tagName === 'TBODY' || this.parentElement.parentElement.tagName === 'TABLE') {
                        this.style.backgroundColor = '
                        this.style.transition = 'background-color 0.2s ease';
                    }
                });
                row.addEventListener('mouseleave', function() {
                    if (this.parentElement.tagName === 'TBODY' || this.parentElement.parentElement.tagName === 'TABLE') {
                        this.style.backgroundColor = '';
                    }
                });
            });

            // Animación de aparición para las secciones
            const sections = document.querySelectorAll('h2, h3');
            const observer = new IntersectionObserver((entries) => {
                entries.forEach(entry => {
                    if (entry.isIntersecting) {
                        entry.target.style.opacity = '1';
                        entry.target.style.transform = 'translateY(0)';
                    }
                });
            });

            sections.forEach(section => {
                section.style.opacity = '0';
                section.style.transform = 'translateY(20px)';
                section.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
                observer.observe(section);
            });
        });
    </script>
</body>
</html>
"@

    return $htmlContent
}
