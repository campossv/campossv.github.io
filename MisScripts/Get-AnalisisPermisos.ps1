function Get-AnalisisPermisos {
    try {
        Write-Progress -Activity "Analizando permisos de carpetas sensibles" -PercentComplete 70
        Write-Host "   Analizando permisos de carpetas sensibles..." -ForegroundColor Yellow

        $analisisPermisos = @()


        $carpetasSensibles = @(
            @{ Ruta = "C:\Windows\System32"; Descripcion = "Archivos del sistema Windows" },
            @{ Ruta = "C:\Windows\SysWOW64"; Descripcion = "Archivos del sistema Windows (32-bit)" },
            @{ Ruta = "C:\Program Files"; Descripcion = "Programas instalados" },
            @{ Ruta = "C:\Program Files (x86)"; Descripcion = "Programas instalados (32-bit)" },
            @{ Ruta = "C:\Windows\Temp"; Descripcion = "Archivos temporales del sistema" },
            @{ Ruta = "C:\ProgramData"; Descripcion = "Datos de aplicaciones" },
            @{ Ruta = "C:\Users\Public"; Descripcion = "Carpeta pública de usuarios" },
            @{ Ruta = "C:\inetpub"; Descripcion = "Sitios web IIS" }
        )


        try {
            $carpetasCompartidas = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "IPC$" -and $_.Name -ne "ADMIN$" -and $_.Name -notlike "*$" }
            foreach ($share in $carpetasCompartidas) {
                $carpetasSensibles += @{ Ruta = $share.Path; Descripcion = "Carpeta compartida: $($share.Name)" }
            }
        } catch {
            Write-Host "   No se pudieron obtener carpetas compartidas" -ForegroundColor Yellow
        }

        foreach ($carpeta in $carpetasSensibles) {
            try {
                if (Test-Path $carpeta.Ruta) {
                    Write-Host "   Analizando: $($carpeta.Ruta)" -ForegroundColor Yellow


                    $acl = Get-Acl $carpeta.Ruta -ErrorAction SilentlyContinue

                    if ($acl) {
                        $permisosProblematicos = @()
                        $permisosNormales = @()

                        foreach ($access in $acl.Access) {
                            $usuario = $access.IdentityReference.Value
                            $permisos = $access.FileSystemRights.ToString()
                            $tipo = $access.AccessControlType.ToString()
                            $herencia = $access.IsInherited


                            $esProblematico = $false
                            $razon = ""


                            if ($usuario -match "(Everyone|Users|Authenticated Users)" -and $tipo -eq "Allow") {
                                if ($permisos -match "(FullControl|Modify|Write)" -and $carpeta.Ruta -match "(System32|Program Files|Windows)") {
                                    $esProblematico = $true
                                    $razon = "Permisos excesivos para grupo amplio en carpeta del sistema"
                                }
                            }


                            if ($permisos -match "(Write|Modify|FullControl)" -and $tipo -eq "Allow" -and $carpeta.Ruta -match "(System32|SysWOW64)") {
                                if ($usuario -notmatch "(SYSTEM|Administrators|TrustedInstaller)") {
                                    $esProblematico = $true
                                    $razon = "Permisos de escritura en carpeta crítica del sistema"
                                }
                            }


                            if ($carpeta.Ruta -match "Temp" -and $permisos -match "FullControl" -and $usuario -match "Everyone") {
                                $esProblematico = $true
                                $razon = "Control total para Everyone en carpeta temporal"
                            }

                            $permisoObj = [PSCustomObject]@{
                                Usuario = $usuario
                                Permisos = $permisos
                                Tipo = $tipo
                                Heredado = $herencia
                                Problematico = $esProblematico
                                Razon = $razon
                            }

                            if ($esProblematico) {
                                $permisosProblematicos += $permisoObj
                            } else {
                                $permisosNormales += $permisoObj
                            }
                        }

                        $analisisPermisos += [PSCustomObject]@{
                            Ruta = $carpeta.Ruta
                            Descripcion = $carpeta.Descripcion
                            Propietario = $acl.Owner
                            PermisosProblematicos = $permisosProblematicos
                            PermisosNormales = $permisosNormales
                            TotalPermisos = $acl.Access.Count
                            PermisosProblematicosCount = $permisosProblematicos.Count
                            Estado = if ($permisosProblematicos.Count -eq 0) { "Normal" }
                                    elseif ($permisosProblematicos.Count -le 2) { "Advertencia" }
                                    else { "Crítico" }
                        }
                    }
                } else {
                    Write-Host "   Carpeta no existe: $($carpeta.Ruta)" -ForegroundColor Gray
                }
            } catch {
                Write-Host "   Error al analizar $($carpeta.Ruta): $_" -ForegroundColor Red
                $analisisPermisos += [PSCustomObject]@{
                    Ruta = $carpeta.Ruta
                    Descripcion = $carpeta.Descripcion
                    Error = $_.Exception.Message
                    Estado = "Error"
                }
            }
        }


        $totalCarpetas = $analisisPermisos.Count
        $carpetasConProblemas = ($analisisPermisos | Where-Object { $_.Estado -in @("Advertencia", "Crítico") }).Count
        $carpetasCriticas = ($analisisPermisos | Where-Object { $_.Estado -eq "Crítico" }).Count

        $resumenPermisos = @{
            TotalCarpetasAnalizadas = $totalCarpetas
            CarpetasConProblemas = $carpetasConProblemas
            CarpetasCriticas = $carpetasCriticas
            CarpetasNormales = $totalCarpetas - $carpetasConProblemas
            PorcentajeSeguras = if ($totalCarpetas -gt 0) {
                [math]::Round((($totalCarpetas - $carpetasConProblemas) / $totalCarpetas) * 100, 1)
            } else { 0 }
            NivelSeguridad = switch ($carpetasCriticas) {
                0 { if ($carpetasConProblemas -eq 0) { "Excelente" } else { "Bueno" } }
                { $_ -le 2 } { "Aceptable" }
                default { "Deficiente" }
            }
        }

        return @{
            AnalisisPermisos = $analisisPermisos
            ResumenPermisos = $resumenPermisos
        }

    } catch {
        Write-Warning "Error en análisis de permisos: $_"
        return @{ Error = $_.Exception.Message }
    }
}