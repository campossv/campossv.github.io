#Vladimir Campos
Add-Type -AssemblyName System.Security

$LPNG = "iVBORw0KGgoAAAANSUhEUgAAAMgAAAAuCAYAAABtRVYBAAAACXBIWXMAAAsTAAALEwEAmpwYAAAGq2lUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNy4xLWMwMDAgNzkuYTg3MzFiOSwgMjAyMS8wOS8wOS0wMDozNzozOCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6cGhvdG9zaG9wPSJodHRwOi8vbnMuYWRvYmUuY29tL3Bob3Rvc2hvcC8xLjAvIiB4bXBNTTpPcmlnaW5hbERvY3VtZW50SUQ9InhtcC5kaWQ6ODIxNzExM2UtNmZjNS0zNjQyLWIzMDctOGEzNDM3ZmY2NWRkIiB4bXBNTTpEb2N1bWVudElEPSJhZG9iZTpkb2NpZDpwaG90b3Nob3A6Mzk4YTY5ZDMtYzljYS0zYzRhLWE4YTctZjhmYmM2MmYxOWU0IiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjEzNDk3YzZlLWVjNTgtMzM0YS1hZWY2LWFhMWFlODRjNGE0YiIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgMjMuMCAoV2luZG93cykiIHhtcDpDcmVhdGVEYXRlPSIyMDI0LTExLTIwVDEzOjU2OjExLTA2OjAwIiB4bXA6TW9kaWZ5RGF0ZT0iMjAyNC0xMS0yMFQxNzo1MTo0My0wNjowMCIgeG1wOk1ldGFkYXRhRGF0ZT0iMjAyNC0xMS0yMFQxNzo1MTo0My0wNjowMCIgZGM6Zm9ybWF0PSJpbWFnZS9wbmciIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDo4MjE3MTEzZS02ZmM1LTM2NDItYjMwNy04YTM0MzdmZjY1ZGQiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6ODIxNzExM2UtNmZjNS0zNjQyLWIzMDctOGEzNDM3ZmY2NWRkIi8+IDx4bXBNTTpIaXN0b3J5PiA8cmRmOlNlcT4gPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjQ4OTAyZGY4LTNjNzQtNzc0MC05YjM1LTBjYjkyODRjYTgyMCIgc3RFdnQ6d2hlbj0iMjAyNC0xMS0yMFQxNzo0NzoyMi0wNjowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDIzLjAgKFdpbmRvd3MpIiBzdEV2dDpjaGFuZ2VkPSIvIi8+IDxyZGY6bGkgc3RFdnQ6YWN0aW9uPSJzYXZlZCIgc3RFdnQ6aW5zdGFuY2VJRD0ieG1wLmlpZDoxMzQ5N2M2ZS1lYzU4LTMzNGEtYWVmNi1hYTFhZTg0YzRhNGIiIHN0RXZ0OndoZW49IjIwMjQtMTEtMjBUMTc6NTE6NDMtMDY6MDAiIHN0RXZ0OnNvZnR3YXJlQWdlbnQ9IkFkb2JlIFBob3Rvc2hvcCAyMy4wIChXaW5kb3dzKSIgc3RFdnQ6Y2hhbmdlZD0iLyIvPiA8L3JkZjpTZXE+IDwveG1wTU06SGlzdG9yeT4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz55i9rdAAAuNUlEQVR4nO2deXyU1fnovzOTmclkZ0IC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"
$logFile = "C:\Windows\PowerShell\datos.log"
$emailConfigFile = "C:\Windows\PowerShell\emailconfig.secure"
$keyFile = "C:\Windows\PowerShell\email_config_key.txt"

# Función para obtener la clave de cifrado
function Get-EncryptionKey {
    if (-not (Test-Path $keyFile)) {
        Write-Log "No se encontró el archivo de clave. Asegúrese de que existe en $keyFile" -Level "ERROR"
        exit
    }
    return Get-Content $keyFile -Encoding Byte
}

# Función para descifrar la configuración de correo
function Unprotect-EmailConfig {
    $key = Get-EncryptionKey
    $encryptedData = Get-Content $emailConfigFile -Encoding Byte
    
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Key = $key
    $aes.IV = $encryptedData[0..15]
    
    $decryptor = $aes.CreateDecryptor()
    $decryptedBytes = $decryptor.TransformFinalBlock($encryptedData, 16, $encryptedData.Length - 16)
    $jsonConfig = [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
    
    return $jsonConfig | ConvertFrom-Json
}

# Función para escribir logs
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"
    Add-Content -Path $logFile -Value $logMessage
    Write-Host $logMessage
}

# Función para cargar la configuración de correo
function Get-EmailConfig {
    if (Test-Path -Path $emailConfigFile) {
        try {
            $config = Unprotect-EmailConfig
            return $config
        }
        catch {
            Write-Log "Error al cargar la configuración: $_" -Level "ERROR"
            exit
        }
    }
    else {
        Write-Log "El archivo de configuración '$emailConfigFile' no existe." -Level "WARN"
        exit
    }
}

try {
    # Cargar la configuración de correo
    $emailConfig = Get-EmailConfig

    # Buscar el último evento 4724
    try {
        $Evento = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4724} -MaxEvents 1
    }
    catch {
        Write-Log "No se encontraron eventos 4724 (cambio de contraseña)." -Level "WARN"
        exit
    }

    # Procesar el mensaje completo para HTML
    $formattedMessage = $Evento.Message
    if ($formattedMessage) {
        $formattedMessage = $formattedMessage -replace "\r\n", "<br>" `
                                               -replace "&", "&amp;" `
                                               -replace "<", "&lt;" `
                                               -replace ">", "&gt;" `
                                               -replace "&lt;br&gt;", "<br>"
    } else {
        $formattedMessage = "Sin mensaje disponible"
    }

    # Crear el cuerpo del correo en HTML
    $htmlBody = @"
    <html>
    <head>
        <style>
            body { font-family: Arial, sans-serif; }
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            td { white-space: normal; }
        </style>
    </head>
    <body>
        <img src="data:image/png;base64,$LPNG" alt="Logo" />
        <h2>Cambio de contraseña detectado</h2>
        <table>
            <tr><th>Fecha y Hora</th><td>$($Evento.TimeCreated)</td></tr>
            <tr><th>ID del Evento</th><td>$($Evento.Id)</td></tr>
            <tr><th>Usuario que cambió la contraseña</th><td>$($Evento.Properties[4].Value)</td></tr>
            <tr><th>Usuario cuya contraseña fue cambiada</th><td>$($Evento.Properties[0].Value)</td></tr>
            <tr><th>Evento completo</th><td>$formattedMessage</td></tr>
        </table>
    </body>
    </html>
"@

    # Enviar correo electrónico
    try {
        $smtpServer = $emailConfig.Servidor
        $smtpPort = $emailConfig.Puerto
        $from = $emailConfig.Remitente
        $toAddresses = $emailConfig.Destinatario -split ',' | ForEach-Object { $_.Trim() }
        $subject = "Informe de Cambio de Contraseña"

        $smtpClient = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
        $smtpClient.EnableSsl = $emailConfig.Protocolo -in @("SSL", "STARTTLS")
        
        if ($emailConfig.Autenticacion -eq "UsuarioContrasena") {
            $smtpClient.Credentials = New-Object System.Net.NetworkCredential($emailConfig.Usuario, $emailConfig.Contrasena)
        }

        $mailMessage = New-Object Net.Mail.MailMessage
        $mailMessage.From = $from
        foreach ($to in $toAddresses) {
            $mailMessage.To.Add($to)
        }
        $mailMessage.Subject = $subject
        $mailMessage.Body = $htmlBody
        $mailMessage.IsBodyHtml = $true
        $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8

        $smtpClient.Send($mailMessage)

        Write-Log "Correo enviado exitosamente a: $($toAddresses -join ', ')" -Level "INFO"
    }
    catch {
        Write-Log "Error al enviar el correo: $_" -Level "ERROR"
    }
}
catch {
    Write-Log "Error inesperado en el script: $_" -Level "ERROR"
}

