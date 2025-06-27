function Get-LogsEventos {
    param([int]$Dias)
    
    try {
        $fechaInicio = (Get-Date).AddDays(-$Dias)
        
        $logsSistema = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Level = 1,2,3
            StartTime = $fechaInicio
        } -MaxEvents 200 | Select-Object TimeCreated, LevelDisplayName, ProviderName, Id, Message
        
        $logsAplicacion = Get-WinEvent -FilterHashtable @{
            LogName = 'Application'
            Level = 1,2,3
            StartTime = $fechaInicio
        } -MaxEvents 200 | Select-Object TimeCreated, LevelDisplayName, ProviderName, Id, Message
        
        $logsSeguridad = Get-WinEvent -LogName "Security" -MaxEvents 100 | Where-Object {
            $_.Id -in @(4625, 4624, 4634, 4648)
        } | Select-Object TimeCreated, Id, @{Name="Account";Expression={$_.Properties[5].Value}}
        
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