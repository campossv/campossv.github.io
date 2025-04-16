<#
AUTOR:
=====
Vladimir Antonio Campos Herrera

REQUISITOS DEL SISTEMA
======================

Sistema Operativo:
------------------
- Windows 7 o superior

PowerShell:
-----------
- Versión 5.1 o superior
- Ejecución de scripts habilitada (Set-ExecutionPolicy RemoteSigned)

.NET Framework:
---------------
- Versión 4.5 o superior

Ensamblados Requeridos:
-----------------------
1. System.Windows.Forms (obligatorio)
2. System.Drawing (obligatorio)
3. System.Windows.Forms.DataVisualization (opcional, solo para funciones gráficas)

CONFIGURACIÓN PREVIA:
---------------------
Los ensamblados requeridos se cargan automáticamente al inicio del script.
Si falta System.Windows.Forms.DataVisualization, las funciones gráficas se deshabilitarán.

Con Nuget:
Install-Package Microsoft.MSChart -ProviderName NuGet
Desde Microsoft:
https://www.microsoft.com/en-us/download/details.aspx?id=14422
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# Verificar y cargar ensamblado para gráficos
$chartAvailable = $false
try {
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization -ErrorAction Stop
    $chartAvailable = $true
} catch {
    Write-Warning "System.Windows.Forms.DataVisualization no está disponible. La función de gráficos estará deshabilitada."
}

# Cargar una imagen en Base64 (opcional)
$LPNG = "iVBORw0KGgoAAAANSUhEUgAAAMgAAAAuCAYAAABtRVYBAAAACXBIWXMAAAsTAAALEwEAmpwYAAAGq2lUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNy4xLWMwMDAgNzkuYTg3MzFiOSwgMjAyMS8wOS8wOS0wMDozNzozOCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6cGhvdG9zaG9wPSJodHRwOi8vbnMuYWRvYmUuY29tL3Bob3Rvc2hvcC8xLjAvIiB4bXBNTTpPcmlnaW5hbERvY3VtZW50SUQ9InhtcC5kaWQ6ODIxNzExM2UtNmZjNS0zNjQyLWIzMDctOGEzNDM3ZmY2NWRkIiB4bXBNTTpEb2N1bWVudElEPSJhZG9iZTpkb2NpZDpwaG90b3Nob3A6Mzk4YTY5ZDMtYzljYS0zYzRhLWE4YTctZjhmYmM2MmYxOWU0IiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjEzNDk3YzZlLWVjNTgtMzM0YS1hZWY2LWFhMWFlODRjNGE0YiIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgMjMuMCAoV2luZG93cykiIHhtcDpDcmVhdGVEYXRlPSIyMDI0LTExLTIwVDEzOjU2OjExLTA2OjAwIiB4bXA6TW9kaWZ5RGF0ZT0iMjAyNC0xMS0yMFQxNzo1MTo0My0wNjowMCIgeG1wOk1ldGFkYXRhRGF0ZT0iMjAyNC0xMS0yMFQxNzo1MTo0My0wNjowMCIgZGM6Zm9ybWF0PSJpbWFnZS9wbmciIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDo4MjE3MTEzZS02ZmM1LTM2NDItYjMwNy04YTM0MzdmZjY1ZGQiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6ODIxNzExM2UtNmZjNS0zNjQyLWIzMDctOGEzNDM3ZmY2NWRkIi8+IDx4bXBNTTpIaXN0b3J5PiA8cmRmOlNlcT4gPHJkZjpsaSBzdEV2dDphY3Rpb249InNhdmVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjQ4OTAyZGY4LTNjNzQtNzc0MC05YjM1LTBjYjkyODRjYTgyMCIgc3RFdnQ6d2hlbj0iMjAyNC0xMS0yMFQxNzo0NzoyMi0wNjowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDIzLjAgKFdpbmRvd3MpIiBzdEV2dDpjaGFuZ2VkPSIvIi8+IDxyZGY6bGkgc3RFdnQ6YWN0aW9uPSJzYXZlZCIgc3RFdnQ6aW5zdGFuY2VJRD0ieG1wLmlpZDoxMzQ5N2M2ZS1lYzU4LTMzNGEtYWVmNi1hYTFhZTg0YzRhNGIiIHN0RXZ0OndoZW49IjIwMjQtMTEtMjBUMTc6NTE6NDMtMDY6MDAiIHN0RXZ0OnNvZnR3YXJlQWdlbnQ9IkFkb2JlIFBob3Rvc2hvcCAyMy4wIChXaW5kb3dzKSIgc3RFdnQ6Y2hhbmdlZD0iLyIvPiA8L3JkZjpTZXE+IDwveG1wTU06SGlzdG9yeT4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz55i9rdAAAuNUlEQVR4nO2deXyU1fnovzOTmclkZ0IC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"
$lenbytes = [Convert]::FromBase64String($LPNG)
$lenmemoria = New-Object System.IO.MemoryStream
$lenmemoria.Write($lenbytes, 0, $lenbytes.Length)
$lenmemoria.Position = 0
$imagenl = [System.Drawing.Image]::FromStream($lenmemoria, $true)

# Función para graficar
function Plot-Function {
    param(
        [string]$functionExpression,
        [double]$xMin = -10,
        [double]$xMax = 10,
        [double]$step = 0.1
    )
    
    try {
        $chart.Series[0].Points.Clear()
        
        # Configurar los ejes
        $chart.ChartAreas[0].AxisX.Minimum = $xMin
        $chart.ChartAreas[0].AxisX.Maximum = $xMax
        $chart.ChartAreas[0].AxisX.Title = "X"
        $chart.ChartAreas[0].AxisY.Title = "Y"
        
        # Graficar la función
        for ($x = $xMin; $x -le $xMax; $x += $step) {
            # Preparar la expresión para evaluación
            $expression = $functionExpression -replace "x", "($x)"
            
            # Evaluar la expresión de manera segura
            try {
                # Usar Invoke-Expression con precaución
                $y = Invoke-Expression $expression
                
                # Agregar el punto al gráfico si es un número válido
                if ($y -is [double] -or $y -is [int]) {
                    $chart.Series[0].Points.AddXY($x, $y) | Out-Null
                }
            }
            catch {
                # Ignorar puntos que causan errores (como divisiones por cero)
                Write-Host "Error en x=$x : $_"
            }
        }
        
        $graphPanel.Visible = $true
        return $true
    }
    catch {
        Write-Error "Error al graficar: $_"
        return $false
    }
}

# Configuración de tema
$theme = @{
    'Light' = @{
        'FormBG' = [System.Drawing.Color]::FromArgb(240, 240, 240)
        'DisplayBG' = [System.Drawing.Color]::White
        'DisplayText' = [System.Drawing.Color]::Black
        'ButtonBG' = [System.Drawing.Color]::FromArgb(220, 220, 220)
        'SpecialBtnBG' = [System.Drawing.Color]::FromArgb(180, 220, 240)
        'FunctionBtnBG' = [System.Drawing.Color]::FromArgb(220, 240, 180)
        'ClearBtnBG' = [System.Drawing.Color]::FromArgb(240, 180, 180)
        'EqualsBtnBG' = [System.Drawing.Color]::FromArgb(180, 240, 180)
    }
    'Dark' = @{
        'FormBG' = [System.Drawing.Color]::FromArgb(45, 45, 48)
        'DisplayBG' = [System.Drawing.Color]::FromArgb(30, 30, 30)
        'DisplayText' = [System.Drawing.Color]::White
        'ButtonBG' = [System.Drawing.Color]::FromArgb(60, 60, 60)
        'SpecialBtnBG' = [System.Drawing.Color]::FromArgb(0, 122, 204)
        'FunctionBtnBG' = [System.Drawing.Color]::FromArgb(63, 72, 204)
        'ClearBtnBG' = [System.Drawing.Color]::FromArgb(204, 62, 68)
        'EqualsBtnBG' = [System.Drawing.Color]::FromArgb(78, 154, 6)
    }
}

# Crear el formulario principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Calculadora Avanzada Pro"
$form.Size = New-Object System.Drawing.Size(450, 840)
$form.StartPosition = "CenterScreen"
#$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\calc.exe")
$form.BackColor = $theme['Light']['FormBG']

# Agregar una imagen
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Size = New-Object System.Drawing.Size(200, 40)
$pictureBox.Location = New-Object System.Drawing.Point(120, 730)
$pictureBox.Image = $imagenl
$form.Controls.Add($pictureBox)

# Variables de estado
$script:lastNumber = $null
$script:operation = $null
$script:memory = 0
$script:history = @()
$script:currentTheme = 'Light'
$script:currentUnit = 'Degrees'

# Pantalla de resultados
$txtDisplay = New-Object System.Windows.Forms.TextBox
$txtDisplay.Location = New-Object System.Drawing.Point(35, 40)
$txtDisplay.Size = New-Object System.Drawing.Size(360, 40)
$txtDisplay.Font = New-Object System.Drawing.Font("Arial", 18)
$txtDisplay.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Right
$txtDisplay.ReadOnly = $true
$txtDisplay.BackColor = $theme['Light']['DisplayBG']
$txtDisplay.ForeColor = $theme['Light']['DisplayText']
$form.Controls.Add($txtDisplay)

# Pantalla de memoria
$lblMemory = New-Object System.Windows.Forms.Label
$lblMemory.Location = New-Object System.Drawing.Point(20, 45)
$lblMemory.Size = New-Object System.Drawing.Size(360, 20)
$lblMemory.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Right
$lblMemory.Font = New-Object System.Drawing.Font("Arial", 10)
$lblMemory.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblMemory)

# Historial de operaciones
$listHistory = New-Object System.Windows.Forms.ListBox
$listHistory.Location = New-Object System.Drawing.Point(35, 85)
$listHistory.Size = New-Object System.Drawing.Size(360, 120)
$listHistory.Font = New-Object System.Drawing.Font("Arial", 10)
$listHistory.BackColor = $theme['Light']['DisplayBG']
$listHistory.ForeColor = $theme['Light']['DisplayText']
$form.Controls.Add($listHistory)

# Gráfico de función (panel)
# Configurar gráficos si están disponibles
if ($chartAvailable) {
    $graphPanel = New-Object System.Windows.Forms.Panel
    $graphPanel.Location = New-Object System.Drawing.Point(35, 205)
    $graphPanel.Size = New-Object System.Drawing.Size(360, 150)
    $graphPanel.BackColor = [System.Drawing.Color]::White
    $graphPanel.Visible = $false
    $form.Controls.Add($graphPanel)

    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.Width = $graphPanel.Width
    $chart.Height = $graphPanel.Height
    $chart.BackColor = [System.Drawing.Color]::White

    $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chart.ChartAreas.Add($chartArea)

    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $series.Color = [System.Drawing.Color]::Blue
    $series.BorderWidth = 2
    $chart.Series.Add($series)

    $graphPanel.Controls.Add($chart)
}

# Botones numéricos y operaciones
$buttonPositions = @{
    "7" = "20,365"; "8" = "100,365"; "9" = "180,365"; "/" = "260,365"; "MC" = "340,365"
    "4" = "20,425"; "5" = "100,425"; "6" = "180,425"; "*" = "260,425"; "MR" = "340,425"
    "1" = "20,485"; "2" = "100,485"; "3" = "180,485"; "-" = "260,485"; "M+" = "340,485"
    "0" = "20,545"; "." = "100,545"; "=" = "180,545"; "+" = "260,545"; "M-" = "340,545"
    "√" = "20,605"; "x²" = "100,605"; "x^y" = "180,605"; "%" = "260,605"; "±" = "340,605"
    "sin" = "20,667"; "cos" = "100,667"; "tan" = "180,667"; "C" = "260,667"; "⌫" = "340,667"
}

# Crear botones
foreach ($key in $buttonPositions.Keys) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $key
    $btn.Font = New-Object System.Drawing.Font("Arial", 12)
    $btn.Size = New-Object System.Drawing.Size(70, 50)
    $pos = $buttonPositions[$key] -split ","
    $btn.Location = New-Object System.Drawing.Point([int]$pos[0], [int]$pos[1])
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    
    # Asignar colores según tipo de botón
    if ($key -match "[+|\-|*|/|%|√|^]") {
        $btn.BackColor = $theme['Light']['SpecialBtnBG']
    } elseif ($key -match "sin|cos|tan") {
        $btn.BackColor = $theme['Light']['FunctionBtnBG']
    } elseif ($key -eq "C") {
        $btn.BackColor = $theme['Light']['ClearBtnBG']
    } elseif ($key -eq "=") {
        $btn.BackColor = $theme['Light']['EqualsBtnBG']
    } elseif ($key -match "MC|MR|M\+|M\-") {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 255)
    } else {
        $btn.BackColor = $theme['Light']['ButtonBG']
    }
    
    $btn.Add_Click({
        $currentText = $txtDisplay.Text
        
        switch ($this.Text) {
            "C" {
                $txtDisplay.Text = ""
                $script:lastNumber = $null
                $script:operation = $null
            }
            "⌫" {
                if ($currentText.Length -gt 0) {
                    $txtDisplay.Text = $currentText.Substring(0, $currentText.Length - 1)
                }
            }
            "±" {
                if ($currentText -ne "") {
                    $txtDisplay.Text = -[double]$currentText
                }
            }
            "MC" { $script:memory = 0; Update-MemoryDisplay }
            "MR" { $txtDisplay.Text = $script:memory }
            "M+" { 
                try { 
                    $script:memory += [double]$currentText
                    Update-MemoryDisplay
                } catch {} 
            }
            "M-" { 
                try { 
                    $script:memory -= [double]$currentText
                    Update-MemoryDisplay
                } catch {} 
            }
            "=" {
                if ($script:operation -and $script:lastNumber -ne $null) {
                    try {
                        if ($script:operation -eq "Pow") {
                            $base = [double]$script:lastNumber
                            $exponent = [double]$currentText
                            $result = [math]::Pow($base, $exponent)
                        }
                        else {
                            $expression = "$script:lastNumber $script:operation $currentText"
                            $result = Invoke-Expression $expression
                        }
            
                        $txtDisplay.Text = $result
                        Add-ToHistory "$script:lastNumber $script:operation $currentText" $result
                        $script:lastNumber = $null
                        $script:operation = $null
                    } catch {
                        $txtDisplay.Text = "Error"
                    }
                }
            }
            "√" {
                try {
                    $num = [double]$currentText
                    if ($num -ge 0) {
                        $result = [math]::Sqrt($num)
                        $txtDisplay.Text = $result
                        Add-ToHistory "√($num)" $result
                    } else {
                        $txtDisplay.Text = "Error"
                    }
                } catch {
                    $txtDisplay.Text = "Error"
                }
            }
            "x²" {
                try {
                    $num = [double]$currentText
                    $result = [math]::Pow($num, 2)
                    $txtDisplay.Text = $result
                    Add-ToHistory "($num)²" $result
                } catch {
                    $txtDisplay.Text = "Error"
                }
            }
            "x^y" {
                $script:lastNumber = $currentText
                $script:operation = "Pow"
                $txtDisplay.Text = ""
            }
            "sin" {
                try {
                    $angle = [double]$currentText
                    if ($script:currentUnit -eq "Degrees") {
                        $angle = [math]::PI * $angle / 180
                    }
                    $result = [math]::Sin($angle)
                    $txtDisplay.Text = $result
                    Add-ToHistory "sin($currentText)" $result
                } catch {
                    $txtDisplay.Text = "Error"
                }
            }
            "cos" {
                try {
                    $angle = [double]$currentText
                    if ($script:currentUnit -eq "Degrees") {
                        $angle = [math]::PI * $angle / 180
                    }
                    $result = [math]::Cos($angle)
                    $txtDisplay.Text = $result
                    Add-ToHistory "cos($currentText)" $result
                } catch {
                    $txtDisplay.Text = "Error"
                }
            }
            "tan" {
                try {
                    $angle = [double]$currentText
                    if ($script:currentUnit -eq "Degrees") {
                        $angle = [math]::PI * $angle / 180
                    }
                    $result = [math]::Tan($angle)
                    $txtDisplay.Text = $result
                    Add-ToHistory "tan($currentText)" $result
                } catch {
                    $txtDisplay.Text = "Error"
                }
            }
            "+" { 
                $script:lastNumber = $currentText
                $script:operation = "+"
                $txtDisplay.Text = ""
            }
            "-" { 
                $script:lastNumber = $currentText
                $script:operation = "-"
                $txtDisplay.Text = ""
            }
            "*" { 
                $script:lastNumber = $currentText
                $script:operation = "*"
                $txtDisplay.Text = ""
            }
            "/" { 
                $script:lastNumber = $currentText
                $script:operation = "/"
                $txtDisplay.Text = ""
            }
            "%" { 
                try {
                    $num = [double]$currentText
                    $result = $num / 100
                    $txtDisplay.Text = $result
                    Add-ToHistory "$num%" $result
                } catch {
                    $txtDisplay.Text = "Error"
                }
            }
            default {
                if ($currentText -eq "Error") {
                    $txtDisplay.Text = $this.Text
                } else {
                    $txtDisplay.Text += $this.Text
                }
            }
        }
    })
    
    $form.Controls.Add($btn)
}

# Menú superior
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

# Menú Funciones Avanzadas
$advancedMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$advancedMenu.Text = "Funciones"

# Submenú Logaritmo
$logItem = New-Object System.Windows.Forms.ToolStripMenuItem
$logItem.Text = "Logaritmo"
$logItem.Add_Click({
    $logForm = New-Object System.Windows.Forms.Form
    $logForm.Text = "Logaritmo"
    $logForm.Size = New-Object System.Drawing.Size(300, 200)
    $logForm.StartPosition = "CenterScreen"
    
    $lblNum = New-Object System.Windows.Forms.Label
    $lblNum.Text = "Número (argumento):"
    $lblNum.Location = New-Object System.Drawing.Point(10, 20)
    $lblNum.Size = New-Object System.Drawing.Size(200, 20)
    $logForm.Controls.Add($lblNum)
    
    $txtNum = New-Object System.Windows.Forms.TextBox
    $txtNum.Location = New-Object System.Drawing.Point(10, 40)
    $txtNum.Size = New-Object System.Drawing.Size(260, 20)
    $logForm.Controls.Add($txtNum)
    
    $lblBase = New-Object System.Windows.Forms.Label
    $lblBase.Text = "Base del logaritmo:"
    $lblBase.Location = New-Object System.Drawing.Point(10, 70)
    $lblBase.Size = New-Object System.Drawing.Size(200, 20)
    $logForm.Controls.Add($lblBase)
    
    $txtBase = New-Object System.Windows.Forms.TextBox
    $txtBase.Location = New-Object System.Drawing.Point(10, 90)
    $txtBase.Size = New-Object System.Drawing.Size(260, 20)
    $txtBase.Text = "10"
    $logForm.Controls.Add($txtBase)
    
    $btnCalculate = New-Object System.Windows.Forms.Button
    $btnCalculate.Text = "Calcular"
    $btnCalculate.Location = New-Object System.Drawing.Point(10, 120)
    $btnCalculate.Size = New-Object System.Drawing.Size(100, 30)
    $btnCalculate.Add_Click({
        try {
            $numVal = [double]$txtNum.Text
            $baseVal = [double]$txtBase.Text
            
            if ($numVal -le 0 -or $baseVal -le 0 -or $baseVal -eq 1) {
                [System.Windows.Forms.MessageBox]::Show("Entradas inválidas. El argumento debe ser positivo y la base debe ser positiva y diferente de 1.", "Error")
            } else {
                $result = [math]::Log($numVal, $baseVal)
                $txtDisplay.Text = $result
                Add-ToHistory "log$baseVal($numVal)" $result
                $logForm.Close()
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Entradas inválidas. Por favor ingrese números válidos.", "Error")
        }
    })
    $logForm.Controls.Add($btnCalculate)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Location = New-Object System.Drawing.Point(120, 120)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.Add_Click({ $logForm.Close() })
    $logForm.Controls.Add($btnCancel)
    
    $logForm.ShowDialog()
})

# Submenú Exponencial
$expItem = New-Object System.Windows.Forms.ToolStripMenuItem
$expItem.Text = "Exponencial (e^x)"
$expItem.Add_Click({
    try {
        $current = [double]$txtDisplay.Text
        $result = [math]::Exp($current)
        $txtDisplay.Text = $result
        Add-ToHistory "e^($current)" $result
    } catch {
        $txtDisplay.Text = "Error"
    }
})

# Submenú Factorial
$factItem = New-Object System.Windows.Forms.ToolStripMenuItem
$factItem.Text = "Factorial"
$factItem.Add_Click({
    try {
        $n = [int]$txtDisplay.Text
        if ($n -lt 0) {
            $txtDisplay.Text = "Error"
        } elseif ($n -gt 20) {
            $txtDisplay.Text = "Muy grande"
        } else {
            $result = 1
            for ($i = 1; $i -le $n; $i++) {
                $result *= $i
            }
            $txtDisplay.Text = $result
            Add-ToHistory "$n!" $result
        }
    } catch {
        $txtDisplay.Text = "Error"
    }
})
# Función para graficar (solo si está disponible)
$graphItem = New-Object System.Windows.Forms.ToolStripMenuItem
$graphItem.Text = "Graficar Función"
$graphItem.Add_Click({
    $graphForm = New-Object System.Windows.Forms.Form
    $graphForm.Text = "Graficar Función"
    $graphForm.Size = New-Object System.Drawing.Size(350, 300)
    $graphForm.StartPosition = "CenterScreen"
    
    $lblFunction = New-Object System.Windows.Forms.Label
    $lblFunction.Text = "Función (use 'x' como variable):"
    $lblFunction.Location = New-Object System.Drawing.Point(10, 20)
    $lblFunction.Size = New-Object System.Drawing.Size(300, 20)
    $graphForm.Controls.Add($lblFunction)
    
    $txtFunction = New-Object System.Windows.Forms.TextBox
    $txtFunction.Text = "[Math]::Sin(x)"
    $txtFunction.Location = New-Object System.Drawing.Point(10, 40)
    $txtFunction.Size = New-Object System.Drawing.Size(300, 20)
    $graphForm.Controls.Add($txtFunction)
    
    $lblRange = New-Object System.Windows.Forms.Label
    $lblRange.Text = "Rango (xMin, xMax):"
    $lblRange.Location = New-Object System.Drawing.Point(10, 70)
    $lblRange.Size = New-Object System.Drawing.Size(300, 20)
    $graphForm.Controls.Add($lblRange)
    
    $txtXMin = New-Object System.Windows.Forms.TextBox
    $txtXMin.Text = "-10"
    $txtXMin.Location = New-Object System.Drawing.Point(10, 90)
    $txtXMin.Size = New-Object System.Drawing.Size(140, 20)
    $graphForm.Controls.Add($txtXMin)
    
    $txtXMax = New-Object System.Windows.Forms.TextBox
    $txtXMax.Text = "10"
    $txtXMax.Location = New-Object System.Drawing.Point(170, 90)
    $txtXMax.Size = New-Object System.Drawing.Size(140, 20)
    $graphForm.Controls.Add($txtXMax)
    
    $btnGraph = New-Object System.Windows.Forms.Button
    $btnGraph.Text = "Graficar"
    $btnGraph.Location = New-Object System.Drawing.Point(10, 130)
    $btnGraph.Size = New-Object System.Drawing.Size(100, 30)
    $btnGraph.Add_Click({
        try {
            $function = $txtFunction.Text
            $xMin = [double]$txtXMin.Text
            $xMax = [double]$txtXMax.Text
            
            $success = Plot-Function -functionExpression $function -xMin $xMin -xMax $xMax
            
            if ($success) {
                $graphForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Error al graficar la función. Verifique la sintaxis.", "Error")
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al graficar la función: $_", "Error")
        }
    })
    $graphForm.Controls.Add($btnGraph)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Location = New-Object System.Drawing.Point(120, 130)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.Add_Click({ $graphForm.Close() })
    $graphForm.Controls.Add($btnCancel)
    
    $graphForm.ShowDialog()
})

$advancedMenu.DropDownItems.AddRange(@($logItem, $expItem, $factItem, $graphItem))

# Menú Temas
$themeMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$themeMenu.Text = "Tema"

$lightThemeItem = New-Object System.Windows.Forms.ToolStripMenuItem
$lightThemeItem.Text = "Claro"
$lightThemeItem.Add_Click({ Set-Theme 'Light' })

$darkThemeItem = New-Object System.Windows.Forms.ToolStripMenuItem
$darkThemeItem.Text = "Oscuro"
$darkThemeItem.Add_Click({ Set-Theme 'Dark' })

$themeMenu.DropDownItems.AddRange(@($lightThemeItem, $darkThemeItem))

# Menú Unidades
$unitMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$unitMenu.Text = "Unidades"

$degreesItem = New-Object System.Windows.Forms.ToolStripMenuItem
$degreesItem.Text = "Grados"
$degreesItem.Add_Click({ $script:currentUnit = 'Degrees' })

$radiansItem = New-Object System.Windows.Forms.ToolStripMenuItem
$radiansItem.Text = "Radianes"
$radiansItem.Add_Click({ $script:currentUnit = 'Radians' })

$unitMenu.DropDownItems.AddRange(@($degreesItem, $radiansItem))

# Menú Conversión
$convertMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$convertMenu.Text = "Conversión"

$degToRadItem = New-Object System.Windows.Forms.ToolStripMenuItem
$degToRadItem.Text = "Grados a Radianes"
$degToRadItem.Add_Click({
    try {
        $degrees = [double]$txtDisplay.Text
        $radians = [math]::PI * $degrees / 180
        $txtDisplay.Text = $radians
        Add-ToHistory "$degrees° a rad" $radians
    } catch {
        $txtDisplay.Text = "Error"
    }
})

$radToDegItem = New-Object System.Windows.Forms.ToolStripMenuItem
$radToDegItem.Text = "Radianes a Grados"
$radToDegItem.Add_Click({
    try {
        $radians = [double]$txtDisplay.Text
        $degrees = $radians * 180 / [math]::PI
        $txtDisplay.Text = $degrees
        Add-ToHistory "$radians rad a °" $degrees
    } catch {
        $txtDisplay.Text = "Error"
    }
})

$convertMenu.DropDownItems.AddRange(@($degToRadItem, $radToDegItem))

$menuStrip.Items.AddRange(@($advancedMenu, $themeMenu, $unitMenu, $convertMenu))
$form.Controls.Add($menuStrip)

# Funciones auxiliares
function Update-MemoryDisplay {
    $lblMemory.Text = if ($script:memory -ne 0) { "M: $($script:memory)" } else { "" }
}

function Update-History {
    $listHistory.Items.Clear()
    $listHistory.Items.AddRange($script:history)
}

function Add-ToHistory {
    param($operation, $result)
    $script:history += "$operation = $result"
    if ($script:history.Count -gt 10) {
        $script:history = $script:history[-10..-1]
    }
    Update-History
}

function Set-Theme {
    param($themeName)
    $script:currentTheme = $themeName
    $colors = $theme[$themeName]
    
    $form.BackColor = $colors['FormBG']
    $txtDisplay.BackColor = $colors['DisplayBG']
    $txtDisplay.ForeColor = $colors['DisplayText']
    $listHistory.BackColor = $colors['DisplayBG']
    $listHistory.ForeColor = $colors['DisplayText']
    
    foreach ($control in $form.Controls) {
        if ($control -is [System.Windows.Forms.Button]) {
            if ($control.Text -match "[+|\-|*|/|%|√|^]") {
                $control.BackColor = $colors['SpecialBtnBG']
            } elseif ($control.Text -match "sin|cos|tan") {
                $control.BackColor = $colors['FunctionBtnBG']
            } elseif ($control.Text -eq "C") {
                $control.BackColor = $colors['ClearBtnBG']
            } elseif ($control.Text -eq "=") {
                $control.BackColor = $colors['EqualsBtnBG']
            } elseif ($control.Text -match "MC|MR|M\+|M\-") {
                $control.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 255)
            } else {
                $control.BackColor = $colors['ButtonBG']
            }
        }
    }
}

# Mostrar el formulario
$form.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $btn = $form.Controls | Where-Object { $_.Text -eq "=" }
        if ($btn) { $btn.PerformClick() }
    }
})

$form.ShowDialog() | Out-Null