# Importa os módulos necessários
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-Message {
    param (
        [string]$message
    )
    $form = New-Object Windows.Forms.Form
    $form.Text = "Mensagem"
    $label = New-Object Windows.Forms.Label
    $label.Text = $message
    $label.AutoSize = $true
    $form.Controls.Add($label)
    $form.StartPosition = "CenterScreen"
    $form.AutoSize = $true
    $form.Topmost = $true
    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = 5000
    $timer.Add_Tick({ $form.Close() })
    $timer.Start()
    $form.ShowDialog()
}

function Is-PowerBI-Running {
    $processes = Get-Process -Name "PBIDesktop" -ErrorAction SilentlyContinue
    return $processes -ne $null
}

function Is-File-Open {
    param (
        [string]$filePath
    )
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $processes = Get-Process -Name "PBIDesktop" -ErrorAction SilentlyContinue
    foreach ($proc in $processes) {
        try {
            $openFiles = (Get-Process -Id $proc.Id -FileVersionInfo).FileName
            if ($openFiles -contains $filePath) {
                return $true
            }
        } catch {
            continue
        }
    }
    return $false
}

function Update-PowerBI {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait("%c")
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.SendKeys]::SendWait("r")
}

function Monitor-PowerBI-And-File {
    param (
        [string]$filePath
    )
    while ($true) {
        if (Is-PowerBI-Running -and (Is-File-Open -filePath $filePath)) {
            Write-Output "Arquivo detectado. Iniciando contagem regressiva..."
            for ($i = 15; $i -gt 0; $i--) {
                Write-Output "$i segundos restantes"
                Start-Sleep -Seconds 1
            }
            Show-Message "Vamos atualizar seu Power BI em 5 segundos, aguarde."
            Start-Sleep -Seconds 5
            Update-PowerBI
            Show-Message "Atualização concluída!"
            return
        }
        Start-Sleep -Seconds 1
    }
}

# Caminho do arquivo a ser monitorado
$filePath = "C:\Users\Lenovo\OneDrive - Excel Engenharia\Documentos\Relatórios - Denis\Gestão de Projetos - BI.pbix"

Monitor-PowerBI-And-File -filePath $filePath
