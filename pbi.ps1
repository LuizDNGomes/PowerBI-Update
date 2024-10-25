# Caminho para o arquivo PBIX
$pbixPath = "C:\Users\Lenovo\OneDrive - Excel Engenharia\Documentos\Relatórios - Denis\Gestão de Projetos - BI.pbix"

# Caminho para o executável do Power BI Desktop instalado pela Microsoft Store
$powerBIPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\Microsoft.MicrosoftPowerBIDesktop_8wekyb3d8bbwe\PBIDesktop.exe"

# Abrir o Power BI Desktop com o arquivo PBIX
Start-Process -FilePath $powerBIPath -ArgumentList $pbixPath

# Esperar o Power BI Desktop abrir (ajuste o tempo conforme necessário)
Start-Sleep -Seconds 30

# Enviar a tecla F5 para atualizar os dados
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::SendWait("{F5}")

# Esperar a atualização completar (ajuste o tempo conforme necessário)
Start-Sleep -Seconds 300

# Fechar o Power BI Desktop
Stop-Process -Name "PBIDesktop"
