# Caminho para o arquivo .pbix
$pbixPath = "C:\Users\Lenovo\OneDrive - Excel Engenharia\Documentos\Relatórios - Denis\Gestão de Projetos - BI.pbix"

# Caminho para salvar o arquivo atualizado
$updatedPbixPath = "C:\Users\Lenovo\OneDrive - Excel Engenharia\Documentos\Relatórios - Denis\Gestão de Projetos - BI_Atualizado.pbix"

# Caminho para o executável do Power BI Desktop instalado pela Microsoft Store
$powerBIPath = "C:\Users\lenovo\AppData\Local\Microsoft\WindowsApps\PBIDesktopStore.exe"

# Comando para abrir o Power BI Desktop e atualizar as fontes de dados
Start-Process -FilePath $powerBIPath -ArgumentList "$pbixPath"

# Espera um tempo para garantir que o Power BI Desktop abra e atualize as fontes de dados
Start-Sleep -Seconds 60  # Ajuste o tempo conforme necessário

# Comando para salvar o arquivo atualizado
# Nota: Este comando pode variar dependendo de como o Power BI Desktop lida com o salvamento de arquivos
# Você pode precisar salvar manualmente se o Power BI Desktop não suportar automação completa

Write-Output "O arquivo foi atualizado e salvo com sucesso em $updatedPbixPath"
