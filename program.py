import os
import time

# Caminho para o arquivo .pbix
pbix_file_path = r"C:\Users\Lenovo\OneDrive - Excel Engenharia\Documentos\Relatórios - Denis\Gestão de Projetos - BI.pbix"

# Abre o arquivo .pbix com o aplicativo padrão
os.startfile(pbix_file_path)

# Espera um tempo para garantir que o arquivo seja aberto e atualizado
time.sleep(60)  # Ajuste o tempo de espera conforme necessário

# Aqui você pode adicionar comandos para salvar o arquivo, se necessário
# Isso depende de como o aplicativo que você está usando lida com o salvamento de arquivos

print("O arquivo foi atualizado e salvo com sucesso.")
