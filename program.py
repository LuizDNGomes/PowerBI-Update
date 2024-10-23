"""
Versão funcional do programa que aguarda a abertura do arquivo: C:\Users\Lenovo\OneDrive - Excel Engenharia\Documentos\Relatórios - Denis\Gestão de Projetos - BI.pbix
Por hora ele apenas analisa e aparece a janela com o 'Olá'

Proximo passo é inserir as teclas que o Power BI precisa para atualizar e informar ao usuário o que está acontecendo

OBS: O programa encerra depois que entende que o Power BI foi aberto.

"""

import psutil
import tkinter as tk
import time
import os

def show_message():
    # Cria uma janela do Tkinter
    root = tk.Tk()
    root.title("Mensagem")
    label = tk.Label(root, text="Olá!")
    label.pack(padx=20, pady=20)
    root.mainloop()

def is_power_bi_running():
    # Verifica se o Power BI está em execução
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'PBIDesktop.exe':
            return True
    return False

def is_file_open(file_path):
    file_name = os.path.basename(file_path)
    # Verifica se o arquivo específico está aberto no Power BI
    for proc in psutil.process_iter(['name', 'open_files']):
        if proc.info['name'] == 'PBIDesktop.exe':
            try:
                for file in proc.open_files():
                    if file_name in file.path:
                        return True
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                continue
    return False

def monitor_power_bi_and_file(file_path):
    while True:
        if is_power_bi_running() and is_file_open(file_path):
            print("Arquivo detectado. Iniciando contagem regressiva...")
            for i in range(15, 0, -1):
                print(f"{i} segundos restantes")
                time.sleep(1)
            show_message()
            return
        time.sleep(1)  # Verifica a cada 1 segundo

if __name__ == "__main__":
    monitor_power_bi_and_file(r"C:\Users\Lenovo\OneDrive - Excel Engenharia\Documentos\Relatórios - Denis\Gestão de Projetos - BI.pbix")
