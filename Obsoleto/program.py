import psutil
import tkinter as tk
import time
import os
import pyautogui

def show_message(message):
    # Cria uma janela do Tkinter
    root = tk.Tk()
    root.title("Mensagem")
    label = tk.Label(root, text=message)
    label.pack(padx=20, pady=20)
    root.after(5000, root.destroy)  # Fecha a janela após 5 segundos
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

def update_power_bi():
    # Simula as teclas para atualizar o Power BI
    pyautogui.press('alt')
    time.sleep(0.5)
    pyautogui.press('c')
    time.sleep(0.5)
    pyautogui.press('r')

def monitor_power_bi_and_file(file_path):
    while True:
        if is_power_bi_running() and is_file_open(file_path):
            print("Arquivo detectado. Iniciando contagem regressiva...")
            for i in range(15, 0, -1):
                print(f"{i} segundos restantes")
                time.sleep(1)
            show_message("Vamos atualizar seu Power BI em 5 segundos, aguarde.")
            time.sleep(5)  # Aguarda 5 segundos antes de atualizar
            update_power_bi()
            show_message("Atualização concluída!")
            return
        time.sleep(1)  # Verifica a cada 1 segundo

if __name__ == "__main__":
    monitor_power_bi_and_file(r"C:\Users\Lenovo\OneDrive - Excel Engenharia\Documentos\Relatórios - Denis\Gestão de Projetos - BI.pbix")

#Funciona