"""
Quando abrir o Power BI ele aparece a janela de 'Olá' inicialmente
"""
import psutil
import tkinter as tk
import time

def show_message():
    # Cria uma janela do Tkinter
    root = tk.Tk()
    root.title("Mensagem")
    label = tk.Label(root, text="Olá!")
    label.pack(padx=20, pady=20)
    root.mainloop()

def monitor_power_bi():
    # Lista de processos já existentes
    existing_pids = {proc.pid for proc in psutil.process_iter(['pid', 'name'])}

    while True:
        # Verifica novos processos
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.pid not in existing_pids:
                if proc.info['name'] == 'PBIDesktop.exe':
                    print("Power BI detectado. Iniciando contagem regressiva...")
                    for i in range(15, 0, -1):
                        print(f"{i} segundos restantes")
                        time.sleep(1)
                    show_message()
                    return
                existing_pids.add(proc.pid)
        time.sleep(1)  # Verifica a cada 1 segundo

if __name__ == "__main__":
    monitor_power_bi()
