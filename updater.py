import tkinter as tk
from tkinter import messagebox
import requests
import os
import sys
import json
import zipfile
import subprocess
from packaging.version import parse as parse_version

# --- CONFIGURAÇÕES ---
# Versão inicial do seu programa.
CURRENT_VERSION = "2.0" 

# !!! MUDE A URL ABAIXO para o caminho do seu repositório !!!
# O link deve apontar para o arquivo "raw" no GitHub.
VERSION_URL = "https://raw.githubusercontent.com/LucasHO94/BolsaoDesktop/main/version.json"

# Nomes da pasta e do executável principal
APP_FOLDER_NAME = "GestorBolsao"
APP_EXE_NAME = "GestorBolsao.exe"

def check_for_updates():
    try:
        response = requests.get(VERSION_URL)
        response.raise_for_status()
        data = response.json()
        server_version_str = data["version"]
        download_url = data["url"]

        if parse_version(server_version_str) > parse_version(CURRENT_VERSION):
            if messagebox.askyesno("Atualização Disponível", 
                                   f"Uma nova versão ({server_version_str}) está disponível.\n\nDeseja baixar e instalar a atualização agora?"):
                
                update_zip_path = "update.zip"
                with requests.get(download_url, stream=True) as r:
                    r.raise_for_status()
                    with open(update_zip_path, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=8192): 
                            f.write(chunk)
                
                with zipfile.ZipFile(update_zip_path, 'r') as zip_ref:
                    zip_ref.extractall(".")

                os.remove(update_zip_path)
                
                messagebox.showinfo("Atualização Concluída", "O programa foi atualizado com sucesso e será reiniciado.")
                # Reinicia o próprio atualizador para que ele possa lançar a nova versão.
                os.execv(sys.executable, ['python'] + sys.argv)

    except requests.RequestException as e:
        print(f"Erro de Rede: {e}") # Não mostra messagebox se falhar silenciosamente
    except Exception as e:
        messagebox.showerror("Erro na Atualização", f"Ocorreu um erro inesperado: {e}")

def main():
    root = tk.Tk()
    root.withdraw()

    check_for_updates()
    
    app_path = os.path.join(APP_FOLDER_NAME, APP_EXE_NAME)
    if os.path.exists(app_path):
        subprocess.Popen([app_path])
    else:
        messagebox.showerror("Erro ao Iniciar", f"Não foi possível encontrar o executável principal em:\n{app_path}")

if __name__ == "__main__":
    main()