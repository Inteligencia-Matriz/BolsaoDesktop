# updater.py
import sys
import os
import requests
import zipfile
import subprocess
import time

def main():
    try:
        # Argumentos passados pelo app.py: [1] URL do zip, [2] Path do exe antigo
        zip_url = sys.argv[1]
        old_exe_path = sys.argv[2]
        
        # O diretório onde o .exe está (ex: a Área de Trabalho)
        install_dir = os.path.dirname(old_exe_path)
        # O nome do .exe (ex: GestorBolsao.exe)
        exe_name = os.path.basename(old_exe_path)

        # 1. Espera um pouco para garantir que o programa principal fechou
        time.sleep(2)

        # 2. Baixa o novo arquivo .zip
        update_zip_path = os.path.join(install_dir, "update.zip")
        with requests.get(zip_url, stream=True) as r:
            r.raise_for_status()
            with open(update_zip_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)

        # 3. Extrai o conteúdo do .zip para o diretório de instalação
        with zipfile.ZipFile(update_zip_path, 'r') as zip_ref:
            zip_ref.extractall(install_dir)
            
        # 4. Remove o arquivo .zip baixado
        os.remove(update_zip_path)

        # 5. Reinicia o programa principal já atualizado
        new_exe_path = os.path.join(install_dir, exe_name)
        subprocess.Popen([new_exe_path])

    except Exception as e:
        # Se algo der errado, cria um log para depuração
        with open("update_error_log.txt", "w") as f:
            f.write(f"Ocorreu um erro durante a atualização:\n{str(e)}\n")
            f.write(f"Argumentos recebidos: {str(sys.argv)}")

if __name__ == "__main__":
    main()