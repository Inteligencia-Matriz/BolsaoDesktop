# updater.py (Versão Final e Robusta)
import sys
import os
import requests
import zipfile
import subprocess
import time

# O nome do executável DENTRO do arquivo .zip. Este nome deve ser consistente.
EXE_NAME_IN_ZIP = "GestorBolsao.exe"

def main():
    try:
        # Argumentos passados pelo app.py: [1] URL do zip, [2] Path do exe antigo
        zip_url = sys.argv[1]
        old_exe_path = sys.argv[2]
        
        install_dir = os.path.dirname(old_exe_path)

        # 1. Espera um pouco para garantir que o programa principal fechou
        time.sleep(2)
        
        # 2. Deleta o executável antigo PRIMEIRO
        if os.path.exists(old_exe_path):
            os.remove(old_exe_path)

        # 3. Baixa o novo arquivo .zip
        update_zip_path = os.path.join(install_dir, "update.zip")
        with requests.get(zip_url, stream=True) as r:
            r.raise_for_status()
            with open(update_zip_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)

        # 4. Extrai o conteúdo do .zip (o novo GestorBolsao.exe)
        with zipfile.ZipFile(update_zip_path, 'r') as zip_ref:
            zip_ref.extractall(install_dir)
            
        # 5. Remove o arquivo .zip baixado
        os.remove(update_zip_path)

        # 6. Reinicia o programa principal usando o NOME PADRÃO que estava no zip
        new_exe_path = os.path.join(install_dir, EXE_NAME_IN_ZIP)
        if os.path.exists(new_exe_path):
            subprocess.Popen([new_exe_path])

    except Exception as e:
        # Se algo der errado, cria um log para depuração
        log_path = os.path.join(os.path.dirname(sys.argv[0]), "update_error_log.txt")
        with open(log_path, "w") as f:
            f.write(f"Ocorreu um erro durante a atualização:\n{str(e)}\n")
            f.write(f"Argumentos recebidos: {str(sys.argv)}")

if __name__ == "__main__":
    main()