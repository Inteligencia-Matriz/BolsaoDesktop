import base64
import json

# Certifique-se de que o nome do arquivo está correto
filename = 'gcp_service_account.json'

with open(filename, 'rb') as f:
    # Lê o arquivo, codifica para base64 e depois decodifica para uma string de texto puro
    encoded_bytes = base64.b64encode(f.read())
    encoded_string = encoded_bytes.decode('utf-8')

print("--- COPIE A STRING ABAIXO E COLE NO SEU backend.py ---")
print(encoded_string)