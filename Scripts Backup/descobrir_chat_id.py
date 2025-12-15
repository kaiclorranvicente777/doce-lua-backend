import requests

# === CONFIGURAÇÕES DO TELEGRAM ===
TOKEN = '8570909503:AAF8h_fee0KOSFg95TT6pOr_dDvTqlLQavk'
URL = f'https://api.telegram.org/bot{TOKEN}/getUpdates'

# === FAZ A REQUISIÇÃO PARA VER MENSAGENS RECENTES ===
response = requests.get(URL)
dados = response.json()

# === EXIBE OS DADOS RECEBIDOS ===
print(dados)
