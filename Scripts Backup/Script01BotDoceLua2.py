import pandas as pd
import requests
import datetime
import os
import time
import win32com.client as win32
from PIL import ImageGrab
import subprocess   # <<< ADICIONADO AQUI

# === CONFIGURAÇÕES DO TELEGRAM ===
TOKEN = '8570909503:AAF8h_fee0KOSFg95TT6pOr_dDvTqlLQavk'
CHAT_ID = -1003380097906
URL = f'https://api.telegram.org/bot{TOKEN}/sendMessage'

# === ARQUIVO DE LOG ===
log_path = r"C:\Users\KICEGA~1\OneDrive\Documentos\log_execucao.txt"

def registrar_log(texto):
    with open(log_path, "a", encoding="utf-8") as log:
        log.write(f"{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - {texto}\n")

# === GERA IMAGEM DO INTERVALO I1:AP27 ===
def gerar_imagem_excel():
    global excel, wb
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = True

        # Fecha qualquer workbook que tenha sido aberto automaticamente
        for wb_extra in list(excel.Workbooks):
            wb_extra.Close(SaveChanges=False)

        # Agora abre apenas a planilha desejada
        wb = excel.Workbooks.Open(arquivo)
        time.sleep(10)

        # Fecha qualquer outra planilha que o Excel tenha aberto automaticamente
        for wb_extra in list(excel.Workbooks):
            if wb_extra.FullName != wb.FullName:
                wb_extra.Close(SaveChanges=False)

        ws = wb.Sheets("Acompanhamento-Geral")
        ws.Activate()
        ws.Range("I1:AP27").Select()
        time.sleep(2)

        ws.Range("I1:AP27").CopyPicture(Appearance=1, Format=2)
        time.sleep(2)

        image = ImageGrab.grabclipboard()
        if image is None:
            raise Exception("Nenhuma imagem encontrada na área de transferência.")

        image.save(imagem_path)
        registrar_log("Imagem do calendário gerada com sucesso.")
    except Exception as e:
        registrar_log(f"Erro ao gerar imagem do calendário: {e}")
    finally:
        # Fecha e encerra o Excel completamente
        try:
            wb.Close(SaveChanges=False)
            excel.Quit()
            time.sleep(2)  # garante que o processo seja encerrado
            del wb
            del excel
            registrar_log("Excel fechado completamente após execução.")
        except Exception as e:
            registrar_log(f"Erro ao fechar Excel: {e}")



# === ENVIA IMAGEM DO CALENDÁRIO COM LEGENDA PERSONALIZADA ===
def enviar_imagem_calendario():
    try:
        with open(imagem_path, 'rb') as img:
            files = {'photo': img}
            data = {'chat_id': CHAT_ID, 'caption': 'Veja seu Calendário Programação Pedidos Doce Lua! 🤖🗓️'}
            response = requests.post(f'https://api.telegram.org/bot{TOKEN}/sendPhoto', files=files, data=data)
            registrar_log(f"Status envio imagem: {response.status_code} | Resposta: {response.text}")
            if response.status_code == 200:
                registrar_log("Imagem do calendário enviada com sucesso.")
            else:
                registrar_log("Falha ao enviar imagem.")
        os.remove(imagem_path)
    except Exception as e:
        registrar_log(f"Erro ao enviar imagem do calendário: {e}")


# === LEITURA DA PLANILHA DE MENSAGENS ===
try:
    arquivo = r"C:\Users\KICEGA~1\OneDrive\Documentos\Sistema.Central_Doce.Lua_11-2025.xlsx"
    imagem_path = r"C:\Users\KICEGA~1\OneDrive\Documentos\excel_intervalo.png"
    aba = "Mensagem-Excel"
    df = pd.read_excel(arquivo, sheet_name=aba, header=None)
    mensagens = df.iloc[0:30, 0].dropna().tolist()
    registrar_log(f"Planilha carregada com {len(mensagens)} mensagens.")
except Exception as e:
    registrar_log(f"Erro ao ler a planilha: {e}")
    mensagens = []

# === GERA E ENVIA IMAGEM DO CALENDÁRIO ===
gerar_imagem_excel()
enviar_imagem_calendario()

# === ENVIO DAS MENSAGENS COM CONTROLE DE FLOOD ===
mensagens_enviadas = 0

for i, mensagem in enumerate(mensagens, start=1):
    tentativas = 0
    sucesso = False

    while not sucesso and tentativas < 3:
        try:
            payload = {'chat_id': CHAT_ID, 'text': str(mensagem)}
            response = requests.post(URL, data=payload)

            if response.status_code == 200:
                registrar_log(f"[{i}] Mensagem enviada: {mensagem}")
                sucesso = True
                mensagens_enviadas += 1
            elif response.status_code == 429:
                registrar_log(f"[{i}] Erro 429 - Muitas requisições. Aguardando 10s... (Tentativa {tentativas + 1})")
                time.sleep(10)
                tentativas += 1
            else:
                registrar_log(f"[{i}] Falha ao enviar: {mensagem} | Status: {response.status_code}")
                sucesso = True
        except Exception as e:
            registrar_log(f"[{i}] Erro ao enviar: {mensagem} | Erro: {e}")
            sucesso = True

    time.sleep(1)

registrar_log(f"Total de mensagens enviadas com sucesso: {mensagens_enviadas} de {len(mensagens)}")

# === MENSAGEM FINAL + CHAMADA DO SCRIPT DE GRAVAÇÃO ===
def finalizar_fluxo():
    try:
        mensagem_final = "Atualizando Sistema para identificar pedidos novos 🤖"
        payload = {'chat_id': CHAT_ID, 'text': mensagem_final}
        response = requests.post(URL, data=payload)
        registrar_log(f"Mensagem final enviada: {mensagem_final} | Status: {response.status_code}")
    except Exception as e:
        registrar_log(f"Erro ao enviar mensagem final: {e}")

    # === CHAMA O SCRIPT DE CONTINUIDADE ===
    try:
        registrar_log("Chamando script bot_continuo...")
        subprocess.run(["python", r"C:\Users\KICEGA~1\OneDrive\Documentos\bot_continuo.py"], check=True)
        registrar_log("Script bot_continuo executado com sucesso.")
    except Exception as e:
        registrar_log(f"Erro ao chamar script bot_continuo: {e}")

# Chama a função de finalização
finalizar_fluxo()
