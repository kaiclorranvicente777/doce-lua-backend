import pandas as pd
import requests
import datetime
import os
import time
import win32com.client as win32
from PIL import ImageGrab
import subprocess
import socket
import re

# === CONFIGURAÇÕES DO TELEGRAM ===
TOKEN = os.getenv("TELEGRAM_TOKEN", "8570909503:AAF8h_fee0KOSFg95TT6pOr_dDvTqlLQavk")
CHAT_ID = -1003380097906
URL = f'https://api.telegram.org/bot{TOKEN}/sendMessage'

# === ARQUIVO DE LOG ===
log_path = r"C:\Users\KICEGA~1\OneDrive\Documentos\log_execucao.txt"

def registrar_log(texto):
    with open(log_path, "a", encoding="utf-8") as log:
        log.write(f"{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - {texto}\n")

# === Função para aguardar internet ===
def aguardar_internet(host="api.telegram.org", port=443, timeout=5, tentativas=30, intervalo=10):
    for i in range(tentativas):
        try:
            socket.create_connection((host, port), timeout=timeout)
            registrar_log("[aguardar_internet] Conexão restabelecida.")
            return True
        except OSError:
            registrar_log(f"[aguardar_internet] Tentativa {i+1}/{tentativas}: sem conexão. Aguardando {intervalo}s...")
            time.sleep(intervalo)
    registrar_log("[aguardar_internet] Não foi possível restabelecer a conexão.")
    return False

# === GERA IMAGEM DO INTERVALO I1:AP27 ===
def gerar_imagem_excel():
    global excel, wb
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = True

        for wb_extra in list(excel.Workbooks):
            wb_extra.Close(SaveChanges=False)

        wb = excel.Workbooks.Open(arquivo)
        time.sleep(10)

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
        registrar_log("[gerar_imagem_excel] Imagem gerada com sucesso.")
    except Exception as e:
        registrar_log(f"[gerar_imagem_excel] Erro: {e}")
    finally:
        try:
            wb.Close(SaveChanges=False)
            excel.Quit()
            time.sleep(2)
            del wb
            del excel
            registrar_log("[gerar_imagem_excel] Excel fechado completamente.")
        except Exception as e:
            registrar_log(f"[gerar_imagem_excel] Erro ao fechar Excel: {e}")

# === ENVIA IMAGEM DO CALENDÁRIO ===
def enviar_imagem_calendario():
    try:
        if not aguardar_internet():
            registrar_log("[enviar_imagem_calendario] Sem internet.")
            return False

        with open(imagem_path, 'rb') as img:
            files = {'photo': img}
            data = {'chat_id': CHAT_ID, 'caption': 'Veja seu Calendário Programação Pedidos Doce Lua! 🤖🗓️'}
            response = requests.post(f'https://api.telegram.org/bot{TOKEN}/sendPhoto', files=files, data=data)
            registrar_log(f"[enviar_imagem_calendario] Status: {response.status_code} | Resposta: {response.text}")
            if response.status_code == 200:
                registrar_log("[enviar_imagem_calendario] Imagem enviada com sucesso.")
                os.remove(imagem_path)
                return True
            else:
                registrar_log("[enviar_imagem_calendario] Falha ao enviar imagem.")
                return False
    except Exception as e:
        registrar_log(f"[enviar_imagem_calendario] Erro: {e}")
        return False

# === LEITURA DA PLANILHA DE MENSAGENS ===
try:
    arquivo = r"C:\Users\KICEGA~1\OneDrive\Documentos\Sistema.Central_Doce.Lua_11-2025.xlsx"
    imagem_path = r"C:\Users\KICEGA~1\OneDrive\Documentos\excel_intervalo.png"
    aba = "Mensagem-Excel"
    df = pd.read_excel(arquivo, sheet_name=aba, header=None)
    mensagens = df.iloc[0:30, 0].dropna().tolist()

    # Ordena por data extraída do texto
    def extrair_data(msg):
        padrao = re.search(r"Data:\s*(\d{2}/\d{2}/\d{4})", msg)
        if padrao:
            return datetime.datetime.strptime(padrao.group(1), "%d/%m/%Y")
        return datetime.datetime.max

    mensagens.sort(key=extrair_data)
    registrar_log(f"[planilha] Carregada com {len(mensagens)} mensagens (ordenadas por data).")
except Exception as e:
    registrar_log(f"[planilha] Erro ao ler: {e}")
    mensagens = []

# === GERA E ENVIA IMAGEM DO CALENDÁRIO ===
gerar_imagem_excel()
enviar_imagem_calendario()

# === ENVIO DAS MENSAGENS EM BLOCO (respeitando limite do Telegram) ===
try:
    if mensagens:
        texto_unico = "Pedidos em Aberto - Sistema Central Doce Lua\n\n" + "\n\n".join(mensagens)
        limite = 4096
        partes = [texto_unico[i:i+limite] for i in range(0, len(texto_unico), limite)]

        for idx, parte in enumerate(partes, start=1):
            if not aguardar_internet():
                registrar_log("[mensagens] Sem internet para enviar.")
                break

            payload = {'chat_id': CHAT_ID, 'text': parte}
            response = requests.post(URL, data=payload)

            if response.status_code == 200:
                registrar_log(f"[mensagens] Parte {idx}/{len(partes)} enviada com sucesso.")
            elif response.status_code == 429:
                wait = response.json().get("parameters", {}).get("retry_after", 10)
                registrar_log(f"[mensagens] Flood detectado. Aguardando {wait}s...")
                time.sleep(wait)
                continue
            else:
                registrar_log(f"[mensagens] Falha ao enviar parte {idx} | Status: {response.status_code}")
            time.sleep(1)
    else:
        registrar_log("[mensagens] Nenhuma mensagem encontrada para enviar.")
except Exception as e:
    registrar_log(f"[mensagens] Erro ao enviar: {e}")

# === FUNÇÕES AUXILIARES PARA VERIFICAR MENSAGENS ===
def buscar_mensagens():
    try:
        if not aguardar_internet():
            registrar_log("[buscar_mensagens] Sem internet.")
            return []
        resp = requests.get(f"https://api.telegram.org/bot{TOKEN}/getUpdates")
        dados = resp.json()
        mensagens = []
        for update in dados.get("result", []):
            msg = update.get("message", {})
            texto = msg.get("text", "")
            data = msg.get("date", 0)
            mensagens.append((texto, datetime.datetime.fromtimestamp(data)))
        return mensagens
    except Exception as e:
        registrar_log(f"[buscar_mensagens] Erro: {e}")
        return []

def verificar_intervalo(mensagens):
    saudacao = None
    despedida = None
    for texto, data in mensagens:
        if "Pedidos em Aberto - Sistema Central Doce Lua" in texto:
            saudacao = data
        if "🤖 Hora da soneca!" in texto:
            despedida = data
    if saudacao and despedida and despedida < saudacao:
        for texto, data in mensagens:
            if despedida < data < saudacao and "Adicionar pedido ao sistema" in texto:
                return True
    return False

# === MENSAGEM FINAL + CHAMADA DO SCRIPT DE GRAVAÇÃO ===
def finalizar_fluxo():
    try:
        if not aguardar_internet():
            registrar_log("[finalizar_fluxo] Sem internet para enviar mensagem final.")
            return
        mensagem_final = "Atualizando Sistema para identificar pedidos novos 🤖"
        payload = {'chat_id': CHAT_ID, 'text': mensagem_final}
        response = requests.post(URL, data=payload)
        registrar_log(f"[finalizar_fluxo] Mensagem final enviada | Status: {response.status_code}")
    except Exception as e:
        registrar_log(f"[finalizar_fluxo] Erro ao enviar mensagem final: {e}")

    try:
        registrar_log("[finalizar_fluxo] Verificando mensagens entre despedida e saudação...")
        mensagens = buscar_mensagens()
        if verificar_intervalo(mensagens):
            registrar_log("[finalizar_fluxo] Mensagem 'Adicionar pedido ao sistema' encontrada. Executando ScriptIntermediario...")
            script_path = r"C:\Users\KICEGA~1\OneDrive\Documentos\ScriptIntermediario.py"
        else:
            registrar_log("[finalizar_fluxo] Nenhuma mensagem encontrada. Executando bot_continuo...")
            script_path = r"C:\Users\KICEGA~1\OneDrive\Documentos\bot_continuo.py"

        # Executa o script escolhido com captura de saída
        result = subprocess.run(
            ["python", script_path],
            capture_output=True, text=True
        )
        registrar_log(f"[finalizar_fluxo] stdout: {result.stdout}")
        registrar_log(f"[finalizar_fluxo] stderr: {result.stderr}")
        if result.returncode == 0:
            registrar_log(f"[finalizar_fluxo] Script {os.path.basename(script_path)} executado com sucesso.")
        else:
            registrar_log(f"[finalizar_fluxo] Script {os.path.basename(script_path)} terminou com erro (código {result.returncode}).")
    except Exception as e:
        registrar_log(f"[finalizar_fluxo] Erro crítico ao executar subprocesso: {e}")

# Chama a função de finalização
finalizar_fluxo()
