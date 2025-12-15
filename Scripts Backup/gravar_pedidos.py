import re
import datetime
import requests
import win32com.client as win32
import time

# === CONFIGURAÇÕES DO TELEGRAM ===
TOKEN = '8570909503:AAF8h_fee0KOSFg95TT6pOr_dDvTqlLQavk'
CHAT_ID = -1003380097906
URL = f'https://api.telegram.org/bot{TOKEN}/sendMessage'

# === ARQUIVO DE LOG ===
log_path = r"C:\Users\KICEGA~1\OneDrive\Documentos\log_execucao.txt"

def registrar_log(texto):
    with open(log_path, "a", encoding="utf-8") as log:
        log.write(f"{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - {texto}\n")

def enviar_mensagem(texto, tentativas=3, intervalo=5):
    for tentativa in range(1, tentativas + 1):
        try:
            payload = {'chat_id': CHAT_ID, 'text': texto}
            response = requests.post(URL, data=payload, timeout=10)
            if response.status_code == 200:
                registrar_log(f"Mensagem enviada: {texto} | Status: {response.status_code}")
                return True
            else:
                registrar_log(f"Tentativa {tentativa} falhou: Status {response.status_code}")
        except Exception as e:
            registrar_log(f"Tentativa {tentativa} falhou: {e}")
        if tentativa < tentativas:
            registrar_log(f"Aguardando {intervalo}s para nova tentativa...")
            time.sleep(intervalo)
    registrar_log(f"Falha ao enviar mensagem após {tentativas} tentativas: {texto}")
    return False

def obter_mensagens_telegram():
    try:
        url_updates = f"https://api.telegram.org/bot{TOKEN}/getUpdates"
        response = requests.get(url_updates, timeout=10)
        if response.status_code == 200:
            dados = response.json()
            if "result" in dados and len(dados["result"]) > 0:
                return dados["result"]
        registrar_log("Nenhuma mensagem encontrada no Telegram.")
    except Exception as e:
        registrar_log(f"Erro ao obter mensagens do Telegram: {e}")
    return []

def interpretar_mensagem(texto):
    dados = {}
    item_match = re.search(r"Item:\s*(.+)", texto)
    if item_match:
        dados["Produto"] = item_match.group(1).strip()
    # ... (mantém o restante da interpretação igual ao seu código)
    dados.setdefault("Horario", datetime.datetime.now().strftime("%H:%M"))
    return dados

def gravar_excel(dados, arquivo, aba="Pedidos-via-bot"):
    try:
        try:
            excel = win32.GetActiveObject("Excel.Application")
            registrar_log("Usando Excel ativo.")
        except Exception:
            excel = win32.gencache.EnsureDispatch("Excel.Application")
            registrar_log("Criando nova instância do Excel.")

        wb = None
        for book in excel.Workbooks:
            if book.FullName.lower() == arquivo.lower():
                wb = book
                break
        if not wb:
            wb = excel.Workbooks.Open(arquivo)
            registrar_log("Workbook aberto pelo caminho.")

        try:
            ws = wb.Sheets(aba)
        except Exception:
            registrar_log(f"Aba {aba} não encontrada no Excel.")
            enviar_mensagem(f"⚠️ Aba {aba} não encontrada na planilha.")
            return None, None, None

        registrar_log(f"Aba encontrada: {ws.Name}")
        excel.Visible = False
        ws.Activate()

        linha_usada = None
        for row in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            if valor_ax == "#0000":
                linha_usada = row
                ws.Range(f"E{row}").Value = dados.get("Data")
                ws.Range(f"F{row}").Value = dados.get("Horario")
                ws.Range(f"G{row}").Value = dados.get("Nome")
                ws.Range(f"H{row}").Value = dados.get("Quantidade")
                ws.Range(f"I{row}").Value = dados.get("UnidadeMedida")
                ws.Range(f"J{row}").Value = dados.get("Produto")
                ws.Range(f"L{row}").Value = dados.get("SaborMassa")
                ws.Range(f"N{row}").Value = dados.get("SaborRecheio")
                ws.Range(f"P{row}").Value = dados.get("DescricaoAdicional")
                ws.Range(f"Q{row}").Value = dados.get("ValorAdicional")
                ws.Range(f"S{row}").Value = dados.get("Endereco")
                ws.Range(f"U{row}").Value = dados.get("Entrega")
                ws.Range(f"V{row}").Value = dados.get("Pagamento")
                ws.Range(f"W{row}").Value = dados.get("RecebimentoFoto")
                ws.Range(f"X{row}").Value = dados.get("PublicacaoInstagram")
                ws.Range(f"Y{row}").Value = dados.get("Preco")
                break

        if not linha_usada:
            registrar_log("Nenhuma linha livre encontrada (sem #0000).")
            enviar_mensagem("⚠️ Nenhuma linha livre encontrada para registrar pedido.")
            return None, None, None

        wb.Save()
        registrar_log("Pedido gravado com sucesso na planilha.")
        wb.Close(SaveChanges=True)
        registrar_log("Workbook fechado após gravação.")
        excel.Quit()
        registrar_log("Excel encerrado.")

        return linha_usada, None, None

    except Exception as e:
        registrar_log(f"Erro ao gravar pedido: {e}")
        enviar_mensagem("⚠️ Não consegui gravar o pedido na planilha.")
        return None, None, None

def main():
    arquivo = r"C:\Users\KICEGA~1\OneDrive\Documentos\Sistema.Central_Doce.Lua_11-2025.xlsx"
    aba = "Pedidos-via-bot"

    mensagens = obter_mensagens_telegram()
    if not mensagens:
        enviar_mensagem("⚠️ Não consegui obter mensagens do Telegram.")
        return

    registrar_log(f"Total de mensagens recebidas do Telegram: {len(mensagens)}")
    ultima_msg = mensagens[-1]["message"]["text"].strip().lower()

    if "adicionar pedido ao sistema" in ultima_msg:
        registrar_log("Comando recebido. Aguardando próxima mensagem como pedido...")
        enviar_mensagem("✅ Comando recebido! Envie agora os detalhes do pedido.")
        return
    else:
        pedido_texto = mensagens[-1]["message"]["text"]
        registrar_log(f"Pedido recebido do Telegram: {pedido_texto}")

    dados = interpretar_mensagem(pedido_texto)
    linha, wb, ws = gravar_excel(dados, arquivo, aba)

    if linha:
        codigo_pedido = "#0001"  # exemplo, pode ajustar conforme sua lógica
        mensagem_final = f"Pedido {codigo_pedido} Adicionado ao sistema! ✅"
        enviar_mensagem(mensagem_final)
        registrar_log(f"Confirmação enviada para pedido {codigo_pedido}.")
        mensagem_encerramento = "Que seu dia seja abençoado, adoçando as vidas das pessoas! 🤖💝"
        enviar_mensagem(mensagem_encerramento)
        registrar_log("Mensagem final de encerramento enviada.")
    else:
        registrar_log("Nenhum pedido foi gravado (linha não encontrada).")
        enviar_mensagem("⚠️ Não foi possível adicionar o pedido ao sistema.")

if __name__ == "__main__":
    main()
