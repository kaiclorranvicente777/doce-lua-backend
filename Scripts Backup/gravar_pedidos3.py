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
log_path = r"C:\Users\Káic e Gabi\OneDrive\Documentos\log_execucao.txt"

def registrar_log(texto):
    with open(log_path, "a", encoding="utf-8") as log:
        log.write(f"{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - {texto}\n")

def enviar_mensagem(texto, tentativas=3, intervalo=5):
    """
    Envia mensagem para o Telegram com tentativas de reenvio.
    tentativas: número máximo de tentativas
    intervalo: tempo em segundos entre cada tentativa
    """
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

def interpretar_mensagem(texto):
    dados = {}
    item_match = re.search(r"Item:\s*(.+)", texto)
    if item_match:
        dados["Produto"] = item_match.group(1).strip()

    sabor_match = re.search(r"Sabor:\s*(.+)\(R\$\s*([\d,]+)\)", texto)
    if sabor_match:
        sabor_texto = sabor_match.group(1).strip()
        preco_unit = sabor_match.group(2).replace(",", ".").strip()
        massa_match = re.search(r"massa de\s+([^\(]+)", sabor_texto, re.IGNORECASE)
        if massa_match:
            dados["SaborMassa"] = massa_match.group(1).strip()
            sabor_recheio = re.sub(r"com\s*massa de\s+[^\(]+", "", sabor_texto, flags=re.IGNORECASE).strip(" ,")
            dados["SaborRecheio"] = sabor_recheio
        else:
            dados["SaborRecheio"] = sabor_texto
        dados["Preco"] = preco_unit

    detalhes_match = re.search(r"Detalhes:\s*(.+)", texto)
    if detalhes_match:
        dados["DescricaoAdicional"] = detalhes_match.group(1).strip()

    adicional_match = re.search(r"Adicional:\s*(.+)", texto)
    if adicional_match:
        adicional_texto = adicional_match.group(1).strip()
        valores = re.findall(r"R\$\s*([\d,]+)", adicional_texto)
        if valores:
            soma = sum(float(v.replace(",", ".")) for v in valores)
            dados["ValorAdicional"] = soma
        if "DescricaoAdicional" in dados:
            dados["DescricaoAdicional"] += f" | {adicional_texto}"
        else:
            dados["DescricaoAdicional"] = adicional_texto

    endereco_match = re.search(r"Endereço.*:\s*(.+)", texto)
    if endereco_match:
        dados["Endereco"] = endereco_match.group(1).strip()
        dados["Entrega"] = "Sim"
    else:
        dados["Entrega"] = "Não"

    data_match = re.search(r"Data:\s*(\d{2}/\d{2})", texto)
    if data_match:
        dados["Data"] = data_match.group(1)

    total_match = re.search(r"Total\s*\(R\$\s*([\d,]+)\)", texto)
    if total_match:
        dados["Preco"] = total_match.group(1).replace(",", ".")

    qtd_un_match = re.search(r"Item:\s*(\d+)\s*Unid\.\s*(.+)", texto)
    if qtd_un_match:
        dados["Quantidade"] = int(qtd_un_match.group(1))
        dados["UnidadeMedida"] = "Unid."
    else:
        dados.setdefault("Quantidade", 1)
        dados.setdefault("UnidadeMedida", "Unid.")

    nome_match = re.search(r"(?:Nome|Cliente)\s*:\s*(.+)", texto, re.IGNORECASE)
    dados["Nome"] = nome_match.group(1).strip() if nome_match else None

    dados.setdefault("Horario", datetime.datetime.now().strftime("%H:%M"))
    dados.setdefault("Pagamento", None)
    dados.setdefault("RecebimentoFoto", None)
    dados.setdefault("PublicacaoInstagram", None)

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

        ws = wb.Sheets(aba)
        registrar_log(f"Aba encontrada: {ws.Name}")

        excel.Visible = True
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
            return None, wb, ws

        wb.Save()
        registrar_log("Pedido gravado com sucesso na planilha.")
        return linha_usada, wb, ws

    except Exception as e:
        registrar_log(f"Erro ao gravar pedido: {e}")
        enviar_mensagem("⚠️ Não consegui gravar o pedido na planilha.")
        return None, None, None

def main():
    arquivo = r"C:\Users\Káic e Gabi\OneDrive\Documentos\Sistema.Central_Doce.Lua_11-2025.xlsx"
    aba = "Pedidos-via-bot"

    mensagem = """Seu pedido Doce Lua confirmado:
🌙 Item: Mini Bolo Confeitado
🌙 Sabor: Pistache (Sabor Premium) com massa de chocolate (R$50,00)
🌙 Detalhes: Branco com trabalho de Chantilly em vermelho e perolas de açúcar brancas
🌙 Adicional: Aplique de escrita em dourado no centro do bolo "Birthday Queen" (R$7,00) + Coroa dourada Shopee (R$11,55)
🌙 Endereço para entrega:
Rua Espirito Santo, 101 - Jacaré
(R$5,00)
🌙 Data: 12/11 (Quarta-Feira)
Total (R$92,55)"""

    dados = interpretar_mensagem(mensagem)
    linha, wb, ws = gravar_excel(dados, arquivo, aba)

    if linha:
        codigo_pedido = str(ws.Range(f"AX{linha}").Value).strip()
        status_atual = ws.Range(f"D{linha}").Value

        # Mensagem de confirmação do pedido
        mensagem_final = f"Pedido {codigo_pedido} Adicionado ao sistema! ✅"
        if status_atual:
            mensagem_final += f" | Status atual do pedido: {status_atual}"
        enviar_mensagem(mensagem_final)
        registrar_log(f"Confirmação enviada para pedido {codigo_pedido}.")

        # === MENSAGEM FINAL DE ENCERRAMENTO ===
        mensagem_encerramento = "Que seu dia seja abençoado, adoçando as vidas das pessoas! 🤖💝"
        enviar_mensagem(mensagem_encerramento)
        registrar_log("Mensagem final de encerramento enviada.")

        # Salvar e fechar workbook
        wb.Save()
        registrar_log("Workbook salvo com sucesso.")
        wb.Close(SaveChanges=True)
        registrar_log("Workbook fechado após gravação.")

    else:
        registrar_log("Nenhum pedido foi gravado (linha não encontrada).")
        enviar_mensagem("⚠️ Não foi possível adicionar o pedido ao sistema.")

if __name__ == "__main__":
    main()
