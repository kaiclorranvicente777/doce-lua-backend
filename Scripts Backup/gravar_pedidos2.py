import datetime
import os
import requests
import re
from openpyxl import load_workbook

# === CONFIGURAÇÕES DO TELEGRAM ===
TOKEN = '8570909503:AAF8h_fee0KOSFg95TT6pOr_dDvTqlLQavk'
CHAT_ID = -1003380097906
URL = f'https://api.telegram.org/bot{TOKEN}/sendMessage'

# === ARQUIVO DE LOG ===
log_path = r"C:\Users\Káic e Gabi\OneDrive\Documentos\log_execucao.txt"

def registrar_log(texto):
    with open(log_path, "a", encoding="utf-8") as log:
        log.write(f"{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - {texto}\n")

# === INTERPRETA MENSAGEM DO TELEGRAM ===
def interpretar_mensagem(texto):
    # Verifica se há mais de um "Item:"
    itens = re.findall(r"\bItem:\b", texto)
    if len(itens) > 1:
        return {"erro": "Mais de um pedido na mesma mensagem, favor corrigir!"}

    dados = {}

    # Produto
    item_match = re.search(r"Item:\s*(.+)", texto)
    if item_match:
        dados["Produto"] = item_match.group(1).strip()

    # Sabor + Massa + Preço
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

        dados.setdefault("Preco", preco_unit)

    # Detalhes
    detalhes_match = re.search(r"Detalhes:\s*(.+)", texto)
    if detalhes_match:
        dados["DescricaoAdicional"] = detalhes_match.group(1).strip()

    # Adicional
    adicional_match = re.search(r"Adicional:\s*(.+)", texto)
    if adicional_match:
        adicional_texto = adicional_match.group(1).strip()
        valores = re.findall(r"R\$\s*([\d,]+)", adicional_texto)
        if valores:
            soma = sum(float(v.replace(",", ".")) for v in valores)
            dados["ValorAdicional"] = soma
        if "DescricaoAdicional" in dados:
            dados["DescricaoAdicional"] = f'{dados["DescricaoAdicional"]} | {adicional_texto}'
        else:
            dados["DescricaoAdicional"] = adicional_texto

    # Endereço (ignora frete)
    endereco_match = re.search(r"Endereço.*:\s*(.+)", texto)
    if endereco_match:
        dados["Endereco"] = endereco_match.group(1).strip()
        dados["Entrega"] = "Sim"
        dados["ColetaAdicional"] = None  # frete não vai para planilha

    # Data
    data_match = re.search(r"Data:\s*(\d{2}/\d{2})", texto)
    if data_match:
        dados["Data"] = data_match.group(1)

    # Total
    total_match = re.search(r"Total\s*\(R\$\s*([\d,]+)\)", texto)
    if total_match:
        dados["Preco"] = total_match.group(1).replace(",", ".")

    # Quantidade / UnidadeMedida
    qtd_un_match = re.search(r"Item:\s*(\d+)\s*Unid\.\s*(.+)", texto)
    if qtd_un_match:
        dados["Quantidade"] = int(qtd_un_match.group(1))
        dados["UnidadeMedida"] = "Unid."
    else:
        dados.setdefault("Quantidade", 1)
        dados.setdefault("UnidadeMedida", "Unid.")

    # Nome do cliente (se houver explicitamente)
    nome_match = re.search(r"(?:Nome|Cliente)\s*:\s*(.+)", texto, re.IGNORECASE)
    dados["Nome"] = nome_match.group(1).strip() if nome_match else None

    # Campos opcionais ficam em branco
    dados.setdefault("Horario", None)
    dados.setdefault("Pagamento", None)
    dados.setdefault("RecebimentoFoto", None)
    dados.setdefault("PublicacaoInstagram", None)

    return dados

# === GRAVAÇÃO NA PLANILHA ===
def gravar_excel(dados, arquivo, aba="Pedidos-via-bot"):
    wb = load_workbook(arquivo, data_only=True)
    ws = wb[aba]
    linha_usada = None

    for row in range(2, ws.max_row + 1):
        valor_ax = str(ws[f"AX{row}"].value).strip()
        if valor_ax == "#0000":
            linha_usada = row
            ws[f"E{row}"] = dados.get("Data")
            ws[f"F{row}"] = dados.get("Horario")
            ws[f"G{row}"] = dados.get("Nome")
            ws[f"H{row}"] = dados.get("Quantidade")
            ws[f"I{row}"] = dados.get("UnidadeMedida")
            ws[f"J{row}"] = dados.get("Produto")
            ws[f"L{row}"] = dados.get("SaborMassa")
            ws[f"M{row}"] = None  # PrepMassa não vem da mensagem
            ws[f"N{row}"] = dados.get("SaborRecheio")
            ws[f"O{row}"] = None  # PrepRecheio não vem da mensagem
            ws[f"P{row}"] = dados.get("DescricaoAdicional")
            ws[f"Q{row}"] = dados.get("ValorAdicional")
            ws[f"R{row}"] = dados.get("ColetaAdicional")
            ws[f"S{row}"] = dados.get("Endereco")
            ws[f"T{row}"] = None  # Contato não vem da mensagem
            ws[f"U{row}"] = dados.get("Entrega")
            ws[f"V{row}"] = dados.get("Pagamento")
            ws[f"W{row}"] = dados.get("RecebimentoFoto")
            ws[f"X{row}"] = dados.get("PublicacaoInstagram")
            ws[f"Y{row}"] = dados.get("Preco")
            break

    wb.save(arquivo)
    return linha_usada

# === ENVIO DE MENSAGENS AO TELEGRAM ===
def enviar_mensagem(texto):
    try:
        payload = {'chat_id': CHAT_ID, 'text': texto}
        response = requests.post(URL, data=payload)
        registrar_log(f"Mensagem enviada: {texto} | Status: {response.status_code}")
    except Exception as e:
        registrar_log(f"Erro ao enviar mensagem: {e}")

# === FLUXO PRINCIPAL ===
def main():
    arquivo = r"C:\Users\Káic e Gabi\OneDrive\Documentos\Sistema.Central_Doce.Lua_11-2025.xlsx"
    aba = "Pedidos-via-bot"

    # Aqui você receberia a mensagem real do Telegram
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

    if "erro" in dados:
        enviar_mensagem(dados["erro"])
        return

    linha = gravar_excel(dados, arquivo, aba)

    if linha:
        # 1ª mensagem: confirmação simples
        enviar_mensagem("Pedido registrado. ✅")

        # 2ª mensagem: código e status
        wb = load_workbook(arquivo, data_only=True)
        ws = wb[aba]
        codigo_pedido = str(ws[f"AX{linha}"].value).strip()
        status_atual = ws[f"D{linha}"].value

        mensagem_final = f"Pedido {codigo_pedido} Adicionado ao sistema! ✅"
        if status_atual:
            mensagem_final += f" | Status atual do pedido: {status_atual}"

        enviar_mensagem(mensagem_final)
    else:
        enviar_mensagem("Nenhuma linha livre encontrada para registrar pedido.")

    # === LÓGICA DE DESLIGAMENTO ===
    hora_atual = datetime.datetime.now().time()
    hora_inicio = datetime.time(7, 0)
    hora_fim = datetime.time(8, 0)

    try:
        if hora_inicio <= hora_atual <= hora_fim:
            mensagem_desligar = "Estou deixando sua máquina descansar automaticamente! Ligue-a após as 8:00 🤖☕"
            enviar_mensagem(mensagem_desligar)
            registrar_log("Desligamento automático acionado.")
            os.system("shutdown /s /t 60")
        else:
            mensagem_ligada = "Sua máquina continuará ligada para que você trabalhe sem interrupções! 🤖🖥️"
            enviar_mensagem(mensagem_ligada)
            registrar_log("Máquina mantida ligada.")
    except Exception as e:
        registrar_log(f"Erro ao verificar hora ou desligar: {e}")

if __name__ == "__main__":
    main()
