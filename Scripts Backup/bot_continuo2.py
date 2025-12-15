import re
import datetime
import logging
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import win32com.client as win32

# === CONFIGURAÇÕES ===
TOKEN = "8570909503:AAF8h_fee0KOSFg95TT6pOr_dDvTqlLQavk"
CHAT_ID = -1003380097906
ARQUIVO = r"C:\Users\KICEGA~1\OneDrive\Documentos\Sistema.Central_Doce.Lua_11-2025.xlsx"
ABA = "Pedidos-via-bot"

# Ativa log detalhado
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.DEBUG
)

# === Função para interpretar pedido ===
def interpretar_mensagem(texto):
    logging.debug(f"Texto recebido para interpretação: {texto}")
    dados = {}

    cliente_match = re.search(r"Cliente:\s*(.+)", texto)
    if cliente_match:
        dados["Nome"] = cliente_match.group(1).strip()

    item_match = re.search(r"Item:\s*(.+)", texto)
    if item_match:
        dados["Produto"] = item_match.group(1).strip()

    qtd_match = re.search(r"(\d+)\s*(Unid\.?|Kg|kg|Caixa|cx)", texto, re.IGNORECASE)
    if qtd_match:
        dados["Quantidade"] = qtd_match.group(1).strip()
        dados["UnidadeMedida"] = qtd_match.group(2).strip()

    sabor_match = re.search(r"Sabor:\s*([^\n]+)", texto)
    if sabor_match:
        valor = sabor_match.group(1).strip()
        valor = re.sub(r"com massa.*", "", valor)
        valor = re.sub(r"\(R\$[0-9.,]+\)", "", valor)
        dados["SaborRecheio"] = valor.strip()

    massa_match = re.search(r"massa\s*de\s*([^\n]+)", texto, re.IGNORECASE)
    if massa_match:
        valor = massa_match.group(1).strip()
        valor = re.sub(r"\(R\$[0-9.,]+\)", "", valor)
        dados["SaborMassa"] = valor.strip()

    adicional_match = re.search(r"Adicional:\s*(.+)", texto)
    if adicional_match:
        dados["DescricaoAdicional"] = adicional_match.group(1).strip()

    adicionais = []
    for match in re.findall(r"\(R\$[0-9.,]+\)", texto):
        if "Total" not in texto and "Sabor" not in texto:
            adicionais.append(match)
    for i, val in enumerate(adicionais):
        if "5,00" in val:
            adicionais[i] = "(R$5,00 = Frete)"
    if adicionais:
        dados["ValorAdicional"] = " + ".join(adicionais)

    endereco_match = re.search(r"Endere[cç]o.*:\s*(.+)", texto)
    if endereco_match:
        dados["Endereco"] = endereco_match.group(1).strip()

    data_match = re.search(r"Data:\s*([0-9]{2}/[0-9]{2})", texto)
    if data_match:
        dados["Data"] = data_match.group(1)

    horario_match = re.search(r"Horário:\s*(.+)", texto)
    if horario_match:
        dados["Horario"] = horario_match.group(1).strip()
    else:
        dados["Horario"] = ""

    preco_match = re.search(r"Total\s*\((R\$[0-9.,]+)\)", texto)
    if preco_match:
        dados["Preco"] = preco_match.group(1).strip()

    dados["Entrega"] = ""

    logging.debug(f"Dados extraídos: {dados}")
    return dados

# === Função para gravar no Excel ===
def gravar_excel(dados, arquivo, aba="Pedidos-via-bot"):
    try:
        logging.debug(f"Tentando abrir arquivo: {arquivo}, aba: {aba}")
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        linha_usada = None
        codigo_pedido = None
        status_atual = None

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            logging.debug(f"Verificando linha {row}, valor AX: {valor_ax}")
            if valor_ax == "#0000":
                linha_usada = row
                logging.debug(f"Linha encontrada para gravação: {linha_usada}")

                # Corrige data para formato brasileiro (dd/mm)
                data_str = dados.get("Data")
                if data_str:
                    try:
                        data_obj = datetime.datetime.strptime(data_str, "%d/%m")
                        ws.Range(f"E{row}").Value = data_obj
                    except Exception as e:
                        logging.warning(f"Falha ao converter data {data_str}: {e}")
                        ws.Range(f"E{row}").Value = "'" + data_str

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
                ws.Range(f"Y{row}").Value = dados.get("Preco")

                codigo_pedido = str(ws.Range(f"AX{row}").Value).strip()
                status_atual = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else "Sem status"
                break

        if linha_usada:
            wb.Save()
            wb.Close(SaveChanges=True)
            excel.Quit()
            logging.info(f"Pedido gravado na linha {linha_usada}, código {codigo_pedido}, status {status_atual}")
            return linha_usada, codigo_pedido, status_atual
        else:
            wb.Close(SaveChanges=False)
            excel.Quit()
            logging.warning("Nenhuma linha disponível encontrada para gravar pedido.")
            return None, None, None

    except Exception as e:
        logging.error(f"Erro ao gravar pedido: {e}", exc_info=True)
        return None, None, None

# Função para listar pedidos abertos (código + cliente)
def listar_pedidos_abertos(arquivo, aba="Pedidos-via-bot"):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(arquivo)
    ws = wb.Sheets(aba)

    pedidos = []
    for row in range(2, ws.UsedRange.Rows.Count + 1):
        codigo = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
        status = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else ""
        cliente = str(ws.Range(f"G{row}").Value).strip() if ws.Range(f"G{row}").Value else ""
        if codigo and codigo != "#0000" and status:  # só lista se tiver status
            pedidos.append(f"{codigo} - {cliente} ({status})")
    wb.Close(SaveChanges=False)
    excel.Quit()
    return pedidos


# Função para listar colunas vazias de um pedido
def listar_colunas_vazias(codigo, arquivo, aba="Pedidos-via-bot"):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(arquivo)
    ws = wb.Sheets(aba)

    colunas_vazias = {}
    mapa = {
        "E": "Data",
        "F": "Horário",
        "G": "Nome",
        "H": "Quantidade",
        "I": "UnidadeMedida",
        "J": "Produto",
        "L": "SaborMassa",
        "M": "Preparação Massa",          # <-- adicionado
        "N": "SaborRecheio",
        "O": "Preparação Recheio",        # <-- adicionado
        "P": "DescricaoAdicional",
        "Q": "ValorAdicional",
        "R": "Coleta Detalhes adicionais",# <-- adicionado
        "S": "Endereco",
        "U": "Entrega",
        "Y": "Preco"
    }

    for row in range(2, ws.UsedRange.Rows.Count + 1):
        valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
        if valor_ax == codigo:
            for col, nome in mapa.items():
                valor = ws.Range(f"{col}{row}").Value
                if not valor:
                    colunas_vazias[nome] = col
            break

    wb.Close(SaveChanges=False)
    excel.Quit()
    return colunas_vazias


# Função para atualizar coluna escolhida e retornar status
def atualizar_coluna(codigo, coluna_excel, valor, arquivo, aba="Pedidos-via-bot"):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(arquivo)
    ws = wb.Sheets(aba)

    for row in range(2, ws.UsedRange.Rows.Count + 1):
        valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
        if valor_ax == codigo:
            ws.Range(f"{coluna_excel}{row}").Value = valor
            status_atual = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else "Sem status"
            wb.Save()
            wb.Close(SaveChanges=True)
            excel.Quit()
            return True, status_atual

    wb.Close(SaveChanges=False)
    excel.Quit()
    return False, None


# === Handler de mensagens ===
async def mensagens(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip()
    logging.info(f"Mensagem recebida: {texto}")

    # Fluxo de atualização
    if "atualizar pedido" in texto.lower():
        pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
        if pedidos_abertos:
            lista = "\n".join(pedidos_abertos)
            await update.message.reply_text(
                f"Qual pedido você quer atualizar?\nPedidos em aberto:\n{lista}"
            )
            context.user_data["esperando_codigo"] = True
        else:
            await update.message.reply_text("⚠️ Nenhum pedido em aberto encontrado.")
        return

    # Receber código do pedido
    if context.user_data.get("esperando_codigo"):
        codigo = texto.strip()
        context.user_data["codigo_pedido"] = codigo
        colunas_vazias = listar_colunas_vazias(codigo, ARQUIVO, ABA)
        if colunas_vazias:
            lista = "\n".join([f"{i+1}. {nome}" for i, nome in enumerate(colunas_vazias.keys())])
            await update.message.reply_text(
                f"As colunas vazias do pedido {codigo} são:\n{lista}\n\nResponda no formato: NomeColuna: Valor"
            )
            context.user_data["colunas_vazias"] = colunas_vazias
            context.user_data["esperando_atualizacao"] = True
        else:
            await update.message.reply_text(f"⚠️ O pedido {codigo} não possui colunas vazias.")
        context.user_data["esperando_codigo"] = False
        return

    # Receber atualização
    if context.user_data.get("esperando_atualizacao"):
        codigo = context.user_data.get("codigo_pedido")
        colunas_vazias = context.user_data.get("colunas_vazias", {})
        if ":" in texto:
            chave, valor = texto.split(":", 1)
            chave = chave.strip()
            valor = valor.strip()

            # Se o usuário digitou um número (ex: "2: ok")
            if chave.isdigit():
                indice = int(chave) - 1
                nomes_colunas = list(colunas_vazias.keys())
                if 0 <= indice < len(nomes_colunas):
                    nome_coluna = nomes_colunas[indice]
                else:
                    await update.message.reply_text("⚠️ Número inválido. Escolha um dos números listados.")
                    return
            else:
                # Se digitou o nome da coluna (ex: "Preparação Recheio: ok")
                nome_coluna = chave

            if nome_coluna in colunas_vazias:
                coluna_excel = colunas_vazias[nome_coluna]
                ok, status_atual = atualizar_coluna(codigo, coluna_excel, valor, ARQUIVO, ABA)
                if ok:
                    await update.message.reply_text(
                         f"Pedido {codigo} atualizado com sucesso! ✅ | {nome_coluna}: {valor}"
                    )
                    await update.message.reply_text(
                        f"Status atual do pedido {codigo}: {status_atual}"
                    )
                else:
                    await update.message.reply_text(f"⚠️ Não foi possível atualizar o pedido {codigo}.")
            else:
                await update.message.reply_text(f"⚠️ A coluna '{nome_coluna}' não está disponível para atualização.")
        else:
            await update.message.reply_text("⚠️ Use o formato: NomeColuna: Valor ou Número listado: Valor")
        context.user_data["esperando_atualizacao"] = False
        return


    # Fluxo normal de adicionar pedido
    if "adicionar pedido ao sistema" in texto.lower():
        await update.message.reply_text("✅ Comando recebido! Me diga os detalhes do pedido!")
        context.user_data["esperando_pedido"] = True
    else:
        if context.user_data.get("esperando_pedido"):
            dados = interpretar_mensagem(texto)
            linha, codigo_pedido, status_atual = gravar_excel(dados, ARQUIVO, ABA)
            if linha:
                mensagem_final = f"Pedido {codigo_pedido} Adicionado ao sistema! ✅ | Status atual do pedido: {status_atual}"
                await update.message.reply_text(mensagem_final)
            else:
                await update.message.reply_text("⚠️ Não foi possível gravar o pedido.")
            context.user_data["esperando_pedido"] = False


# === Handlers ===
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logging.info("Comando /start recebido")
    await update.message.reply_text("🤖 Bot ativo! Envie 'Adicionar pedido ao sistema' para começar.")

# === Mensagem inicial automática via post_init ===
async def avisar_inicio(app):
    await app.bot.send_message(
        chat_id=CHAT_ID,
        text="🤖 Bot ativo! Estou aguardando suas interações para trabalharmos juntos!"
    )

# === Main ===
def main():
    app = ApplicationBuilder().token(TOKEN).post_init(avisar_inicio).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, mensagens))

    print("Bot rodando... Ctrl+C para parar.")
    app.run_polling()

if __name__ == "__main__":
    main()
