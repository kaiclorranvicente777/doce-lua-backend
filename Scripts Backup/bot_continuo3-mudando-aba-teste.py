import re
import datetime
import logging
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import win32com.client as win32
import requests
import signal
import sys
import win32api
import win32con

# === CONFIGURAÇÕES ===
TOKEN = "8570909503:AAF8h_fee0KOSFg95TT6pOr_dDvTqlLQavk"
CHAT_ID = -1003380097906
ARQUIVO = r"C:\Users\KICEGA~1\OneDrive\Documentos\Sistema.Central_Doce.Lua_11-2025.xlsx"
ABA = "Acompanhamento-Pedidos"
URL = f"https://api.telegram.org/bot{TOKEN}/sendMessage"

# Ativa log detalhado
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.DEBUG
)

# === Handler de fechamento da janela do CMD ===
def on_console_event(event):
    if event in (win32con.CTRL_CLOSE_EVENT, win32con.CTRL_LOGOFF_EVENT, win32con.CTRL_SHUTDOWN_EVENT):
        try:
            logging.info("Evento de fechamento de console capturado. Enviando despedida...")
            enviar_despedida()
        except Exception as e:
            logging.error(f"Erro ao enviar despedida no fechamento do console: {e}")
        return True
    return False

win32api.SetConsoleCtrlHandler(on_console_event, True)

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

# === Funções Excel ===
def gravar_excel(dados, arquivo, aba="Acompanhamento-Pedidos"):
    excel = None
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        linha_usada = None
        codigo_pedido = None
        status_atual = None

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            if valor_ax == "#0000":
                linha_usada = row

                data_str = dados.get("Data")
                if data_str:
                    try:
                        ano_atual = datetime.datetime.now().year
                        data_obj = datetime.datetime.strptime(f"{data_str}/{ano_atual}", "%d/%m/%Y")
                        ws.Range(f"E{row}").Value = data_obj
                    except Exception:
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
            return linha_usada, codigo_pedido, status_atual
        else:
            wb.Close(SaveChanges=False)
            return None, None, None
    except Exception as e:
        logging.error(f"Erro ao gravar pedido: {e}", exc_info=True)
        return None, None, None
    finally:
        if excel:
            excel.Quit()

def listar_pedidos_abertos(arquivo, aba="Acompanhamento-Pedidos"):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(arquivo)
    ws = wb.Sheets(aba)
    pedidos = []
    for row in range(2, ws.UsedRange.Rows.Count + 1):
        codigo = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
        status = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else ""
        cliente = str(ws.Range(f"G{row}").Value).strip() if ws.Range(f"G{row}").Value else ""
        if codigo and codigo != "#0000" and status:
            pedidos.append(f"{codigo} - {cliente} ({status})")
    wb.Close(SaveChanges=False)
    excel.Quit()
    return pedidos

def listar_colunas_vazias(codigo, arquivo, aba="Acompanhamento-Pedidos"):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(arquivo)
    ws = wb.Sheets(aba)
    colunas_vazias = {}
    mapa = {
        "E": "Data","F": "Horário","G": "Nome","H": "Quantidade","I": "UnidadeMedida","J": "Produto",
        "L": "SaborMassa","M": "Preparação Massa","N": "SaborRecheio","O": "Preparação Recheio",
        "P": "DescricaoAdicional","Q": "ValorAdicional","R": "Coleta Detalhes adicionais",
        "S": "Endereco","U": "Entrega","Y":        "Preco"
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

def atualizar_coluna(codigo, coluna_excel, valor, arquivo, aba="Acompanhamento-Pedidos"):
    excel = None
    try:
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
                return True, status_atual

        wb.Close(SaveChanges=False)
        return False, None
    except Exception as e:
        logging.error(f"Erro ao atualizar coluna: {e}", exc_info=True)
        return False, None
    finally:
        if excel:
            excel.Quit()

# === Função para cancelar pedido ===
def cancelar_pedido(codigo, arquivo, aba="Acompanhamento-Pedidos"):
    excel = None
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        mapa = {
            "E": "Data","F": "Horário","G": "Nome","H": "Quantidade","I": "UnidadeMedida","J": "Produto",
            "L": "SaborMassa","M": "Preparação Massa","N": "SaborRecheio","O": "Preparação Recheio",
            "P": "DescricaoAdicional","Q": "ValorAdicional","R": "Coleta Detalhes adicionais",
            "S": "Endereco","U": "Entrega","Y": "Preco"
        }

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            if valor_ax == codigo:
                for col in mapa.keys():
                    valor = ws.Range(f"{col}{row}").Value
                    if not valor:
                        ws.Range(f"{col}{row}").Value = "Cancelado"
                status_atual = "Cancelado"
                wb.Save()
                wb.Close(SaveChanges=True)
                return True, status_atual

        wb.Close(SaveChanges=False)
        return False, None
    except Exception as e:
        logging.error(f"Erro ao cancelar pedido: {e}", exc_info=True)
        return False, None
    finally:
        if excel:
            excel.Quit()

# === Handler de mensagens ===
async def mensagens(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip()
    logging.info(f"Mensagem recebida: {texto}")

    # Fluxo de pedidos em aberto (consulta rápida)
    if "me mostre os pedidos em aberto" in texto.lower() or "pedidos em aberto" in texto.lower():
        pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
        if pedidos_abertos:
            lista = "\n".join(pedidos_abertos)
            await update.message.reply_text(f"📋 Aqui estão os pedidos em aberto:\n{lista}")
            await update.message.reply_text(
                "🤖 Caso queira atualizar algum pedido, me envie \"Atualizar pedido\" ou \"Cancelar pedido\" que trabalharemos juntos novamente!"
            )
        else:
            await update.message.reply_text("⚠️ Nenhum pedido em aberto encontrado.")
        return

    # Fluxo de atualização
    if "atualizar pedido" in texto.lower():
        pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
        if pedidos_abertos:
            lista = "\n".join(pedidos_abertos)
            await update.message.reply_text(f"Qual pedido você quer atualizar?\nPedidos em aberto:\n{lista}")
            context.user_data["esperando_codigo"] = True
            context.user_data["tempo_inicio_codigo"] = datetime.datetime.now()
        else:
            await update.message.reply_text("⚠️ Nenhum pedido em aberto encontrado.")
        return

    # Fluxo de cancelamento
    if "cancelar pedido" in texto.lower():
        pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
        if pedidos_abertos:
            lista = "\n".join(pedidos_abertos)
            await update.message.reply_text(
                f"Qual pedido devemos cancelar?\nPedidos em aberto:\n{lista}"
            )
            context.user_data["esperando_cancelamento"] = True
            context.user_data["tempo_inicio_cancelamento"] = datetime.datetime.now()
        else:
            await update.message.reply_text("⚠️ Nenhum pedido em aberto encontrado.")
        return

    # Confirmação de cancelamento
    if context.user_data.get("confirmando_cancelamento"):
        codigo = context.user_data.get("codigo_cancelamento")
        resposta = texto.strip().lower()   # normaliza para minúsculo

        if resposta == "sim":
            ok, status_atual = cancelar_pedido(codigo, ARQUIVO, ABA)
            if ok:
                await update.message.reply_text(
                    f"Pedido {codigo} cancelado com sucesso! ❌ | Status atual do pedido: {status_atual}"
                )
            else:
                await update.message.reply_text(f"⚠️ Não foi possível cancelar o pedido {codigo}.")
        elif resposta == "não":
            await update.message.reply_text(
                "🤖 Entendido, estou desconsiderando o cancelamento deste pedido! 👍"
            )
        else:
            await update.message.reply_text("⚠️ Responda apenas com 'sim' ou 'não'.")
            return

        context.user_data["confirmando_cancelamento"] = False
        context.user_data["codigo_cancelamento"] = None
        return

    # Receber código para cancelamento
    if context.user_data.get("esperando_cancelamento"):
        inicio = context.user_data.get("tempo_inicio_cancelamento")
        if inicio and (datetime.datetime.now() - inicio).total_seconds() > 600:
            await update.message.reply_text(
                "⏳ Parece que demoraram para escolher o pedido! Se quiser realmente cancelar, envie novamente 'Cancelar pedido' 🤖"
            )
            context.user_data["esperando_cancelamento"] = False
            context.user_data["tempo_inicio_cancelamento"] = None
            return

        codigo = texto.strip()
        context.user_data["codigo_cancelamento"] = codigo
        await update.message.reply_text(
            f"⚠️ Você selecionou o pedido {codigo}. Tem certeza que deseja marcar como CANCELADO no sistema? (responda 'sim' ou 'não')"
        )
        context.user_data["confirmando_cancelamento"] = True
        context.user_data["esperando_cancelamento"] = False  # desativa o estado de espera
        return

    # Receber código do pedido para atualização
    if context.user_data.get("esperando_codigo"):
        inicio = context.user_data.get("tempo_inicio_codigo")
        if inicio and (datetime.datetime.now() - inicio).total_seconds() > 600:
            await update.message.reply_text(
                "⏳ Parece que demoraram para escolher o pedido! Se quiser realmente atualizar, envie novamente 'Atualizar pedido' 🤖"
            )
            context.user_data["esperando_codigo"] = False
            context.user_data["tempo_inicio_codigo"] = None
            return

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
            context.user_data["tempo_inicio_atualizacao"] = datetime.datetime.now()
        else:
            await update.message.reply_text(f"⚠️ O pedido {codigo} não possui colunas vazias.")
        context.user_data["esperando_codigo"] = False
        return


    # Receber atualização
    if context.user_data.get("esperando_atualizacao"):
        inicio = context.user_data.get("tempo_inicio_atualizacao")
        if inicio and (datetime.datetime.now() - inicio).total_seconds() > 600:
            await update.message.reply_text(
                "🤔 Acho que ficaram ocupados e não enviaram a atualização! Se quiser realmente atualizar o pedido, envie novamente 'Atualizar pedido' 🤖"
            )
            context.user_data["esperando_atualizacao"] = False
            context.user_data["tempo_inicio_atualizacao"] = None
            return

        if ":" in texto:
            chave, valor = texto.split(":", 1)
            chave = chave.strip()
            valor = valor.strip()

            codigo = context.user_data.get("codigo_pedido")
            colunas_vazias = context.user_data.get("colunas_vazias", {})

            if chave.isdigit():
                indice = int(chave) - 1
                nomes_colunas = list(colunas_vazias.keys())
                if 0 <= indice < len(nomes_colunas):
                    nome_coluna = nomes_colunas[indice]
                else:
                    await update.message.reply_text("⚠️ Número inválido. Escolha um dos números listados.")
                    return
            else:
                nome_coluna = chave

            if nome_coluna in colunas_vazias:
                coluna_excel = colunas_vazias[nome_coluna]
                ok, status_atual = atualizar_coluna(codigo, coluna_excel, valor, ARQUIVO, ABA)
                if ok:
                    await update.message.reply_text(
                        f"Pedido {codigo} atualizado com sucesso! ✅ | {nome_coluna}: {valor}"
                    )
                    await update.message.reply_text(f"Status atual do pedido {codigo}: {status_atual}")
                else:
                    await update.message.reply_text(f"⚠️ Não foi possível atualizar o pedido {codigo}.")
            else:
                await update.message.reply_text(f"⚠️ A coluna '{nome_coluna}' não está disponível para atualização.")
        else:
            await update.message.reply_text("⚠️ Use o formato: NomeColuna: Valor ou Número listado: Valor")

        context.user_data["esperando_atualizacao"] = False
        context.user_data["tempo_inicio_atualizacao"] = None
        return

    # Fluxo normal de adicionar pedido
    if "adicionar pedido ao sistema" in texto.lower():
        await update.message.reply_text("✅ Comando recebido! Me diga os detalhes do pedido!")
        context.user_data["esperando_pedido"] = True
        context.user_data["tempo_inicio_pedido"] = datetime.datetime.now()
    else:
        if context.user_data.get("esperando_pedido"):
            inicio = context.user_data.get("tempo_inicio_pedido")
            if inicio and (datetime.datetime.now() - inicio).total_seconds() > 600:
                await update.message.reply_text(
                    "🤔 Acredito que estejam ocupados no momento! Caso queira realmente adicionar um novo pedido, me envie 'Adicionar pedido ao sistema' novamente, que estarei pronto para te ajudar! 🤖"
                )
                context.user_data["esperando_pedido"] = False
                context.user_data["tempo_inicio_pedido"] = None
                return

            dados = interpretar_mensagem(texto)
            linha, codigo_pedido, status_atual = gravar_excel(dados, ARQUIVO, ABA)
            if linha:
                mensagem_final = f"Pedido {codigo_pedido} Adicionado ao sistema! ✅ | Status atual do pedido: {status_atual}"
                await update.message.reply_text(mensagem_final)
            else:
                await update.message.reply_text("⚠️ Não foi possível gravar o pedido.")

            context.user_data["esperando_pedido"] = False
            context.user_data["tempo_inicio_pedido"] = None


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


def enviar_despedida():
    mensagem = (
        "🤖 Hora da soneca! 😴 Vou buscar as últimas mensagens enviadas no grupo "
        "e continuarei o trabalho assim que o bot_continuo for colocado para rodar novamente na máquina principal."
    )
    try:
        requests.post(URL, data={'chat_id': CHAT_ID, 'text': mensagem})
        logging.info("Mensagem de despedida enviada.")
    except Exception as e:
        logging.error(f"Erro ao enviar mensagem de despedida: {e}")


def handle_exit(signum=None, frame=None):
    enviar_despedida()
    sys.exit(0)


# Registrar sinais de interrupção
signal.signal(signal.SIGINT, handle_exit)   # Ctrl+C
signal.signal(signal.SIGTERM, handle_exit)  # Encerramento normal

def main():
    app = ApplicationBuilder().token(TOKEN).post_init(avisar_inicio).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, mensagens))

    print("Bot rodando... Ctrl+C para parar.")
    try:
        app.run_polling()
    finally:
        enviar_despedida()

if __name__ == "__main__":
    main()

