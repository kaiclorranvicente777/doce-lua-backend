import telegram
import asyncio
import os
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
import difflib
import time
import functools
import unicodedata
import fitz  # PyMuPDF
from PIL import Image
from telegram import InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import CallbackQueryHandler

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

# === Funções utilitárias ===
async def enviar_mensagem_segura(bot, chat_id, texto, tentativas=3, intervalo=5):
    for tentativa in range(tentativas):
        try:
            await bot.send_message(chat_id=chat_id, text=texto)
            return True  # sucesso
        except Exception as e:
            logging.error(f"Erro ao enviar mensagem (tentativa {tentativa+1}): {e}")

            # 🚨 Alerta no Telegram sobre falha de rede em cada tentativa
            try:
                await bot.send_message(
                    chat_id=chat_id,
                    text=f"⚠️ Problema de rede detectado (tentativa {tentativa+1}). "
                         f"O bot vai tentar novamente em {intervalo}s..."
                )
            except Exception as alerta_erro:
                logging.error(f"Falha ao enviar alerta de rede: {alerta_erro}")

            await asyncio.sleep(intervalo)  # espera alguns segundos antes de tentar de novo

    # 🚨 Se todas as tentativas falharem, avisa no Telegram
    try:
        await bot.send_message(
            chat_id=chat_id,
            text="❌ Todas as tentativas de envio falharam. "
                 "O bot não conseguiu entregar a mensagem devido a problema de rede."
        )
    except Exception as alerta_final:
        logging.error(f"Falha ao enviar alerta final: {alerta_final}")

    return False  # falhou todas as tentativas


COMANDOS = [
    "enviar calendário",
    "pedidos em aberto",
    "atualizar pedido",
    "editar pedido",
    "cancelar pedido",
    "adicionar pedido ao sistema",
    "pagamentos pendentes"
]


def comando_mais_proximo(texto):
    texto = texto.lower().strip()
    match = difflib.get_close_matches(texto, COMANDOS, n=1, cutoff=0.5)
    return match[0] if match else None

def normalizar(texto):
    if not texto:
        return ""
    # Remove acentos
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    # Converte para minúsculo
    return texto.lower().strip()

def converter_data_excel(valor):
    """
    Converte um valor vindo do Excel em um objeto datetime.date.
    Aceita datetime, número serial do Excel ou string no formato dd/mm/yyyy.
    Retorna None se não conseguir converter.
    """
    if not valor:
        return None

    # Caso já seja datetime
    if isinstance(valor, datetime.datetime):
        return valor.date()

    # Caso seja número serial do Excel (float)
    if isinstance(valor, (int, float)):
        try:
            return datetime.datetime.fromordinal(
                datetime.datetime(1899, 12, 30).toordinal() + int(valor)
            ).date()
        except Exception:
            return None

    # Caso seja string
    if isinstance(valor, str):
        try:
            return datetime.datetime.strptime(valor.strip(), "%d/%m/%Y").date()
        except ValueError:
            return None

    return None

def limpar_valor_monetario(valor):
    """
    Remove 'R$', espaços e converte para número float.
    Exemplo: 'R$10,00' -> 10.00
    """
    if not valor:
        return None
    valor_str = str(valor).replace("R$", "").replace("r$", "").strip()
    valor_str = valor_str.replace(",", ".")
    try:
        return float(valor_str)
    except:
        return valor  # se não conseguir converter, retorna original


# === CONFIGURAÇÕES ===
TOKEN = "8570909503:AAF8h_fee0KOSFg95TT6pOr_dDvTqlLQavk"
CHAT_ID = -1003380097906
ARQUIVO = r"C:\Users\KICEGA~1\OneDrive\Documentos\Sistema.Central_Doce.Lua_11-2025.xlsx"
ABA = "Acompanhamento Doce Lua"
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

# === Verificação inicial do Excel ===
def verificar_excel_aberto():
    try:
        excel = win32.GetActiveObject("Excel.Application")
        if excel:
            # Desativa alertas de salvar
            excel.DisplayAlerts = False

            # Envia aviso no Telegram
            requests.post(
                f"https://api.telegram.org/bot{TOKEN}/sendMessage",
                data={
                    "chat_id": CHAT_ID,
                    "text": "Ops, encontrei o sistema Excel aberto na máquina principal, "
                            "mas estou fechando para que não dê nenhuma divergência no fluxo 😉"
                }
            )

            # Fecha todas as planilhas sem salvar
            for wb in list(excel.Workbooks):
                wb.Close(SaveChanges=False)

            # Fecha o Excel
            excel.Quit()
            time.sleep(2)
            logging.info("Excel aberto foi detectado e fechado com sucesso.")
        else:
            logging.info("Nenhuma instância do Excel estava ativa.")
    except Exception as e:
        logging.warning(f"Não foi possível capturar instância do Excel: {e}")


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

    preco_match = re.search(r"Total[:\s]*\(?(R\$[0-9.,]+)\)?", texto)
    if preco_match:
        dados["Preco"] = preco_match.group(1).strip()

    dados["Entrega"] = ""

    logging.debug(f"Dados extraídos: {dados}")
    return dados

# === Funções Excel ===
def gravar_excel(dados, arquivo, aba="Acompanhamento Doce Lua"):
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
                ws.Range(f"Q{row}").Value = limpar_valor_monetario(dados.get("ValorAdicional"))
                ws.Range(f"S{row}").Value = dados.get("Endereco")
                ws.Range(f"U{row}").Value = dados.get("Entrega")
                ws.Range(f"V{row}").Value = dados.get("Pagamento")
                ws.Range(f"W{row}").Value = dados.get("RecebimentoFotoComPedido")
                ws.Range(f"X{row}").Value = dados.get("PublicaçãoInstagram")
                ws.Range(f"Y{row}").Value = limpar_valor_monetario(dados.get("Preco"))

                codigo_pedido = str(ws.Range(f"AX{row}").Value).strip()
                status_atual = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else "Sem status"
                break
            
                # ✅ Ajuste para coluna Q (Valor Adicional)
                valor_adicional = dados.get("ValorAdicional")
                if valor_adicional:
                    valor_num = str(valor_adicional).replace("R$", "").replace(",", ".").strip()
                    try:
                        ws.Range(f"Q{row}").Value = float(valor_num)
                    except:
                        ws.Range(f"Q{row}").Value = valor_adicional
                else:
                    ws.Range(f"Q{row}").Value = None

                ws.Range(f"S{row}").Value = dados.get("Endereco")
                ws.Range(f"U{row}").Value = dados.get("Entrega")
                ws.Range(f"V{row}").Value = dados.get("Pagamento")
                ws.Range(f"W{row}").Value = dados.get("RecebimentoFotoComPedido")
                ws.Range(f"X{row}").Value = dados.get("PublicaçãoInstagram")

                # ✅ Ajuste para coluna Y (Preço)
                preco = dados.get("Preco")
                if preco:
                    preco_num = str(preco).replace("R$", "").replace(",", ".").strip()
                    try:
                        ws.Range(f"Y{row}").Value = float(preco_num)
                    except:
                        ws.Range(f"Y{row}").Value = preco
                else:
                    ws.Range(f"Y{row}").Value = None

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
# Corte 2

def listar_pedidos_abertos(arquivo, aba="Acompanhamento Doce Lua"):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(arquivo)
    ws = wb.Sheets(aba)
    pedidos = []
    for row in range(2, ws.UsedRange.Rows.Count + 1):
        codigo = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
        status = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else ""
        cliente = str(ws.Range(f"G{row}").Value).strip() if ws.Range(f"G{row}").Value else ""

        # Só adiciona se não estiver concluído
        if codigo and codigo != "#0000" and status and "concluído" not in status.lower():
            pedidos.append(f"{codigo} - {cliente} ({status})")

    wb.Close(SaveChanges=False)
    excel.Quit()
    return pedidos


def listar_colunas_vazias(codigo, arquivo, aba="Acompanhamento Doce Lua"):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(arquivo)
    ws = wb.Sheets(aba)
    colunas_vazias = {}
    mapa = {
        "E": "Data","F": "Horário","G": "Nome","H": "Quantidade","I": "Unidade Medida","J": "Produto",
            "L": "Sabor Massa","M": "Preparação Massa","N": "Sabor Recheio","O": "Preparação Recheio",
            "P": "Descrição Adicional","Q": "Valor Adicional","R": "Coleta Detalhes adicionais",
            "S": "Endereco","U": "Entrega",
            "V": "Pagamento","W": "Recebimento foto com pedido","X": "Publicação Instagram",
            "Y": "Preço"
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

def atualizar_coluna(codigo, coluna_excel, valor, arquivo, aba="Acompanhamento Doce Lua"):
    excel = None
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            if valor_ax == codigo:
                # ✅ Tratamento especial para colunas Q e Y
                if coluna_excel in ["Q", "Y"]:
                    ws.Range(f"{coluna_excel}{row}").Value = limpar_valor_monetario(valor)
                else:
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
def cancelar_pedido(codigo, arquivo, aba="Acompanhamento Doce Lua"):
    excel = None
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        mapa = {
            "E": "Data","F": "Horário","G": "Nome","H": "Quantidade","I": "Unidade Medida","J": "Produto",
            "L": "SaborMassa","M": "Preparação Massa","N": "Sabor Recheio","O": "Preparação Recheio",
            "P": "Descrição Adicional","Q": "Valor Adicional","R": "Coleta Detalhes adicionais",
            "S": "Endereço","U": "Entrega",
            "V": "Pagamento","W": "Recebimento foto com pedido","X": "Publicação Instagram",
            "Y": "Preço"
        }


        for row in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            if valor_ax == codigo:
                for col in mapa.keys():
                    valor = ws.Range(f"{col}{row}").Value
                    if not valor:
                        ws.Range(f"{col}{row}").Value = "Cancelado"

                ws.Range(f"D{row}").Value = "Cancelado"
                
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

IMAGEM_PATH = r"C:\Users\KICEGA~1\OneDrive\Documentos\excel_intervalo.png"

def gerar_imagem_excel():
    global excel, wb
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # Fecha qualquer workbook extra
        for wb_extra in list(excel.Workbooks):
            wb_extra.Close(SaveChanges=False)

        wb = excel.Workbooks.Open(ARQUIVO)
        ws = wb.Sheets("Acompanhamento-Geral")
        ws.Activate()

        # ✅ Garante que a área de impressão está correta e cabe em uma página
        ws.PageSetup.PrintArea = "I1:AP27"
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1

        pdf_path = r"C:\Users\KICEGA~1\OneDrive\Documentos\calendario.pdf"

        # Exporta só o intervalo definido
        ws.ExportAsFixedFormat(Type=0, Filename=pdf_path)

        # Converte PDF para PNG com PyMuPDF
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)
        zoom = 200 / 72  # escala baseada no DPI desejado
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        pix.save(IMAGEM_PATH)
        doc.close()

        logging.info("Imagem do calendário gerada com sucesso (PDF ajustado).")

    except Exception as e:
        logging.error(f"Erro ao gerar imagem do calendário: {e}", exc_info=True)
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
            time.sleep(2)
            del wb
            del excel
            logging.info("Excel fechado completamente após execução.")
        except Exception as e:
            logging.error(f"Erro ao fechar Excel: {e}", exc_info=True)



# corte 3
         
def enviar_imagem_calendario():
    try:
        if not os.path.exists(IMAGEM_PATH):
            logging.error(f"Arquivo de imagem não encontrado: {IMAGEM_PATH}")
            return False  # retorna falha

        with open(IMAGEM_PATH, "rb") as img:
            files = {"photo": img}
            data = {"chat_id": CHAT_ID, "caption": "Veja seu Calendário Programação de Pedidos Doce Lua! 🤖🗓️"}
            response = requests.post(
                f"https://api.telegram.org/bot{TOKEN}/sendPhoto",
                files=files, data=data
            )
            logging.info(f"Status envio: {response.status_code} | Resposta: {response.text}")

        if response.status_code == 200:
            logging.info("Imagem enviada com sucesso.")
            os.remove(IMAGEM_PATH)  # só remove se deu certo
            return True
        else:
            logging.error("Falha ao enviar imagem.")
            return False
    except Exception as e:
        logging.error(f"Erro ao enviar imagem: {e}", exc_info=True)
        return False
    
async def verificar_usuario_fluxo(update, context):
    """
    Verifica se quem está respondendo é diferente do usuário que iniciou o fluxo.
    Se for diferente, envia uma saudação e segue com o fluxo normalmente.
    """
    usuario_atual = update.effective_user.first_name
    usuario_fluxo = context.user_data.get("usuario_fluxo")

    if usuario_fluxo and usuario_atual != usuario_fluxo:
        await enviar_mensagem_segura(context.bot, update.effective_chat.id, 
            f"Olá {usuario_atual}, vamos seguir com a sua resposta da solicitação anterior de {usuario_fluxo} 😉"
        )


# === Handler de mensagens ===

async def enviar_menu_comandos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mensagem = (
        "🤖 Comandos disponíveis para interagir comigo:\n"
        "- \"Enviar calendário\" → Veja seu Calendário Programação de Pedidos Doce Lua! 🤖🗓️\n"
        "- \"Pedidos em aberto\" → Pedidos que ainda não foram concluídos 📋\n"
        "- \"Atualizar pedido\" → Preencher colunas vazias de um pedido ✏️\n"
        "- \"Editar pedido\" → Alterar colunas já preenchidas de um pedido 📝\n"
        "- \"Cancelar pedido\" → Inicia o processo para cancelamento ❌\n"
        "- \"Adicionar pedido ao sistema\" → Fluxo para registrar um novo pedido ✅\n"
        "- \"Pagamentos pendentes\" → Textos prontos para cobrar clientes com pagamento em aberto ⚠️\n\n"
        "Use qualquer um desses comandos a qualquer momento e eu cuidarei do restante! 🐱"
    )
    await enviar_mensagem_segura(context.bot, update.effective_chat.id, mensagem)


pedidos_enviados = []  # memória simples para comparar pedidos anteriores


async def fluxo_inicial(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # 1. Pedidos em aberto
    pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
    if pedidos_abertos:
        lista = "\n".join(pedidos_abertos)
        await enviar_mensagem_segura(context.bot, update.effective_chat.id, f"📋 Pedidos em aberto:\n{lista}")
    else:
        await enviar_mensagem_segura(context.bot, update.effective_chat.id, "Nenhum pedido em aberto ✅")

    # 2. Verificar se houve atualização
    global pedidos_enviados
    if pedidos_abertos != pedidos_enviados:
        pedidos_enviados = pedidos_abertos
        await enviar_mensagem_segura(context.bot, update.effective_chat.id, "🔄 Houve atualização nos pedidos.")
    else:
        await enviar_mensagem_segura(context.bot, update.effective_chat.id, "Nenhuma atualização desde a última verificação.")

    # 3. Pagamentos pendentes
    pendentes = []
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(ARQUIVO)
    ws = wb.Sheets(ABA)
    for row in range(2, ws.UsedRange.Rows.Count + 1):
        codigo = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
        status = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else ""
        cliente = str(ws.Range(f"G{row}").Value).strip() if ws.Range(f"G{row}").Value else ""
        if "pagamento" in status.lower():
            pendentes.append((codigo, cliente))
    wb.Close(SaveChanges=False)
    excel.Quit()

    if pendentes:
        aviso = (
            "⚠️ Identifiquei ausencia de pagamentos dos pedidos! Não esqueça de cobrar as pendências!\n"
            "Posso te encaminhar o texto pronto para enviar aos nossos clientes, "
            "basta solicitar com o comando \"Pagamentos pendentes\""
        )
        await enviar_mensagem_segura(context.bot, update.effective_chat.id, aviso)


    # 4. Menu de comandos
    await enviar_menu_comandos(update, context)

async def mensagens(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip()

    # 1️⃣ Primeiro: se estamos aguardando atualização de coluna
    if context.user_data.get("esperando_coluna"):
        codigo_pedido = context.user_data.get("codigo_pedido")
        colunas_vazias = listar_colunas_vazias(codigo_pedido, ARQUIVO, ABA)

        # cria mapeamento numérico (1, 2, 3...) → nome da coluna
        mapa_numerico = {str(i+1): nome for i, nome in enumerate(colunas_vazias.keys())}

        if ":" in texto:
            chave, valor = texto.split(":", 1)
            chave = chave.strip()
            valor = valor.strip()

            # Se o usuário digitou número, converte para nome da coluna
            if chave in mapa_numerico:
                nome_coluna = mapa_numerico[chave]
            else:
                nome_coluna = chave  # assume que já é o nome da coluna

            coluna_excel = colunas_vazias.get(nome_coluna)
            if coluna_excel:
                sucesso, status = atualizar_coluna(codigo_pedido, coluna_excel, valor, ARQUIVO, ABA)
                if sucesso:
                    await enviar_mensagem_segura(
                        context.bot,
                        update.effective_chat.id,
                        f"✅ Coluna '{nome_coluna}' atualizada com: {valor}"
                    )
                    await enviar_mensagem_segura(
                        context.bot,
                        update.effective_chat.id,
                        f"📌 Status atual do pedido {codigo_pedido}: {status}"
                    )

                    # Relista colunas restantes
                    colunas_restantes = listar_colunas_vazias(codigo_pedido, ARQUIVO, ABA)
                    if colunas_restantes:
                        lista_formatada = "\n".join([f"{i+1}. {nome}" for i, nome in enumerate(colunas_restantes.keys())])
                        await enviar_mensagem_segura(
                            context.bot,
                            update.effective_chat.id,
                            f"Ainda faltam:\n{lista_formatada}\n\nResponda no formato: Número: Valor ou NomeColuna: Valor"
                        )
                    else:
                        await enviar_mensagem_segura(
                            context.bot,
                            update.effective_chat.id,
                            f"🎉 Todas as colunas do pedido {codigo_pedido} foram preenchidas!"
                        )
                        context.user_data["esperando_coluna"] = False
                else:
                    await enviar_mensagem_segura(
                        context.bot,
                        update.effective_chat.id,
                        f"⚠️ Não consegui atualizar a coluna '{nome_coluna}'."
                    )
            else:
                await enviar_mensagem_segura(
                    context.bot,
                    update.effective_chat.id,
                    "⚠️ Coluna inválida. Use o número da lista ou o nome correto."
                )
        return  # 🔴 Sai da função aqui, não continua para checar comandos

    # 2️⃣ Se não estamos em fluxo de atualização, segue para comandos
    texto_lower = texto.lower()

    # Se houver comando confirmado
    if context.user_data.get("comando_confirmado"):
        comando = context.user_data["comando_confirmado"]

    else:
        # Se aguardando confirmação
        if context.user_data.get("esperando_confirmacao_comando"):
            resposta = texto.lower().strip()
            comando_sugerido = context.user_data["esperando_confirmacao_comando"].lower().strip()

            if resposta in ["sim", "s", "yes", "y"] or resposta == comando_sugerido:
                comando = context.user_data["esperando_confirmacao_comando"]
                context.user_data["comando_confirmado"] = comando
                context.user_data["esperando_confirmacao_comando"] = None

                # Reexecuta já com comando confirmado
                await mensagens(update, context)
                return

            elif resposta in ["não", "nao", "n", "no"]:
                await enviar_mensagem_segura(
                    context.bot,
                    update.effective_chat.id,
                    "Ok, comando cancelado. Estarei pronto para trabalhar com você assim que me enviar algum dos comandos abaixo:"
                )
                await enviar_menu_comandos(update, context)
                context.user_data["esperando_confirmacao_comando"] = None
                return

            else:
                await enviar_mensagem_segura(context.bot, update.effective_chat.id, "⚠️ Responda apenas com 'sim' ou 'não'.")
                return

        # Se não há confirmação, tenta sugerir comando
        comando = comando_mais_proximo(texto)
        if comando and texto not in COMANDOS:
            context.user_data["esperando_confirmacao_comando"] = comando
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                f"🤔 Você quis dizer \"{comando}\"? Responda SIM ou NÃO."
            )
            return
        else:
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                "❌ Não entendi esse comando 🤔. Tente novamente ou veja a lista abaixo 👇"
            )
            await enviar_menu_comandos(update, context)
            return


    # === Fluxo: Enviar calendário ===
    if comando == "enviar calendário":
        try:
            gerar_imagem_excel()
            if enviar_imagem_calendario():
                await enviar_mensagem_segura(context.bot, update.effective_chat.id, "📤 Calendário enviado com sucesso!")
            else:
                await enviar_mensagem_segura(context.bot, update.effective_chat.id, "⚠️ Ocorreu um erro ao enviar o calendário.")
        except Exception as e:
            logging.error(f"Erro ao gerar/enviar calendário: {e}", exc_info=True)
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, "⚠️ Ocorreu um erro ao gerar/enviar o calendário.")
        await enviar_menu_comandos(update, context)
        context.user_data["comando_confirmado"] = None
        return

    # === Fluxo: Pedidos em aberto ===
    if comando == "pedidos em aberto":
        pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
        if pedidos_abertos:
            lista = "\n".join(pedidos_abertos)
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, f"📋 Aqui estão os pedidos em aberto:\n{lista}")
            await enviar_mensagem_segura(context.bot, update.effective_chat.id,
                "🤖 Caso queira atualizar algum pedido, me envie \"Atualizar pedido\" ou \"Cancelar pedido\" que trabalharemos juntos novamente!"
            )
        else:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, "⚠️ Nenhum pedido em aberto encontrado.")
        await enviar_menu_comandos(update, context)
        context.user_data["comando_confirmado"] = None
        return

# corte 4

    # === Fluxo: Pagamentos pendentes ===
    if comando == "pagamentos pendentes":
        pendentes = []
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(ARQUIVO)
        ws = wb.Sheets(ABA)

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            codigo = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            status = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else ""
            cliente = str(ws.Range(f"G{row}").Value).strip() if ws.Range(f"G{row}").Value else ""

            # captura a data da coluna E
            data_entrega = converter_data_excel(ws.Range(f"E{row}").Value)

            if status.lower() == "aguardando - pagamento":
                pendentes.append((codigo, cliente, data_entrega))

        wb.Close(SaveChanges=False)
        excel.Quit()

        if pendentes:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, "⚠️ Não esqueça de cobrar os pagamentos pendentes!")
            hoje = datetime.datetime.now().date()

            for codigo, cliente, data_entrega in pendentes:
                if data_entrega and hoje <= data_entrega:
                    # Pedido ainda não chegou → cobrar adiantamento
                    msg = (
                        f"Olá, {cliente} tudo bem? Sou o bot da Confeitaria Doce Lua 😊🤖\n"
                        f"Seu pedido {codigo} está confirmado e marcado como *Aguardando - Pagamento*.\n\n"
                        "Pedimos que nos envie o comprovante de pagamento de pelo menos 50% do seu pedido.\n\n"
                        "Caso já tenha realizado, por favor desconsidere esta mensagem.\n"
                        "Estamos à disposição para qualquer dúvida 💜"
                    )
                else:
                    # Pedido já entregue → cobrar pagamento total
                    msg = (
                        f"Olá, {cliente} tudo bem? Sou o bot da Confeitaria Doce Lua 😊🤖\n"
                        f"Verificamos que seu pedido {codigo} já foi entregue 🍰✨,\n"
                        "mas ainda consta como *Aguardando - Pagamento* em nosso sistema.\n\n"
                        "Pedimos a gentileza de nos enviar o comprovante de pagamento para concluirmos o processo.\n"
                        "Caso já tenha realizado, por favor desconsidere esta mensagem.\n\n"
                        "Agradecemos pela confiança 💜"
                    )

                await enviar_mensagem_segura(context.bot, update.effective_chat.id, msg)
        else:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, "✅ Nenhum pagamento pendente encontrado.")

        await enviar_menu_comandos(update, context)
        context.user_data["comando_confirmado"] = None
        return



    # === Fluxo: Editar pedido ===
    if comando == "editar pedido":
        pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
        if pedidos_abertos:
            lista = "\n".join(pedidos_abertos)
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, f"Qual pedido você quer editar?\nPedidos em aberto:\n{lista}")
            context.user_data["usuario_fluxo"] = update.effective_user.first_name
            context.user_data["esperando_codigo_edicao"] = True
            context.user_data["tempo_inicio_codigo_edicao"] = datetime.datetime.now()
        else:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id,"⚠️ Nenhum pedido em aberto encontrado.")
            await enviar_menu_comandos(update, context)
        context.user_data["comando_confirmado"] = None
        return

    # === Fluxo: Cancelar pedido ===
    if comando == "cancelar pedido":
        pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
        if pedidos_abertos:
            lista = "\n".join(pedidos_abertos)
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, f"Qual pedido devemos cancelar?\nPedidos em aberto:\n{lista}")
            context.user_data["usuario_fluxo"] = update.effective_user.first_name
            context.user_data["esperando_cancelamento"] = True
            context.user_data["tempo_inicio_cancelamento"] = datetime.datetime.now()
        else:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id,"⚠️ Nenhum pedido em aberto encontrado.")
            await enviar_menu_comandos(update, context)
        context.user_data["comando_confirmado"] = None
        return

    # === Fluxo: Atualizar pedido ===
    if comando == "atualizar pedido":
        pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
        if pedidos_abertos:
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                "📋 Qual pedido você quer atualizar?\n\nCódigo pedido Individual - (Status atual)"
            )

            # Para cada pedido, envia mensagem + botão
            for pedido in pedidos_abertos:
                # ignora cabeçalho/legenda
                if "codigo pedido" in pedido.lower() or "status atual" in pedido.lower():
                    continue

                # separa código, cliente e status
                partes = pedido.split(" - ")
                codigo = partes[0].strip()
                resto = " - ".join(partes[1:])  # cliente + status

                # monta botão
                keyboard = [[InlineKeyboardButton(text=f"Selecionar {codigo}", callback_data=codigo)]]
                reply_markup = InlineKeyboardMarkup(keyboard)

                # envia texto + botão juntos
                await context.bot.send_message(
                    chat_id=update.effective_chat.id,
                    text=f"{codigo} - {resto}",
                    reply_markup=reply_markup
                )


        else:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, "Nenhum pedido em aberto ✅")

        context.user_data["comando_confirmado"] = None
        return

    # === Callback: Selecionar pedido ===
    async def selecionar_pedido(update: Update, context: ContextTypes.DEFAULT_TYPE):
        query = update.callback_query
        await query.answer()
        codigo_pedido = query.data

        # guarda no contexto
        context.user_data["codigo_pedido"] = codigo_pedido
        context.user_data["esperando_coluna"] = True

        # lista colunas vazias
        colunas_vazias = listar_colunas_vazias(codigo_pedido, ARQUIVO, ABA)
        context.user_data["colunas_vazias"] = colunas_vazias

        mapa_numerico = {str(i+1): nome for i, nome in enumerate(colunas_vazias.keys())}
        lista_formatada = "\n".join([f"{i+1}. {nome}" for i, nome in enumerate(colunas_vazias.keys())])

        await enviar_mensagem_segura(
            context.bot,
            update.effective_chat.id,
            f"Você selecionou o pedido {codigo_pedido}.\n\n"
            f"As colunas vazias são:\n{lista_formatada}\n\n"
            "Responda no formato: Número: Valor ou NomeColuna: Valor"
        )

# corte 5

    # === Receber código para cancelamento ===
    if context.user_data.get("esperando_cancelamento"):
        codigo = texto.strip()

        # ✅ Valida se o código existe em AX antes de pedir confirmação
        colunas_vazias = listar_colunas_vazias(codigo, ARQUIVO, ABA)
        if not colunas_vazias:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"⚠️ O pedido {codigo} não foi encontrado."
            )
            context.user_data["esperando_cancelamento"] = False
            return

        if not context.user_data.get("confirmando_cancelamento"):
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"Você confirma o cancelamento do pedido {codigo}? Responda SIM ou NÃO."
            )
            context.user_data["codigo_cancelamento"] = codigo
            context.user_data["confirmando_cancelamento"] = True
            return

        resposta = texto.lower().strip()
        if resposta in ["sim", "s", "yes", "y"]:
            codigo = context.user_data.get("codigo_cancelamento", codigo)
            ok, status_atual = cancelar_pedido(codigo, ARQUIVO, ABA)
            if ok:
                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    f"❌ Pedido {codigo} cancelado com sucesso! | Status: {status_atual}"
                )
            else:
                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    f"⚠️ Não foi possível cancelar o pedido {codigo}."
                )
        elif resposta in ["não", "nao", "n", "no"]:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "Cancelamento abortado. Se precisar, envie novamente \"Cancelar pedido\"."
            )
        else:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "⚠️ Responda apenas com 'sim' ou 'não' para confirmar o cancelamento."
            )
            return

        # Limpa estados
        context.user_data["esperando_cancelamento"] = False
        context.user_data["confirmando_cancelamento"] = False
        context.user_data["codigo_cancelamento"] = None

        await enviar_menu_comandos(update, context)
        return


    # === Receber código para edição ===
    if context.user_data.get("esperando_codigo_edicao"):
        codigo = texto.strip()

        # ✅ Valida se o código existe
        colunas_vazias = listar_colunas_vazias(codigo, ARQUIVO, ABA)
        if not colunas_vazias:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"⚠️ O pedido {codigo} não existe ou já foi concluído."
            )
            context.user_data["esperando_codigo_edicao"] = False
            return

        context.user_data["codigo_edicao"] = codigo
        mapa_edicao = {
            "E": "Data",
            "F": "Horário",
            "G": "Nome",
            "H": "Quantidade",
            "I": "Unidade Medida",
            "J": "Produto",
            "L": "Sabor Massa",
            "M": "Preparação Massa",
            "N": "Sabor Recheio",
            "O": "Preparação Recheio",
            "P": "Descrição Adicional",
            "Q": "Valor Adicional",
            "R": "Coleta Detalhes adicionais",
            "S": "Endereco",
            "U": "Entrega",
            "V": "Pagamento",
            "W": "Recebimento foto com pedido",
            "X": "Publicação Instagram",
            "Y": "Preço"
        }
        context.user_data["mapa_edicao"] = mapa_edicao

        nomes = list(mapa_edicao.values())
        lista = "\n".join([f"{i+1}. {nome}" for i, nome in enumerate(nomes)])

        await enviar_mensagem_segura(
            context.bot, update.effective_chat.id,
            f"Quais colunas deseja editar no pedido {codigo}?\n"
            f"Você pode responder por número ou nome no formato: NomeColuna: Valor\n\n{lista}"
        )

        context.user_data["esperando_coluna_edicao"] = True
        context.user_data["esperando_codigo_edicao"] = False
        return


    # === Receber edição múltipla (pré-visualização) ===
    if context.user_data.get("esperando_coluna_edicao"):
        codigo = context.user_data.get("codigo_edicao")
        mapa_edicao = context.user_data.get("mapa_edicao", {})

        if texto.lower().strip() == "fim":
            # Apenas sai do modo de edição de colunas e vai para confirmação
            context.user_data["esperando_coluna_edicao"] = False
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "Digite 'sim' para salvar as alterações ou 'não' para cancelar."
            )
            context.user_data["aguardando_confirmacao_edicao"] = True
            return

        # 👉 Pré-processa múltiplas linhas sem gravar ainda
        entradas = texto.splitlines()
        pendentes = []
        for entrada in entradas:
            if ":" not in entrada:
                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "⚠️ Formato inválido. Use NomeColuna: Valor ou número da lista."
                )
                continue

            chave, valor = entrada.split(":", 1)
            chave = chave.strip()
            valor = valor.strip()

            # Se for número
            if chave.isdigit():
                indice = int(chave) - 1
                nomes = list(mapa_edicao.values())
                if 0 <= indice < len(nomes):
                    nome_coluna = nomes[indice]
                else:
                    await enviar_mensagem_segura(
                        context.bot, update.effective_chat.id,
                        "⚠️ Número inválido. Escolha um número da lista."
                    )
                    continue
            else:
                nome_coluna = chave

            pendentes.append((nome_coluna, valor))

        if not pendentes:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "⚠️ Nenhuma coluna válida foi informada."
            )
            return

        # 👉 Salva pendentes para confirmar depois
        context.user_data["pendentes_edicao"] = pendentes

        # 👉 Mostra prévia ao usuário
        preview = "\n".join([f"{col}: {val}" for col, val in pendentes])
        await enviar_mensagem_segura(
            context.bot, update.effective_chat.id,
            f"Você solicitou atualizar o pedido {codigo} com os seguintes valores:\n\n{preview}\n\n"
            f"Digite 'sim' para confirmar ou 'não' para cancelar."
        )
        context.user_data["aguardando_confirmacao_edicao"] = True
        context.user_data["esperando_coluna_edicao"] = False
        return

# corte 6

    # === Receber código do pedido ===
    if context.user_data.get("esperando_codigo"):
        codigo = texto.strip()
        context.user_data["codigo_pedido"] = codigo
        colunas_vazias = listar_colunas_vazias(codigo, ARQUIVO, ABA)
        if colunas_vazias:
            lista = "\n".join([f"{i+1}. {nome}" for i, nome in enumerate(colunas_vazias.keys())])
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                f"As colunas vazias do pedido {codigo} são:\n{lista}\n\nResponda no formato: NomeColuna: Valor"
            )
            context.user_data["colunas_vazias"] = colunas_vazias
            context.user_data["esperando_atualizacao"] = True
        else:
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                f"⚠️ O pedido {codigo} não possui colunas vazias."
            )
        context.user_data["esperando_codigo"] = False
        return

    # === Receber atualização múltipla (pré-visualização) ===
    if context.user_data.get("esperando_atualizacao"):
        codigo = context.user_data.get("codigo_pedido")
        colunas_vazias = context.user_data.get("colunas_vazias", {})

        if texto.lower().strip() == "fim":
            context.user_data["esperando_atualizacao"] = False
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "Digite 'sim' para salvar as alterações ou 'não' para cancelar."
            )
            context.user_data["aguardando_confirmacao_atualizacao"] = True
            return

        entradas = texto.splitlines()
        pendentes = []
        for entrada in entradas:
            if ":" not in entrada:
                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "⚠️ Formato inválido. Use NomeColuna: Valor ou número da lista."
                )
                continue

            chave, valor = entrada.split(":", 1)
            chave = chave.strip()
            valor = valor.strip()

            # Se for número
            if chave.isdigit():
                indice = int(chave) - 1
                nomes = list(colunas_vazias.keys())
                if 0 <= indice < len(nomes):
                    nome_coluna = nomes[indice]
                else:
                    await enviar_mensagem_segura(
                        context.bot, update.effective_chat.id,
                        "⚠️ Número inválido. Escolha um número da lista."
                    )
                    continue
            else:
                nome_coluna = chave

            pendentes.append((nome_coluna, valor))

        if not pendentes:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "⚠️ Nenhuma coluna válida foi informada."
            )
            return

        context.user_data["pendentes_atualizacao"] = pendentes

        preview = "\n".join([f"{col}: {val}" for col, val in pendentes])
        await enviar_mensagem_segura(
            context.bot, update.effective_chat.id,
            f"Você solicitou atualizar o pedido {codigo} com os seguintes valores:\n\n{preview}\n\n"
            f"Digite 'sim' para confirmar ou 'não' para cancelar."
        )
        context.user_data["aguardando_confirmacao_atualizacao"] = True
        context.user_data["esperando_atualizacao"] = False
        return

    # === Confirmação da atualização múltipla ===
    if context.user_data.get("aguardando_confirmacao_atualizacao"):
        resposta = texto.lower().strip()
        if resposta in ["sim", "s", "yes", "y"]:
            codigo = context.user_data.get("codigo_pedido")
            pendentes = context.user_data.get("pendentes_atualizacao", [])
            alteracoes = []

            for nome_coluna, valor in pendentes:
                coluna_excel = None
                for nome, col_letra in context.user_data.get("colunas_vazias", {}).items():
                    if normalizar(nome_coluna) == normalizar(nome):
                        coluna_excel = col_letra
                        break

                if coluna_excel:
                    ok, status_atual = atualizar_coluna(codigo, coluna_excel, valor.strip(), ARQUIVO, ABA)
                    if ok:
                        alteracoes.append(f"{nome_coluna}: {valor}")
                    else:
                        alteracoes.append(f"{nome_coluna}: ⚠️ erro ao atualizar")
                else:
                    alteracoes.append(f"{nome_coluna}: ⚠️ coluna não disponível")

            # 👉 Envia apenas uma mensagem consolidada
            mensagem_final = (
                f"📝 Pedido {codigo} atualizado!\n"
                + "\n".join(alteracoes)
                + f"\n\n✅ Todas as alterações foram aplicadas ao pedido {codigo}."
            )
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, mensagem_final)

            # 👉 Buscar o status atual na coluna D
            excel = win32.gencache.EnsureDispatch("Excel.Application")
            wb = excel.Workbooks.Open(ARQUIVO)
            ws = wb.Sheets(ABA)

            status_atual = "Sem status"
            for r in range(2, ws.UsedRange.Rows.Count + 1):
                valor_ax = str(ws.Range(f"AX{r}").Value).strip() if ws.Range(f"AX{r}").Value else ""
                if valor_ax == codigo:
                    status_atual = str(ws.Range(f"D{r}").Value).strip() if ws.Range(f"D{r}").Value else "Sem status"
                    break

            wb.Close(SaveChanges=False)
            excel.Quit()

            # 👉 Enviar status atual ao usuário
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                f"📌 Status Atual: {status_atual}"
            )

        elif resposta in ["não", "nao", "n", "no"]:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "❌ Alterações canceladas."
            )
        else:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "⚠️ Responda apenas com 'sim' ou 'não'."
            )
            return

        # Limpa estado
        context.user_data["aguardando_confirmacao_atualizacao"] = False
        context.user_data["pendentes_atualizacao"] = None
        context.user_data["esperando_atualizacao"] = False
        return


    # === Receber código para edição ===
    if context.user_data.get("esperando_codigo_edicao"):
        codigo = texto.strip()

        # ✅ Valida se o código existe
        colunas_vazias = listar_colunas_vazias(codigo, ARQUIVO, ABA)
        if not colunas_vazias:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"⚠️ O pedido {codigo} não existe ou já foi concluído."
            )
            context.user_data["esperando_codigo_edicao"] = False
            return

        context.user_data["codigo_edicao"] = codigo
        mapa_edicao = {
            "E": "Data",
            "F": "Horário",
            "G": "Nome",
            "H": "Quantidade",
            "I": "Unidade Medida",
            "J": "Produto",
            "L": "Sabor Massa",
            "M": "Preparação Massa",
            "N": "Sabor Recheio",
            "O": "Preparação Recheio",
            "P": "Descrição Adicional",
            "Q": "Valor Adicional",
            "R": "Coleta Detalhes adicionais",
            "S": "Endereco",
            "U": "Entrega",
            "V": "Pagamento",
            "W": "Recebimento foto com pedido",
            "X": "Publicação Instagram",
            "Y": "Preço"
        }
        context.user_data["mapa_edicao"] = mapa_edicao

        nomes = list(mapa_edicao.values())
        lista = "\n".join([f"{i+1}. {nome}" for i, nome in enumerate(nomes)])

        await enviar_mensagem_segura(
            context.bot, update.effective_chat.id,
            f"Quais colunas deseja editar no pedido {codigo}?\n"
            f"Você pode responder por número ou nome no formato: NomeColuna: Valor\n\n{lista}"
        )

        context.user_data["esperando_coluna_edicao"] = True
        context.user_data["esperando_codigo_edicao"] = False   # 🔧 só desativa depois de receber o código
        return

    # === Receber edição múltipla (pré-visualização) ===
    if context.user_data.get("esperando_coluna_edicao"):
        codigo = context.user_data.get("codigo_edicao")
        mapa_edicao = context.user_data.get("mapa_edicao", {})

        if texto.lower().strip() == "fim":
            context.user_data["esperando_coluna_edicao"] = False
            context.user_data["codigo_edicao"] = None
            context.user_data["mapa_edicao"] = None
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "Digite 'sim' para salvar as alterações ou 'não' para cancelar."
            )
            context.user_data["aguardando_confirmacao_edicao"] = True
            return

        # 👉 Pré-processa múltiplas linhas sem gravar ainda
        entradas = texto.splitlines()
        pendentes = []
        for entrada in entradas:
            if ":" not in entrada:
                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "⚠️ Formato inválido. Use NomeColuna: Valor ou número da lista."
                )
                continue

            chave, valor = entrada.split(":", 1)
            chave = chave.strip()
            valor = valor.strip()

            # Se for número
            if chave.isdigit():
                indice = int(chave) - 1
                nomes = list(mapa_edicao.values())
                if 0 <= indice < len(nomes):
                    nome_coluna = nomes[indice]
                else:
                    await enviar_mensagem_segura(
                        context.bot, update.effective_chat.id,
                        "⚠️ Número inválido. Escolha um número da lista."
                    )
                    continue
            else:
                nome_coluna = chave

            pendentes.append((nome_coluna, valor))

        if not pendentes:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "⚠️ Nenhuma coluna válida foi informada."
            )
            return

        # 👉 Salva pendentes para confirmar depois
        context.user_data["pendentes_edicao"] = pendentes

        # 👉 Mostra prévia ao usuário
        preview = "\n".join([f"{col}: {val}" for col, val in pendentes])
        await enviar_mensagem_segura(
            context.bot, update.effective_chat.id,
            f"Você solicitou atualizar o pedido {codigo} com os seguintes valores:\n\n{preview}\n\n"
            f"Digite 'sim' para confirmar ou 'não' para cancelar."
        )
        context.user_data["aguardando_confirmacao_edicao"] = True
        context.user_data["esperando_coluna_edicao"] = False   # 🔧 DESATIVA edição aqui
        

    # === Confirmação da edição múltipla ===
    if context.user_data.get("aguardando_confirmacao_edicao"):
        resposta = texto.lower().strip()
        if resposta in ["sim", "s", "yes", "y"]:
            codigo = context.user_data.get("codigo_edicao")
            pendentes = context.user_data.get("pendentes_edicao", [])
            for nome_coluna, valor in pendentes:
                coluna_excel = None
                for col_letra, nome in context.user_data.get("mapa_edicao", {}).items():
                    if normalizar(nome_coluna) == normalizar(nome):
                        coluna_excel = col_letra
                        break
                if coluna_excel:
                    # 👉 grava exatamente como o usuário digitou
                    atualizar_coluna(codigo, coluna_excel, valor, ARQUIVO, ABA)
                    await enviar_mensagem_segura(
                        context.bot, update.effective_chat.id,
                        f"📝 Pedido {codigo} atualizado! | {nome_coluna}: {valor}"
                    )
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"✅ Todas as alterações foram aplicadas ao pedido {codigo}."
            )

            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"✅ Todas as alterações foram aplicadas ao pedido {codigo}."
            )
        elif resposta in ["não", "nao", "n", "no"]:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "❌ Alterações canceladas."
            )
        else:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "⚠️ Responda apenas com 'sim' ou 'não' para confirmar a edição."
            )
            return

        # Limpa estado somente após confirmar/cancelar
        context.user_data["aguardando_confirmacao_edicao"] = False
        context.user_data["pendentes_edicao"] = None
        context.user_data["codigo_edicao"] = None
        context.user_data["mapa_edicao"] = None
        await enviar_menu_comandos(update, context)
        return


    # === Fallback para comandos não reconhecidos ===
    await verificar_usuario_fluxo(update, context)
    comando_sugerido = comando_mais_proximo(texto)

    if context.user_data.get("esperando_confirmacao_comando"):
        resposta = texto.lower().strip()
        if resposta in ["sim", "s", "yes", "y"]:
            comando = context.user_data["esperando_confirmacao_comando"]
            context.user_data["comando_confirmado"] = comando
            context.user_data["esperando_confirmacao_comando"] = None

            # 👉 reexecuta já com comando confirmado
            await mensagens(update, context)
            return

        elif resposta in ["não", "nao", "n", "no"]:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, 
                "Ah, tudo bem! Estarei pronto para trabalhar com você assim que me enviar algum dos comandos abaixo:"
            )
            await enviar_menu_comandos(update, context)
            context.user_data["esperando_confirmacao_comando"] = None
            return

        else:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, "⚠️ Responda apenas com 'sim' ou 'não'.")
            return

    if comando_sugerido and texto not in COMANDOS:
        await enviar_mensagem_segura(context.bot, update.effective_chat.id, 
            f"🤔 Você quis dizer \"{comando_sugerido}\"? Responda SIM ou NÃO."
        )
        context.user_data["esperando_confirmacao_comando"] = comando_sugerido
        return

    # Se não sugeriu nada
    await enviar_mensagem_segura(context.bot, update.effective_chat.id, 
        "Ok, comando cancelado. Digite novamente."
    )
    await enviar_menu_comandos(update, context)
    return

# === Handler para clique no botão ===
async def selecionar_pedido(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    codigo = query.data
    context.user_data["pedido_selecionado"] = codigo
    context.user_data["codigo_pedido"] = codigo

    colunas_vazias = listar_colunas_vazias(codigo, ARQUIVO, ABA)
    if colunas_vazias:
        lista = "\n".join([f"{i+1}. {nome}" for i, nome in enumerate(colunas_vazias.keys())])
        await query.edit_message_text(
            text=f"Você selecionou o pedido {codigo}.\n\nAs colunas vazias são:\n{lista}\n\nResponda no formato: NomeColuna: Valor"
        )
        context.user_data["colunas_vazias"] = colunas_vazias
        context.user_data["esperando_atualizacao"] = True
    else:
        await query.edit_message_text(
            text=f"⚠️ O pedido {codigo} não possui colunas vazias."
        )


# === Handlers ===
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logging.info("Comando /start recebido")
    await enviar_menu_comandos(update, context)

# === Mensagem inicial automática via post_init ===
async def avisar_inicio(app: Application):
    # Mensagem inicial
    await enviar_mensagem_segura(
        app.bot,
        CHAT_ID,
        "🤖 Bot ativo! Estou aguardando suas interações para trabalharmos juntos!\n"
        "Aguarde enquanto eu ajusto alguns detalhes para trabalhar na confeitaria!🍰"
    )
        
    # === Verificar pagamentos pendentes logo no início ===
    pendentes = []
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(ARQUIVO)
        ws = wb.Sheets(ABA)

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            codigo = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            status = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else ""
            cliente = str(ws.Range(f"G{row}").Value).strip() if ws.Range(f"G{row}").Value else ""

            if "pagamento" in status.lower():
                pendentes.append((codigo, cliente))

        wb.Close(SaveChanges=False)
        excel.Quit()
        del excel
    except Exception as e:
        logging.error(f"Erro ao acessar Excel: {e}")
        return

    if pendentes:
        aviso = (
            "⚠️ Identifiquei ausência de pagamentos dos pedidos! Não esqueça de cobrar as pendências!\n"
            "Posso te encaminhar o texto pronto para enviar aos nossos clientes, "
            "basta solicitar com o comando \"Pagamentos pendentes\""
        )
        await enviar_mensagem_segura(app.bot, CHAT_ID, aviso)

    # ✅ Menu inicial
    await enviar_mensagem_segura(
        app.bot,
        CHAT_ID,
        "🤖 Comandos disponíveis para interagir comigo:\n"
        "- \"Enviar calendário\" → Imagem do calendário atualizado 🗓️\n"
        "- \"Pedidos em aberto\" → Pedidos que ainda não foram concluídos 📋\n"
        "- \"Atualizar pedido\" → Preencher colunas vazias de um pedido ✏️\n"
        "- \"Editar pedido\" → Alterar colunas já preenchidas de um pedido 📝\n"
        "- \"Cancelar pedido\" → Inicia o processo para cancelamento ❌\n"
        "- \"Adicionar pedido ao sistema\" → Fluxo para registrar um novo pedido ✅\n"
        "- \"Pagamentos pendentes\" → Textos prontos para cobrar clientes ⚠️\n\n"
        "Use qualquer um desses comandos a qualquer momento e eu cuidarei do restante! 🐱"
    )


# === Mensagem de despedida automática via post_stop ===
async def enviar_despedida(app: Application):
    mensagem = (
        "🤖 Hora da soneca! 😴 Vou buscar as últimas mensagens enviadas no grupo "
        "e continuarei o trabalho assim que o bot_continuo for colocado para rodar novamente na máquina principal."
    )
    await enviar_mensagem_segura(app.bot, CHAT_ID, mensagem)
    logging.info("Mensagem de despedida enviada.")


def main():
    # 1. Verifica se o Excel está aberto e fecha para evitar conflitos
    verificar_excel_aberto()

    # lista de timeouts progressivos
    timeouts = [15, 30, 60]

    for t in timeouts:
        try:
            logging.info(f"🔄 Tentando iniciar bot com connect_timeout={t}s e read_timeout={t}s...")

            # 2. Inicializa o bot com timeout atual
            app = (
                ApplicationBuilder()
                .token(TOKEN)
                .post_init(avisar_inicio)
                .post_stop(enviar_despedida)
                .connect_timeout(t)
                .read_timeout(t)
                .build()
            )

            # 3. Registra os handlers
            app.add_handler(CommandHandler("start", fluxo_inicial))
            app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, mensagens))
            app.add_handler(CallbackQueryHandler(selecionar_pedido))

            logging.info("✅ Bot rodando... Ctrl+C para parar.")
            # 4. Executa o bot em modo polling
            app.run_polling(
                poll_interval=3,
                allowed_updates=Update.ALL_TYPES,
                close_loop=False
            )
            break  # se rodar sem erro, sai do loop

        except telegram.error.NetworkError as e:
            logging.error(f"❌ Erro de rede com timeout={t}s: {e}")
            logging.info("⏳ Tentando novamente com timeout maior...")
            time.sleep(5)  # espera antes de tentar de novo
            continue

    else:
        logging.critical("🚨 Não foi possível iniciar o bot com nenhum dos timeouts configurados.")

if __name__ == "__main__":
    main()
