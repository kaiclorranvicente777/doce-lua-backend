import sys
import asyncio
import os
import re
import datetime
import logging
import requests
import signal
import difflib
import time
import functools
import unicodedata
import fitz  # PyMuPDF
from PIL import Image

# Bibliotecas do Telegram
import telegram
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters
)

# Imports exclusivos do Windows (condicionados)
if sys.platform == "win32":
    import pythoncom
    import win32com.client as win32
    import win32api
    import win32con


logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    datefmt="%d/%m/%Y %H:%M:%S"   # força formato local
)


# === Funções utilitárias ===
async def enviar_mensagem_segura(bot, chat_id, texto, tentativas=3, intervalo=5, **kwargs):
    for tentativa in range(tentativas):
        try:
            await bot.send_message(chat_id=chat_id, text=texto, **kwargs)
            return True  # sucesso
        except Exception as e:
            logging.error(f"Erro ao enviar mensagem (tentativa {tentativa+1}): {e}")
            await asyncio.sleep(intervalo)

    try:
        await bot.send_message(
            chat_id=chat_id,
            text="Desculpe, algum erro ocorreu na rede ou servidores do Telegram. "
                 "Pode me enviar novamente o que precisa que eu faça?"
        )
    except Exception as alerta_final:
        logging.error(f"Falha ao enviar mensagem de erro padrão: {alerta_final}")

    return False


async def enviar_menu_comandos_chat(bot, chat_id):
    mensagem = (
        "🤖 Comandos disponíveis para interagir comigo:\n"
        "- \"Enviar calendário\" →  🗓️\n"
        "- \"Pedidos em aberto\" →  📋\n"
        "- \"Atualizar pedido\" →  ✏️\n"
        "- \"Editar pedido\" →  📝\n"
        "- \"Cancelar pedido\" → ❌\n"
        "- \"Adicionar pedido ao sistema\" → ✅\n"
        "- \"Endereços do dia\" → 🚚\n"
        "- \"Pagamentos pendentes\" → ⚠️\n\n"
        "Use qualquer um desses comandos a qualquer momento e eu cuidarei do restante! 🐱"
    )
    await enviar_mensagem_segura(bot, chat_id, mensagem)


COMANDOS = [
    "enviar calendário",
    "pedidos em aberto",
    "atualizar pedido",
    "editar pedido",
    "cancelar pedido",
    "adicionar pedido ao sistema",
    "endereços do dia",
    "endereços",
    "pagamentos pendentes"
]


def comando_mais_proximo(texto):
    texto = texto.lower().strip()

    # ✅ Ajuste: se o usuário digitar "adicionar pedido" (com ou sem complemento),
    # sempre mapeia para "adicionar pedido ao sistema"
    if texto.startswith("adicionar pedido"):
        return "adicionar pedido ao sistema"
    if texto == "enviar calendario":  # sem acento
        return "enviar calendário"
    if texto == "enderecos":          # sem acento
        return "endereços do dia"

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


def calcular_etiqueta(valor_total, valor_adicional):
    try:
        total = float(valor_total) if valor_total else 0.0
        adicional = float(valor_adicional) if valor_adicional else 0.0
        base = total - adicional

        if 15 <= base <= 29:
            return "⚫️"  # Lua Simples
        elif 30 <= base <= 49:
            return "⚪️"  # Lua Prata
        elif 50 <= base <= 99:
            return "🟡"  # Lua Dourada
        elif base >= 100:
            return "🟣"  # Lua Furta-cor
        else:
            return ""  # sem pontuação
    except Exception:
        return ""

def listar_enderecos_do_dia(arquivo, aba="Acompanhamento Doce Lua"):
    excel = None
    enderecos = []
    pythoncom.CoInitialize()   # ✅ inicializa COM nesta thread
    try:
        hoje = datetime.datetime.now().date()
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            codigo = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            status = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else ""
            cliente = str(ws.Range(f"G{row}").Value).strip() if ws.Range(f"G{row}").Value else ""
            endereco = str(ws.Range(f"S{row}").Value).strip() if ws.Range(f"S{row}").Value else ""
            horario = formatar_valor_excel(ws.Range(f"F{row}").Value)
            produto = str(ws.Range(f"J{row}").Value).strip() if ws.Range(f"J{row}").Value else ""  # ✅ novo
            data_entrega = converter_data_excel(ws.Range(f"E{row}").Value)

            if codigo and codigo != "#0000" and status and "concluído" not in status.lower():
                if data_entrega == hoje:
                    enderecos.append((codigo, cliente, endereco, horario, produto))

        wb.Close(SaveChanges=False)
    except Exception as e:
        logging.error(f"Erro ao listar endereços do dia: {e}", exc_info=True)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()   # ✅ finaliza COM para liberar recursos

    return enderecos


# Função que envia mensagem amigável após timeout
async def fluxo_timeout(context: ContextTypes.DEFAULT_TYPE):
    chat_id = context.job.chat_id
    await enviar_mensagem_segura(
        context.bot,
        chat_id,
        "🐱💭 Parece que você ficou um tempinho sem responder…\n"
        "Tudo bem, imagino que estejam ocupados agora!\n\n"
        "Vou encerrar este fluxo por enquanto e deixar o menu de comandos à mão para quando quiser reiniciar um comando: ⚡🐈\n\n"
        "🤖 Comandos disponíveis para interagir comigo:\n"
        "- \"Enviar calendário\" →  🗓️\n"
        "- \"Pedidos em aberto\" →  📋\n"
        "- \"Atualizar pedido\" →  ✏️\n"
        "- \"Editar pedido\" →  📝\n"
        "- \"Cancelar pedido\" → ❌\n"
        "- \"Adicionar pedido ao sistema\" → ✅\n"
        "- \"Pagamentos pendentes\" → ⚠️\n\n"
        "Use qualquer um desses comandos a qualquer momento e eu cuidarei do restante! 🐱"
    )
    # limpa flags de fluxo
    context.user_data.clear()


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


def interpretar_mensagem(texto):
    logging.debug(f"Texto recebido para interpretação: {texto}")
    dados = {}

    # Nome ou Cliente
    cliente_match = re.search(r"(Cliente|Nome):\s*(.+)", texto, re.IGNORECASE)
    if cliente_match:
        dados["Nome"] = cliente_match.group(2).strip()

    # Item
    item_match = re.search(r"Item:\s*(.+)", texto, re.IGNORECASE)
    if item_match:
        dados["Produto"] = item_match.group(1).strip()

    # Quantidade + Unidade
    qtd_match = re.search(r"(\d+)\s*(Unid\.?|Kg|kg|Caixa|Cx)", texto, re.IGNORECASE)
    if qtd_match:
        dados["Quantidade"] = qtd_match.group(1).strip()
        dados["UnidadeMedida"] = qtd_match.group(2).strip()

    # Sabor/Recheio
    sabor_match = re.search(r"Sabor:\s*([^\n]+)", texto, re.IGNORECASE)
    if sabor_match:
        valor = sabor_match.group(1).strip()
        valor = re.sub(r"com massa.*", "", valor)
        valor = re.sub(r"\(R\$[0-9.,]+\)", "", valor)
        dados["SaborRecheio"] = valor.strip()

    # Massa
    massa_match = re.search(r"massa\s*de\s*([^\n]+)", texto, re.IGNORECASE)
    if massa_match:
        valor = massa_match.group(1).strip()
        valor = re.sub(r"\(R\$[0-9.,]+\)", "", valor)
        dados["SaborMassa"] = valor.strip()

    # Detalhes adicionais (separa descrição e valor, trata frete como descrição)
    adicional_match = re.search(r"Adicional:\s*(.+)", texto, re.IGNORECASE)
    if adicional_match:
        valor = adicional_match.group(1).strip()
        valor_monetario = re.search(r"\(R\$[0-9.,]+\)", valor)
        if valor_monetario:
            if "5,00" in valor_monetario.group(0):  # frete
                dados["DescricaoAdicional"] = valor  # mantém descrição + frete
            else:
                dados["DescricaoAdicional"] = valor.replace(valor_monetario.group(0), "").strip()
                dados["ValorAdicional"] = valor_monetario.group(0).replace("(", "").replace(")", "")
        else:
            dados["DescricaoAdicional"] = valor

    # Valores adicionais extras (ex.: frete no endereço)
    adicionais = []
    for match in re.findall(r"\(R\$[0-9.,]+\)", texto):
        if "Total" not in texto and "Sabor" not in texto:
            if "5,00" in match:  # frete
                if dados.get("DescricaoAdicional"):
                    dados["DescricaoAdicional"] += f" | Frete {match}"
                else:
                    dados["DescricaoAdicional"] = f"Frete {match}"
            else:
                adicionais.append(match)

    if adicionais:
        if dados.get("ValorAdicional"):
            dados["ValorAdicional"] += " + " + " + ".join(adicionais)
        else:
            dados["ValorAdicional"] = " + ".join(adicionais)

    # Endereço
    endereco_match = re.search(r"(Endere[cç]o|Endereço para (Entrega|Coleta)):\s*(.+)", texto, re.IGNORECASE)
    if endereco_match:
        dados["Endereco"] = endereco_match.group(3).strip()

    # Data (aceita 1 ou 2 dígitos + remove extras como (Sábado))
    data_match = re.search(r"Data:\s*([0-9]{1,2}/[0-9]{1,2})", texto, re.IGNORECASE)
    if data_match:
        dados["Data"] = data_match.group(1).strip()

    # Horário
    horario_match = re.search(r"Horário:\s*(.+)", texto, re.IGNORECASE)
    if horario_match:
        dados["Horario"] = horario_match.group(1).strip()
    else:
        dados["Horario"] = ""

    # Preço total
    preco_match = re.search(r"Total[:\s]*\(?(R\$[0-9.,]+)\)?", texto, re.IGNORECASE)
    if preco_match:
        dados["Preco"] = preco_match.group(1).strip()

    dados["Entrega"] = ""

    logging.debug(f"Dados extraídos: {dados}")
    return dados


# === Funções Excel ===
def gravar_excel(dados, arquivo, aba="Acompanhamento Doce Lua"):
    excel = None
    linha_usada, codigo_pedido, status_atual = None, None, None
    pythoncom.CoInitialize()   # ✅ inicializa COM nesta thread
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            if valor_ax == "#0000":
                linha_usada = row

                # --- Data ---
                data_str = dados.get("Data")
                if data_str:
                    try:
                        data_str = re.sub(r"\(.*?\)", "", data_str).strip()
                        ano_atual = datetime.datetime.now().year

                        if re.match(r"^\d{1,2}/\d{1,2}$", data_str):
                            data_str = f"{data_str}/{ano_atual}"
                        elif re.match(r"^\d{1,2}/\d{1,2}/\d{2}$", data_str):
                            dia, mes, ano = data_str.split("/")
                            ano = f"20{ano}"
                            data_str = f"{dia}/{mes}/{ano}"

                        data_obj = datetime.datetime.strptime(data_str, "%d/%m/%Y").date()
                        excel_serial = (data_obj - datetime.date(1899, 12, 30)).days
                        ws.Range(f"E{row}").Value2 = excel_serial
                        ws.Range(f"E{row}").NumberFormatLocal = "dd/mm/yyyy"
                    except Exception:
                        ws.Range(f"E{row}").Value = data_str

                # --- Demais campos ---
                ws.Range(f"F{row}").Value = dados.get("Horario")
                ws.Range(f"G{row}").Value = dados.get("Nome")
                ws.Range(f"H{row}").Value = dados.get("Quantidade")
                ws.Range(f"I{row}").Value = dados.get("UnidadeMedida")
                ws.Range(f"J{row}").Value = dados.get("Produto")
                ws.Range(f"L{row}").Value = dados.get("SaborMassa")
                ws.Range(f"N{row}").Value = dados.get("SaborRecheio")
                ws.Range(f"P{row}").Value = dados.get("DescricaoAdicional")

                # --- Valor adicional ---
                valor_adicional = dados.get("ValorAdicional")
                if valor_adicional:
                    valor_num = str(valor_adicional).replace("R$", "").replace(",", ".").strip()
                    try:
                        ws.Range(f"Q{row}").Value = float(valor_num)
                    except Exception:
                        ws.Range(f"Q{row}").Value = valor_adicional
                else:
                    ws.Range(f"Q{row}").Value = None

                # --- Endereço e outros ---
                ws.Range(f"S{row}").Value = dados.get("Endereco")
                ws.Range(f"U{row}").Value = dados.get("Entrega")
                ws.Range(f"V{row}").Value = dados.get("Pagamento")
                ws.Range(f"W{row}").Value = dados.get("RecebimentoFotoComPedido")
                ws.Range(f"X{row}").Value = dados.get("PublicaçãoInstagram")
                ws.Range(f"Y{row}").Value = limpar_valor_monetario(dados.get("Preco"))

                codigo_pedido = str(ws.Range(f"AX{row}").Value).strip()
                status_atual = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else "Sem status"
                break

        if linha_usada:
            wb.Save()
            wb.Close(SaveChanges=True)
        else:
            wb.Close(SaveChanges=False)
    except Exception as e:
        logging.error(f"Erro ao gravar pedido: {e}", exc_info=True)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return linha_usada, codigo_pedido, status_atual

        
# Corte 2

def listar_pedidos_abertos(arquivo, aba="Acompanhamento Doce Lua"):
    excel = None
    pedidos = []
    pythoncom.CoInitialize()   # ✅ inicializa COM nesta thread
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            codigo = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            status = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else ""
            cliente = str(ws.Range(f"G{row}").Value).strip() if ws.Range(f"G{row}").Value else ""
            data_entrega = converter_data_excel(ws.Range(f"E{row}").Value)

            # pega valores monetários
            valor_total = limpar_valor_monetario(ws.Range(f"Y{row}").Value)
            valor_adicional = limpar_valor_monetario(ws.Range(f"Q{row}").Value)

            # calcula etiqueta fidelidade
            etiqueta = calcular_etiqueta(valor_total, valor_adicional)

            # Só adiciona se não estiver concluído
            if codigo and codigo != "#0000" and status and "concluído" not in status.lower():
                data_str = data_entrega.strftime("%d/%m/%Y") if data_entrega else " "
                pedidos.append(f"{codigo} - {data_str} - {etiqueta} {cliente}\n({status})")

        wb.Close(SaveChanges=False)
    except Exception as e:
        logging.error(f"Erro ao listar pedidos abertos: {e}", exc_info=True)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()   # ✅ finaliza COM para liberar recursos

    return pedidos


def listar_colunas_vazias(codigo, arquivo, aba="Acompanhamento Doce Lua"):
    excel = None
    colunas_vazias = {}
    pythoncom.CoInitialize()   # ✅ inicializa COM nesta thread
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        mapa = {
            "E": "Data","F": "Horário","G": "Nome","H": "Quantidade","I": "Unidade Medida","J": "Produto",
            "L": "Sabor Massa","M": "Preparação Massa","N": "Sabor Recheio","O": "Preparação Recheio",
            "P": "Descrição Adicional","Q": "Valor Adicional","R": "Coleta Detalhes adicionais",
            "S": "Endereco","U": "Entrega","V": "Pagamento","W": "Recebimento foto com pedido",
            "X": "Publicação Instagram","Y": "Preço"
        }

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            if valor_ax == codigo:
                for col, nome in mapa.items():
                    valor = ws.Range(f"{col}{row}").Value
                    if not valor or str(valor).strip() == "":
                        colunas_vazias[nome] = col
                break

        wb.Close(SaveChanges=False)
    except Exception as e:
        logging.error(f"Erro ao listar colunas vazias: {e}", exc_info=True)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()   # ✅ finaliza COM para liberar recursos

    return colunas_vazias


def atualizar_coluna(codigo, coluna_excel, valor, arquivo, aba="Acompanhamento Doce Lua"):
    excel = None
    sucesso, status_atual = False, None
    pythoncom.CoInitialize()   # ✅ inicializa COM nesta thread
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            if valor_ax == codigo:
                # ✅ Tratamento especial para colunas Q e Y (valores monetários)
                if coluna_excel in ["Q", "Y"]:
                    ws.Range(f"{coluna_excel}{row}").Value = limpar_valor_monetario(valor)
                else:
                    ws.Range(f"{coluna_excel}{row}").Value = valor

                status_atual = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else "Sem status"
                wb.Save()
                sucesso = True
                break

        wb.Close(SaveChanges=True)
    except Exception as e:
        logging.error(f"Erro ao atualizar coluna: {e}", exc_info=True)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()   # ✅ finaliza COM para liberar recursos

    return sucesso, status_atual


# === Função para cancelar pedido ===
def cancelar_pedido(codigo, arquivo, aba="Acompanhamento Doce Lua"):
    excel = None
    sucesso, status_atual = False, None
    pythoncom.CoInitialize()   # ✅ inicializa COM nesta thread
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(arquivo)
        ws = wb.Sheets(aba)

        mapa = {
            "E": "Data","F": "Horário","G": "Nome","H": "Quantidade","I": "Unidade Medida","J": "Produto",
            "L": "Sabor Massa","M": "Preparação Massa","N": "Sabor Recheio","O": "Preparação Recheio",
            "P": "Descrição Adicional","Q": "Valor Adicional","R": "Coleta Detalhes adicionais",
            "S": "Endereço","U": "Entrega","V": "Pagamento","W": "Recebimento foto com pedido",
            "X": "Publicação Instagram","Y": "Preço"
        }

        for row in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            if valor_ax == codigo:
                # Marca todas as colunas vazias como "Cancelado"
                for col in mapa.keys():
                    valor = ws.Range(f"{col}{row}").Value
                    if not valor:
                        ws.Range(f"{col}{row}").Value = "Cancelado"

                status_atual = "Cancelado"
                wb.Save()
                sucesso = True
                break

        wb.Close(SaveChanges=True)
    except Exception as e:
        logging.error(f"Erro ao cancelar pedido: {e}", exc_info=True)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()   # ✅ finaliza COM

    return sucesso, status_atual


IMAGEM_PATH = r"C:\Users\KICEGA~1\OneDrive\Documentos\excel_intervalo.png"

def gerar_imagem_excel():
    excel, wb = None, None
    pdf_path = r"C:\Users\KICEGA~1\OneDrive\Documentos\calendario.pdf"

    pythoncom.CoInitialize()   # ✅ inicializa COM nesta thread
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

        # ✅ Ajusta área de impressão para caber em uma página
        ws.PageSetup.PrintArea = "I1:AP27"
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1

        # Exporta intervalo como PDF
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
            if wb is not None:
                wb.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
            time.sleep(2)
            logging.info("Excel fechado completamente após execução.")
        except Exception as e:
            logging.error(f"Erro ao fechar Excel: {e}", exc_info=True)
        finally:
            pythoncom.CoUninitialize()   # ✅ finaliza COM para liberar recursos

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
        "- \"Enviar calendário\" → 🗓️\n"
        "- \"Pedidos em aberto\" → 📋\n"
        "- \"Atualizar pedido\" → ✏️\n"
        "- \"Editar pedido\" → 📝\n"
        "- \"Cancelar pedido\" → ❌\n"
        "- \"Adicionar pedido ao sistema\" → ✅\n"
        "- \"Endereços do dia\" → 🚚\n"
        "- \"Pagamentos pendentes\" → ⚠️\n\n"
        "Use qualquer um desses comandos a qualquer momento e eu cuidarei do restante! 🐱"
    )
    await enviar_mensagem_segura(context.bot, update.effective_chat.id, mensagem)


pedidos_enviados = []  # memória simples para comparar pedidos anteriores


async def fluxo_inicial(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # 1. Pedidos em aberto
    pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
    if pedidos_abertos:
        lista = "\n\n".join(pedidos_abertos)
        await enviar_mensagem_segura(
            context.bot,
            update.effective_chat.id,
            f"📋 Aqui estão os pedidos em aberto:\n{lista}"
        )
    else:
        await enviar_mensagem_segura(
            context.bot,
            update.effective_chat.id,
            "Nenhum pedido em aberto ✅"
        )

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

    # ✅ Cancela timeout se o usuário respondeu dentro de um fluxo
    if context.user_data.get("esperando_coluna"):
        jobs = context.job_queue.get_jobs_by_name(
            f"timeout_atualizar_{context.user_data.get('codigo_pedido')}"
        )
        for job in jobs:
            job.schedule_removal()

    if context.user_data.get("esperando_coluna_edicao"):
        jobs = context.job_queue.get_jobs_by_name(
            f"timeout_editar_{context.user_data.get('codigo_edicao')}"
        )
        for job in jobs:
            job.schedule_removal()

    if context.user_data.get("esperando_cancelamento"):
        jobs = context.job_queue.get_jobs_by_name("timeout_cancelar")
        for job in jobs:
            job.schedule_removal()

    if context.user_data.get("esperando_pedido"):
        jobs = context.job_queue.get_jobs_by_name("timeout_adicionar")
        for job in jobs:
            job.schedule_removal()
        
    # 🔴 Fluxo de cancelamento de pedido (PRIORITÁRIO)
    if context.user_data.get("esperando_cancelamento"):
        logging.debug(
            f"[CANCELAMENTO] esperando={context.user_data.get('esperando_cancelamento')}, "
            f"confirmando={context.user_data.get('confirmando_cancelamento')}, texto={texto}"
        )

        # 👉 Fase 2: já estamos aguardando confirmação (sim/não)
        if context.user_data.get("confirmando_cancelamento"):
            resposta = texto.lower().strip()
            if resposta in ["sim", "s", "yes", "y"]:
                codigo = context.user_data.get("codigo_cancelamento")
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

        # 👉 Fase 1: usuário acabou de digitar o código do pedido
        codigo = texto.strip()
        colunas_vazias = listar_colunas_vazias(codigo, ARQUIVO, ABA)
        if not colunas_vazias:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"⚠️ O pedido {codigo} não foi encontrado."
            )
            context.user_data["esperando_cancelamento"] = False
            return

        await enviar_mensagem_segura(
            context.bot, update.effective_chat.id,
            f"Você confirma o cancelamento do pedido {codigo}? Responda SIM ou NÃO."
        )
        context.user_data["codigo_cancelamento"] = codigo
        context.user_data["confirmando_cancelamento"] = True
        return


        # 1️⃣ Fluxo de atualização de colunas (único)
        if context.user_data.get("esperando_coluna"):
            # 👉 Intercepta "sim" logo no início
            if texto.lower().strip() in ["sim", "s", "yes", "y"]:
                if context.user_data.get("comando_confirmado"):
                    # 👉 limpa estado de atualização antes de trocar
                    context.user_data["esperando_coluna"] = False
                    context.user_data["codigo_pedido"] = None
                    context.user_data["colunas_vazias"] = None
                    context.user_data["pendentes_atualizacao"] = None
                    context.user_data["aguardando_confirmacao_atualizacao"] = False

                    await enviar_mensagem_segura(
                        context.bot, update.effective_chat.id,
                        "Acho que quer trocar para outro comando…"
                    )
                    # 🔥 dispara o comando confirmado
                    await mensagens(update, context)
                    context.user_data["comando_confirmado"] = None  # 🔧 resetar para não loopar
                    return

            codigo_pedido = context.user_data.get("codigo_pedido")
            colunas_vazias = listar_colunas_vazias(codigo_pedido, ARQUIVO, ABA)

            entradas = texto.splitlines()
            pendentes = []
            for entrada in entradas:
                if ":" not in entrada:
                    sugestao = comando_mais_proximo(entrada)
                    if sugestao:
                        context.user_data["comando_confirmado"] = sugestao

                        # 👉 limpa estados ativos para não cair de novo no fluxo errado
                        for flag in [
                            "esperando_coluna",
                            "aguardando_confirmacao_atualizacao",
                            "esperando_coluna_edicao",
                            "aguardando_confirmacao_edicao",
                            "esperando_pedido",
                            "esperando_cancelamento"
                        ]:
                            context.user_data[flag] = False

                        await enviar_mensagem_segura(
                            context.bot, update.effective_chat.id,
                            "Acho que quer trocar para outro comando…"
                        )
                        await enviar_mensagem_segura(
                            context.bot, update.effective_chat.id,
                            f"Você quis dizer \"{sugestao}\"?"
                        )
                        return  # 🔴 sai aqui para não continuar no fluxo de atualização
                    else:
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
                    nomes = context.user_data.get("lista_colunas_vazias", [])
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
            context.user_data["pendentes_atualizacao"] = pendentes

            # Se só veio uma coluna, grava direto
            if len(pendentes) == 1:
                nome_coluna, valor = pendentes[0]
                coluna_excel = colunas_vazias.get(nome_coluna)
                if coluna_excel:
                    sucesso, status = atualizar_coluna(codigo_pedido, coluna_excel, valor, ARQUIVO, ABA)
                    if sucesso:
                        await enviar_mensagem_segura(
                            context.bot, update.effective_chat.id,
                            f"✅ Coluna '{nome_coluna}' atualizada com: {valor}"
                        )
                        await enviar_mensagem_segura(
                            context.bot, update.effective_chat.id,
                            f"📌 Status atual do pedido {codigo_pedido}: {status}"
                        )


                    # Relista colunas restantes
                    colunas_restantes = listar_colunas_vazias(codigo_pedido, ARQUIVO, ABA)
                    if colunas_restantes:
                        lista_formatada = "\n".join([f"{i+1}. {nome}" for i, nome in enumerate(colunas_restantes.keys())])
                        await enviar_mensagem_segura(
                            context.bot,
                            update.effective_chat.id,
                            f"Ainda faltam:\n{lista_formatada}\n\n"
                            "Responda no formato: Número: Valor ou NomeColuna: Valor\n\n"
                            "👉 Se quiser continuar atualizando pedidos, ative o comando atualizar pedido"
                            " ou envie um novo comando para iniciar outro fluxo."
                        )
                    else:
                        await enviar_mensagem_segura(
                            context.bot, update.effective_chat.id,
                            f"🎉 Todas as colunas selecionadas do pedido {codigo_pedido} foram preenchidas!"
                        )
                    context.user_data["esperando_coluna"] = False
            else:
                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "⚠️ Coluna inválida. Use o número da lista ou o nome correto."
                )

        # 👉 Se veio mais de uma coluna, pede confirmação
        elif len(pendentes) > 1:
            context.user_data["pendentes_atualizacao"] = pendentes
            preview = "\n".join([f"{col}: {val}" for col, val in pendentes])
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"Você solicitou atualizar o pedido {codigo_pedido} com os seguintes valores:\n\n{preview}\n\n"
                "Digite 'sim' para confirmar ou 'não' para cancelar."
            )
            context.user_data["aguardando_confirmacao_atualizacao"] = True
            context.user_data["esperando_coluna"] = False

        return  # 🔴 Sai aqui e não continua para checar comandos

    # # === Confirmação da atualização múltipla ===
    if context.user_data.get("aguardando_confirmacao_atualizacao"):
        resposta = texto.lower().strip()

        # 👉 Intercepta "sim" logo no início
        if resposta in ["sim", "s", "yes", "y"]:
            if context.user_data.get("comando_confirmado"):
                # 👉 limpa estado antes de trocar de comando
                for flag in [
                    "aguardando_confirmacao_atualizacao",
                    "pendentes_atualizacao",
                    "codigo_pedido",
                    "esperando_coluna",
                    "esperando_coluna_edicao",
                    "aguardando_confirmacao_edicao",
                    "esperando_pedido",
                    "esperando_cancelamento"
                ]:
                    context.user_data[flag] = False

                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "Acho que quer trocar para outro comando…"
                )
                # 🔥 dispara o comando confirmado
                await mensagens(update, context)
                context.user_data["comando_confirmado"] = None  # 🔧 resetar para não loopar
                return

            # 👉 Se não há comando alternativo, aplica as alterações normalmente
            codigo = context.user_data.get("codigo_pedido")
            pendentes = context.user_data.get("pendentes_atualizacao", [])
            alteracoes = []

            colunas_vazias = context.user_data.get("colunas_vazias", {})

            for nome_coluna, valor in pendentes:
                # ✅ pega direto do dicionário nome → letra
                coluna_excel = colunas_vazias.get(nome_coluna)

                if coluna_excel:
                    ok, status_atual = atualizar_coluna(codigo, coluna_excel, valor.strip(), ARQUIVO, ABA)
                    if ok:
                        alteracoes.append(f"{nome_coluna}: {valor}")
                    else:
                        alteracoes.append(f"{nome_coluna}: ⚠️ erro ao atualizar")
                else:
                    alteracoes.append(f"{nome_coluna}: ⚠️ coluna não disponível")


            # 👉 Envia mensagem consolidada
            if alteracoes:
                mensagem_final = (
                    f"📝 Pedido {codigo} atualizado!\n"
                    + "\n".join(alteracoes)
                    + f"\n\n🎉 Todas as colunas selecionadas do pedido {codigo} foram preenchidas!"
                )
                await enviar_mensagem_segura(context.bot, update.effective_chat.id, mensagem_final)

            # 👉 Buscar status atual na coluna D
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

            # 👉 limpa estado após aplicar
            context.user_data["aguardando_confirmacao_atualizacao"] = False
            context.user_data["pendentes_atualizacao"] = None
            context.user_data["esperando_coluna"] = False
            return

        elif resposta in ["não", "nao", "n", "no"]:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "❌ Alterações canceladas."
            )
            # 👉 limpa estado
            context.user_data["aguardando_confirmacao_atualizacao"] = False
            context.user_data["pendentes_atualizacao"] = None
            context.user_data["esperando_coluna"] = False
            return

        # 👉 Se o usuário digitar outro comando
        sugestao = comando_mais_proximo(resposta)
        if sugestao:
            context.user_data["comando_confirmado"] = sugestao  # 🔧 salva sugestão
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "Acho que quer trocar para outro comando…"
            )
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"Você quis dizer \"{sugestao}\"?"
            )
            return

        # 👉 Se não for comando nem sim/não
        await enviar_mensagem_segura(
            context.bot, update.effective_chat.id,
            "⚠️ Responda apenas com 'sim' ou 'não'."
        )


    # === Receber edição múltipla (pré-visualização) ===
    if context.user_data.get("esperando_coluna_edicao"):
        # 👉 Intercepta "sim" logo no início
        if texto.lower().strip() in ["sim", "s", "yes", "y"]:
            if context.user_data.get("comando_confirmado"):
                # 👉 limpa estado antes de trocar
                context.user_data["esperando_coluna_edicao"] = False
                context.user_data["aguardando_confirmacao_edicao"] = False
                context.user_data["pendentes_edicao"] = None
                context.user_data["codigo_edicao"] = None
                context.user_data["mapa_edicao"] = None

                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "Acho que quer trocar para outro comando…"
                )
                # 🔥 dispara o comando confirmado
                await mensagens(update, context)
                context.user_data["comando_confirmado"] = None  # 🔧 resetar
                return

        # 👉 fluxo normal da edição
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
                sugestao = comando_mais_proximo(entrada)
                if sugestao:
                    context.user_data["comando_confirmado"] = sugestao  # 🔧 salva sugestão

                    # 👉 limpa estados ativos para não cair de novo no fluxo errado
                    for flag in [
                        "esperando_coluna",
                        "aguardando_confirmacao_atualizacao",
                        "esperando_coluna_edicao",
                        "aguardando_confirmacao_edicao",
                        "esperando_pedido",
                        "esperando_cancelamento"
                    ]:
                        context.user_data[flag] = False

                    await enviar_mensagem_segura(
                        context.bot, update.effective_chat.id,
                        "Acho que quer trocar para outro comando…"
                    )
                    await enviar_mensagem_segura(
                        context.bot, update.effective_chat.id,
                        f"Você quis dizer \"{sugestao}\"?"
                    )
                    return  # 🔴 sai aqui para não continuar no fluxo de edição
                else:
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
        return


      

    # === Confirmação da edição múltipla ===
    if context.user_data.get("aguardando_confirmacao_edicao"):
        resposta = texto.lower().strip()

        # 👉 Intercepta "sim" logo no início
        if resposta in ["sim", "s", "yes", "y"]:
            if context.user_data.get("comando_confirmado"):
                # 👉 limpa estado antes de trocar de comando
                context.user_data["aguardando_confirmacao_edicao"] = False
                context.user_data["pendentes_edicao"] = None
                context.user_data["codigo_edicao"] = None
                context.user_data["mapa_edicao"] = None
                context.user_data["esperando_coluna_edicao"] = False

                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "Acho que quer trocar para outro comando…"
                )
                # 🔥 dispara o comando confirmado
                await mensagens(update, context)
                context.user_data["comando_confirmado"] = None  # 🔧 resetar para não loopar
                return

            # 👉 Se não há comando alternativo, aplica as alterações normalmente
            codigo = context.user_data.get("codigo_edicao")
            pendentes = context.user_data.get("pendentes_edicao", [])
            mapa_edicao = context.user_data.get("mapa_edicao", {})

            alteracoes = []
            for nome_coluna, valor in pendentes:
                # ✅ pega direto do dicionário nome → letra
                coluna_excel = {v: k for k, v in mapa_edicao.items()}.get(nome_coluna)

                if coluna_excel:
                    ok, status_atual = atualizar_coluna(codigo, coluna_excel, valor.strip(), ARQUIVO, ABA)
                    if ok:
                        alteracoes.append(f"{nome_coluna}: {valor}")
                    else:
                        alteracoes.append(f"{nome_coluna}: ⚠️ erro ao atualizar")
                else:
                    alteracoes.append(f"{nome_coluna}: ⚠️ coluna não disponível")

            # 👉 Mensagem consolidada
            mensagem_final = (
                f"📝 Pedido {codigo} atualizado!\n"
                + "\n".join(alteracoes)
                + f"\n\n🎉 Todas as colunas selecionadas do pedido {codigo} foram preenchidas!"
            )
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, mensagem_final)

            # 👉 limpa estado após aplicar
            context.user_data["aguardando_confirmacao_edicao"] = False
            context.user_data["pendentes_edicao"] = None
            context.user_data["codigo_edicao"] = None
            context.user_data["mapa_edicao"] = None
            context.user_data["esperando_coluna_edicao"] = False
            await enviar_menu_comandos(update, context)
            return

        elif resposta in ["não", "nao", "n", "no"]:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "❌ Alterações canceladas."
            )
            # 👉 limpa estado
            context.user_data["aguardando_confirmacao_edicao"] = False
            context.user_data["pendentes_edicao"] = None
            context.user_data["codigo_edicao"] = None
            context.user_data["mapa_edicao"] = None
            context.user_data["esperando_coluna_edicao"] = False
            await enviar_menu_comandos(update, context)
            return

        # 👉 Se o usuário digitar outro comando
        sugestao = comando_mais_proximo(resposta)
        if sugestao:
            context.user_data["comando_confirmado"] = sugestao  # 🔧 salva sugestão
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "Acho que quer trocar para outro comando…"
            )
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"Você quis dizer \"{sugestao}\"?"
            )
            return

        # 👉 Se não for comando nem sim/não
        await enviar_mensagem_segura(
            context.bot, update.effective_chat.id,
            "⚠️ Responda apenas com 'sim' ou 'não' para confirmar a edição."
        )


    # === Receber dados do novo pedido ===
    if context.user_data.get("esperando_pedido"):
        # 👉 Intercepta "sim" logo no início
        if texto.lower().strip() in ["sim", "s", "yes", "y"]:
            if context.user_data.get("comando_confirmado"):
                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "Acho que quer trocar para outro comando…"
                )
                # 🔥 dispara o comando confirmado
                await mensagens(update, context)
                context.user_data["comando_confirmado"] = None  # 🔧 resetar para não loopar
            return

        # 👉 Sugestão de comando alternativo (sem formato esperado)
        sugestao = comando_mais_proximo(texto)
        if sugestao and ":" not in texto:
            context.user_data["comando_confirmado"] = sugestao

            # 👉 limpa estados ativos para não cair de novo no fluxo errado
            for flag in [
                "esperando_coluna",
                "aguardando_confirmacao_atualizacao",
                "esperando_coluna_edicao",
                "aguardando_confirmacao_edicao",
                "esperando_pedido",
                "esperando_cancelamento"
            ]:
                context.user_data[flag] = False

            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "Acho que quer trocar para outro comando…"
            )
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"Você quis dizer \"{sugestao}\"?"
            )
            return

        # 👉 fluxo normal de novo pedido
        dados = interpretar_mensagem(texto)
        linha, codigo_pedido, status_atual = gravar_excel(dados, ARQUIVO, ABA)
        if linha:
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                f"✅ Pedido {codigo_pedido} adicionado ao sistema! | Status: {status_atual}"
            )
        else:
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                "⚠️ Não foi possível gravar o pedido. Verifique os dados enviados."
            )
        context.user_data["esperando_pedido"] = False
        await enviar_menu_comandos(update, context)
        return



    # 2️⃣ Fluxos normais de comando
    texto_lower = texto.lower()

    if context.user_data.get("comando_confirmado"):
        comando = context.user_data["comando_confirmado"]
    else:
        if context.user_data.get("esperando_confirmacao_comando"):
            resposta = texto_lower
            comando_sugerido = context.user_data["esperando_confirmacao_comando"].lower().strip()

            if resposta in ["sim", "s", "yes", "y"] or resposta == comando_sugerido:
                comando = context.user_data["esperando_confirmacao_comando"]
                context.user_data["comando_confirmado"] = comando
                context.user_data["esperando_confirmacao_comando"] = None

                # 🔧 Limpa todos os estados de fluxo para não conflitar
                for flag in [
                    "esperando_coluna",
                    "aguardando_confirmacao_atualizacao",
                    "esperando_coluna_edicao",
                    "aguardando_confirmacao_edicao",
                    "esperando_pedido",
                    "esperando_cancelamento"
                ]:
                    context.user_data[flag] = False

                await mensagens(update, context)
                return


            elif resposta in ["não", "nao", "n", "no"]:
                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "Ok, comando cancelado. Estarei pronto para trabalhar com você assim que me enviar algum dos comandos abaixo:"
                )
                await enviar_menu_comandos(update, context)
                context.user_data["esperando_confirmacao_comando"] = None
                return

            else:
                await enviar_mensagem_segura(context.bot, update.effective_chat.id, "⚠️ Responda apenas com 'sim' ou 'não'.")
                return

        comando = comando_mais_proximo(texto_lower)
        if comando and texto_lower not in COMANDOS:
            context.user_data["esperando_confirmacao_comando"] = comando
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"🤔 Você quis dizer \"{comando}\"? Responda SIM ou NÃO."
            )
            return
        elif not comando:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "❌ Não entendi esse comando 🤔. Tente novamente ou veja a lista abaixo 👇"
            )
            await enviar_menu_comandos(update, context)
            return

    # === Aqui seguem os fluxos normais (enviar calendário, pedidos em aberto, editar pedido, cancelar pedido, etc.
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
            lista = "\n\n".join(pedidos_abertos)
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

    # === Fluxo: Endereços do dia ===
    if comando in ["endereços do dia", "endereços"]:
        enderecos = listar_enderecos_do_dia(ARQUIVO, ABA)
        if enderecos:
            for codigo, cliente, endereco, horario, produto in enderecos:
                keyboard = [
                    [InlineKeyboardButton(text=f"Marcar {codigo} como entregue", callback_data=f"entregar:{codigo}")]
                ]
                reply_markup = InlineKeyboardMarkup(keyboard)
                await enviar_mensagem_segura(
                    context.bot,
                    update.effective_chat.id,
                    f"{codigo} - {cliente}\n{endereco}\nHorário: {horario}\nProduto: {produto}",
                    reply_markup=reply_markup
                )
        else:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, "✅ Nenhum pedido para entrega hoje.")
        context.user_data["comando_confirmado"] = None
        return


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

    # === Fluxo: Cancelar pedido ===
    if comando == "cancelar pedido":
        pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
        if pedidos_abertos:
            lista = "\n\n".join(pedidos_abertos)
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                f"Qual pedido devemos cancelar?\nPedidos em aberto:\n{lista}"
            )
            context.user_data["usuario_fluxo"] = update.effective_user.first_name
            context.user_data["esperando_cancelamento"] = True
            context.user_data["tempo_inicio_cancelamento"] = datetime.datetime.now()

            # ✅ agenda timeout em 180 segundos (3 minutos)
            context.job_queue.run_once(
                fluxo_timeout,
                180,
                chat_id=update.effective_chat.id,
                name="timeout_cancelar"
            )

            logging.info("[CANCELAMENTO] Flag esperando_cancelamento foi setada como True")
        else:
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                "⚠️ Nenhum pedido em aberto encontrado."
            )
            await enviar_menu_comandos(update, context)

        # 👉 Zera comando confirmado no início
        context.user_data["comando_confirmado"] = None
        return

    # === Receber confirmação de cancelamento ===
    if context.user_data.get("esperando_cancelamento"):
        # 👉 Intercepta "sim" logo no início
        if texto.lower().strip() in ["sim", "s", "yes", "y"]:
            if context.user_data.get("comando_confirmado"):
                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "Acho que quer trocar para outro comando…"
                )
                # 🔥 dispara o comando confirmado
                await mensagens(update, context)
                context.user_data["comando_confirmado"] = None  # 🔧 resetar para não loopar
            return

        resposta = texto.lower().strip()
        if resposta in ["não", "nao", "n", "no"]:
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "❌ Cancelamento abortado."
            )
            context.user_data["esperando_cancelamento"] = False
            return

        # 👉 Se o usuário digitar outro comando
        sugestao = comando_mais_proximo(resposta)
        if sugestao:
            context.user_data["comando_confirmado"] = sugestao  # 🔧 salva sugestão
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                "Acho que quer trocar para outro comando…"
            )
            await enviar_mensagem_segura(
                context.bot, update.effective_chat.id,
                f"Você quis dizer \"{sugestao}\"?"
            )
            return

        # 👉 Se não for comando nem sim/não
        await enviar_mensagem_segura(
            context.bot, update.effective_chat.id,
            "⚠️ Responda apenas com 'sim' ou 'não'."
        )


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
                await enviar_mensagem_segura(
                    context.bot,
                    update.effective_chat.id,
                    f"{codigo} - {resto}",
                    reply_markup=reply_markup
                )

        else:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, "Nenhum pedido em aberto ✅")

        context.user_data["comando_confirmado"] = None
        return


    # === Fluxo: Editar pedido ===
    if comando == "editar pedido":
        pedidos_abertos = listar_pedidos_abertos(ARQUIVO, ABA)
        if pedidos_abertos:
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                "📋 Qual pedido você quer editar?\n\nCódigo pedido Individual - (Status atual)"
            )

            for pedido in pedidos_abertos:
                partes = pedido.split(" - ")
                codigo = partes[0].strip()
                resto = " - ".join(partes[1:])  # cliente + status

                keyboard = [[InlineKeyboardButton(text=f"Selecionar {codigo}", callback_data=f"editar:{codigo}")]]
                reply_markup = InlineKeyboardMarkup(keyboard)

                await enviar_mensagem_segura(
                    context.bot,
                    update.effective_chat.id,
                    f"{codigo} - {resto}",
                    reply_markup=reply_markup
                )
        else:
            await enviar_mensagem_segura(context.bot, update.effective_chat.id, "Nenhum pedido em aberto ✅")

        context.user_data["comando_confirmado"] = None
        return

    
# corte 5

    # === Fluxo: Adicionar pedido ao sistema ===
    if comando == "adicionar pedido ao sistema":
        # mensagem1
        await enviar_mensagem_segura(
            context.bot,
            update.effective_chat.id,
            "Vamos adicionar um novo pedido ao sistema! Por favor, envie os detalhes no formato padrão:"
        )

        # mensagem2
        await enviar_mensagem_segura(
            context.bot,
            update.effective_chat.id,
            "Seu pedido Doce Lua confirmado:\n"
            " \n"
            "🌙 Nome:\n"
            "🌙 Item:\n"
            "🌙 Sabor: # (R$#)\n"
            "🌙 Detalhes:\n"
            "🌙 Adicional: # (R$#)\n"
            "🌙 Horário:\n"
            "🌙 Endereço para Entrega:\n"
            "Rua Argentina, 930 - Fazendinha Real (R$#)\n\n"
            "🌙 Data:\n"
            "Total: R$#"
        )

        # 👉 Marca estado de espera
        context.user_data["esperando_pedido"] = True
        context.user_data["comando_confirmado"] = None

        # ✅ agenda timeout em 600 segundos (20 minutos)
        context.job_queue.run_once(
            fluxo_timeout,
            1200,
            chat_id=update.effective_chat.id,
            name="timeout_adicionar"
        )

        return

    # === Receber dados do novo pedido ===
    if context.user_data.get("esperando_pedido"):
        # 👉 Intercepta "sim" logo no início
        if texto.lower().strip() in ["sim", "s", "yes", "y"]:
            if context.user_data.get("comando_confirmado"):
                await enviar_mensagem_segura(
                    context.bot, update.effective_chat.id,
                    "Acho que quer trocar para outro comando…"
                )
                # 🔥 dispara o comando confirmado
                await mensagens(update, context)
                context.user_data["comando_confirmado"] = None  # 🔧 resetar para não loopar
            return

        # 👉 fluxo normal de novo pedido
        dados = interpretar_mensagem(texto)
        linha, codigo_pedido, status_atual = gravar_excel(dados, ARQUIVO, ABA)
        if linha:
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                f"✅ Pedido {codigo_pedido} adicionado ao sistema! | Status: {status_atual}"
            )
        else:
            await enviar_mensagem_segura(
                context.bot,
                update.effective_chat.id,
                "⚠️ Não foi possível gravar o pedido. Verifique os dados enviados."
            )
        context.user_data["esperando_pedido"] = False
        await enviar_menu_comandos(update, context)
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

    # === Fallback para comandos não reconhecidos ===
    await verificar_usuario_fluxo(update, context)
    comando_sugerido = comando_mais_proximo(texto)

    if context.user_data.get("esperando_confirmacao_comando"):
        resposta = texto.lower().strip()
        if resposta in ["sim", "s", "yes", "y"]:
            comando = context.user_data["esperando_confirmacao_comando"]
            context.user_data["comando_confirmado"] = comando
            context.user_data["esperando_confirmacao_comando"] = None

            # 🔧 Limpa todos os estados de fluxo para não conflitar
            for flag in [
                "esperando_coluna",
                "aguardando_confirmacao_atualizacao",
                "esperando_coluna_edicao",
                "aguardando_confirmacao_edicao",
                "esperando_pedido",
                "esperando_cancelamento"
            ]:
                context.user_data[flag] = False

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

def formatar_valor_excel(valor):
    """Formata valores vindos do Excel para exibição amigável."""
    if not valor:
        return "(vazio)"

    # Se for datetime → dd/mm/yyyy
    if isinstance(valor, datetime.datetime):
        return valor.strftime("%d/%m/%Y")

    # Se for float representando horário (fração do dia)
    if isinstance(valor, float) and 0 <= valor < 1:
        minutos_totais = int(round(valor * 24 * 60))
        horas, minutos = divmod(minutos_totais, 60)
        return f"{horas:02d}:{minutos:02d}"

    # Se for float inteiro → mostra sem .0
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))

    return str(valor)


# === Callback: Selecionar pedido (ATUALIZAR) ===
async def selecionar_pedido(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    codigo = query.data

    # guarda no contexto
    context.user_data["codigo_pedido"] = codigo
    context.user_data["esperando_coluna"] = True

    pythoncom.CoInitialize()  # inicializa COM nesta thread
    excel = None
    try:
        # Abre Excel e busca a linha do pedido
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(ARQUIVO)
        ws = wb.Sheets(ABA)

        row = None
        for r in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{r}").Value).strip() if ws.Range(f"AX{r}").Value else ""
            if valor_ax == codigo:
                row = r
                break

        if not row:
            await query.edit_message_text(text=f"⚠️ Pedido {codigo} não encontrado.")
            wb.Close(SaveChanges=False)
            excel.Quit()
            return

        # Mapa de colunas
        mapa_vazias = {
            "E": "Data","F": "Horário","G": "Nome","H": "Quantidade","I": "Unidade Medida","J": "Produto",
            "L": "Sabor Massa","M": "Preparação Massa","N": "Sabor Recheio","O": "Preparação Recheio",
            "P": "Descrição Adicional","Q": "Valor Adicional","R": "Coleta Detalhes adicionais",
            "S": "Endereco","U": "Entrega","V": "Pagamento","W": "Recebimento foto com pedido",
            "X": "Publicação Instagram","Y": "Preço"
        }

        # ✅ Só lista colunas realmente vazias e salva ordem
        lista_nomes = [
            nome for col, nome in mapa_vazias.items()
            if not ws.Range(f"{col}{row}").Value or str(ws.Range(f"{col}{row}").Value).strip() == ""
        ]
        context.user_data["lista_colunas_vazias"] = lista_nomes
        context.user_data["colunas_vazias"] = {
            nome: col for col, nome in mapa_vazias.items()
            if not ws.Range(f"{col}{row}").Value or str(ws.Range(f"{col}{row}").Value).strip() == ""
        }

        lista_formatada = "\n".join([f"{i+1}: {nome} → (vazio)" for i, nome in enumerate(lista_nomes)])


        data_entrega = formatar_valor_excel(ws.Range(f"E{row}").Value)
        cliente = formatar_valor_excel(ws.Range(f"G{row}").Value)
        status = ws.Range(f"D{row}").Value if ws.Range(f"D{row}").Value else "Sem status"
        valor_total = limpar_valor_monetario(ws.Range(f"Y{row}").Value)
        valor_adicional = limpar_valor_monetario(ws.Range(f"Q{row}").Value)
        etiqueta = calcular_etiqueta(valor_total, valor_adicional)

        wb.Close(SaveChanges=False)
        excel.Quit()

        await query.edit_message_text(
            text=f"📌 Você selecionou o pedido {codigo}\n"
                 f"Data: {data_entrega if data_entrega else '(vazio)'}\n"
                 f"Cliente: {cliente if cliente else '(vazio)'}\n"
                 f"Status atual: {status}\n"
                 f"Pontuação fidelidade: {etiqueta}\n\n"
                 f"As colunas disponíveis para atualização são:\n{lista_formatada}\n\n"
                 "👉 Responda no formato: Número: Valor ou NomeColuna: Valor\n"
                 "👉 Digite 'fim' quando terminar para confirmar todas as alterações."
        )

        # agenda timeout em 180 segundos (3 minutos)
        context.job_queue.run_once(
            fluxo_timeout,
            180,
            chat_id=update.effective_chat.id,
            name=f"timeout_atualizar_{codigo}"
        )
    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()  # finaliza COM

# === Callback: Selecionar pedido para edição (EDITAR) ===
async def selecionar_pedido_edicao(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    codigo = query.data.replace("editar:", "")

    pythoncom.CoInitialize()  # inicializa COM nesta thread
    excel = None
    try:
        # Abre Excel e busca a linha do pedido
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(ARQUIVO)
        ws = wb.Sheets(ABA)

        row = None
        for r in range(2, ws.UsedRange.Rows.Count + 1):
            valor_ax = str(ws.Range(f"AX{r}").Value).strip() if ws.Range(f"AX{r}").Value else ""
            if valor_ax == codigo:
                row = r
                break

        if not row:
            await query.edit_message_text(text=f"⚠️ Pedido {codigo} não encontrado.")
            wb.Close(SaveChanges=False)
            excel.Quit()
            return

        # Mapa de colunas
        mapa_edicao = {
            "E": "Data","F": "Horário","G": "Nome","H": "Quantidade","I": "Unidade Medida","J": "Produto",
            "L": "Sabor Massa","M": "Preparação Massa","N": "Sabor Recheio","O": "Preparação Recheio",
            "P": "Descrição Adicional","Q": "Valor Adicional","R": "Coleta Detalhes adicionais",
            "S": "Endereco","U": "Entrega","V": "Pagamento","W": "Recebimento foto com pedido",
            "X": "Publicação Instagram","Y": "Preço"
        }
        context.user_data["mapa_edicao"] = mapa_edicao
        context.user_data["codigo_edicao"] = codigo
        context.user_data["esperando_coluna_edicao"] = True

        # Captura valores atuais com formatação
        lista_formatada = "\n".join([
            f"{i+1}: {nome} → {formatar_valor_excel(ws.Range(f'{col}{row}').Value)}"
            for i, (col, nome) in enumerate(mapa_edicao.items())
        ])

        # Dados principais do pedido com formatação
        data_entrega = formatar_valor_excel(ws.Range(f"E{row}").Value)
        cliente = formatar_valor_excel(ws.Range(f"G{row}").Value)
        status = ws.Range(f"D{row}").Value if ws.Range(f"D{row}").Value else "Sem status"
        valor_total = limpar_valor_monetario(ws.Range(f"Y{row}").Value)
        valor_adicional = limpar_valor_monetario(ws.Range(f"Q{row}").Value)
        etiqueta = calcular_etiqueta(valor_total, valor_adicional)

        wb.Close(SaveChanges=False)
        excel.Quit()

        await query.edit_message_text(
            text=f"📌 Você selecionou o pedido {codigo}\n"
                 f"Data: {data_entrega if data_entrega else '(vazio)'}\n"
                 f"Cliente: {cliente if cliente else '(vazio)'}\n"
                 f"Status atual: {status}\n"
                 f"Pontuação fidelidade: {etiqueta}\n\n"
                 f"As colunas disponíveis para edição são:\n{lista_formatada}\n\n"
                 "👉 Responda no formato: Número: Valor ou NomeColuna: Valor\n"
                 "👉 Digite 'fim' quando terminar para confirmar todas as alterações."
        )


        # agenda timeout em 180 segundos (3 minutos)
        context.job_queue.run_once(
            fluxo_timeout,
            180,
            chat_id=update.effective_chat.id,
            name=f"timeout_editar_{codigo}"
        )
    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()  # finaliza COM
        
# === Callback: Confirmar entrega ===
async def confirmar_entrega(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    codigo = query.data.replace("entregar:", "")

    keyboard = [
        [InlineKeyboardButton("✅ Sim, entregue", callback_data=f"confirmar_entrega:{codigo}")],
        [InlineKeyboardButton("❌ Não", callback_data="cancelar_entrega")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    # edita a mensagem diretamente, sem query.answer()
    await query.edit_message_text(
        text=f"O pedido {codigo} acabou de ser entregue?\nConfirme abaixo 👇",
        reply_markup=reply_markup
    )

    # Timeout de 5 minutos (300 segundos)
    context.job_queue.run_once(
        fluxo_timeout,
        300,
        chat_id=update.effective_chat.id,
        name=f"timeout_entrega_{codigo}"
    )

# === Callback: Marcar como entregue ===
async def marcar_entregue(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    codigo = query.data.replace("confirmar_entrega:", "")

    sucesso, status = atualizar_coluna(codigo, "U", "Entregue! 🐱", ARQUIVO, ABA)
    if sucesso:
        await query.edit_message_text(f"✅ Pedido {codigo} marcado como Entregue! 🐱")
    else:
        await query.edit_message_text(f"⚠️ Não foi possível marcar o pedido {codigo} como entregue.")

# === Callback: Cancelar entrega ===
async def cancelar_entrega(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text("❌ Cancelamento da entrega confirmado. Nenhuma alteração foi feita.")


# === Handler /start ===
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logging.info("Comando /start recebido")
    await enviar_menu_comandos(update, context)


# === Mensagem inicial automática via post_init ===
async def avisar_inicio(app: Application):
    # Mensagem inicial no grupo
    await enviar_mensagem_segura(
        app.bot,
        CHAT_ID,
        "🤖 Bot ativo! Estou aguardando suas interações para trabalharmos juntos!\n"
        "Aguarde enquanto eu ajusto alguns detalhes para trabalhar na confeitaria!🍰"
    )

    pendentes = []
    excel = None
    pythoncom.CoInitialize()   # ✅ inicializa COM nesta thread
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(ARQUIVO)
        ws = wb.Sheets(ABA)

        # Verifica pagamentos pendentes
        for row in range(2, ws.UsedRange.Rows.Count + 1):
            codigo = str(ws.Range(f"AX{row}").Value).strip() if ws.Range(f"AX{row}").Value else ""
            status = str(ws.Range(f"D{row}").Value).strip() if ws.Range(f"D{row}").Value else ""
            cliente = str(ws.Range(f"G{row}").Value).strip() if ws.Range(f"G{row}").Value else ""

            if "pagamento" in status.lower():
                pendentes.append((codigo, cliente))

        wb.Close(SaveChanges=False)
    except Exception as e:
        logging.error(f"Erro ao acessar Excel em avisar_inicio: {e}", exc_info=True)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()   # ✅ finaliza COM para liberar recursos

    # Se houver pendências, envia aviso
    if pendentes:
        aviso = (
            "⚠️ Identifiquei ausência de pagamentos dos pedidos! Não esqueça de cobrar as pendências!\n"
            "Posso te encaminhar o texto pronto para enviar aos nossos clientes, "
            "basta solicitar com o comando \"Pagamentos pendentes\""
        )
        await enviar_mensagem_segura(app.bot, CHAT_ID, aviso)

    # Menu inicial
    await enviar_menu_comandos_chat(app.bot, CHAT_ID)



# === Mensagem de despedida automática via post_stop ===
async def enviar_despedida(app: Application):
    mensagem = (
        "🤖 Hora da soneca! 😴 Vou buscar as últimas mensagens enviadas no grupo "
        "e continuarei o trabalho assim que o bot_continuo for colocado para rodar novamente na máquina principal."
    )
    await enviar_mensagem_segura(app.bot, CHAT_ID, mensagem)
    logging.info("Mensagem de despedida enviada.")


def main():
    verificar_excel_aberto()
    timeouts = [15, 30, 60]

    for t in timeouts:
        try:
            logging.info(f"🔄 Tentando iniciar bot com connect_timeout={t}s e read_timeout={t}s...")

            app = (
                Application.builder()
                .token(TOKEN)
                .post_init(avisar_inicio)
                .post_stop(enviar_despedida)
                .connect_timeout(t)
                .read_timeout(t)
                .build()
            )

            # Handlers
            app.add_handler(CommandHandler("start", fluxo_inicial))
            app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, mensagens))
            app.add_handler(CallbackQueryHandler(selecionar_pedido, pattern="^[^:]+$"))  # atualização
            app.add_handler(CallbackQueryHandler(selecionar_pedido_edicao, pattern="^editar:"))  # edição
            app.add_handler(CallbackQueryHandler(confirmar_entrega, pattern="^entregar:"))
            app.add_handler(CallbackQueryHandler(marcar_entregue, pattern="^confirmar_entrega:"))
            app.add_handler(CallbackQueryHandler(cancelar_entrega, pattern="^cancelar_entrega$"))

            logging.info("✅ Bot rodando... Ctrl+C para parar.")
            app.run_polling(
                poll_interval=3,
                allowed_updates=Update.ALL_TYPES
            )
            break

        except telegram.error.NetworkError as e:
            logging.error(f"❌ Erro de rede com timeout={t}s: {e}")
            logging.info("⏳ Tentando novamente com timeout maior...")
            time.sleep(5)
            continue

    else:
        logging.critical("🚨 Não foi possível iniciar o bot com nenhum dos timeouts configurados.")


if __name__ == "__main__":
    main()
