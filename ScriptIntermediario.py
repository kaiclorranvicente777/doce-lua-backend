import requests
import datetime
import logging
import subprocess

# Importa funções e configs do bot_continuo
from bot_continuo import interpretar_mensagem, gravar_excel, ARQUIVO, ABA, TOKEN, CHAT_ID

# === CONFIGURAÇÕES ===
GET_UPDATES_URL = f"https://api.telegram.org/bot{TOKEN}/getUpdates"
SEND_MESSAGE_URL = f"https://api.telegram.org/bot{TOKEN}/sendMessage"

logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

def registrar_log(msg):
    logging.info(msg)

def buscar_mensagens():
    """Busca histórico de mensagens via Telegram e filtra por chat_id."""
    try:
        resp = requests.get(GET_UPDATES_URL, timeout=15)
        dados = resp.json()
        mensagens = []
        for update in dados.get("result", []):
            msg = update.get("message", {})
            if not msg:
                continue
            chat = msg.get("chat", {})
            chat_id = chat.get("id")
            texto = msg.get("text", "")
            data = msg.get("date", 0)

            if chat_id == CHAT_ID and texto:
                mensagens.append((texto, datetime.datetime.fromtimestamp(data)))
        mensagens = sorted(list(set(mensagens)), key=lambda x: x[1])
        return mensagens
    except Exception as e:
        registrar_log(f"Erro ao buscar mensagens: {e}")
        return []

def filtrar_intervalo_pendente(mensagens):
    """Retorna mensagens entre a última despedida e a saudação atual."""
    saudacao = None
    despedida = None

    for texto, data in mensagens:
        if "Pedidos em Aberto - Sistema Central Doce Lua" in texto:
            saudacao = data
        if "🤖 Hora da soneca!" in texto:
            despedida = data

    if saudacao and despedida and despedida < saudacao:
        return [(t, d) for (t, d) in mensagens if despedida < d < saudacao]
    return []

def extrair_pedidos(mensagens_intervalo):
    """Agrupa cada 'Adicionar pedido ao sistema' com suas mensagens subsequentes de detalhes."""
    pedidos = []
    coletando = False
    buffer = []

    stop_markers = {
        "Pedidos em Aberto - Sistema Central Doce Lua",
        "🤖 Hora da soneca!"
    }

    for texto, _data in mensagens_intervalo:
        t_lower = texto.lower()
        if "adicionar pedido ao sistema" in t_lower:
            if buffer:
                pedidos.append("\n".join(buffer))
                buffer = []
            coletando = True
            buffer.append(texto)
        elif coletando:
            # Evita incluir novos marcadores de sistema no mesmo bloco
            if any(m in texto for m in stop_markers):
                # fecha bloco atual ao encontrar um marcador
                if buffer:
                    pedidos.append("\n".join(buffer))
                    buffer = []
                coletando = False
                continue
            buffer.append(texto)

    if buffer:
        pedidos.append("\n".join(buffer))

    return pedidos

def notificar(texto):
    for tentativa in range(2):
        try:
            requests.post(SEND_MESSAGE_URL, data={'chat_id': CHAT_ID, 'text': texto}, timeout=10)
            return
        except Exception as e:
            if tentativa == 1:
                registrar_log(f"Falha ao notificar: {e}")

def processar_pedidos_pendentes():
    registrar_log("Buscando mensagens para identificar pedidos pendentes...")
    todos = buscar_mensagens()
    intervalo = filtrar_intervalo_pendente(todos)
    pedidos = extrair_pedidos(intervalo)

    if not pedidos:
        registrar_log("Nenhum 'Adicionar pedido ao sistema' encontrado no intervalo.")
        return False

    notificar("Processando pedidos pendentes do intervalo 🤖")

    sucesso_total = 0
    falhas = 0

    for bloco in pedidos:
        dados = interpretar_mensagem(bloco)

        # Validação mínima de conteúdo antes de gravar
        essenciais = any([
            bool(dados.get("Nome")),
            bool(dados.get("Produto")),
            bool(dados.get("Quantidade")),
            bool(dados.get("Data")),
        ])
        if not essenciais:
            falhas += 1
            registrar_log(f"Bloco sem dados essenciais, ignorado:\n{bloco}")
            continue

        linha, codigo, status = gravar_excel(dados, ARQUIVO, ABA)
        if linha:
            sucesso_total += 1
            registrar_log(f"Pedido {codigo} adicionado | Status: {status}")
            notificar(f"Pedido {codigo} adicionado automaticamente do intervalo ✅ | Status: {status}")
        else:
            falhas += 1
            registrar_log(f"Falha ao gravar pedido pendente. Bloco:\n{bloco}")

    registrar_log(f"Pedidos processados com sucesso: {sucesso_total} de {len(pedidos)} | Falhas: {falhas}")
    notificar(f"Resumo: {sucesso_total} pedidos adicionados automaticamente. {falhas} falhas.")
    return sucesso_total > 0

def main():
    processar_pedidos_pendentes()
    try:
        registrar_log("Chamando bot_continuo.py...")
        subprocess.run(["python", r"C:\Users\KICEGA~1\OneDrive\Documentos\bot_continuo.py"], check=True)
        registrar_log("bot_continuo.py executado com sucesso.")
    except Exception as e:
        registrar_log(f"Erro ao chamar bot_continuo.py: {e}")

if __name__ == "__main__":
    main()
