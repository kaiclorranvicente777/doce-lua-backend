import re
from openpyxl import load_workbook

# === Função para extrair os campos da mensagem ===
def extrair_campos(texto):
    dados = {
        "Data": "Não informado",
        "Horario": "Não informado",
        "Nome": "Não informado",
        "Quantidade": "Não informado",
        "Produto": "Não informado",
        "SaborMassa": "Não informado",
        "PrepMassa": "Não informado",
        "SaborRecheio": "Não informado",
        "PrepRecheio": "Não informado",
        "DescricaoAdicional": "Não informado",
        "ValorAdicional": "Não informado",
        "ColetaAdicional": "Não informado",
        "Endereco": "Não informado",
        "Contato": "Não informado",
        "Entrega": "Não informado",
        "Pagamento": "Não informado",
        "RecebimentoFoto": "Não informado",
        "PublicacaoInstagram": "Não informado",
        "Preco": "Não informado"
    }

    # Regex simples para extrair dados
    item = re.search(r"Item:\s*(.+)", texto)
    if item:
        dados["Produto"] = item.group(1).strip()
        qtd = re.search(r"(\d+)\s*Kg", item.group(1))
        if qtd:
            dados["Quantidade"] = qtd.group(1)

    massa = re.search(r"Sabor Massa:\s*(.+)", texto)
    if massa:
        dados["SaborMassa"] = massa.group(1).strip()

    recheio = re.search(r"Sabor Cobertura/Recheio:\s*(.+)", texto)
    if recheio:
        dados["SaborRecheio"] = recheio.group(1).strip()

    detalhes = re.search(r"Detalhes Adicionais:\s*(.+)", texto)
    if detalhes:
        dados["DescricaoAdicional"] = detalhes.group(1).strip()

    endereco = re.search(r"Endereço.*:\s*(.+)", texto)
    if endereco:
        dados["Endereco"] = endereco.group(1).strip()

    data = re.search(r"Data:\s*(\d{2}/\d{2})", texto)
    if data:
        dados["Data"] = data.group(1)

    horario = re.search(r"Horário:\s*(\d{2}:\d{2})", texto)
    if horario:
        dados["Horario"] = horario.group(1)

    valor = re.search(r"Valor total pedido:\s*R\$\s*([\d,]+)", texto)
    if valor:
        dados["Preco"] = valor.group(1)

    return dados

# === Função para gravar os dados no Excel ===
def gravar_excel(dados, arquivo, aba="Pedidos-via-bot"):
    wb = load_workbook(arquivo)
    ws = wb[aba]

    # Adiciona na próxima linha disponível
    ws.append([
        dados["Data"], dados["Horario"], dados["Nome"], dados["Quantidade"], dados["Produto"],
        dados["SaborMassa"], dados["PrepMassa"], dados["SaborRecheio"], dados["PrepRecheio"],
        dados["DescricaoAdicional"], dados["ValorAdicional"], dados["ColetaAdicional"],
        dados["Endereco"], dados["Contato"], dados["Entrega"], dados["Pagamento"],
        dados["RecebimentoFoto"], dados["PublicacaoInstagram"], dados["Preco"]
    ])

    wb.save(arquivo)
    print("Pedido gravado com sucesso na aba 'Pedidos-via-bot'!")

# === TESTE COM UMA MENSAGEM DE EXEMPLO ===
mensagem = """
☽ Item: 4 Kg Bolo Confeitado - Gourmet R$70xKg
☽ Sabor Massa: Baunilha
☽ Sabor Cobertura/Recheio: Ninho com Pêssego
☽ Detalhes Adicionais: Bolo de corte
☽ Endereço: Av. Espanha, 1240 - Vilarejo
☽ Data: 07/07 | Horário: 19:00
☽ Valor total pedido: R$1120
"""

dados = extrair_campos(mensagem)
gravar_excel(dados, r"C:\Users\Káic e Gabi\OneDrive\Documentos\Sistema.Central_Doce.Lua_11-2025.xlsx")
