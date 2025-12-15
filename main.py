from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware   # ✅ importa CORS
from pydantic import BaseModel
from typing import List, Dict
from bot_continuo import (
    listar_pedidos_abertos,
    gravar_excel,
    cancelar_pedido,
    listar_enderecos_do_dia,
    listar_colunas_vazias,
    atualizar_coluna,
    gerar_imagem_excel,
    interpretar_mensagem,
    IMAGEM_PATH   # ✅ importa o caminho da imagem gerada
)

app = FastAPI(
    title="API Doce Lua",
    description="Gerenciamento de pedidos Doce Lua",
    version="1.0"
)

# ✅ Middleware CORS para permitir acesso do Flutter Web
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # ou especifique ["http://localhost:8080", "https://otto-milanaise-tate.ngrok-free.dev"]
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

ARQUIVO = r"C:\Users\Káic e Gabi\OneDrive\Documentos\Sistema.Central_Doce.Lua_11-2025.xlsx"

# Modelo Pydantic para pedidos
class Pedido(BaseModel):
    Nome: str
    Produto: str
    Quantidade: str
    UnidadeMedida: str
    SaborMassa: str
    SaborRecheio: str
    DescricaoAdicional: str
    Endereco: str
    Entrega: str
    Pagamento: str
    Preco: str
    Data: str
    Horario: str

# Modelo para atualização múltipla
class Atualizacao(BaseModel):
    campos: Dict[str, str]

# Endpoint raiz
@app.get("/", response_model=Dict[str, str])
def root():
    return {"msg": "API Doce Lua funcionando!"}

# Listar pedidos abertos
@app.get("/pedidos", response_model=List[str])
def pedidos_abertos():
    return listar_pedidos_abertos(ARQUIVO)

# Adicionar pedido
@app.post("/novo", response_model=Dict[str, str])
def novo_pedido(dados: Pedido):
    linha, codigo, status = gravar_excel(dados.dict(), ARQUIVO)
    return {"linha": str(linha), "codigo": str(codigo), "status": str(status)}

# Adicionar pedido via texto bruto (usa interpretar_mensagem)
@app.post("/novo-texto", response_model=Dict[str, str])
def novo_pedido_texto(texto: str):
    dados = interpretar_mensagem(texto)
    linha, codigo, status = gravar_excel(dados, ARQUIVO)
    return {"linha": str(linha), "codigo": str(codigo), "status": str(status)}

# Cancelar pedido
@app.post("/cancelar/{codigo}", response_model=Dict[str, str])
def cancelar(codigo: str):
    ok, status = cancelar_pedido(codigo, ARQUIVO)
    return {"sucesso": str(ok), "status": str(status)}

# Listar endereços do dia
@app.get("/enderecos", response_model=List[Dict[str, str]])
def enderecos_do_dia():
    enderecos = listar_enderecos_do_dia(ARQUIVO)
    return [{"codigo": c, "cliente": cl, "endereco": e, "horario": h} for c, cl, e, h in enderecos]

# Listar colunas vazias de um pedido
@app.get("/pedido/{codigo}/colunas-vazias", response_model=Dict[str, str])
def colunas_vazias(codigo: str):
    return listar_colunas_vazias(codigo, ARQUIVO)

# Atualizar uma coluna específica
@app.put("/pedido/{codigo}/coluna", response_model=Dict[str, str])
def atualizar(codigo: str, coluna: str, valor: str):
    ok, status = atualizar_coluna(codigo, coluna, valor, ARQUIVO)
    return {"sucesso": str(ok), "status": str(status)}

# Atualizar várias colunas de uma vez
@app.put("/pedido/{codigo}", response_model=Dict[str, str])
def atualizar_varias(codigo: str, dados: Atualizacao):
    resultados = []
    for coluna, valor in dados.campos.items():
        ok, status = atualizar_coluna(codigo, coluna, valor, ARQUIVO)
        resultados.append(f"{coluna}: {valor} ({'ok' if ok else 'erro'})")
    return {"resultado": "; ".join(resultados), "status": str(status)}

# ✅ Gerar calendário e retornar a imagem
@app.get("/calendario")
def calendario():
    gerar_imagem_excel()
    return FileResponse(
        path=IMAGEM_PATH,
        media_type="image/png",
        filename="calendario.png"
    )

# Marcar pedido como entregue
@app.post("/pedido/{codigo}/entregar", response_model=Dict[str, str])
def marcar_entregue(codigo: str):
    ok, status = atualizar_coluna(codigo, "U", "Entregue! 🐱", ARQUIVO)
    return {"sucesso": str(ok), "status": str(status)}

# Listar pagamentos pendentes
@app.get("/pagamentos-pendentes", response_model=List[Dict[str, str]])
def pagamentos_pendentes():
    pedidos = listar_pedidos_abertos(ARQUIVO)
    pendentes = []
    for pedido in pedidos:
        if "pagamento" in pedido.lower():
            partes = pedido.split(" - ")
            pendentes.append({"codigo": partes[0], "detalhes": pedido})
    return pendentes
