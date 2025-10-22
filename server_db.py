import os
import re
from datetime import datetime
from typing import Literal, Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from dotenv import load_dotenv

from db_backend import DB as _DB


load_dotenv()
ALLOWED_ORIGINS = [o.strip() for o in os.getenv("ALLOWED_ORIGINS", "").split(",") if o.strip()]

if not _DB.is_ready():
    raise RuntimeError("DATABASE_URL não configurado ou indisponível para server_db")
try:
    _DB.init_schema_portable()
except Exception:
    try:
        _DB.init_schema()
    except Exception:
        pass

# garante administrador padrão
try:
    _DB.ensure_admin(
        usuario="Leandro",
        nome="Leandro Rodrigo da Silva Rosa",
        email="leandro.rosa@audaces.com",
        setor="ADM",
        cargo="Supervisor",
        senha="1234",
    )
except Exception:
    pass


def pt(n: float) -> str:
    return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def data_tokens(now: Optional[datetime] = None):
    d = now or datetime.now()
    return {
        "data_compacta": d.strftime("%d%m%Y"),
        "combinado": d.strftime("%d/%m/%Y %H:%M:%S"),
    }


def sigla_tipo(tipo_servico: str) -> Literal["IM", "DG"]:
    return "IM" if tipo_servico.lower().startswith("imp") else "DG"


def validar_email(email: str) -> bool:
    return re.match(r'^[\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})*$', email or "") is not None


def formatar_cnpj(cnpj: str) -> str:
    d = re.sub(r'\D', '', cnpj or "")[:14]
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}" if len(d)==14 else cnpj


class OrcamentoIn(BaseModel):
    tipo_servico: Literal["Impressão", "Digitalização"]
    cliente: str
    cnpj: str
    email: str
    status: Literal["Sem desconto", "Novo", "Ativo", "Inativo"]
    unidade: Literal["Centímetros", "Metro"]
    quantidade: str


class OrcamentoOut(BaseModel):
    id_orcamento: str
    data_hora: str
    tipo_servico: str
    cliente: str
    cnpj: str
    email: str
    status: str
    quantidade: str
    unidade: str
    metros: str
    preco_por_metro: str
    valor_total: str


app = FastAPI(title="Integração Orçamento API (DB)")

if ALLOWED_ORIGINS:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=ALLOWED_ORIGINS,
        allow_methods=["*"],
        allow_headers=["*"],
    )

class UsuarioIn(BaseModel):
    usuario: str
    nome: str
    email: str
    setor: str
    cargo: str
    senha: str | None = None
    is_admin: bool = False
    permissoes: str | None = None


class LoginIn(BaseModel):
    usuario: str
    senha: str


@app.get("/api/proximo-id")
async def proximo_id(tipo_servico: Literal["Impressão", "Digitalização"]):
    sigla = sigla_tipo(tipo_servico)
    dtok = data_tokens()
    prefix = f"OR-{sigla}"
    rows = _DB.list_orcamentos_excel()
    count = 0
    for r in rows:
        try:
            v = str(r.get("ID Orçamento") or r.get("ID") or "")
            if v.startswith(prefix):
                count += 1
        except Exception:
            pass
    seq = count + 1
    return {"id": f"{prefix}{seq}{dtok['data_compacta']}"}


@app.post("/api/orcamentos", response_model=OrcamentoOut)
async def criar_orcamento(body: OrcamentoIn):
    if not validar_email(body.email):
        raise HTTPException(400, "E-mail inválido")
    d = re.sub(r'\D', '', body.cnpj)
    if len(d) != 14:
        raise HTTPException(400, "CNPJ deve ter 14 dígitos")
    try:
        qtd = float((body.quantidade or "").replace(".", "").replace(",", "."))
    except Exception:
        qtd = 0.0
    if qtd <= 0:
        raise HTTPException(400, "Quantidade deve ser > 0")

    sigla = sigla_tipo(body.tipo_servico)
    dtok = data_tokens()
    prefix = f"OR-{sigla}"
    rows = _DB.list_orcamentos_excel()
    count = 0
    for r in rows:
        try:
            v = str(r.get("ID Orçamento") or r.get("ID") or "")
            if v.startswith(prefix):
                count += 1
        except Exception:
            pass
    seq = count + 1
    id_orc = f"{prefix}{seq}{dtok['data_compacta']}"

    metros = (qtd/100.0) if body.unidade == "Centímetros" else qtd
    preco  = 8.00 if body.status in ["Novo", "Ativo"] else 8.50
    total  = metros * preco
    cnpj_fmt = formatar_cnpj(body.cnpj)

    row_by_db = {
        "id_orcamento": id_orc,
        "data_hora": dtok["combinado"],
        "tipo_servico": body.tipo_servico,
        "cliente_label": "Razão Social",
        "cliente_valor": body.cliente,
        "documento": "CNPJ",
        "cnpj_cpf": cnpj_fmt,
        "email": body.email,
        "vendedor": "",
        "desconto": body.status,
        "quantidade": pt(qtd),
        "unidade": body.unidade,
        "metros": pt(metros),
        "preco_por_metro": pt(preco),
        "forma_pagamento": "",
        "valor_total": pt(total),
    }
    dados_excel = { _DB.REV_ORC.get(k, k): v for k, v in row_by_db.items() }
    _DB.salvar_orcamento(dados_excel)

    return OrcamentoOut(
        id_orcamento=id_orc,
        data_hora=dtok["combinado"],
        tipo_servico=body.tipo_servico,
        cliente=body.cliente,
        cnpj=cnpj_fmt,
        email=body.email,
        status=body.status,
        quantidade=pt(qtd),
        unidade=body.unidade,
        metros=pt(metros),
        preco_por_metro=pt(preco),
        valor_total=pt(total),
    )


@app.get("/api/orcamentos")
async def listar_orcamentos(
    id: Optional[str] = None,
    cnpj: Optional[str] = None,
    vendedor: Optional[str] = None,
    start: Optional[str] = None,
    end: Optional[str] = None,
):
    if id:
        d = _DB.get_orcamento_by_id(id)
        rows = [d] if d else []
        return {"count": len(rows), "rows": rows}
    rows = _DB.list_orcamentos_excel(start=start, end=end, vendedor=vendedor, cnpj_digits=cnpj)
    return {"count": len(rows), "rows": rows}


@app.get("/api/usuarios")
async def listar_usuarios():
    try:
        return {"rows": _DB.list_usuarios()}
    except Exception as ex:
        raise HTTPException(500, f"Erro ao listar usuários: {ex}")


@app.post("/api/usuarios")
async def criar_atualizar_usuario(body: UsuarioIn):
    try:
        _DB.upsert_usuario(body.dict())
        return {"ok": True}
    except Exception as ex:
        raise HTTPException(500, f"Erro ao salvar usuário: {ex}")


@app.post("/api/login")
async def login(body: LoginIn):
    u = _DB.check_login(body.usuario, body.senha)
    if not u:
        raise HTTPException(401, "Usuário ou senha inválidos")
    return {
        "usuario": u.get("usuario"),
        "nome": u.get("nome"),
        "email": u.get("email"),
        "is_admin": bool(u.get("is_admin")),
        "permissoes": u.get("permissoes") or "*",
    }


@app.post("/api/usuarios/change-senha")
async def change_senha(usuario: str, senha_atual: str | None = None, senha_nova: str = "", force: bool = False):
    if not force:
        u = _DB.check_login(usuario, senha_atual or "")
        if not u:
            raise HTTPException(401, "Senha atual inválida")
    _DB.set_password(usuario, senha_nova)
    return {"ok": True}


@app.get("/api/orcamentos/{orc_id}")
async def obter_orcamento(orc_id: str):
    d = _DB.get_orcamento_by_id(orc_id)
    if d:
        return d
    raise HTTPException(404, "Orçamento não encontrado")


@app.get("/api/cadastros")
async def listar_cadastros(
    cnpj: Optional[str] = None,
    vendedor: Optional[str] = None,
    start: Optional[str] = None,
    end: Optional[str] = None,
):
    rows = _DB.list_cadastros_excel(start=start, end=end, vendedor=vendedor, cnpj_digits=cnpj)
    return {"count": len(rows), "rows": rows}


@app.get("/api/pedidos")
async def listar_pedidos(
    cnpj: Optional[str] = None,
    vendedor: Optional[str] = None,
    start: Optional[str] = None,
    end: Optional[str] = None,
):
    rows = _DB.list_pedidos_excel(start=start, end=end, vendedor=vendedor, cnpj_digits=cnpj)
    return {"count": len(rows), "rows": rows}

@app.get("/api/info")
async def api_info():
    return {"storage": "db", "backend": "postgres", "db_ready": True}


# ====== Criação/Atualização (DB) ======
class CadastroDBIn(BaseModel):
    documento: Optional[str] = None
    cnpj_cpf: str
    razao_social_nome: Optional[str] = None
    nome_fantasia: Optional[str] = None
    contato: Optional[str] = None
    email_cnpj: Optional[str] = None
    email_manual: Optional[str] = None
    cep: Optional[str] = None
    endereco: Optional[str] = None
    numero: Optional[str] = None
    complemento: Optional[str] = None
    bairro: Optional[str] = None
    municipio: Optional[str] = None
    uf: Optional[str] = None
    entrega_cep: Optional[str] = None
    entrega_endereco: Optional[str] = None
    entrega_numero: Optional[str] = None
    entrega_complemento: Optional[str] = None
    entrega_bairro: Optional[str] = None
    entrega_municipio: Optional[str] = None
    entrega_uf: Optional[str] = None
    desconto_duracao: Optional[str] = None
    desconto_unidade: Optional[str] = None
    telefone1: Optional[str] = None
    telefone2: Optional[str] = None
    vendedor: Optional[str] = None


@app.post("/api/cadastros")
async def criar_ou_atualizar_cadastro(body: CadastroDBIn):
    digits = re.sub(r"\D", "", body.cnpj_cpf or "")
    if len(digits) not in (11, 14):
        raise HTTPException(400, "CNPJ/CPF deve ter 11 ou 14 dígitos")
    dados_db = body.dict()
    # Preenche documento se ausente
    if not dados_db.get("documento"):
        dados_db["documento"] = "CNPJ" if len(digits) == 14 else "CPF"
    # Converte para labels Excel usando REV_CAD
    dados_excel = { _DB.REV_CAD.get(k, k): v for k, v in dados_db.items() if v is not None }
    try:
        _DB.salvar_cadastro(dados_excel)
        return {"ok": True}
    except Exception as ex:
        raise HTTPException(500, f"Erro ao salvar cadastro: {ex}")


class PedidoDBIn(BaseModel):
    id: str
    pedido: Optional[int] = None
    tipo_servico: Optional[str] = None
    status_cliente: Optional[str] = None
    quantidade_m: Optional[str] = None
    valor_unitario: Optional[str] = None
    valor_total: Optional[str] = None
    data_hora_criacao: Optional[str] = None
    id_orcamento: Optional[str] = None
    documento: Optional[str] = None
    cnpj_cpf: str
    cliente: Optional[str] = None
    vendedor: Optional[str] = None
    forma_pgto_orcamento: Optional[str] = None
    forma_pgto_contrato: Optional[str] = None
    pct_comissao_vendedor: Optional[str] = None
    valor_comissao_vendedor: Optional[str] = None
    pct_comissao_adm: Optional[str] = None
    valor_comissao_adm: Optional[str] = None


@app.post("/api/pedidos")
async def criar_pedido(body: PedidoDBIn):
    digits = re.sub(r"\D", "", body.cnpj_cpf or "")
    if len(digits) not in (11, 14):
        raise HTTPException(400, "CNPJ/CPF deve ter 11 ou 14 dígitos")
    dados_db = body.dict()
    # Preenche documento se ausente
    if not dados_db.get("documento"):
        dados_db["documento"] = "CNPJ" if len(digits) == 14 else "CPF"
    dados_excel = { _DB.REV_PED.get(k, k): v for k, v in dados_db.items() if v is not None }
    try:
        _DB.salvar_pedido(dados_excel)
        return {"ok": True}
    except Exception as ex:
        raise HTTPException(500, f"Erro ao salvar pedido: {ex}")
