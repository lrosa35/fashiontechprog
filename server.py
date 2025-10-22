import os, json, re, pickle, urllib.parse
from datetime import datetime
from typing import Literal, Optional, List, Dict
import time
import unicodedata

import httpx
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from dotenv import load_dotenv
import msal
import threading, socket, json

# ====== Config ======
load_dotenv()
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
EXCEL_REL_PATH = os.getenv("EXCEL_RELATIVE_PATH")  # ex: "01 LEANDRO/IMPRESSÃ•ES/Controle de ImpressÃµes.xlsx"
TABLE_NAME = os.getenv("EXCEL_TABLE_NAME", "Orcamentos")
ALLOWED_ORIGINS = [o.strip() for o in os.getenv("ALLOWED_ORIGINS", "").split(",") if o.strip()]
# Preferimos 'db' como padrão para ambientes de nuvem (evita exigir OneDrive/Graph na importação)
STORAGE_BACKEND = os.getenv("STORAGE_BACKEND", "db").strip().lower()

if STORAGE_BACKEND != "db":
    if not (TENANT_ID and CLIENT_ID and EXCEL_REL_PATH):
        raise RuntimeError("Configure TENANT_ID, CLIENT_ID e EXCEL_RELATIVE_PATH no .env")

SCOPES = ["Files.ReadWrite.All", "offline_access"]  # Delegated scopes
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# ====== Token cache (persistente) ======
CACHE_FILE = "token_cache.bin"

def _load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE, "r", encoding="utf-8").read())
    return cache

def _save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        open(CACHE_FILE, "w", encoding="utf-8").write(cache.serialize())

def acquire_token():
    cache = _load_cache()
    app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    # Tenta silencioso (usa cache/refresh)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]

    # Device Code (primeira vez)
    flow = app.initiate_device_flow(scopes=[f"https://graph.microsoft.com/{s}" for s in SCOPES])
    if "user_code" not in flow:
        raise RuntimeError("Falha ao iniciar Device Code Flow")
    print("\n=== Autorização necessÃ¡ria ===")
    print(flow["message"])  # abre https://microsoft.com/devicelogin e informe o cÃ³digo
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Erro ao obter token: {result.get('error_description')}")
    _save_cache(cache)
    return result["access_token"]

GRAPH = "https://graph.microsoft.com/v1.0"
_HTTP_CLIENT: httpx.AsyncClient | None = None
_EXCEL_ITEM_ID: str | None = None
_SESSION_ID: str | None = None
_SESSION_TS: float = 0.0
_SESSION_TTL_SECONDS = 300

# ====== DB opcional ======
try:
    from db_backend import DB as _DB
    _DB_READY = _DB.is_ready()
    if _DB_READY:
        try:
            _DB.init_schema()
        except Exception:
            try:
                _DB.init_schema_portable()
            except Exception:
                pass
except Exception:
    _DB_READY = False

# ====== Helpers Graph ======
async def graph_client(token: str):
    return httpx.AsyncClient(headers={"Authorization": f"Bearer {token}"}, timeout=30.0)

async def get_drive_item_id(token: str) -> str:
    # localiza o item pelo caminho no OneDrive do usuÃ¡rio logado
    encoded = urllib.parse.quote(EXCEL_REL_PATH)
    url = f"{GRAPH}/me/drive/root:/{encoded}"
    async with await graph_client(token) as client:
        r = await client.get(url)
        if r.status_code != 200:
            raise HTTPException(500, f"Não achei o arquivo no OneDrive ({r.text})")
        return r.json()["id"]

async def create_session(token: str, item_id: str) -> str:
    url = f"{GRAPH}/me/drive/items/{item_id}/workbook/createSession"
    async with await graph_client(token) as client:
        r = await client.post(url, json={"persistChanges": True})
        if r.status_code not in (200, 201):
            raise HTTPException(500, f"Erro ao criar sessão do Excel: {r.text}")
        return r.json()["id"]  # workbook-session-id

async def list_rows(token: str, item_id: str, session_id: str) -> list:
    url = f"{GRAPH}/me/drive/items/{item_id}/workbook/tables/{TABLE_NAME}/rows"
    async with await graph_client(token) as client:
        r = await client.get(url, headers={"workbook-session-id": session_id})
        if r.status_code != 200:
            raise HTTPException(500, f"Erro ao listar linhas: {r.text}")
        data = r.json()
        # Cada row tem "values": [[col1, col2, ...]]
        return [row["values"][0] for row in data.get("value", [])]

async def list_columns(token: str, item_id: str, session_id: str) -> List[str]:
    url = f"{GRAPH}/me/drive/items/{item_id}/workbook/tables/{TABLE_NAME}/columns"
    async with await graph_client(token) as client:
        r = await client.get(url, headers={"workbook-session-id": session_id})
        if r.status_code != 200:
            raise HTTPException(500, f"Erro ao listar colunas: {r.text}")
        data = r.json()
        cols = [c.get("name") for c in data.get("value", [])]
        return [str(x) if x is not None else "" for x in cols]

def _norm(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("/", "_").replace(" ", "_")
    return re.sub(r"[^a-z0-9_]+", "", s)

def _find_col(cols: List[str], candidates: List[str]) -> Optional[int]:
    norm_cols = [_norm(c) for c in cols]
    norm_cands = [_norm(x) for x in candidates]
    for i, nc in enumerate(norm_cols):
        if nc in norm_cands:
            return i
    return None

async def list_rows_dicts(token: str, item_id: str, session_id: str) -> List[Dict[str, str]]:
    cols = await list_columns(token, item_id, session_id)
    rows = await list_rows(token, item_id, session_id)
    out: List[Dict[str, str]] = []
    for r in rows:
        d = {}
        for i, name in enumerate(cols):
            d[name] = r[i] if i < len(r) else None
        out.append(d)
    return out

async def add_row(token: str, item_id: str, session_id: str, values: list):
    url = f"{GRAPH}/me/drive/items/{item_id}/workbook/tables/{TABLE_NAME}/rows/add"
    body = {"values": [values]}  # uma linha; pode enviar vÃ¡rias
    async with await graph_client(token) as client:
        r = await client.post(url, json=body, headers={"workbook-session-id": session_id})
        if r.status_code not in (200, 201):
            raise HTTPException(500, f"Erro ao inserir linha: {r.text}")
        return r.json()

# ====== DomÃ­nio (mesma regra do app) ======
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
    return re.match(r'^[\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})*$', email) is not None

def formatar_cnpj(cnpj: str) -> str:
    d = re.sub(r'\D', '', cnpj)[:14]
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}" if len(d)==14 else cnpj

def formatar_cpf(cpf: str) -> str:
    d = re.sub(r'\D', '', cpf)[:11]
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}" if len(d)==11 else cpf

def proximo_seq_por_rows(rows: list, prefixo: str) -> int:
    # rows: lista de linhas completas da tabela (inclui cabeÃ§alho? Graph rows jÃ¡ desconsidera cabeÃ§alho)
    # Assumindo col 0 = "ID OrÃ§amento"
    count = 0
    for r in rows:
        try:
            idv = str(r[0])
            if idv.startswith(prefixo):
                count += 1
        except Exception:
            raise HTTPException(500, "DB indisponível")
    return count + 1

# ====== FastAPI ======
class OrcamentoIn(BaseModel):
    tipo_servico: Literal["Impressão", "Digitalização"]
    cliente: str
    cnpj: str
    email: str
    status: Literal["Sem desconto", "Novo", "Ativo", "Inativo"]
    unidade: Literal["Centímetros", "Metro"]
    quantidade: str = Field(..., description="Em cm (inteiro) ou m (decimal). Usa vÃ­rgula ou ponto.")
    # Opcionais (UI Web): se informados, sobrescrevem cÃ¡lculos/salvam no registro
    vendedor: Optional[str] | None = None
    forma_pagamento: Optional[str] | None = None
    preco_por_metro_opc: Optional[str] | None = None
    metros_opc: Optional[str] | None = None

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

app = FastAPI(title="Integração Orçamento API")

if ALLOWED_ORIGINS:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=ALLOWED_ORIGINS,
        allow_methods=["*"],
        allow_headers=["*"],
    )

@app.get("/api/proximo-id")
async def proximo_id(tipo_servico: Literal["Impressao", "Digitalizacao"]):
    sigla = sigla_tipo(tipo_servico)
    dtok = data_tokens()
    prefix = f"OR-{sigla}"
    if STORAGE_BACKEND == "db" and _DB_READY:
        try:
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
        except Exception:
            pass
    token = acquire_token()
    item_id = await get_drive_item_id_cached(token)
    session_id = await get_session_id_cached(token, item_id)
    rows = await list_rows(token, item_id, session_id)
    seq = proximo_seq_por_rows(rows, prefix)
    return {"id": f"{prefix}{seq}{dtok['data_compacta']}"}
@app.post("/api/orcamentos", response_model=OrcamentoOut)
async def criar_orcamento(body: OrcamentoIn):
    # validaÃ§Ãµes
    if not validar_email(body.email):
        raise HTTPException(400, "E-mail inválido")
    d = re.sub(r'\D', '', body.cnpj)
    if len(d) not in (11, 14):
        raise HTTPException(400, "Documento deve ter 11 (CPF) ou 14 (CNPJ) dígitos")
    # quant
    qtd = float((body.quantidade or "").replace(",", "."))
    if qtd <= 0:
        raise HTTPException(400, "Quantidade deve ser > 0")

    sigla = sigla_tipo(body.tipo_servico)
    if STORAGE_BACKEND == "db" and _DB_READY:
        try:
            rows = _DB.list_orcamentos_excel()
        except Exception:
            rows = []
        seq = 1
        for r in rows:
            try:
                v = str(r.get("ID Orçamento") or r.get("ID") or "")
                if v.startswith(f"OR-{sigla}"):
                    seq += 1
            except Exception:
                pass
    else:
        token = acquire_token()
        item_id = await get_drive_item_id_cached(token)
        session_id = await get_session_id_cached(token, item_id)
        rows = await list_rows(token, item_id, session_id)
        seq = proximo_seq_por_rows(rows, f"OR-{sigla}")
    dtok = data_tokens()

    id_orc = f"OR-{sigla}{seq}{dtok['data_compacta']}"
    metros = (qtd/100.0) if body.unidade == "Centímetros" else qtd
    preco  = 8.00 if body.status in ["Novo", "Ativo"] else 8.50
    # Overrides opcionais vindos do formulÃ¡rio
    try:
        if body.metros_opc:
            metros = float(str(body.metros_opc).replace(",", "."))
    except Exception:
        pass
    try:
        if body.preco_por_metro_opc:
            preco = float(str(body.preco_por_metro_opc).replace(",", "."))
    except Exception:
        pass
    total  = metros * preco

    cnpj_fmt = formatar_cnpj(body.cnpj) if len(d)==14 else formatar_cpf(body.cnpj)

    # Monta linha nos MESMOS cabeÃ§alhos que vocÃª jÃ¡ usa:
    linha = [
        id_orc,
        dtok["combinado"],
        body.tipo_servico,
        body.cliente,
        cnpj_fmt,
        body.email,
        body.status,
        pt(qtd),
        body.unidade,
        pt(metros),
        pt(preco),
        pt(total),
    ]

    if STORAGE_BACKEND == "db" and _DB_READY:
        _DB.salvar_orcamento({
            "ID Orçamento": id_orc,
            "Data/Hora": dtok["combinado"],
            "Tipo de Serviço": body.tipo_servico,
            "CLIENTE (Etiqueta PDF)": ("Nome" if len(d)==11 else "Razão Social"),
            "CLIENTE (Valor)": body.cliente,
            "Documento": ("CPF" if len(d)==11 else "CNPJ"),
            "CNPJ/CPF": cnpj_fmt,
            "E-mail": body.email,
            "Vendedor": str(body.vendedor or ""),
            "Status": body.status,
            "Quantidade": pt(qtd),
            "Unidade": body.unidade,
            "Metros": pt(metros),
            "Preço por metro": pt(preco),
            "Forma de Pagamento": str(body.forma_pagamento or ""),
            "Valor Total": pt(total),
        })
    else:
        # Insere via Graph/Excel
        await add_row(token, item_id, session_id, linha)

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
    start: Optional[str] = None,  # dd/mm/yyyy
    end: Optional[str] = None,    # dd/mm/yyyy
):
    if STORAGE_BACKEND == "db" and _DB_READY:
        rows = _DB.list_orcamentos_excel(start=start, end=end, vendedor=vendedor, cnpj_digits=cnpj)
        return {"count": len(rows), "rows": rows}
    token = acquire_token()
    item_id = await get_drive_item_id_cached(token)
    session_id = await get_session_id_cached(token, item_id)
    rows = await list_rows_dicts(token, item_id, session_id)

    # Identify columns by normalized names
    cols = await list_columns(token, item_id, session_id)
    idx_id = _find_col(cols, ["id_orcamento", "id_orc", "idorcamento", "id_orcamento"])
    idx_cnpj = _find_col(cols, ["cnpj_cpf", "cnpjcpf", "cnpj"])
    idx_vend = _find_col(cols, ["vendedor"])
    idx_dh = _find_col(cols, ["data_hora", "datahora", "data_hora"])

    def _digits(s: Optional[str]) -> str:
        return re.sub(r"\D", "", s or "")

    def _parse_date(d: Optional[str]):
        if not d:
            return None
        try:
            return datetime.strptime(d, "%d/%m/%Y")
        except Exception:
            return None

    ds = _parse_date(start)
    de = _parse_date(end)
    out = []
    for d in rows:
        # Apply filters using indices if found
        if id and idx_id is not None and str(list(d.values())[idx_id]) != id:
            continue
        if cnpj and idx_cnpj is not None and _digits(str(list(d.values())[idx_cnpj])) != _digits(cnpj):
            continue
        if vendedor and idx_vend is not None and str(list(d.values())[idx_vend]).strip().lower() != vendedor.strip().lower():
            continue
        if (ds or de) and idx_dh is not None:
            dt_txt = str(list(d.values())[idx_dh] or "").strip()
            try:
                dt = datetime.strptime(dt_txt, "%d/%m/%Y %H:%M:%S")
            except Exception:
                dt = None
            if dt is None:
                continue
            if ds and dt.date() < ds.date():
                continue
            if de and dt.date() > de.date():
                continue
        out.append(d)
    return {"count": len(out), "rows": out}

@app.get("/api/orcamentos/{orc_id}")
async def obter_orcamento(orc_id: str):
    if STORAGE_BACKEND == "db" and _DB_READY:
        d = _DB.get_orcamento_by_id(orc_id)
        if d:
            return d
    token = acquire_token()
    item_id = await get_drive_item_id_cached(token)
    session_id = await get_session_id_cached(token, item_id)
    rows = await list_rows_dicts(token, item_id, session_id)
    cols = await list_columns(token, item_id, session_id)
    idx_id = _find_col(cols, ["id_orcamento", "id_orc", "idorcamento", "id_orcamento"])
    for d in rows:
        if idx_id is not None and str(list(d.values())[idx_id]) == orc_id:
            return d
    raise HTTPException(404, "OrÃ§amento não encontrado")

# ====== Cache helpers (item_id e sessão) ======
async def get_drive_item_id_cached(token: str) -> str:
    global _EXCEL_ITEM_ID
    if _EXCEL_ITEM_ID:
        return _EXCEL_ITEM_ID
    encoded = urllib.parse.quote(EXCEL_REL_PATH)
    url = f"{GRAPH}/me/drive/root:/{encoded}"
    async with await graph_client(token) as client:
        r = await client.get(url)
        if r.status_code != 200:
            raise HTTPException(500, f"Não achei o arquivo no OneDrive ({r.text})")
        data = r.json(); _EXCEL_ITEM_ID = str(data.get("id") or "")
        if not _EXCEL_ITEM_ID:
            raise HTTPException(500, "ID do arquivo Excel não retornado")
        return _EXCEL_ITEM_ID

async def get_session_id_cached(token: str, item_id: str) -> str:
    global _SESSION_ID, _SESSION_TS
    now = time.time()
    if _SESSION_ID and (now - _SESSION_TS) < _SESSION_TTL_SECONDS:
        return _SESSION_ID
    url = f"{GRAPH}/me/drive/items/{item_id}/workbook/createSession"
    async with await graph_client(token) as client:
        r = await client.post(url, json={"persistChanges": True})
        if r.status_code not in (200, 201):
            raise HTTPException(500, f"Erro ao criar sessão do Excel: {r.text}")
        data = r.json(); _SESSION_ID = str(data.get("id") or "")
        if not _SESSION_ID:
            raise HTTPException(500, "Sessão do Excel não retornada")
        _SESSION_TS = now
        return _SESSION_ID
# ====== Descoberta na rede (UDP) para clientes auto-configurarem a URL ======
_DISCOVERY_PORT = 56789

def _get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"


def _discovery_loop():
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        sock.bind(("", _DISCOVERY_PORT))
    except Exception:
        return
    while True:
        try:
            data, addr = sock.recvfrom(1024)
            if not data:
                continue
            if b"AUDACES_DISCOVERY" in data:
                ip = _get_local_ip()
                url = f"http://{ip}:8000"
                payload = json.dumps({"url": url}).encode("utf-8")
                sock.sendto(payload, addr)
        except Exception:
            continue


@app.on_event("startup")
async def _start_discovery():
    try:
        t = threading.Thread(target=_discovery_loop, daemon=True)
        t.start()
    except Exception:
        pass






