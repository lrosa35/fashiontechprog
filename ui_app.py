import os
from typing import Optional

import os
from datetime import datetime
from typing import Optional

from fastapi import FastAPI, Request, Form, status, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import ValidationError
from fastapi.security import HTTPBasic, HTTPBasicCredentials
import secrets

# Reuse existing API app and logic
from server import app as api_app
from server import OrcamentoIn, criar_orcamento, listar_orcamentos, obter_orcamento
import orcamento as orc


app = FastAPI(title="Orçamentos Web UI")

# Mount the JSON API under /api (no changes required in server.py)
app.mount("/api", api_app)


BASE_DIR = os.path.dirname(__file__)
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")

templates = Jinja2Templates(directory=TEMPLATES_DIR)

if os.path.isdir(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Basic auth (optional). Keep above route definitions.
security = HTTPBasic()
UI_USER = os.getenv("UI_BASIC_USER")
UI_PASS = os.getenv("UI_BASIC_PASS")

def require_auth(credentials: HTTPBasicCredentials = Depends(security)):
    if not (UI_USER and UI_PASS):
        return  # auth disabled when not configured
    is_user = secrets.compare_digest(credentials.username, UI_USER)
    is_pass = secrets.compare_digest(credentials.password, UI_PASS)
    if not (is_user and is_pass):
        raise HTTPException(status_code=401, detail="Unauthorized")
    return


@app.get("/", response_class=HTMLResponse)
async def index(request: Request, _auth=Depends(require_auth)):
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "defaults": {
                "tipo_servico": "Impressão",
                "status": "Sem desconto",
                "unidade": "Centímetros",
            },
        },
    )


@app.post("/orcamentos", response_class=HTMLResponse)
async def create_orcamento(
    request: Request,
    tipo_servico: str = Form(...),
    cliente: str = Form(...),
    cnpj: str = Form(...),
    email: str = Form(...),
    status_: str = Form(alias="status", default="Sem desconto"),
    unidade: str = Form(...),
    quantidade: str = Form(...),
    _auth=Depends(require_auth),
):
    try:
        body = OrcamentoIn(
            tipo_servico=tipo_servico,
            cliente=cliente,
            cnpj=cnpj,
            email=email,
            status=status_,
            unidade=unidade,
            quantidade=quantidade,
        )
    except ValidationError as ve:
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "error": "Erro de validação: verifique os campos.",
                "details": ve.errors(),
                "defaults": {
                    "tipo_servico": tipo_servico,
                    "cliente": cliente,
                    "cnpj": cnpj,
                    "email": email,
                    "status": status_,
                    "unidade": unidade,
                    "quantidade": quantidade,
                },
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )

    try:
        result = await criar_orcamento(body)  # reuse API business logic
        return templates.TemplateResponse(
            "result.html",
            {"request": request, "orc": result.model_dump()},
            status_code=status.HTTP_201_CREATED,
        )
    except Exception as ex:
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "error": f"Falha ao criar orçamento: {ex}",
                "defaults": body.model_dump(),
            },
            status_code=status.HTTP_400_BAD_REQUEST,
        )


@app.get("/buscar", response_class=HTMLResponse)
async def buscar_page(request: Request, id: Optional[str] = None, cnpj: Optional[str] = None, _auth=Depends(require_auth)):
    rows = []
    error = None
    if id or cnpj:
        try:
            resp = await listar_orcamentos(id=id, cnpj=cnpj, vendedor=None, start=None, end=None)
            rows = resp.get("rows", []) if isinstance(resp, dict) else []
        except Exception as ex:
            error = f"Falha ao buscar: {ex}"
    return templates.TemplateResponse(
        "list.html",
        {"request": request, "rows": rows, "id": id or "", "cnpj": cnpj or "", "error": error},
    )


@app.get("/orcamentos/{orc_id}", response_class=HTMLResponse)
async def detalhe_orcamento(request: Request, orc_id: str, _auth=Depends(require_auth)):
    try:
        d = await obter_orcamento(orc_id)
        return templates.TemplateResponse("detail.html", {"request": request, "orc": d})
    except Exception as ex:
        return templates.TemplateResponse(
            "detail.html",
            {"request": request, "error": f"Não foi possível obter o orçamento: {ex}"},
            status_code=status.HTTP_404_NOT_FOUND,
        )


def _infer_doc_label_and_value(d: dict) -> tuple[str, str, str]:
    # returns (documento, label, value)
    doc_raw = d.get("CNPJ/CPF") or d.get("cnpj") or d.get("cnpj_cpf") or ""
    digits = "".join([ch for ch in str(doc_raw) if ch.isdigit()])
    if len(digits) == 11:
        return ("CPF", "Nome", orc.formatar_cpf(digits))
    if len(digits) == 14:
        return ("CNPJ", "Razão Social", orc.formatar_cnpj(digits))
    # fallback
    return ("Documento", "Documento", str(doc_raw))


@app.get("/orcamentos/{orc_id}/pdf")
async def baixar_pdf(orc_id: str, _auth=Depends(require_auth)):
    # get data
    d = await obter_orcamento(orc_id)
    if not isinstance(d, dict):
        raise HTTPException(404, "Orçamento não encontrado")

    id_orc = d.get("ID Orçamento") or d.get("id_orcamento") or orc_id
    datahora = d.get("Data/Hora") or d.get("data_hora") or datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    tipo_servico = d.get("Tipo de Serviço") or d.get("tipo_servico") or "Impressão"
    cliente_val = d.get("CLIENTE (Valor)") or d.get("Cliente") or d.get("cliente") or ""
    vendedor = d.get("Vendedor") or ""
    status_val = d.get("Status") or d.get("status") or "Sem desconto"
    qtd = d.get("Quantidade") or d.get("quantidade") or ""
    unidade = d.get("Unidade") or d.get("unidade") or "Centímetros"
    metros = d.get("Metros") or d.get("metros") or ""
    preco = d.get("Preço por metro") or d.get("preco_por_metro") or ""
    forma_pgto = d.get("Forma de Pagamento") or ""
    total = d.get("Valor Total") or d.get("valor_total") or ""

    documento, cliente_label, doc_fmt = _infer_doc_label_and_value(d)

    dados_seq = [
        id_orc,
        datahora,
        tipo_servico,
        cliente_label,
        str(cliente_val),
        documento,
        doc_fmt,
        str(d.get("E-mail") or d.get("email") or ""),
        vendedor,
        status_val,
        str(qtd),
        unidade,
        str(metros),
        str(preco),
        str(forma_pgto),
        str(total),
    ]

    export_dir = os.path.join(os.path.dirname(__file__), "data", "exports")
    os.makedirs(export_dir, exist_ok=True)
    pdf_path = orc.gerar_pdf_orcamento(dados_seq, export_dir)
    filename = os.path.basename(pdf_path)
    return FileResponse(pdf_path, media_type="application/pdf", filename=filename)


@app.get("/healthz")
async def healthz():
    return {"ok": True}
