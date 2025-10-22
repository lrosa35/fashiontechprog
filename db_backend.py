import os
import re
from datetime import datetime, timedelta

try:
    from sqlalchemy import create_engine, text
    _SA_OK = True
except Exception:
    _SA_OK = False


def _snake(s: str) -> str:
    s = s.strip().lower()
    s = s.replace("ã", "a").replace("õ", "o").replace("ç", "c").replace("é", "e").replace("ê", "e").replace("á", "a").replace("í", "i").replace("ú", "u").replace("ó", "o")
    return re.sub(r"[^a-z0-9]+", "_", s).strip("_")


class DB:
    _engine = None

    @classmethod
    def is_ready(cls) -> bool:
        url = os.environ.get("DATABASE_URL")
        if not (_SA_OK and url):
            return False
        try:
            if cls._engine is None:
                cls._engine = create_engine(url, pool_pre_ping=True, future=True)
            return True
        except Exception:
            return False

    @classmethod
    def init_schema(cls):
        """Cria as tabelas se não existirem."""
        ddl = """
        create table if not exists orcamentos (
            id_orcamento text primary key,
            data_hora text,
            tipo_servico text,
            cliente_label text,
            cliente_valor text,
            documento text,
            cnpj_cpf text,
            email text,
            vendedor text,
            desconto text,
            quantidade text,
            unidade text,
            metros text,
            preco_por_metro text,
            forma_pagamento text,
            valor_total text
        );
        create index if not exists idx_orc_cnpj on orcamentos((regexp_replace(cnpj_cpf,'\D','','g')));
        
        create table if not exists cadastros (
            cnpj_cpf text primary key,
            documento text,
            razao_social_nome text,
            nome_fantasia text,
            contato text,
            email_cnpj text,
            email_manual text,
            cep text,
            endereco text,
            numero text,
            complemento text,
            bairro text,
            municipio text,
            uf text,
            entrega_cep text,
            entrega_endereco text,
            entrega_numero text,
            entrega_complemento text,
            entrega_bairro text,
            entrega_municipio text,
            entrega_uf text,
            desconto_duracao text,
            desconto_unidade text,
            telefone1 text,
            telefone2 text,
            vendedor text,
            criado_em text,
            atualizado_em text
        );
        create index if not exists idx_cad_cnpj on cadastros((regexp_replace(cnpj_cpf,'\D','','g')));

        create table if not exists pedidos (
            id text primary key,
            pedido int,
            tipo_servico text,
            status_cliente text,
            quantidade_m text,
            valor_unitario text,
            valor_total text,
            data_hora_criacao text,
            id_orcamento text,
            documento text,
            cnpj_cpf text,
            cliente text,
            vendedor text,
            forma_pgto_orcamento text,
            forma_pgto_contrato text,
            pct_comissao_vendedor text,
            valor_comissao_vendedor text,
            pct_comissao_adm text,
            valor_comissao_adm text
        );
        create index if not exists idx_ped_cnpj on pedidos((regexp_replace(cnpj_cpf,'\D','','g')));
        """
        with cls._engine.begin() as c:
            c.execute(text(ddl))
        # Views tipadas para Power Query
        try:
            with cls._engine.begin() as c:
                c.execute(text("""
                create or replace view vw_orcamentos_typed as
                select
                  id_orcamento,
                  to_timestamp(nullif(data_hora,''),'DD/MM/YYYY HH24:MI:SS') as data_hora_ts,
                  tipo_servico, cliente_label, cliente_valor, documento, cnpj_cpf, email, vendedor, desconto,
                  nullif(replace(replace(quantidade,'.',''),',','.'),'')::numeric as quantidade_num,
                  unidade,
                  nullif(replace(replace(metros,'.',''),',','.'),'')::numeric as metros_num,
                  nullif(replace(replace(preco_por_metro,'.',''),',','.'),'')::numeric as preco_num,
                  forma_pagamento,
                  nullif(replace(replace(valor_total,'.',''),',','.'),'')::numeric as valor_total_num
                from orcamentos;

                create or replace view vw_pedidos_typed as
                select
                  id, pedido,
                  to_timestamp(nullif(data_hora_criacao,''),'DD/MM/YYYY HH24:MI:SS') as data_hora_ts,
                  tipo_servico, status_cliente,
                  nullif(replace(replace(quantidade_m,'.',''),',','.'),'')::numeric as quantidade_m_num,
                  nullif(replace(replace(valor_unitario,'.',''),',','.'),'')::numeric as valor_unitario_num,
                  nullif(replace(replace(valor_total,'.',''),',','.'),'')::numeric as valor_total_num,
                  id_orcamento, documento, cnpj_cpf, cliente, vendedor,
                  forma_pgto_orcamento, forma_pgto_contrato,
                  pct_comissao_vendedor, valor_comissao_vendedor,
                  pct_comissao_adm, valor_comissao_adm
                from pedidos;
                """))
        except Exception:
            pass

    @classmethod
    def init_schema_portable(cls):
        """Cria o esquema de forma tolerante a SQLite/Postgres (executando statement a statement)."""
        try:
            is_sqlite = (cls._engine.dialect.name == "sqlite")
        except Exception:
            is_sqlite = False

        stmts = [
            """
            create table if not exists orcamentos (
                id_orcamento text primary key,
                data_hora text,
                tipo_servico text,
                cliente_label text,
                cliente_valor text,
                documento text,
                cnpj_cpf text,
                email text,
                vendedor text,
                desconto text,
                quantidade text,
                unidade text,
                metros text,
                preco_por_metro text,
                forma_pagamento text,
                valor_total text
            );
            """,
            """
            create table if not exists cadastros (
                cnpj_cpf text primary key,
                documento text,
                razao_social_nome text,
                nome_fantasia text,
                contato text,
                email_cnpj text,
                email_manual text,
                cep text,
                endereco text,
                numero text,
                complemento text,
                bairro text,
                municipio text,
                uf text,
                entrega_cep text,
                entrega_endereco text,
                entrega_numero text,
                entrega_complemento text,
                entrega_bairro text,
                entrega_municipio text,
                entrega_uf text,
                desconto_duracao text,
                desconto_unidade text,
                telefone1 text,
                telefone2 text,
                vendedor text,
                criado_em text,
                atualizado_em text
            );
            """,
            """
            create table if not exists pedidos (
                id text primary key,
                pedido int,
                tipo_servico text,
                status_cliente text,
                quantidade_m text,
                valor_unitario text,
                valor_total text,
                data_hora_criacao text,
                id_orcamento text,
                documento text,
                cnpj_cpf text,
                cliente text,
                vendedor text,
                forma_pgto_orcamento text,
                forma_pgto_contrato text,
                pct_comissao_vendedor text,
                valor_comissao_vendedor text,
                pct_comissao_adm text,
                valor_comissao_adm text
            );
            """,
        ]
        if is_sqlite:
            stmts += [
                "create index if not exists idx_orc_cnpj on orcamentos(replace(replace(replace(cnpj_cpf,'/',''),'.',''),'-',''));",
                "create index if not exists idx_cad_cnpj on cadastros(replace(replace(replace(cnpj_cpf,'/',''),'.',''),'-',''));",
                "create index if not exists idx_ped_cnpj on pedidos(replace(replace(replace(cnpj_cpf,'/',''),'.',''),'-',''));",
            ]
        else:
            stmts += [
                "create index if not exists idx_orc_cnpj on orcamentos((regexp_replace(cnpj_cpf,'\\D','','g')));",
                "create index if not exists idx_cad_cnpj on cadastros((regexp_replace(cnpj_cpf,'\\D','','g')));",
                "create index if not exists idx_ped_cnpj on pedidos((regexp_replace(cnpj_cpf,'\\D','','g')));",
            ]
        with cls._engine.begin() as conn:
            for sql in stmts:
                try:
                    conn.execute(text(sql))
                except Exception:
                    pass

    # Mapeamentos Excel -> DB
    ORC_MAP = {
        "ID Orçamento": "id_orcamento",
        "Data/Hora": "data_hora",
        "Tipo de Serviço": "tipo_servico",
        "Cliente (Etiqueta PDF)": "cliente_label",
        "Cliente (Valor)": "cliente_valor",
        "Documento": "documento",
        "CNPJ/CPF": "cnpj_cpf",
        "E-mail": "email",
        "Vendedor": "vendedor",
        "Status": "desconto",
        "Desconto": "desconto",
        "Quantidade": "quantidade",
        "Unidade": "unidade",
        "Metros": "metros",
        "Preço por metro": "preco_por_metro",
        "Forma de Pagamento": "forma_pagamento",
        "Valor Total": "valor_total",
    }

    CAD_MAP = {
        "Documento": "documento",
        "CNPJ/CPF": "cnpj_cpf",
        "Razão Social/Nome": "razao_social_nome",
        "Nome Fantasia": "nome_fantasia",
        "Contato": "contato",
        "E-mail (CNPJ)": "email_cnpj",
        "E-mail (Manual)": "email_manual",
        "CEP": "cep",
        "Endereço": "endereco",
        "Número": "numero",
        "Complemento": "complemento",
        "Bairro": "bairro",
        "Município": "municipio",
        "UF": "uf",
        "Entrega CEP": "entrega_cep",
        "Entrega Endereço": "entrega_endereco",
        "Entrega Número": "entrega_numero",
        "Entrega Complemento": "entrega_complemento",
        "Entrega Bairro": "entrega_bairro",
        "Entrega Município": "entrega_municipio",
        "Entrega UF": "entrega_uf",
        "Desconto Duração": "desconto_duracao",
        "Desconto Unidade": "desconto_unidade",
        "Telefone 1": "telefone1",
        "Telefone 2": "telefone2",
        "Vendedor": "vendedor",
        "Criado em": "criado_em",
        "Atualizado em": "atualizado_em",
    }

    PED_MAP = {
        "ID": "id",
        "Pedido": "pedido",
        "Tipo de Serviço": "tipo_servico",
        "Status do Cliente": "status_cliente",
        "Quantidade (m)": "quantidade_m",
        "Valor Unitário": "valor_unitario",
        "Valor Total": "valor_total",
        "Data/Hora da criação do pedido": "data_hora_criacao",
        "ID Orçamento": "id_orcamento",
        "Documento": "documento",
        "CNPJ/CPF": "cnpj_cpf",
        "Cliente": "cliente",
        "Vendedor": "vendedor",
        "Forma de Pagamento Orçamento": "forma_pgto_orcamento",
        "Forma de Pagamento Contrato": "forma_pgto_contrato",
        "% Comissão Vendedor": "pct_comissao_vendedor",
        "Valor Comissão Vendedor": "valor_comissao_vendedor",
        "% Comissão ADM": "pct_comissao_adm",
        "Valor Comissão ADM": "valor_comissao_adm",
    }

    REV_ORC = {v: k for k, v in ORC_MAP.items()}
    REV_CAD = {v: k for k, v in CAD_MAP.items()}
    REV_PED = {v: k for k, v in PED_MAP.items()}

    @classmethod
    def _map_payload(cls, d: dict, mapping: dict) -> dict:
        out = {}
        for k, v in mapping.items():
            out[v] = d.get(k)
        return out

    @classmethod
    def _row_to_excel_orc(cls, row: dict) -> dict:
        return {cls.REV_ORC.get(k, k): v for k, v in row.items()}

    # ============ ORÇAMENTOS ============
    @classmethod
    def salvar_orcamento(cls, dados: dict):
        payload = cls._map_payload(dados, cls.ORC_MAP)
        cols = ",".join(payload.keys())
        params = ",".join(f":{k}" for k in payload.keys())
        with cls._engine.begin() as c:
            c.execute(text(f"insert into orcamentos ({cols}) values ({params}) on conflict (id_orcamento) do update set data_hora=excluded.data_hora"), payload)

    @classmethod
    def get_orcamento_by_id(cls, id_orc: str):
        with cls._engine.connect() as c:
            row = c.execute(text("select * from orcamentos where id_orcamento=:id limit 1"), {"id": id_orc}).mappings().first()
            return cls._row_to_excel_orc(dict(row)) if row else None

    @classmethod
    def get_orcamentos_list(cls, doc_formatado: str | None = None, id_orc: str | None = None):
        where, params = [], {}
        if id_orc:
            where.append("id_orcamento = :id_orc")
            params["id_orc"] = id_orc
        if doc_formatado:
            digits = re.sub(r"\D", "", doc_formatado or "")
            where.append("regexp_replace(cnpj_cpf,'\\D','','g') = :digits")
            params["digits"] = digits
        sql = "select * from orcamentos"
        if where:
            sql += " where " + " and ".join(where)
        sql += " order by data_hora desc"
        with cls._engine.connect() as c:
            res = c.execute(text(sql), params).mappings().all()
            return [cls._row_to_excel_orc(dict(r)) for r in res]

    @classmethod
    def list_orcamentos_excel(cls, start: str | None = None, end: str | None = None, vendedor: str | None = None, cnpj_digits: str | None = None) -> list[dict]:
        where = []
        params = {}
        is_sqlite = cls._engine.dialect.name == "sqlite"
        if vendedor:
            if is_sqlite:
                where.append("lower(coalesce(vendedor,'')) like :vend")
                params["vend"] = f"%{(vendedor or '').lower()}%"
            else:
                where.append("coalesce(vendedor,'') ilike :vend")
                params["vend"] = f"%{vendedor}%"
        if cnpj_digits:
            if is_sqlite:
                where.append("replace(replace(replace(cnpj_cpf,'/',''),'.',''),'-','') = :digits")
            else:
                where.append("regexp_replace(cnpj_cpf,'\\D','','g') = :digits")
            params["digits"] = re.sub(r"\D","", cnpj_digits)
        sql = "select * from orcamentos"
        if where:
            sql += " where " + " and ".join(where)
        sql += " order by data_hora desc"
        with cls._engine.connect() as c:
            res = c.execute(text(sql), params).mappings().all()
            rows = [cls._row_to_excel_orc(dict(r)) for r in res]
        # Filtro de data (formato DD/MM/YYYY) feito em Python para compatibilidade
        if start or end:
            dstart = datetime.strptime(start, "%d/%m/%Y") if start else None
            dend = datetime.strptime(end, "%d/%m/%Y") if end else None
            def parse_dt(s: str):
                try:
                    return datetime.strptime((s or "").split()[0], "%d/%m/%Y")
                except Exception:
                    return None
            def ok(r):
                ts = parse_dt(r.get("Data/Hora"))
                if ts is None:
                    return False if (dstart or dend) else True
                if dstart and ts < dstart: return False
                if dend and ts >= (dend + timedelta(days=1)): return False
                return True
            rows = [r for r in rows if ok(r)]
        return rows

    # ============ CADASTROS ============
    @classmethod
    def salvar_cadastro(cls, dados: dict):
        payload = cls._map_payload(dados, cls.CAD_MAP)
        payload["cnpj_cpf"] = re.sub(r"\D", "", payload.get("cnpj_cpf") or "")
        payload.setdefault("criado_em", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        payload["atualizado_em"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        cols = ",".join(payload.keys())
        params = ",".join(f":{k}" for k in payload.keys())
        with cls._engine.begin() as c:
            c.execute(text(f"insert into cadastros ({cols}) values ({params}) on conflict (cnpj_cpf) do update set documento=excluded.documento, razao_social_nome=excluded.razao_social_nome, nome_fantasia=excluded.nome_fantasia, contato=excluded.contato, email_cnpj=excluded.email_cnpj, email_manual=excluded.email_manual, cep=excluded.cep, endereco=excluded.endereco, numero=excluded.numero, complemento=excluded.complemento, bairro=excluded.bairro, municipio=excluded.municipio, uf=excluded.uf, entrega_cep=excluded.entrega_cep, entrega_endereco=excluded.entrega_endereco, entrega_numero=excluded.entrega_numero, entrega_complemento=excluded.entrega_complemento, entrega_bairro=excluded.entrega_bairro, entrega_municipio=excluded.entrega_municipio, entrega_uf=excluded.entrega_uf, desconto_duracao=excluded.desconto_duracao, desconto_unidade=excluded.desconto_unidade, telefone1=excluded.telefone1, telefone2=excluded.telefone2, vendedor=excluded.vendedor, atualizado_em=excluded.atualizado_em"), payload)

    @classmethod
    def atualizar_cadastro(cls, doc_formatado: str, dados: dict) -> bool:
        cls.salvar_cadastro(dados)
        return True

    @classmethod
    def buscar_cadastro_por_documento(cls, tipo: str, valor_digitado: str):
        digits = re.sub(r"\D", "", valor_digitado or "")
        is_sqlite = cls._engine.dialect.name == "sqlite"
        with cls._engine.connect() as c:
            if is_sqlite:
                sql = "select * from cadastros where replace(replace(replace(cnpj_cpf,'/',''),'.',''),'-','')=:d order by atualizado_em desc limit 1"
            else:
                sql = "select * from cadastros where regexp_replace(cnpj_cpf,'\\D','','g')=:d order by atualizado_em desc limit 1"
            row = c.execute(text(sql), {"d": digits}).mappings().first()
            if not row:
                return None
            d = dict(row)
            # Converte para chaves usadas no app
            out = {k: d.get(v) for k, v in cls.CAD_MAP.items()}
            out["CNPJ/CPF"] = d.get("cnpj_cpf")
            return out

    @classmethod
    def list_cadastros_excel(cls, start: str | None = None, end: str | None = None, vendedor: str | None = None, cnpj_digits: str | None = None) -> list[dict]:
        where = []
        params = {}
        is_sqlite = cls._engine.dialect.name == "sqlite"
        if vendedor:
            if is_sqlite:
                where.append("lower(coalesce(vendedor,'')) like :vend")
                params["vend"] = f"%{(vendedor or '').lower()}%"
            else:
                where.append("coalesce(vendedor,'') ilike :vend")
                params["vend"] = f"%{vendedor}%"
        if cnpj_digits:
            if is_sqlite:
                where.append("replace(replace(replace(cnpj_cpf,'/',''),'.',''),'-','') = :digits")
            else:
                where.append("regexp_replace(cnpj_cpf,'\\D','','g') = :digits")
            params["digits"] = re.sub(r"\D","", cnpj_digits)
        sql = "select * from cadastros"
        if where:
            sql += " where " + " and ".join(where)
        sql += " order by atualizado_em desc"
        with cls._engine.connect() as c:
            res = c.execute(text(sql), params); rows = []
            for m in res.mappings().all():
                d = dict(m)
                out = {label: d.get(col) for label, col in cls.CAD_MAP.items()}
                out["CNPJ/CPF"] = d.get("cnpj_cpf")
                rows.append(out)
        # Filtro por data (formato DD/MM/YYYY) aplicado em Python
        if start or end:
            dstart = datetime.strptime(start, "%d/%m/%Y") if start else None
            dend = datetime.strptime(end, "%d/%m/%Y") if end else None
            def parse_dt(s: str):
                try:
                    return datetime.strptime((s or "").split()[0], "%d/%m/%Y")
                except Exception:
                    return None
            def ok(r):
                ts = parse_dt(r.get("Atualizado em") or r.get("Criado em") or "")
                if ts is None:
                    return False if (dstart or dend) else True
                if dstart and ts < dstart: return False
                if dend and ts >= (dend + timedelta(days=1)): return False
                return True
            rows = [r for r in rows if ok(r)]
            return rows
        return rows

    # ============ PEDIDOS / OUTROS ============
    @classmethod
    def get_ultimo_pedido_data(cls, doc_formatado: str):
        digits = re.sub(r"\D", "", doc_formatado or "")
        is_sqlite = cls._engine.dialect.name == "sqlite"
        with cls._engine.connect() as c:
            if is_sqlite:
                res = c.execute(text("select data_hora_criacao from pedidos where replace(replace(replace(cnpj_cpf,'/',''),'.',''),'-','')=:d"), {"d": digits}).fetchall()
                vals = [r[0] for r in res if r and r[0]]
                return max(vals) if vals else None
            else:
                row = c.execute(text("select data_hora_criacao from pedidos where regexp_replace(cnpj_cpf,'\\D','','g')=:d order by data_hora_criacao desc limit 1"), {"d": digits}).first()
                return row[0] if row else None

    @classmethod
    def get_proximo_pedido_numero(cls) -> int:
        with cls._engine.connect() as c:
            row = c.execute(text("select coalesce(max(pedido),0) from pedidos")).first()
            return int(row[0]) + 1

    @classmethod
    def salvar_pedido(cls, dados: dict):
        payload = cls._map_payload(dados, cls.PED_MAP)
        payload["cnpj_cpf"] = re.sub(r"\D", "", payload.get("cnpj_cpf") or "")
        cols = ",".join(payload.keys())
        params = ",".join(f":{k}" for k in payload.keys())
        with cls._engine.begin() as c:
            c.execute(text(f"insert into pedidos ({cols}) values ({params}) on conflict (id) do nothing"), payload)

    @classmethod
    def list_pedidos_excel(cls, start: str | None = None, end: str | None = None, vendedor: str | None = None, cnpj_digits: str | None = None) -> list[dict]:
        where = []
        params = {}
        is_sqlite = cls._engine.dialect.name == "sqlite"
        if vendedor:
            if is_sqlite:
                where.append("lower(coalesce(vendedor,'')) like :vend")
                params["vend"] = f"%{(vendedor or '').lower()}%"
            else:
                where.append("coalesce(vendedor,'') ilike :vend")
                params["vend"] = f"%{vendedor}%"
        if cnpj_digits:
            if is_sqlite:
                where.append("replace(replace(replace(cnpj_cpf,'/',''),'.',''),'-','') = :digits")
            else:
                where.append("regexp_replace(cnpj_cpf,'\\D','','g') = :digits")
            params["digits"] = re.sub(r"\D","", cnpj_digits)
        sql = "select * from pedidos"
        if where:
            sql += " where " + " and ".join(where)
        sql += " order by data_hora_criacao desc"
        with cls._engine.connect() as c:
            res = c.execute(text(sql), params); rows = []
            for m in res.mappings().all():
                d = dict(m)
                # reverse map
                out = {}
                for label, col in cls.PED_MAP.items():
                    out[label] = d.get(col)
                rows.append(out)
        if start or end:
            dstart = datetime.strptime(start, "%d/%m/%Y") if start else None
            dend = datetime.strptime(end, "%d/%m/%Y") if end else None
            def parse_dt(s: str):
                try:
                    return datetime.strptime(s, "%d/%m/%Y %H:%M:%S")
                except Exception:
                    return None
            def ok(r):
                ts = parse_dt(r.get("Data/Hora da criação do pedido") or r.get("data_hora_criacao") or "")
                if ts is None:
                    return False if (dstart or dend) else True
                if dstart and ts < dstart: return False
                if dend and ts >= (dend + timedelta(days=1)): return False
                return True
            rows = [r for r in rows if ok(r)]
            return rows
        return rows

    # ============ USUARIOS / ACESSO ============
    @staticmethod
    def _hash_password(raw: str) -> str:
        import hashlib
        raw = (raw or "").encode("utf-8")
        return hashlib.sha256(raw).hexdigest()

    @classmethod
    def upsert_usuario(cls, u: dict):
        usuario = (u.get("usuario") or "").strip()
        if not usuario:
            raise ValueError("usuario obrigatorio")
        payload = {
            "usuario": usuario,
            "nome": u.get("nome"),
            "email": u.get("email"),
            "setor": u.get("setor"),
            "cargo": u.get("cargo"),
            "senha_hash": None,
            "is_admin": 1 if (u.get("is_admin") in (1, True, "1", "true", "True")) else 0,
            "permissoes": u.get("permissoes"),
        }
        senha = u.get("senha")
        if senha:
            payload["senha_hash"] = cls._hash_password(senha)
        with cls._engine.begin() as c:
            if payload["senha_hash"] is None:
                c.execute(text(
                    "insert into usuarios (usuario,nome,email,setor,cargo,is_admin,permissoes) values (:usuario,:nome,:email,:setor,:cargo,:is_admin,:permissoes) "
                    "on conflict(usuario) do update set nome=excluded.nome,email=excluded.email,setor=excluded.setor,cargo=excluded.cargo,is_admin=excluded.is_admin,permissoes=excluded.permissoes"
                ), payload)
            else:
                c.execute(text(
                    "insert into usuarios (usuario,nome,email,setor,cargo,senha_hash,is_admin,permissoes) values (:usuario,:nome,:email,:setor,:cargo,:senha_hash,:is_admin,:permissoes) "
                    "on conflict(usuario) do update set nome=excluded.nome,email=excluded.email,setor=excluded.setor,cargo=excluded.cargo,senha_hash=excluded.senha_hash,is_admin=excluded.is_admin,permissoes=excluded.permissoes"
                ), payload)

    @classmethod
    def list_usuarios(cls) -> list[dict]:
        with cls._engine.connect() as c:
            res = c.execute(text("select usuario,nome,email,setor,cargo,is_admin,coalesce(permissoes,'') as permissoes from usuarios order by usuario")).mappings().all()
            return [dict(r) for r in res]

    @classmethod
    def get_usuario(cls, usuario: str) -> dict | None:
        with cls._engine.connect() as c:
            r = c.execute(text("select * from usuarios where usuario=:u"), {"u": usuario}).mappings().first()
            return dict(r) if r else None

    @classmethod
    def check_login(cls, usuario: str, senha: str) -> dict | None:
        u = cls.get_usuario(usuario)
        if not u:
            return None
        ok = (u.get("senha_hash") == cls._hash_password(senha or ""))
        return u if ok else None

    @classmethod
    def set_password(cls, usuario: str, nova_senha: str):
        with cls._engine.begin() as c:
            c.execute(text("update usuarios set senha_hash=:h where usuario=:u"), {"h": cls._hash_password(nova_senha), "u": usuario})

    @classmethod
    def ensure_admin(cls, usuario: str, nome: str, email: str, setor: str, cargo: str, senha: str):
        with cls._engine.begin() as c:
            r = c.execute(text("select 1 from usuarios where usuario=:u"), {"u": usuario}).first()
            exists = bool(r)
        if not exists:
            cls.upsert_usuario({
                "usuario": usuario,
                "nome": nome,
                "email": email,
                "setor": setor,
                "cargo": cargo,
                "senha": senha,
                "is_admin": 1,
                "permissoes": "*",
            })
