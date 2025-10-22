-- PostgreSQL schema for FashionTech ERP (initial)

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

-- Views com tipos normalizados (úteis para Power Query)
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
