// Simple static frontend for Netlify, calling the backend API.
// Configure API base in js/config.js (see config.js.example).

const API_BASE = (window.API_BASE || '').replace(/\/$/, '');

function qs(id) { return document.getElementById(id); }
function showMsg(html, cls='') {
  const el = qs('msg');
  if (!el) return;
  el.className = cls;
  el.innerHTML = html;
}

async function criarOrcamento(e) {
  e.preventDefault();
  showMsg('');
  const body = {
    tipo_servico: qs('tipo_servico').value,
    cliente: qs('cliente').value,
    cnpj: qs('cnpj').value,
    email: qs('email').value,
    status: qs('status').value,
    unidade: qs('unidade').value,
    quantidade: qs('quantidade').value,
  };
  try {
    const r = await fetch(`${API_BASE}/api/orcamentos`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    const data = await r.json();
    if (!r.ok) throw new Error(data?.detail || 'Falha ao criar orçamento');
    showMsg('Orçamento criado com sucesso.', 'success');
    const pre = qs('result-pre');
    const box = qs('result');
    if (pre && box) { pre.textContent = JSON.stringify(data, null, 2); box.style.display = 'block'; }
  } catch (ex) {
    showMsg(String(ex), 'error');
  }
}

async function buscar(e) {
  e.preventDefault();
  showMsg('');
  const id = qs('id')?.value?.trim();
  const cnpj = qs('cnpj')?.value?.trim();
  const params = new URLSearchParams();
  if (id) params.set('id', id);
  if (cnpj) params.set('cnpj', cnpj);
  try {
    const r = await fetch(`${API_BASE}/api/orcamentos?${params.toString()}`);
    const data = await r.json();
    if (!r.ok) throw new Error(data?.detail || 'Falha na busca');
    const rows = data?.rows || [];
    const tbl = qs('tbl');
    const tbody = tbl?.querySelector('tbody');
    if (!tbody) return;
    tbody.innerHTML = '';
    for (const d of rows) {
      const idv = d['ID Orçamento'] || d['id_orcamento'] || d['ID'] || '';
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${idv}</td>
        <td>${d['Data/Hora'] || d['data_hora'] || ''}</td>
        <td>${d['Cliente'] || d['cliente'] || ''}</td>
        <td>${d['CNPJ/CPF'] || d['cnpj'] || ''}</td>
        <td>${d['Valor Total'] || d['valor_total'] || ''}</td>
        <td><a class="btn secondary" href="${API_BASE}/api/orcamentos/${idv}">JSON</a></td>
      `;
      tbody.appendChild(tr);
    }
    tbl.style.display = rows.length ? 'table' : 'none';
    showMsg(`${rows.length} registro(s) encontrados.`, 'success');
  } catch (ex) {
    showMsg(String(ex), 'error');
  }
}

window.addEventListener('DOMContentLoaded', () => {
  const f1 = qs('form-orc');
  const f2 = qs('form-busca');
  if (f1) f1.addEventListener('submit', criarOrcamento);
  if (f2) f2.addEventListener('submit', buscar);
});

