Como publicar o Frontend (Netlify)

- Esta pasta contém um frontend estático simples (HTML + JS) que consome a API JSON do backend.
- Antes de publicar, copie `js/config.js.example` para `js/config.js` e ajuste `window.API_BASE` para a URL do Heroku.
- No Netlify, a raiz de publicação deve apontar para a pasta `frontend/` (já configurado em `netlify.toml`).

Arquivos
- `index.html`: criar orçamento (POST /api/orcamentos)
- `buscar.html`: buscar/listar orçamentos (GET /api/orcamentos)
- `js/config.js`: define `window.API_BASE` (não comitar credenciais)
- `js/app.js`: lógica de chamadas ao backend

