Guia de Publicação Online (Netlify + Heroku)

Objetivo
- Frontend 100% online no Netlify (estático, HTML/JS)
- Backend + Banco de Dados no Heroku (FastAPI + Postgres)
- Escolher entre ambiente local (teste) e online (produção)

1) Backend (Heroku)
- Crie um app no Heroku: heroku create SEU-HEROKU-APP
- Adicione Postgres: heroku addons:create heroku-postgresql:hobby-dev -a SEU-HEROKU-APP
- Configure variáveis:
  - STORAGE_BACKEND=db
  - DATABASE_URL (já vem do add-on Postgres)
  - ALLOWED_ORIGINS=https://SEU-SITE-NETLIFY.netlify.app
  - (opcional) UI_BASIC_USER, UI_BASIC_PASS para proteger a UI do FastAPI
- Deploy (duas opções):
  1. GitHub deploy automático (recomendado)
  2. Heroku CLI: `git push heroku main` (Procfile já incluso)
- O app vai iniciar `uvicorn ui_app:app` (API JSON em /api e UI simples server-side).

2) Frontend (Netlify)
- Em `netlify.toml` ajuste o domínio do Heroku na regra de redirect.
- Em `frontend/js/config.js.example` copie para `frontend/js/config.js` e ajuste `window.API_BASE` para a URL do Heroku.
- Publique a pasta `frontend/` no Netlify (Deploy via Git ou arrastar e soltar a pasta).

3) Rodar localmente (teste)
- Backend local: `python -m uvicorn ui_app:app --reload --port 8000`
- Frontend local: abra `frontend/index.html` no navegador e defina `window.API_BASE` como `http://localhost:8000`.
- App GUI local (Flet): use `run_ui_web.bat` (já abre o navegador em http://localhost:8090).

4) Alternar Local vs Nuvem
- App Flet/desktop usa a prioridade:
  1) arquivo `cloud_api_url.txt` (coloque a URL do Heroku aqui para forçar nuvem)
  2) variável `CLOUD_API_DEFAULT_URL`
  3) padrão `http://localhost:8000`
- Para testar local: remova/renomeie `cloud_api_url.txt` e não defina `CLOUD_API_DEFAULT_URL`.
- Para produção: crie `cloud_api_url.txt` com a URL do Heroku (ex: https://SEU-HEROKU-APP.herokuapp.com).

5) Template do Contrato online
- O programa busca o template DOCX em `data/CONTRATO PARA ATUALIZAÇÃO/`.
- Inclua seu arquivo "CONTRATO COMERCIAL Impressão.docx" nessa pasta e faça o deploy.
- Em Heroku (Linux), a conversão para PDF pode não funcionar; o DOCX é gerado normalmente.
- Alternativa: ajustar `LOCAL_FILES_BASE` para um volume/armazenamento externo e sincronizar o DOCX.

6) Renomeação de módulo
- O rótulo do módulo foi alterado para "Cadastro de Clientes" no app.

7) Requisitos
- `requirements.txt` cobre FastAPI/Jinja2/Uvicorn e demais libs.
- Para banco, o backend usa SQLAlchemy com `DATABASE_URL` (Heroku Postgres).

8) Segurança e CORS
- Defina `ALLOWED_ORIGINS` no Heroku com a URL do Netlify para liberar o CORS do /api.
- Se quiser proteger a UI server-side (ui_app), defina `UI_BASIC_USER` e `UI_BASIC_PASS`.

