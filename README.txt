Orçamento - Guia de Uso (API Excel OneDrive)
===========================================

Resumo
- Este projeto permite que vários usuários usem um único servidor para gravar em um arquivo Excel no OneDrive.
- Clientes rodam o aplicativo e se conectam à API; apenas o servidor escreve no Excel (via Microsoft Graph), evitando corrupção.

Arquivos importantes
- .env.example: modelo de variáveis de ambiente (copie para .env e preencha).
- iniciar_api_excel.bat: inicia a API Excel (server:app) em console.
- instalar_servico_excel.bat: instala a API como serviço do Windows.
- remover_servico_excel.bat: remove o serviço.
- iniciar_servidor.bat / instalar_servico.bat / remover_servico.bat: opção para backend SQL (futuro).

Servidor (OneDrive do dono do Excel)
1) Registrar o app no Entra ID (Azure AD)
   - App registrations → New registration → copie Tenant ID e Client ID.
   - API permissions → Microsoft Graph (Delegated): Files.ReadWrite.All e offline_access. Faça “Grant admin consent” se necessário.

2) Configurar .env
   - Copie .env.example para .env e preencha:
     - TENANT_ID=...
     - CLIENT_ID=...
     - EXCEL_RELATIVE_PATH=01 LEANDRO/IMPRESSÕES/BANCO_DE_DADOS_ORCAMENTO.xlsx
     - EXCEL_TABLE_NAME=Orcamentos
     - STORAGE_BACKEND=excel

3) Preparar venv e dependências (no diretório do projeto)
   - python -m venv .venv
   - .venv\Scripts\pip install -U pip uvicorn fastapi httpx msal python-dotenv

4) Subir a API (primeira vez)
   - Dê duplo clique em iniciar_api_excel.bat
   - Siga o Device Code mostrado no console (microsoft.com/devicelogin), faça login e aceite permissões.
   - O token será salvo em token_cache.bin (reutilizado em execuções futuras).

5) Rodar como Serviço (opcional)
   - Clique direito em instalar_servico_excel.bat → Executar como administrador.
   - Abra a porta TCP 8000 no firewall (entrada) para a rede local.

Excel no OneDrive
- O caminho apontado em EXCEL_RELATIVE_PATH deve existir no OneDrive da conta do servidor.
- A planilha precisa ter uma Tabela (não apenas células). O nome deve ser EXCEL_TABLE_NAME (padrão Orcamentos).
- Os cabeçalhos devem ser compatíveis com o aplicativo (ex.: ID Orçamento, Data/Hora, Tipo de Serviço, Cliente (Etiqueta PDF), Cliente (Valor), Documento, CNPJ/CPF, E-mail, Vendedor, Status, Quantidade, Unidade, Metros, Preço por metro, Forma de Pagamento, Valor Total).

Clientes (outros PCs)
1) Execute o aplicativo (EXE gerado via PyInstaller ou python orcamento.py).
2) Na aba Integração API:
   - Clique “Detectar API” se estiver no mesmo PC do servidor, OU
   - Informe API Base: http://IP_DO_SERVIDOR:8000 → Aplicar → Testar conexão.
3) Com a API ativa, todos os cadastros/orçamentos passam a gravar no Excel do OneDrive via API.

Distribuição facilitada da URL
- Gere o arquivo com seu IP local: gerar_cloud_url_txt.bat (cria cloud_api_url.txt com http://SEU_IP:8000).
- Ao distribuir o EXE, inclua esse cloud_api_url.txt ao lado do executável; o app vai pré-preencher a URL usando esse arquivo.
- Alternativa: defina a variável de ambiente CLOUD_API_DEFAULT_URL=http://SEU_IP:8000 nos clientes.

Firewall
- Execute abrir_firewall_8000.bat no servidor para liberar a porta 8000 (entrada TCP).

Gerar executável (opcional)
- .venv\Scripts\pip install pyinstaller
- pyinstaller orcamento.spec
- Distribua dist\orcamento.exe aos usuários.

Alternativa: Backend SQL
- Se o volume crescer, mude para STORAGE_BACKEND=db e use server_db:app com db_backend.py.
- Ajuste DATABASE_URL no .env para Postgres e use os scripts iniciar_servidor.bat/instalar_servico.bat.

Suporte
- Em caso de erro de conexão dos clientes: verifique porta 8000 no firewall e o IP do servidor.
- Em caso de erro de acesso ao Excel: confira EXCEL_RELATIVE_PATH, EXCEL_TABLE_NAME e permissões no Entra ID.
