$ErrorActionPreference = 'Stop'

# Substitua com o nome do seu app Heroku
$App = "SEU-HEROKU-APP"

# CORS: Libera o Netlify
$Netlify = "https://SEU-SITE-NETLIFY.netlify.app"

# Modo de armazenamento: banco (Postgres)
heroku config:set -a $App STORAGE_BACKEND=db

# CORS
heroku config:set -a $App ALLOWED_ORIGINS=$Netlify

# Opcional: autenticação básica da UI
#heroku config:set -a $App UI_BASIC_USER=admin UI_BASIC_PASS=trocar123

Write-Host "Configure também o add-on Postgres: heroku addons:create heroku-postgresql:hobby-dev -a $App"
Write-Host "DATABASE_URL será definido automaticamente pelo Heroku."
