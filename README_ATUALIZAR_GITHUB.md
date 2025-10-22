Atualizar arquivos a partir do GitHub (lrosa35/fashiontechprog)

Opção 1 — usando git (recomendado)
- Pré-requisito: Git instalado e disponível no PATH
- Execute: `scripts/atualizar_github.bat`
- O script vai clonar na pasta `fashiontechprog/` ou, se já existir, fará `fetch/checkout/pull` do branch `main`.

Opção 2 — sem git (fallback automático)
- O mesmo `scripts/atualizar_github.bat` chamará `scripts/atualizar_github.ps1` para baixar o ZIP do GitHub e extrair na pasta `fashiontechprog/`.

Comandos equivalentes (manuais)
- git clone https://github.com/lrosa35/fashiontechprog.git
- cd fashiontechprog
- git checkout main

Observações
- Esses scripts não sobrescrevem automaticamente os arquivos do projeto atual; eles criam/atualizam a pasta `fashiontechprog/` no mesmo nível do repositório atual.
- Copie/mescle os arquivos que você deseja a partir dessa pasta, ou adapte o script para sincronizar arquivos específicos conforme sua necessidade.

