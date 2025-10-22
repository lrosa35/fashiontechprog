Coloque aqui o template do contrato (.docx)

Como o programa encontra o arquivo:
- A pasta local padrão é `data/CONTRATO PARA ATUALIZAÇÃO` (esta).
- O arquivo recomendado: "CONTRATO COMERCIAL Impressão.docx".
- Alternativamente, qualquer .docx nesta pasta contendo no nome "CONTRATO COMERCIAL Impressão" será detectado.

Publicação online:
- Em servidores (Heroku), inclua o arquivo .docx neste diretório no repositório privado ou ajuste a variável de ambiente `LOCAL_FILES_BASE` para apontar para um volume persistente e subpasta `CONTRATO PARA ATUALIZAÇÃO`.

Observação:
- A conversão para PDF com `docx2pdf` requer Microsoft Word (Windows/macOS). Em Linux/Heroku, gere o DOCX; o PDF pode não estar disponível.
