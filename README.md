ESTRUTURA DOS ARQUIVOS
-----------------------
  app.py      → Código do aplicativo (interface + coleta + Excel)
  build.bat   → Script para gerar o .exe (rodar no seu PC)

------------------------------------------------------------
OPÇÃO 1 — Gerar o .exe (recomendado para distribuir)
------------------------------------------------------------

No SEU computador (que tem Python):

  1. Coloque app.py e build.bat na mesma pasta
  2. Dê dois cliques em build.bat
  3. Aguarde 1-2 minutos
  4. O arquivo  dist\VilaDigital_UFU.exe  estará pronto

Esse .exe pode ser:
  ✔ Copiado por pendrive para qualquer PC Windows
  ✔ Compartilhado via e-mail ou Google Drive
  ✔ Executado sem instalar Python ou qualquer coisa

ATENÇÃO: O .exe precisa de conexão com a internet
para acessar o site do PROPLAD e baixar o ChromeDriver
automaticamente na primeira execução em cada máquina.

------------------------------------------------------------
OPÇÃO 2 — Rodar direto com Python (sem gerar .exe)
------------------------------------------------------------

  pip install selenium openpyxl webdriver-manager
  python app.py

------------------------------------------------------------
COMO USAR O APLICATIVO
------------------------------------------------------------

  1. Abra o VilaDigital_UFU.exe
  2. Escolha a data inicial e final do relatório
  3. Marque as salas desejadas (padrão: todas 4)
  4. Escolha onde salvar o arquivo Excel
  5. Clique em "Gerar Relatório"
  6. Aguarde — o log mostra o progresso em tempo real
  7. Uma mensagem confirma quando terminar

O Excel gerado tem 3 abas:
  • Relatório    → resumo no formato oficial
  • Dados        → todos os registros individuais
  • Por Mês      → alunos por sala por mês

------------------------------------------------------------
DÚVIDAS
------------------------------------------------------------

"Chrome não encontrado":
  → O Chrome precisa estar instalado na máquina
  → Download: https://www.google.com/chrome

"Erro de conexão":
  → Verifique a internet e tente novamente

O app trava ou fecha sozinho:
  → Desmarque "Modo silencioso" para ver o Chrome
    e identificar onde está parando

============================================================
