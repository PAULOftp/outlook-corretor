# Instruções para o Claude neste repositório

- Falar com o utilizador (Paulo) em **português europeu**, num tom direto e colegial, sem preâmbulos nem elogios.
- Este repositório tem dois projetos:
  1. **A.P.IAv1**: assistente de voz pessoal (`apiav1.html`, `antena.py`, `server.js`, `iniciar-apiav1.bat`, `criar-atalho.vbs`, ícones e manifesto). **Antes de mexer, lê `APIAv1-CONTEXTO.md`**: tem a arquitetura, as decisões e o estado atual.
  2. Add-in do Outlook "corretor" (`index.html`, `taskpane.*`, `commands.*`, `core.js`, `manifest`).
- **Gravar sempre no GitHub:** cada alteração é um commit seguido de push no ramo de trabalho. Entregar também os ficheiros alterados ao Paulo (um ZIP quando forem vários).
- **A.P.IAv1:**
  - Não mudar os nomes das chaves do localStorage (`jarvis_*`, `anthropic_key`), senão os dados do Paulo perdem-se.
  - Manter "sem bibliotecas": HTML/CSS/JS puro e só a biblioteca padrão do Python/Node.
  - Envolver as funções existentes em vez de as reescrever.
  - Nunca enviar segredos para o git: `*.txt` (chave da API), `m365.json` e `antena_tokens.json` estão no `.gitignore`.
  - Testar antes de entregar: `node --check` em cada `<script>`, `python3 -m py_compile antena.py` e um teste no browser com respostas simuladas.
