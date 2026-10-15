# A.P.IAv1: contexto do projeto (passagem entre contas Claude)

Assistente de voz pessoal do **António Paulo (Paulo)**, Diretor Comercial na FTP Porto, 49 anos, vive em Avanca.
Começou como "Jarvis" (tutorial em 2 partes) e evoluiu por iterações: Rissa → PauloIA → **A.P.IAv1**.
O assistente chama-se **A.P.IAv1** e é chamado em voz alta com **"Olá IA"**. É feminina, com personalidade descontraída, e trata o Paulo por tu.

> Este repositório também contém um add-in do Outlook (`index.html`, `taskpane.*`, `commands.*`, `core.js`, `manifest`), que é um projeto **separado**. Os ficheiros da A.P.IAv1 são só os listados abaixo.

## Preferências do Paulo (respeitar sempre)
- Responder em **português europeu**, num tom direto e colegial, sem preâmbulos nem elogios. Nível avançado.
- O conteúdo da app e a voz também são em **pt-PT** (nunca pt-BR).
- **Gravar sempre no GitHub**: cada alteração é um commit seguido de push. Ele também quer os ficheiros entregues na conversa (de preferência num ZIP).
- Instalação no PC dele: pasta **`C:\A.P.IAv1`**, Windows, Edge em modo aplicação, Python instalado. A tarefa agendada do Outlook corre na **outra conta Claude** (a que tem o conector Microsoft 365) e grava `C:\A.P.IAv1\m365.json`.

## Ficheiros
| Ficheiro | Função |
|---|---|
| `apiav1.html` | A app inteira (HTML + CSS + JS puro, sem bibliotecas). Inclui a fotografia da agente embutida em base64 (`<script type="text/plain" id="avatarData">`). |
| `antena.py` | Antena local em **Python (só biblioteca padrão)**, em `127.0.0.1:8080`. Serve a página e funciona como proxy de agenda, notícias, YouTube e IMAP, além de ler `m365.json` e a chave de um `.txt`. |
| `server.js` | A mesma antena em **Node puro** (porta 4242), para abrir a página com duplo clique (`file://`). É secundária. |
| `iniciar-apiav1.bat` | Fecha antenas antigas (PowerShell: só processos python com `antena.py`), arranca `pythonw antena.py` sem janela e abre `msedge --app=http://localhost:8080/apiav1.html`. |
| `criar-atalho.vbs` | Cria o atalho "A.P.IAv1" (com ícone) no Ambiente de Trabalho e no menu Iniciar. |
| `apiav1.webmanifest`, `apiav1-192/512.png`, `apiav1.ico` | Manifesto e ícones da aplicação (instalável no Edge). |
| `apiav1-avatar.jpg` | Fotografia da agente (1000×1300). A página usa a cópia embutida. |
| Fora do git (`.gitignore`) | `*.txt` (chave da API, ex. `cheveapi.txt`), `m365.json`, `antena_tokens.json`. |

## Arquitetura da página (`apiav1.html`)
- **Script 1 (Parte 1, "Jarvis")**:
  - `CONFIG` (nome, wake word e variantes, `lang: "pt-PT"`, `model`, `geminiModels`), `AREA`, `NOTES` (14 notas de base) e `REL`.
  - Voz: `speak`, `say`, reconhecimento (`SpeechRecognition` contínuo) e `splitWake`/`handleHeard`.
  - Cérebro: `ask` → `callGemini`/`callClaude`. O fornecedor é detetado pela chave: `AIza…` = Gemini, `sk-ant-…` = Claude.
  - Second Brain: grafo SVG injetado via `innerHTML` de uma DIV. Memória viva com `[[SAVE:area|titulo|texto]]`.
  - Notas: `{id, area, title, body (manual), auto: [{t, fonte, d}] (aprendizagem automática)}`. O editor tem dois separadores.
- **Script 2 (IIFE de upgrade)**: **envolve** (não reescreve) `processCommand`, `buildSystem`, `speak`, `handleHeard`, `setUser` e `setBot`.
  - **Agenda:** Outlook via `m365.json` + iCal pela antena. Parser ICS próprio (TZID, UTC, dia inteiro, RRULE, EXDATE).
  - **E-mails:** Outlook via `m365.json` + Gmail por IMAP (senha de app). Triagem AÇÃO/INFO/RUÍDO numa chamada por lote, com cache `jarvis_email_triage`. A mesma chamada extrai factos sobre a FTP Porto.
  - **Notícias:** RSS do Google News PT (`hl=pt-PT&gl=PT&ceid=PT:pt-150`), 5 temas.
  - **Morning Digest:** uma chamada por dia, com o tempo em Avanca (Open-Meteo) e um modelo local como alternativa.
  - **APRENDER:**
    - ZIP do LinkedIn: CSV lidos localmente, sem IA, e a IA enriquece se disponível.
    - ZIP do Facebook: HTML ou JSON, só perfil, publicações, páginas e grupos; nunca mensagens.
    - Texto colado.
    - Leitor de ZIP nativo (`DecompressionStream`).
  - **Música:** "põe música de X" → `/api/youtube` (pesquisa sem chave) → leitor `youtube-nocookie` embutido. Baixa o volume quando ela fala e, com música a tocar, só reage à wake word.
  - **Voz:**
    - Motor "Automático": voz Natural do Edge (Raquel pt-PT) se existir; senão Gemini TTS (voz Kore, PCM com Web Audio, a 1.ª frase em paralelo); senão a voz do browser.
    - Indicador `#vozChip`.
    - Anti-eco: `abort()` do reconhecimento enquanto fala, filtro de eco e margem de 1,5 s.
  - **Avatar** (⚙ → VOZ → AVATAR): **Fotografia** por defeito, com a boca animada pelo volume real da voz (o lábio inferior e o queixo descem por colunas), cabeça viva (oscilações, acenos, inclinação), olhar com sacadas e piscar. Alternativas: rosto de partículas e esfera.
  - **Definições ⚙:** motor (económico, equilibrado ou máximo; Gemini Flash-Lite/Flash/Pro ou Claude Haiku/Sonnet/Opus), modo economia, voz, Microsoft 365, agendas iCal, Gmail, notícias, digest e **Exportar/Importar dados** (cópia do localStorage `jarvis_*` + `anthropic_key`).
  - **Chave da API:** lida de qualquer `.txt` da pasta pela rota `/api/chave` (ex.: `cheveapi.txt`).
- **localStorage**: as chaves são `jarvis_*` e `anthropic_key`. **Não mudar estes nomes**, senão os dados guardados do Paulo perdem-se.

## Decisões importantes (e porquê)
- **Office 365 sem Azure:** o registo de app no Azure exigia aprovação do administrador da FTP Porto, por isso foi abandonado. Uma **tarefa agendada na app Claude do PC** (conector Microsoft 365, já aprovado) grava `m365.json` na pasta de hora a hora. O texto do pedido está em ⚙ → COPIAR PEDIDO e inclui a triagem e o campo `conhecimento`.
- **Gemini gratuito** em vez da API da Anthropic, porque o Paulo não tem chave paga. A quota gratuita da voz Gemini é baixa, e por isso existe o motor Automático com a voz Natural do Edge.
- **App do Windows** = Edge em modo `--app` + manifesto. A antena corre com `pythonw`, sem janela.
- **Segurança da antena:** só escuta em 127.0.0.1, valida o cabeçalho Host e só aceita a origem localhost ou `null`. Tem allowlist de domínios no proxy e de hosts IMAP, e não regista credenciais nos logs.
- **LinkedIn e Facebook:** só pela **cópia oficial dos dados** (ZIP). Não há acesso automático à conta, porque essas redes bloqueiam e os termos o proíbem.

## Estado e o que falta confirmar no PC do Paulo
- Tudo foi testado no Chromium headless com respostas simuladas. **Ainda não foi confirmado com contas e serviços reais:**
  - a tarefa agendada do Claude a gravar `m365.json`;
  - o Gmail por IMAP;
  - o leitor do YouTube;
  - o `.bat` no Windows (fechar a antena antiga);
  - a voz Natural do Edge;
  - o aspeto final dos movimentos do rosto.
- **Avatar:** as coordenadas (boca: `mx0/mx1/split/end`; olhos: `eyes[]`) estão afinadas para `apiav1-avatar.jpg`. Para mudar de cara, é preciso recalibrar.
- **Histórico:** o histórico completo está nos commits deste ramo e no PR #1 (`PAULOftp/outlook-corretor`).

## Como testar (padrão usado)
1. Arrancar a antena: `python3 antena.py` (porta 8080).
2. Correr o Playwright/Chromium (`executablePath: '/opt/pw-browsers/chromium'` no ambiente cloud), com respostas simuladas via `page.route` para Gemini, proxy, `/emails` e TTS.
3. Verificar a sintaxe de cada `<script>` com `node --check` e do `antena.py` com `python3 -m py_compile`.
