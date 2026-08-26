/* Corretor IA — núcleo partilhado pelo painel e pelo botão do friso.
   Tudo o que define COMO se corrige vive aqui, para não haver duas versões. */
"use strict";

var CORRETOR = (function () {

  var DEFAULT_MODEL = {
    google: "gemini-3.6-flash",
    groq: "llama-3.3-70b-versatile",
    openrouter: "meta-llama/llama-3.3-70b-instruct:free",
    openai: "gpt-4.1-mini",
    anthropic: "claude-haiku-4-5-20251001"
  };
  
  var INSTRUCTIONS = {
    corrigir: "Corrige com rigor todos os erros de ortografia, gramática, concordância, regência, acentuação e pontuação, em português europeu. Não alteres o estilo, o tom, o vocabulário nem a estrutura. Se o texto estiver noutra língua, mantém essa língua.",
    melhorar: "Corrige todos os erros e melhora a fluidez e a clareza da escrita, em português europeu correto e natural, mantendo o sentido e o tom. Se o texto estiver noutra língua, mantém essa língua.",
    formal: "Reescreve num registo mais formal e profissional, adequado a correspondência de trabalho. Mantém o sentido e a língua original.",
    simples: "Reescreve de forma mais clara e direta: frases curtas, sem redundâncias nem palavras desnecessárias. Mantém o sentido e a língua original.",
    encurtar: "Reduz o texto para cerca de metade do comprimento, mantendo toda a informação essencial, o tom e a língua original.",
    en: "Traduz o texto para inglês, com registo natural e adequado a correspondência profissional.",
    pt: "Traduz o texto para português europeu (norma de Portugal), com registo natural e adequado a correspondência profissional."
  };
  
  var SYSTEM_PROMPT =
    "És um revisor profissional de texto de emails, especialista em português europeu.\n\n" +
    "NORMA LINGUÍSTICA (regra absoluta): quando o texto está em português, o resultado tem de " +
    "estar em PORTUGUÊS EUROPEU — norma de Portugal, Acordo Ortográfico de 1990 tal como " +
    "aplicado em Portugal. Nunca devolvas português do Brasil. Em concreto:\n" +
    "- Gerúndio: usa \"estou a fazer\", \"continuamos a analisar\" (nunca \"estou fazendo\").\n" +
    "- Colocação dos pronomes: ênclise por defeito (\"envio-lhe\", \"chamo-me\"); próclise só " +
    "quando há atrator (negação, advérbio, conjunção subordinativa, pronome relativo, " +
    "interrogativo): \"não lhe envio\", \"já lhe enviei\", \"que me disse\".\n" +
    "- Vocabulário de Portugal: ecrã, ficheiro, telemóvel, autocarro, comboio, morada, " +
    "encomenda, equipa, casa de banho, rececionista, utilizador, gestor, faturação, " +
    "IVA, sítio (web), anexo, reunião, receção.\n" +
    "- Formas de tratamento de Portugal: \"o Senhor\"/\"a Senhora\", \"V. Exa.\", 3.ª pessoa; " +
    "nunca \"você\" à brasileira nem \"a gente\" com valor de \"nós\".\n" +
    "- Ortografia AO90 na variante de Portugal: receção, direção, setor, projeto, atual, " +
    "objetivo, exceção, adoção, ótimo, contacto, facto, teto, húmido, connosco.\n" +
    "- Pontuação e espaçamento à portuguesa; datas 19/08/2026; decimais com vírgula; " +
    "milhares com espaço; € depois do valor (1 250,00 €).\n\n" +
    "RIGOR GRAMATICAL: corrige concordância nominal e verbal, regência verbal e nominal, " +
    "uso de crase/contrações (à, às, ao, aos, do, no, pelo), tempos e modos verbais, " +
    "conjuntivo depois de \"esperar que\", \"caso\", \"embora\", \"para que\", acentuação, " +
    "hífens, maiúsculas e minúsculas, e pontuação. Elimina pleonasmos e concordâncias " +
    "erradas do tipo \"houveram\", \"há-de haver muitos\", \"a nível de\".\n\n" +
    "FORMATO DA RESPOSTA: devolves EXCLUSIVAMENTE o texto resultante — sem introduções, " +
    "sem comentários, sem aspas à volta, sem marcadores de código. Preservas as quebras de " +
    "linha e a estrutura de parágrafos do original. Não inventas conteúdo novo nem " +
    "acrescentas saudações ou despedidas que não existam. NÃO REMOVES nenhuma frase, linha " +
    "nem informação do original — todas as frases têm de aparecer no resultado. Se o texto " +
    "já estiver correto, devolve-o inalterado.";
  
  /* Marcadores que indicam o inicio da assinatura, aviso legal ou historico
     da conversa. Tudo a partir do primeiro marcador encontrado fica INTOCADO. */
  var MARCADORES_FIM = [
    /^[ \t]*--[ \t]*$/m,
    /^[ \t]*_{5,}[ \t]*$/m,
    /^[ \t]*-{5,}[ \t]*(mensagem original|original message|forwarded message)/im,
    /^[ \t]*(com os |com |os )?(meus |nossos )?(melhores )?cumprimentos\b/im,
    /^[ \t]*(atenciosamente|com estima|melhores saudacoes|melhores saudações)\b/im,
    /^[ \t]*(best regards|kind regards|regards|sincerely|yours (faithfully|sincerely))\b/im,
    /^[ \t]*aviso legal\b/im,
    /^[ \t]*(disclaimer|confidencialidade|privileged and confidential)\b/im,
    /^[ \t]*esta (mensagem|comunicação) é confidencial\b/im,
    /^[ \t]*(de|from|remetente):[ \t]*\S+/im,
    /^[ \t]*(enviada?|sent|em)[ \t]*:[ \t]*\S+/im
  ];
  
  /* Le a fonte usada no trecho selecionado, para o texto corrigido ficar igual. */
  function estiloDaSelecao(html) {
    var out = "";
    var f = html.match(/font-family\s*:\s*([^;"'<>]+)/i);
    var t = html.match(/font-size\s*:\s*([^;"'<>]+)/i);
    var c = html.match(/(?:^|[;"'\s])color\s*:\s*([^;"'<>]+)/i);
    if (f) out += "font-family:" + f[1].trim() + ";";
    if (t) out += "font-size:" + t[1].trim() + ";";
    if (c) out += "color:" + c[1].trim() + ";";
    return out;
  }
  
  /* Converte o texto corrigido em HTML: uma linha = um paragrafo.
     E assim que as mudancas de linha (Enter) do original sao respeitadas. */
  function textoParaHtml(t, estilo) {
    var st = ' style="margin:0;' + (estilo || "") + '"';
    var linhas = String(t).replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    var partes = [];
    for (var i = 0; i < linhas.length; i++) {
      var e = linhas[i]
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
      partes.push("<p" + st + ">" + (e.trim() ? e : "<br>") + "</p>");
    }
    return partes.join("");
  }
  
  /* Devolve o indice onde comeca a assinatura, ou -1 se nao encontrar. */
  function inicioDaAssinatura(texto) {
    var melhor = -1;
    for (var i = 0; i < MARCADORES_FIM.length; i++) {
      var m = MARCADORES_FIM[i].exec(texto);
      if (m && m.index > 0 && (melhor === -1 || m.index < melhor)) melhor = m.index;
    }
    return melhor;
  }

  var MODOS_LINHA_A_LINHA = {
    corrigir: 1, melhorar: 1, formal: 1, simples: 1, en: 1, pt: 1
  };

  /* Devolve { prompt, linhas }. `linhas` é null quando não se numera. */
  function montarPrompt(mode, texto, instrucaoCustom) {
    var instruction = mode === "custom"
      ? (instrucaoCustom || INSTRUCTIONS.corrigir)
      : (INSTRUCTIONS[mode] || INSTRUCTIONS.corrigir);

    var linhas = String(texto).replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    var comTexto = linhas.filter(function (l) { return l.trim(); });

    if (MODOS_LINHA_A_LINHA[mode] && comTexto.length > 1) {
      var numeradas = comTexto.map(function (l, i) { return (i + 1) + "| " + l; }).join("\n");
      return {
        linhas: linhas,
        prompt: instruction +
          "\n\nO texto vem numerado, uma linha por número. Devolve EXATAMENTE o mesmo " +
          "número de linhas, cada uma com o seu prefixo \"N| \" igual ao original, pela " +
          "mesma ordem. Não juntes duas linhas numa só, não dividas uma linha em duas e " +
          "não acrescentes nem removas linhas. Corrige apenas o conteúdo de cada linha.\n\n" +
          "--- TEXTO ---\n" + numeradas
      };
    }
    return { linhas: null, prompt: instruction + "\n\n--- TEXTO ---\n" + texto };
  }

  function reconstruir(out, linhas) {
    if (!linhas) return null;
    var mapa = {};
    var re = /^\s*(\d+)\s*\|\s?(.*)$/;
    var linhasOut = String(out).replace(/\r\n/g, "\n").split("\n");
    for (var i = 0; i < linhasOut.length; i++) {
      var m = re.exec(linhasOut[i]);
      if (m) mapa[m[1]] = m[2];
    }
    var n = 0, res = [];
    for (var j = 0; j < linhas.length; j++) {
      if (!linhas[j].trim()) { res.push(""); continue; }
      n++;
      if (mapa[String(n)] === undefined) return null;
      res.push(mapa[String(n)]);
    }
    return res.join("\n");
  }

  /* cfg = { provider, key, model } */
  function pedirCorrecao(cfg, prompt) {
    var provider = cfg.provider || "google";
    var key = cfg.key || "";
    var model = cfg.model || DEFAULT_MODEL[provider];
    if (!key) return Promise.reject(new Error("Falta a chave de API (Definições)."));

    var url, headers, payload, pick;

    if (provider === "google") {
      url = "https://generativelanguage.googleapis.com/v1beta/models/" +
            encodeURIComponent(model) + ":generateContent";
      headers = { "content-type": "application/json", "x-goog-api-key": key };
      payload = {
        system_instruction: { parts: [{ text: SYSTEM_PROMPT }] },
        contents: [{ role: "user", parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0.2,
          maxOutputTokens: 2048,
          thinkingConfig: { thinkingLevel: "low" }
        }
      };
      pick = function (d) {
        var c = (d.candidates || [])[0];
        if (!c) throw new Error("Resposta sem conteúdo" + (d.promptFeedback ? " (bloqueada pelo filtro)" : "") + ".");
        return ((c.content && c.content.parts) || [])
          .filter(function (p) { return p.text && !p.thought; })
          .map(function (p) { return p.text; }).join("");
      };
    } else if (provider === "groq" || provider === "openrouter") {
      url = provider === "groq"
        ? "https://api.groq.com/openai/v1/chat/completions"
        : "https://openrouter.ai/api/v1/chat/completions";
      headers = { "content-type": "application/json", "authorization": "Bearer " + key };
      payload = {
        model: model, temperature: 0.2,
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: prompt }
        ]
      };
      pick = function (d) { return d.choices[0].message.content; };
    } else if (provider === "anthropic") {
      url = "https://api.anthropic.com/v1/messages";
      headers = {
        "content-type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true"
      };
      payload = {
        model: model, max_tokens: 4000, system: SYSTEM_PROMPT,
        messages: [{ role: "user", content: prompt }]
      };
      pick = function (d) { return (d.content || []).map(function (c) { return c.text || ""; }).join(""); };
    } else {
      url = "https://api.openai.com/v1/chat/completions";
      headers = { "content-type": "application/json", "authorization": "Bearer " + key };
      payload = {
        model: model, temperature: 0.2,
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: prompt }
        ]
      };
      pick = function (d) { return d.choices[0].message.content; };
    }

    function enviar(corpo) {
      return fetch(url, { method: "POST", headers: headers, body: JSON.stringify(corpo) })
        .then(function (r) {
          return r.text().then(function (t) {
            if (!r.ok) {
              var msg = t;
              try { msg = JSON.parse(t).error.message; } catch (e) {}
              throw new Error("API " + r.status + ": " + msg);
            }
            return pick(JSON.parse(t));
          });
        });
    }

    return enviar(payload)
      .catch(function (e) {
        var temThinking = payload.generationConfig && payload.generationConfig.thinkingConfig;
        if (temThinking && /thinking|unknown name|invalid json payload|not supported/i.test(e.message)) {
          delete payload.generationConfig.thinkingConfig;
          return enviar(payload);
        }
        throw e;
      })
      .then(function (out) { return String(out || "").trim(); });
  }

  /* Lê a seleção (texto + estilo) e valida. Devolve Promise<{texto, estilo}>. */
  function lerSelecao(item, marcadorManual) {
    return new Promise(function (resolve, reject) {
      item.getSelectedDataAsync(Office.CoercionType.Html, function (rh) {
        var selHtml = (rh.status === Office.AsyncResultStatus.Succeeded &&
                       rh.value && rh.value.data) || "";
        item.getSelectedDataAsync(Office.CoercionType.Text, function (res) {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error("Não foi possível ler a seleção: " + res.error.message));
            return;
          }
          var sel = (res.value && res.value.data) || "";
          if (!sel.trim()) {
            reject(new Error("Selecione o texto que escreveu e tente outra vez."));
            return;
          }
          var corte = -1;
          var sig = (marcadorManual || "").trim();
          if (sig) corte = sel.indexOf(sig);
          if (corte < 0) corte = inicioDaAssinatura(sel);
          if (corte > -1) {
            reject(new Error("A seleção inclui a assinatura ou a mensagem citada. Selecione apenas o texto que escreveu."));
            return;
          }
          resolve({ texto: sel, estilo: estiloDaSelecao(selHtml) });
        });
      });
    });
  }

  /* Escreve o resultado por cima da seleção, em HTML. */
  function escreverSelecao(item, texto, estilo) {
    return new Promise(function (resolve, reject) {
      var html = textoParaHtml(texto, estilo);
      item.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error("Não foi possível substituir: " + r.error.message));
      });
    });
  }

  /* Fluxo completo: ler → corrigir → escrever. Devolve Promise<texto corrigido>. */
  function corrigirSelecao(item, cfg, mode, instrucaoCustom) {
    var estilo = "";
    return lerSelecao(item, cfg.signature)
      .then(function (s) {
        estilo = s.estilo;
        var m = montarPrompt(mode, s.texto, instrucaoCustom);
        return pedirCorrecao(cfg, m.prompt).then(function (out) {
          if (!out) throw new Error("A resposta veio vazia.");
          return reconstruir(out, m.linhas) || out;
        });
      })
      .then(function (corrigido) {
        return escreverSelecao(item, corrigido, estilo).then(function () { return corrigido; });
      });
  }

  return {
    DEFAULT_MODEL: DEFAULT_MODEL,
    INSTRUCTIONS: INSTRUCTIONS,
    SYSTEM_PROMPT: SYSTEM_PROMPT,
    inicioDaAssinatura: inicioDaAssinatura,
    estiloDaSelecao: estiloDaSelecao,
    textoParaHtml: textoParaHtml,
    montarPrompt: montarPrompt,
    reconstruir: reconstruir,
    pedirCorrecao: pedirCorrecao,
    lerSelecao: lerSelecao,
    escreverSelecao: escreverSelecao,
    corrigirSelecao: corrigirSelecao
  };
})();

if (typeof module !== "undefined") module.exports = CORRETOR;
