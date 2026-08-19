/* Corretor IA — painel do add-in do Outlook */
"use strict";

var settings = null;
var lastSource = null; // { kind: "selection" | "body", tail: string }

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
  "acrescentas saudações ou despedidas que não existam. Se o texto já estiver correto, " +
  "devolve-o inalterado.";

/* ---------- arranque ---------- */

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) return;

  settings = {
    get: function (k) { return Office.context.roamingSettings.get(k); },
    set: function (k, v) { Office.context.roamingSettings.set(k, v); },
    save: function (cb) { Office.context.roamingSettings.saveAsync(cb); }
  };

  migrarModeloAntigo();

  el("provider").value = settings.get("provider") || "google";
  el("apiKey").value = settings.get("apiKey") || "";
  el("model").value = settings.get("model") || DEFAULT_MODEL[el("provider").value];
  el("signature").value = settings.get("signature") || "";

  if (!settings.get("apiKey")) {
    el("settings").open = true;
    setStatus("Comece por guardar a sua chave de API nas definições.", "err");
  }

  el("provider").onchange = function () {
    el("model").value = DEFAULT_MODEL[this.value];
  };
  el("mode").onchange = function () {
    el("customWrap").classList.toggle("hidden", this.value !== "custom");
  };
  el("save").onclick = saveSettings;
  el("run").onclick = run;
  el("apply").onclick = applyResult;
  el("copy").onclick = copyResult;
});

function el(id) { return document.getElementById(id); }

/* Modelos descontinuados guardados nas definicoes: substitui pelo atual. */
function migrarModeloAntigo() {
  var m = settings.get("model");
  if (!m) return;
  var obsoleto = /^(gemini-(1\.|2\.)|models\/gemini-(1\.|2\.))/.test(m);
  if (!obsoleto) return;
  settings.set("model", DEFAULT_MODEL.google);
  settings.save(function () {});
}

function setStatus(msg, cls) {
  var s = el("status");
  s.textContent = msg || "";
  s.className = "status" + (cls ? " " + cls : "");
}

function saveSettings() {
  settings.set("provider", el("provider").value);
  settings.set("apiKey", el("apiKey").value.trim());
  settings.set("model", el("model").value.trim() || DEFAULT_MODEL[el("provider").value]);
  settings.set("signature", el("signature").value.trim());
  settings.save(function (r) {
    if (r.status === Office.AsyncResultStatus.Succeeded) {
      setStatus("Definições guardadas.", "ok");
      el("settings").open = false;
    } else {
      setStatus("Não foi possível guardar: " + r.error.message, "err");
    }
  });
}

/* ---------- ler o texto ---------- */

function getSourceText() {
  return new Promise(function (resolve, reject) {
    var item = Office.context.mailbox.item;

    item.getSelectedDataAsync(Office.CoercionType.Text, function (res) {
      var sel = (res.status === Office.AsyncResultStatus.Succeeded && res.value && res.value.data) || "";
      if (sel.trim().length > 0) {
        lastSource = { kind: "selection", tail: "" };
        resolve(sel);
        return;
      }
      item.body.getAsync(Office.CoercionType.Text, function (b) {
        if (b.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Não foi possível ler o email: " + b.error.message));
          return;
        }
        var full = b.value || "";
        var sig = (settings.get("signature") || "").trim();
        var head = full, tail = "";
        if (sig) {
          var i = full.indexOf(sig);
          if (i > -1) { head = full.slice(0, i); tail = full.slice(i); }
        }
        if (!head.trim()) { reject(new Error("O email está vazio.")); return; }
        lastSource = { kind: "body", tail: tail };
        resolve(head);
      });
    });
  });
}

/* ---------- chamada à API ---------- */

function buildPrompt(text) {
  var mode = el("mode").value;
  var instruction = mode === "custom"
    ? (el("customPrompt").value.trim() || INSTRUCTIONS.corrigir)
    : INSTRUCTIONS[mode];
  return instruction + "\n\n--- TEXTO ---\n" + text;
}

function callApi(prompt) {
  var provider = settings.get("provider") || "google";
  var key = settings.get("apiKey") || "";
  var model = settings.get("model") || DEFAULT_MODEL[provider];
  if (!key) return Promise.reject(new Error("Falta a chave de API (Definições)."));

  var url, headers, payload, pick;

  if (provider === "google") {
    url = "https://generativelanguage.googleapis.com/v1beta/models/" +
          encodeURIComponent(model) + ":generateContent";
    headers = { "content-type": "application/json", "x-goog-api-key": key };
    payload = {
      system_instruction: { parts: [{ text: SYSTEM_PROMPT }] },
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2 }
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
      model: model,
      temperature: 0.2,
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
      model: model,
      max_tokens: 4000,
      system: SYSTEM_PROMPT,
      messages: [{ role: "user", content: prompt }]
    };
    pick = function (d) { return (d.content || []).map(function (c) { return c.text || ""; }).join(""); };
  } else {
    url = "https://api.openai.com/v1/chat/completions";
    headers = { "content-type": "application/json", "authorization": "Bearer " + key };
    payload = {
      model: model,
      temperature: 0.2,
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: prompt }
      ]
    };
    pick = function (d) { return d.choices[0].message.content; };
  }

  return fetch(url, { method: "POST", headers: headers, body: JSON.stringify(payload) })
    .then(function (r) {
      return r.text().then(function (t) {
        if (!r.ok) {
          var msg = t;
          try { msg = JSON.parse(t).error.message; } catch (e) {}
          throw new Error("API " + r.status + ": " + msg);
        }
        return pick(JSON.parse(t));
      });
    })
    .then(function (out) { return String(out || "").trim(); });
}

/* ---------- fluxo principal ---------- */

function run() {
  el("run").disabled = true;
  el("result").classList.add("hidden");
  el("actions").classList.add("hidden");
  setStatus("A processar…");

  getSourceText()
    .then(function (text) {
      setStatus("A processar… (" + (lastSource.kind === "selection" ? "texto selecionado" : "email completo") + ")");
      return callApi(buildPrompt(text));
    })
    .then(function (out) {
      if (!out) throw new Error("A resposta veio vazia.");
      el("result").textContent = out;
      el("result").classList.remove("hidden");
      el("actions").classList.remove("hidden");
      setStatus("Pronto. Reveja e substitua se concordar.", "ok");
    })
    .catch(function (e) {
      setStatus(e.message, "err");
    })
    .then(function () { el("run").disabled = false; });
}

function applyResult() {
  var text = el("result").textContent;
  var item = Office.context.mailbox.item;

  function done(r) {
    if (r.status === Office.AsyncResultStatus.Succeeded) {
      setStatus("Texto substituído no email.", "ok");
    } else {
      setStatus("Não foi possível substituir: " + r.error.message, "err");
    }
  }

  if (lastSource && lastSource.kind === "selection") {
    item.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, done);
  } else {
    var full = text + (lastSource && lastSource.tail ? "\n" + lastSource.tail : "");
    item.body.setAsync(full, { coercionType: Office.CoercionType.Text }, done);
  }
}

function copyResult() {
  var text = el("result").textContent;
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text)
      .then(function () { setStatus("Copiado.", "ok"); })
      .catch(function () { selectResult(); });
  } else {
    selectResult();
  }
}

function selectResult() {
  var range = document.createRange();
  range.selectNodeContents(el("result"));
  var sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  setStatus("Selecionado — prima Ctrl+C para copiar.");
}
