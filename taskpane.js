/* Corretor IA — painel. Só interface; as regras estão em core.js. */
"use strict";

var settings = null;

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) return;
  if (!document.getElementById("run")) return;   /* não é o painel */

  settings = {
    get: function (k) { return Office.context.roamingSettings.get(k); },
    set: function (k, v) { Office.context.roamingSettings.set(k, v); },
    save: function (cb) { Office.context.roamingSettings.saveAsync(cb); }
  };

  migrarModeloAntigo();

  el("provider").value = settings.get("provider") || "google";
  el("apiKey").value = settings.get("apiKey") || "";
  el("model").value = settings.get("model") || CORRETOR.DEFAULT_MODEL[el("provider").value];
  el("signature").value = settings.get("signature") || "";

  if (!settings.get("apiKey")) {
    el("settings").open = true;
    setStatus("Comece por guardar a sua chave de API nas definições.", "err");
  }

  el("provider").onchange = function () {
    el("model").value = CORRETOR.DEFAULT_MODEL[this.value];
  };
  el("mode").onchange = function () {
    el("customWrap").classList.toggle("hidden", this.value !== "custom");
  };
  el("save").onclick = saveSettings;
  el("run").onclick = run;
  el("copy").onclick = copyResult;
});

function el(id) { return document.getElementById(id); }

/* Modelos descontinuados guardados nas definições: substitui pelo atual. */
function migrarModeloAntigo() {
  var m = settings.get("model");
  if (!m) return;
  if (!/^(gemini-(1\.|2\.)|models\/gemini-(1\.|2\.))/.test(m)) return;
  settings.set("model", CORRETOR.DEFAULT_MODEL.google);
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
  settings.set("model", el("model").value.trim() || CORRETOR.DEFAULT_MODEL[el("provider").value]);
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

function run() {
  el("run").disabled = true;
  el("result").classList.add("hidden");
  el("actions").classList.add("hidden");
  setStatus("A processar o texto selecionado…");

  var cfg = {
    provider: settings.get("provider") || "google",
    key: settings.get("apiKey") || "",
    model: settings.get("model") || CORRETOR.DEFAULT_MODEL[settings.get("provider") || "google"],
    signature: settings.get("signature") || ""
  };

  CORRETOR.corrigirSelecao(
    Office.context.mailbox.item, cfg, el("mode").value, el("customPrompt").value.trim()
  )
    .then(function (corrigido) {
      el("result").textContent = corrigido;
      el("result").classList.remove("hidden");
      el("actions").classList.remove("hidden");
      setStatus("Substituído no email. Ctrl+Z anula.", "ok");
    })
    .catch(function (e) { setStatus(e.message, "err"); })
    .then(function () { el("run").disabled = false; });
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
