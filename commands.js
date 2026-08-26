/* Corretor IA — botão do friso que corrige diretamente, sem abrir o painel.
   Usa sempre o modo "Corrigir ortografia e gramática". */
"use strict";

var CHAVE_AVISO = "corretorIA";

function aviso(item, texto, erro) {
  var msg = erro
    ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: texto.slice(0, 150) }
    : {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: texto.slice(0, 150),
        icon: "icon16",
        persistent: false
      };
  try { item.notificationMessages.replaceAsync(CHAVE_AVISO, msg, function () {}); }
  catch (e) { try { item.notificationMessages.addAsync(CHAVE_AVISO, msg, function () {}); } catch (e2) {} }
}

function corrigirTextoDireto(event) {
  var item = Office.context.mailbox.item;
  var s = Office.context.roamingSettings;

  var cfg = {
    provider: s.get("provider") || "google",
    key: s.get("apiKey") || "",
    model: s.get("model") || CORRETOR.DEFAULT_MODEL[s.get("provider") || "google"],
    signature: s.get("signature") || ""
  };

  if (!cfg.key) {
    aviso(item, "Falta a chave de API. Abra o painel Corretor IA e guarde-a nas Definições.", true);
    event.completed();
    return;
  }

  aviso(item, "A corrigir o texto selecionado…", false);

  CORRETOR.corrigirSelecao(item, cfg, "corrigir")
    .then(function () {
      aviso(item, "Texto corrigido. Ctrl+Z anula.", false);
    })
    .catch(function (e) {
      aviso(item, e.message || "Não foi possível corrigir.", true);
    })
    .then(function () { event.completed(); });
}

Office.onReady(function () {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("corrigirTextoDireto", corrigirTextoDireto);
  }
});

/* Compatibilidade com o runtime antigo, que procura a função no global. */
if (typeof window !== "undefined") window.corrigirTextoDireto = corrigirTextoDireto;
