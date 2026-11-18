/*
 * Conferência de cadastro — painel com o checklist obrigatório.
 *
 * "Enviar e-mail":
 *   - se faltar alguma conferência, mostra o aviso e não envia;
 *   - com as 3 marcadas, grava a confirmação em sessionData e dispara o envio
 *     (item.sendAsync, quando disponível). O OnMessageSend encontra a confirmação
 *     e libera o envio. Sem sendAsync, o usuário clica em "Enviar" no Outlook.
 *
 * "Revisar antes de enviar": descarta qualquer confirmação e fecha o painel.
 */

/* global Office */

(function () {
  "use strict";

  var SESSION_KEY_CONFIRMED = "conferenciaCadastroConfirmada";
  var SESSION_KEY_FOUND = "conferenciaCadastroPalavras";

  var MSG_OBRIGATORIO =
    "Para enviar este e-mail, é necessário confirmar todas as conferências obrigatórias.";

  var checkboxes, btnEnviar, btnRevisar, alertEl, progressFill, progressLabel, keywordsEl;

  function $(id) { return document.getElementById(id); }

  function item() {
    return typeof Office !== "undefined" && Office.context && Office.context.mailbox
      ? Office.context.mailbox.item
      : null;
  }

  /* ------------------------------- UI helpers ------------------------------- */

  function mostrarAlerta(texto, tipo) {
    alertEl.textContent = texto;
    alertEl.className = "alert alert--" + tipo;
    alertEl.hidden = false;
  }

  function ocultarAlerta() {
    alertEl.hidden = true;
    alertEl.textContent = "";
  }

  function marcadas() {
    return checkboxes.filter(function (c) { return c.checked; }).length;
  }

  function atualizarProgresso() {
    var total = checkboxes.length;
    var n = marcadas();
    progressFill.style.width = (n / total) * 100 + "%";
    progressLabel.textContent = n + " de " + total;

    checkboxes.forEach(function (c) {
      if (c.checked) c.closest(".check").classList.remove("is-missing");
    });
    if (n === total && alertEl.classList.contains("alert--error")) ocultarAlerta();
  }

  function destacarPendentes() {
    checkboxes.forEach(function (c) {
      c.closest(".check").classList.toggle("is-missing", !c.checked);
    });
    var primeira = checkboxes.filter(function (c) { return !c.checked; })[0];
    if (primeira) primeira.focus();
  }

  function setOcupado(ocupado) {
    btnEnviar.disabled = ocupado;
    btnRevisar.disabled = ocupado;
    btnEnviar.textContent = ocupado ? "Enviando…" : "Enviar e-mail";
  }

  /* ---------------------------- Office helpers ----------------------------- */

  function sessionSet(chave, valor) {
    return new Promise(function (resolve, reject) {
      var it = item();
      if (!it || !it.sessionData) return reject(new Error("sessionData indisponível"));
      it.sessionData.setAsync(chave, valor, function (r) {
        r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error);
      });
    });
  }

  function sessionGet(chave) {
    return new Promise(function (resolve) {
      var it = item();
      if (!it || !it.sessionData) return resolve(null);
      it.sessionData.getAsync(chave, function (r) {
        resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null);
      });
    });
  }

  function sessionRemove(chave) {
    return new Promise(function (resolve) {
      var it = item();
      if (!it || !it.sessionData) return resolve();
      it.sessionData.removeAsync(chave, function () { resolve(); });
    });
  }

  function enviarItem() {
    return new Promise(function (resolve, reject) {
      var it = item();
      if (!it || typeof it.sendAsync !== "function") return resolve(false);
      it.sendAsync(function (r) {
        r.status === Office.AsyncResultStatus.Succeeded ? resolve(true) : reject(r.error);
      });
    });
  }

  function fecharPainel() {
    try {
      if (Office.context.ui && typeof Office.context.ui.closeContainer === "function") {
        Office.context.ui.closeContainer();
        return true;
      }
    } catch (e) { /* host não suporta fechar o painel */ }
    return false;
  }

  /* -------------------------------- Ações --------------------------------- */

  function onEnviar() {
    if (marcadas() < checkboxes.length) {
      mostrarAlerta(MSG_OBRIGATORIO, "error");
      destacarPendentes();
      return;
    }

    ocultarAlerta();
    setOcupado(true);

    sessionSet(SESSION_KEY_CONFIRMED, "true")
      .then(enviarItem)
      .then(function (enviado) {
        if (enviado) {
          mostrarAlerta("Conferências confirmadas. E-mail enviado.", "success");
        } else {
          setOcupado(false);
          mostrarAlerta(
            "Conferências confirmadas. Clique em \"Enviar\" no Outlook para concluir o envio.",
            "success"
          );
        }
      })
      .catch(function (erro) {
        setOcupado(false);
        console.error(erro);
        mostrarAlerta(
          "Não foi possível concluir o envio automaticamente. Clique em \"Enviar\" no Outlook.",
          "info"
        );
      });
  }

  function onRevisar() {
    sessionRemove(SESSION_KEY_CONFIRMED).then(function () {
      checkboxes.forEach(function (c) { c.checked = false; });
      atualizarProgresso();
      if (!fecharPainel()) {
        mostrarAlerta("Revise o e-mail e clique em \"Enviar\" quando estiver pronto.", "info");
      }
    });
  }

  function mostrarPalavrasEncontradas() {
    sessionGet(SESSION_KEY_FOUND).then(function (valor) {
      if (!valor) return;
      keywordsEl.textContent = "Termos identificados: ";
      valor.split(",").forEach(function (p) {
        var tag = document.createElement("strong");
        tag.textContent = p;
        keywordsEl.appendChild(tag);
      });
      keywordsEl.hidden = false;
    });
  }

  /* ------------------------------ Inicialização ----------------------------- */

  function init() {
    checkboxes = Array.prototype.slice.call(document.querySelectorAll("input[name='conferencia']"));
    btnEnviar = $("btnEnviar");
    btnRevisar = $("btnRevisar");
    alertEl = $("alert");
    progressFill = $("progressFill");
    progressLabel = $("progressLabel");
    keywordsEl = $("keywords");

    checkboxes.forEach(function (c) { c.addEventListener("change", atualizarProgresso); });
    btnEnviar.addEventListener("click", onEnviar);
    btnRevisar.addEventListener("click", onRevisar);

    // Sempre começa desmarcado: cada envio exige nova conferência.
    sessionRemove(SESSION_KEY_CONFIRMED);
    mostrarPalavrasEncontradas();
    atualizarProgresso();
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(init);
  } else {
    document.addEventListener("DOMContentLoaded", init);
  }
})();
