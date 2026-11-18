/*
 * Versão de diagnóstico: ao clicar em Enviar, mostra no alerta do Outlook
 * o resultado de cada etapa. Não envia nenhum dado para fora do Outlook.
 */
/* global Office */
function diagHandler(event) {
  var feito = false;
  var etapas = ["1) Suplemento carregou: OK"];
  function fim() {
    if (feito) return;
    feito = true;
    event.completed({ allowEvent: false, errorMessage: "DIAGNÓSTICO — " + etapas.join(" | ") });
  }
  setTimeout(function () { etapas.push("tempo total esgotado"); fim(); }, 12000);

  function passo(nome, fn) {
    return new Promise(function (resolve) {
      var t0 = Date.now();
      var timer = setTimeout(function () { etapas.push(nome + ": TRAVOU"); resolve(); }, 3000);
      try {
        fn(function (r) {
          clearTimeout(timer);
          var ms = Date.now() - t0;
          etapas.push(nome + ": " + (r.status === Office.AsyncResultStatus.Succeeded ? "OK" : "ERRO " + (r.error && r.error.message)) + " (" + ms + "ms)");
          resolve();
        });
      } catch (e) {
        clearTimeout(timer);
        etapas.push(nome + ": EXCEÇÃO " + e.message);
        resolve();
      }
    });
  }

  var item;
  try {
    item = Office.context.mailbox.item;
    var req = Office.context.requirements;
    etapas.push("2) Mailbox 1.12/1.14/1.15: " +
      ["1.12", "1.14", "1.15"].map(function (v) { return req.isSetSupported("Mailbox", v) ? "sim" : "não"; }).join("/"));
    etapas.push("sessionData: " + (item.sessionData ? "sim" : "não"));
  } catch (e) {
    etapas.push("2) EXCEÇÃO " + e.message);
    fim();
    return;
  }

  passo("3) Assunto", function (cb) { item.subject.getAsync(cb); })
    .then(function () { return passo("4) Corpo", function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); }); })
    .then(function () {
      if (!item.sessionData) return;
      return passo("5) Sessão (ler)", function (cb) { item.sessionData.getAsync("diagTeste", cb); });
    })
    .then(function () {
      if (!item.sessionData) return;
      return passo("6) Sessão (gravar)", function (cb) { item.sessionData.setAsync("diagTeste", "1", cb); });
    })
    .then(fim, fim);
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onMessageSendHandler", diagHandler);
}
