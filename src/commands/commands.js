/*
 * Conferência de cadastro — manipulador do evento OnMessageSend (Smart Alerts).
 *
 * Fluxo:
 *  1. O usuário clica em "Enviar".
 *  2. Lemos assunto e corpo. Se não houver palavra-chave, o envio segue normalmente.
 *  3. Se houver, verificamos em sessionData se as 3 conferências já foram confirmadas
 *     pelo painel. Se sim, o envio segue (e a confirmação é consumida).
 *  4. Caso contrário, o envio é bloqueado e o alerta oferece o botão
 *     "Conferência de cadastro", que abre o painel com os checkboxes.
 *
 * Este arquivo roda tanto no runtime HTML (Outlook na Web / novo Outlook)
 * quanto no runtime somente-JS (Outlook clássico no Windows), por isso não usa DOM
 * nem módulos ES.
 */

/* global Office, module */

var CONFERENCIA = {
  // Palavras monitoradas (comparação sem acento e sem diferenciar maiúsculas).
  // Plurais incluídos para não deixar passar "cadastros" ou "inserções".
  KEYWORDS: [
    "cadastro", "cadastros",
    "cadastrar",
    "cadastramento", "cadastramentos",
    "insercao", "insercoes",
    "inserir"
  ],
  SESSION_KEY_CONFIRMED: "conferenciaCadastroConfirmada",
  SESSION_KEY_FOUND: "conferenciaCadastroPalavras",
  TASKPANE_COMMAND_ID: "msgComposeOpenPaneButton"
};

/** Remove acentos e coloca em minúsculas: "Inserção" -> "insercao". */
function normalizarTexto(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase();
}

var KEYWORD_REGEX = new RegExp("\\b(" + CONFERENCIA.KEYWORDS.join("|") + ")\\b", "g");

/**
 * Retorna as palavras-chave encontradas (sem repetição, na forma normalizada).
 * Usa limites de palavra para evitar falsos positivos como "cadastrado" ou "reinserir".
 */
function encontrarPalavrasChave(assunto, corpo) {
  var texto = normalizarTexto(assunto) + "\n" + normalizarTexto(corpo);
  var encontradas = [];
  var match;
  KEYWORD_REGEX.lastIndex = 0;
  while ((match = KEYWORD_REGEX.exec(texto)) !== null) {
    if (encontradas.indexOf(match[1]) === -1) {
      encontradas.push(match[1]);
    }
  }
  return encontradas;
}

/* ---------- Helpers que transformam callbacks do Office.js em Promises ---------- */

// Cada chamada ao Outlook tem limite de tempo: se não responder, segue o fluxo com erro.
function officeAsync(fn, limiteMs) {
  return new Promise(function (resolve, reject) {
    var timer = setTimeout(function () { reject(new Error("timeout")); }, limiteMs || 4000);
    fn(function (result) {
      clearTimeout(timer);
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

function lerAssunto(item) {
  return officeAsync(function (cb) { item.subject.getAsync(cb); });
}

function lerCorpo(item) {
  return officeAsync(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); });
}

function lerSessao(item, chave) {
  if (!item.sessionData) return Promise.resolve(null);
  return officeAsync(function (cb) { item.sessionData.getAsync(chave, cb); })
    .catch(function () { return null; }); // chave inexistente gera erro: tratamos como vazio
}

function gravarSessao(item, chave, valor) {
  if (!item.sessionData) return Promise.resolve();
  return officeAsync(function (cb) { item.sessionData.setAsync(chave, valor, cb); })
    .catch(function () { /* não impede o fluxo */ });
}

function removerSessao(item, chave) {
  if (!item.sessionData) return Promise.resolve();
  return officeAsync(function (cb) { item.sessionData.removeAsync(chave, cb); })
    .catch(function () { /* não impede o fluxo */ });
}

/* ---------------------------- Manipulador do evento ---------------------------- */

function onMessageSendHandler(event) {
  // Garante uma única resposta ao Outlook, mesmo em caso de erro ou demora.
  var concluido = false;
  var original = event.completed.bind(event);
  event = {
    completed: function (opcoes) {
      if (concluido) return;
      concluido = true;
      original(opcoes);
    }
  };
  setTimeout(function () {
    if (!concluido) {
      console.error("Conferência de cadastro: tempo esgotado ao analisar o e-mail.");
      event.completed({ allowEvent: true });
    }
  }, 10000);

  var item;
  try {
    item = Office.context.mailbox.item;
  } catch (e) {
    event.completed({ allowEvent: true });
    return;
  }

  Promise.all([lerAssunto(item), lerCorpo(item)])
    .then(function (valores) {
      var palavras = encontrarPalavrasChave(valores[0], valores[1]);

      // Nenhuma palavra-chave: envio normal, sem alerta.
      if (palavras.length === 0) {
        event.completed({ allowEvent: true });
        return;
      }

      return lerSessao(item, CONFERENCIA.SESSION_KEY_CONFIRMED).then(function (confirmado) {
        if (confirmado === "true") {
          // Confirmação válida para um único envio.
          return removerSessao(item, CONFERENCIA.SESSION_KEY_CONFIRMED).then(function () {
            event.completed({ allowEvent: true });
          });
        }

        return gravarSessao(item, CONFERENCIA.SESSION_KEY_FOUND, palavras.join(",")).then(function () {
          bloquearEnvio(event);
        });
      });
    })
    .catch(function (erro) {
      // Se não for possível ler o item, não travamos todos os e-mails do usuário.
      console.error("Conferência de cadastro: falha ao analisar o e-mail.", erro);
      event.completed({ allowEvent: true });
    });
}

function bloquearEnvio(event) {
  var mensagem =
    "Este e-mail parece tratar de cadastro ou inserção. " +
    "Antes de enviar, confirme as conferências obrigatórias em \"Conferência de cadastro\".";

  var opcoes = {
    allowEvent: false,
    errorMessage: mensagem
  };

  // Mailbox 1.14+: o botão do alerta passa a abrir o painel com o checklist.
  if (Office.context.requirements.isSetSupported("Mailbox", "1.14")) {
    opcoes.cancelLabel = "Conferência de cadastro";
    opcoes.commandId = CONFERENCIA.TASKPANE_COMMAND_ID;
  } else {
    opcoes.errorMessage +=
      " Abra o painel pelo botão \"Conferência de cadastro\" na faixa de opções.";
  }

  event.completed(opcoes);
}

/* --------------------------------- Registro ---------------------------------- */

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}

// Permite testar a lógica de detecção com Node.js (ignorado pelo Outlook).
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    CONFERENCIA: CONFERENCIA,
    normalizarTexto: normalizarTexto,
    encontrarPalavrasChave: encontrarPalavrasChave
  };
}
