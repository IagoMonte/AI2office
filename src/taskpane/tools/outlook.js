const MAX_BODY_CHARS = 8000;
const MAX_SUBJECT_CHARS = 500;

let lastReadBody = null;
let lastReadSubject = null;

function clearOutlookCache() {
  lastReadBody = null;
  lastReadSubject = null;
}

function requireMailboxItem() {
  if (typeof Office === "undefined" || !Office.context || !Office.context.mailbox) {
    throw new Error("Outlook não detectado: abra este painel dentro do Outlook.");
  }
  const item = Office.context.mailbox.item;
  if (!item) {
    throw new Error("Nenhum email ou compromisso aberto. Abra um item para começar.");
  }
  return item;
}

function obterModoDoItem(item) {
  const ehComposicao = typeof item.subject !== "string";
  const tipo = item.itemType === "appointment" ? "appointment" : "message";
  return { ehComposicao, tipo };
}

function obterCorpoHtml() {
  return new Promise((resolve, reject) => {
    try {
      const item = requireMailboxItem();
      item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(String(result.value || ""));
        } else {
          reject(new Error((result.error && result.error.message) || "Falha ao ler o corpo."));
        }
      });
    } catch (error) {
      reject(error);
    }
  });
}

function obterAssunto() {
  return new Promise((resolve, reject) => {
    try {
      const item = requireMailboxItem();
      if (typeof item.subject === "string") {
        resolve(item.subject || "");
        return;
      }
      item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(String(result.value || ""));
        } else {
          reject(new Error((result.error && result.error.message) || "Falha ao ler o assunto."));
        }
      });
    } catch (error) {
      reject(error);
    }
  });
}

function stripHtml(html) {
  return String(html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li|h[1-6]|tr)>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function truncate(text, max) {
  const s = String(text || "");
  if (s.length <= max) {
    return { text: s, truncated: false, totalChars: s.length };
  }
  return { text: s.substring(0, max), truncated: true, totalChars: s.length };
}

function emailOf(recipient) {
  if (!recipient) return null;
  if (typeof recipient === "string") return recipient;
  if (typeof recipient.emailAddress === "string") return recipient.emailAddress;
  if (recipient.emailAddress && typeof recipient.emailAddress.emailAddress === "string") {
    return recipient.emailAddress.emailAddress;
  }
  if (typeof recipient.address === "string") return recipient.address;
  return null;
}

function listEmails(list) {
  if (Array.isArray(list)) {
    return list.map(emailOf).filter(Boolean);
  }
  return [];
}

function getRecipientsAsync(field) {
  return new Promise((resolve) => {
    try {
      if (!field) {
        resolve([]);
        return;
      }
      if (Array.isArray(field)) {
        resolve(listEmails(field));
        return;
      }
      if (typeof field.getAsync === "function") {
        field.getAsync((result) => {
          if (result && result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(listEmails(result.value));
          } else {
            resolve([]);
          }
        });
        return;
      }
      resolve([]);
    } catch {
      resolve([]);
    }
  });
}

function getFromAsync(item) {
  return new Promise((resolve) => {
    try {
      const from = item && item.from;
      if (!from) {
        resolve(null);
        return;
      }
      if (typeof from.getAsync === "function") {
        from.getAsync((result) => {
          if (result && result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(emailOf(result.value) || null);
          } else {
            resolve(null);
          }
        });
        return;
      }
      resolve(emailOf(from) || null);
    } catch {
      resolve(null);
    }
  });
}

async function getEmailInfo() {
  const item = requireMailboxItem();
  const { ehComposicao, tipo } = obterModoDoItem(item);
  const [corpoHtmlIntegral, assuntoIntegral, remetente, to, cc, required, optional] =
    await Promise.all([
      obterCorpoHtml(),
      obterAssunto(),
      getFromAsync(item),
      getRecipientsAsync(item.to),
      getRecipientsAsync(item.cc),
      getRecipientsAsync(item.requiredAttendees),
      getRecipientsAsync(item.optionalAttendees),
    ]);

  lastReadBody = corpoHtmlIntegral;
  lastReadSubject = assuntoIntegral;

  const corpoTexto = stripHtml(corpoHtmlIntegral);
  const corpoCortado = truncate(corpoTexto, MAX_BODY_CHARS);
  const htmlCortado = truncate(corpoHtmlIntegral, MAX_BODY_CHARS);
  const assuntoCortado = truncate(assuntoIntegral, MAX_SUBJECT_CHARS);

  return {
    mode: ehComposicao ? "compose" : "read",
    tipo: tipo,
    assunto: assuntoCortado.text,
    assuntoTruncado: assuntoCortado.truncated,
    remetente: remetente,
    destinatarios: to.concat(required),
    cc: cc.concat(optional),
    corpoTexto: corpoCortado.text,
    corpoHtml: htmlCortado.text,
    truncado: corpoCortado.truncated || htmlCortado.truncated,
    totalChars: corpoCortado.totalChars,
    itemId: item.itemId || null,
    conversationId: item.conversationId || null,
  };
}

function ensureCompose(item, acao) {
  const { ehComposicao } = obterModoDoItem(item);
  if (!ehComposicao) {
    throw new Error(
      "Não é possível " +
        acao +
        ": o item está em modo leitura (somente-consulta). " +
        "Use compose_reply, compose_forward, compose_new ou compose_new_appointment para abrir uma janela preenchida."
    );
  }
}

function ensureRead(item, acao) {
  const { ehComposicao } = obterModoDoItem(item);
  if (ehComposicao) {
    throw new Error(
      "Não é possível " +
        acao +
        ": você já está redigindo um rascunho. " +
        "Use set_email_body ou set_subject para editar o rascunho atual."
    );
  }
}

function ensureMessage(item, acao) {
  const { tipo } = obterModoDoItem(item);
  if (tipo !== "message") {
    throw new Error(
      "Não é possível " +
        acao +
        ": só vale para emails (message). " +
        "Este item é um compromisso (appointment); use compose_new_appointment."
    );
  }
}

async function revalidarRascunho() {
  if (lastReadBody !== null) {
    const atual = await obterCorpoHtml();
    if (atual !== lastReadBody) {
      lastReadBody = atual;
      lastReadSubject = await obterAssunto().catch(() => lastReadSubject);
      throw new Error(
        "O rascunho foi alterado desde a última leitura. Chame get_email_info de novo, reapresente a proposta e só aplique após confirmação."
      );
    }
  }
  if (lastReadSubject !== null) {
    const assuntoAtual = await obterAssunto().catch(() => null);
    if (assuntoAtual !== null && assuntoAtual !== lastReadSubject) {
      lastReadSubject = assuntoAtual;
      throw new Error(
        "O assunto foi alterado desde a última leitura. Chame get_email_info de novo, reapresente a proposta e só aplique após confirmação."
      );
    }
  }
}

function looksLikeHtml(s) {
  return /<\s*(p|div|br|ul|ol|li|b|i|u|a|h[1-6])[\s>]/i.test(String(s || ""));
}

function escapeText(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function normalizeHtml(html) {
  const s = String(html || "");
  if (!s.trim()) {
    throw new Error('Parâmetro "html" vazio: informe o corpo em HTML simples.');
  }
  if (looksLikeHtml(s)) {
    return s;
  }
  return s
    .split(/\r\n|\n|\r/)
    .map((line) => "<p>" + (escapeText(line) || "<br>") + "</p>")
    .join("");
}

function setEmailBody(html, mode) {
  return new Promise((resolve, reject) => {
    (async () => {
      try {
        const item = requireMailboxItem();
        ensureCompose(item, "escrever no corpo");
        const m = String(mode || "prepend").toLowerCase();
        if (m !== "replace" && m !== "cursor" && m !== "prepend") {
          throw new Error('Parâmetro "mode" inválido: use "prepend", "cursor" ou "replace".');
        }
        const htmlFinal = normalizeHtml(html);
        await revalidarRascunho();

        if (m === "cursor") {
          try {
            await new Promise((res, rej) => {
              item.body.setSelectedDataAsync(
                htmlFinal,
                { coercionType: Office.CoercionType.Html },
                (result) => {
                  if (result.status === Office.AsyncResultStatus.Succeeded) {
                    res();
                  } else {
                    rej(
                      new Error(
                        (result.error && result.error.message) ||
                          "Falha ao inserir na posição do cursor."
                      )
                    );
                  }
                }
              );
            });
          } catch (error) {
            throw new Error(
              error.message +
                ' (sem cursor/seleção no rascunho? Tente mode: "prepend", que insere no topo preservando assinatura e histórico, ou peça ao usuário para posicionar o cursor.)'
            );
          }
          lastReadBody = await obterCorpoHtml();
          resolve("Texto inserido na posição do cursor com sucesso.");
          return;
        }

        if (m === "prepend") {
          const corpoAtual = await obterCorpoHtml();
          const combinado = htmlFinal + "<br>" + corpoAtual;
          await new Promise((res, rej) => {
            item.body.setAsync(combinado, { coercionType: Office.CoercionType.Html }, (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                res();
              } else {
                rej(
                  new Error((result.error && result.error.message) || "Falha ao inserir no topo.")
                );
              }
            });
          });
          lastReadBody = await obterCorpoHtml();
          resolve("Texto inserido no topo do rascunho, preservando assinatura e histórico.");
          return;
        }

        await new Promise((res, rej) => {
          item.body.setAsync(htmlFinal, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              res();
            } else {
              rej(
                new Error((result.error && result.error.message) || "Falha ao substituir o corpo.")
              );
            }
          });
        });
        lastReadBody = await obterCorpoHtml();
        resolve("Corpo do rascunho substituído com sucesso.");
      } catch (error) {
        reject(error);
      }
    })();
  });
}

function setSubject(subject) {
  return new Promise((resolve, reject) => {
    (async () => {
      try {
        const item = requireMailboxItem();
        ensureCompose(item, "alterar o assunto");
        const s = String(subject || "");
        if (!s.trim()) {
          throw new Error('Parâmetro "subject" vazio: informe o novo assunto.');
        }
        await revalidarRascunho();
        await new Promise((res, rej) => {
          item.subject.setAsync(s, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              res();
            } else {
              rej(
                new Error((result.error && result.error.message) || "Falha ao definir o assunto.")
              );
            }
          });
        });
        lastReadSubject = s;
        resolve("Assunto atualizado com sucesso.");
      } catch (error) {
        reject(error);
      }
    })();
  });
}

function asEmailArray(value) {
  if (value === undefined || value === null) return [];
  const list = Array.isArray(value) ? value : [value];
  return list.map((e) => String(e).trim()).filter(Boolean);
}

function composeNew(args) {
  const item = requireMailboxItem();
  ensureRead(item, "abrir um novo email");
  const a = args || {};
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.displayNewMessageFormAsync({
        toRecipients: asEmailArray(a.to),
        ccRecipients: asEmailArray(a.cc),
        subject: a.subject !== undefined && a.subject !== null ? String(a.subject) : "",
        htmlBody: a.htmlBody !== undefined && a.htmlBody !== null ? normalizeHtml(a.htmlBody) : "",
      });
      resolve("Nova janela de email aberta. O usuário revisa e envia manualmente.");
    } catch (error) {
      reject(error);
    }
  });
}

function composeReply(htmlBody) {
  const item = requireMailboxItem();
  ensureRead(item, "responder");
  ensureMessage(item, "responder");
  if (!String(htmlBody || "").trim()) {
    throw new Error('Parâmetro "htmlBody" vazio: informe o texto da resposta em HTML simples.');
  }
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.displayReplyFormAsync(
        { htmlBody: normalizeHtml(htmlBody) },
        (result) => {
          if (result && result.status === Office.AsyncResultStatus.Failed) {
            reject(
              new Error((result.error && result.error.message) || "Falha ao abrir a resposta.")
            );
            return;
          }
          resolve("Janela de resposta aberta. O usuário revisa e envia manualmente.");
        }
      );
    } catch (error) {
      reject(error);
    }
  });
}

function composeForward(htmlBody) {
  const item = requireMailboxItem();
  ensureRead(item, "encaminhar");
  ensureMessage(item, "encaminhar");
  if (!String(htmlBody || "").trim()) {
    throw new Error(
      'Parâmetro "htmlBody" vazio: informe o texto adicional do encaminhamento em HTML simples.'
    );
  }
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.displayForwardFormAsync(
        { htmlBody: normalizeHtml(htmlBody) },
        (result) => {
          if (result && result.status === Office.AsyncResultStatus.Failed) {
            reject(
              new Error(
                (result.error && result.error.message) || "Falha ao abrir o encaminhamento."
              )
            );
            return;
          }
          resolve("Janela de encaminhamento aberta. O usuário revisa e envia manualmente.");
        }
      );
    } catch (error) {
      reject(error);
    }
  });
}

function composeNewAppointment(args) {
  const item = requireMailboxItem();
  ensureRead(item, "abrir um novo compromisso");
  const a = args || {};
  const params = {
    subject: a.subject !== undefined && a.subject !== null ? String(a.subject) : "",
    htmlBody: a.htmlBody !== undefined && a.htmlBody !== null ? normalizeHtml(a.htmlBody) : "",
  };
  if (a.location !== undefined && a.location !== null && String(a.location).trim()) {
    params.location = String(a.location);
  }
  if (a.start !== undefined && a.start !== null && String(a.start).trim()) {
    const d = new Date(String(a.start));
    if (isNaN(d.getTime())) {
      throw new Error('Parâmetro "start" inválido: use ISO 8601 (ex: 2026-09-10T14:00:00).');
    }
    params.start = d;
  }
  if (a.end !== undefined && a.end !== null && String(a.end).trim()) {
    const d = new Date(String(a.end));
    if (isNaN(d.getTime())) {
      throw new Error('Parâmetro "end" inválido: use ISO 8601 (ex: 2026-09-10T15:00:00).');
    }
    params.end = d;
  }
  return new Promise((resolve, reject) => {
    try {
      if (typeof Office.context.mailbox.displayNewAppointmentFormAsync !== "function") {
        throw new Error("Este cliente Outlook não suporta criar compromissos via add-in.");
      }
      Office.context.mailbox.displayNewAppointmentFormAsync(params);
      resolve("Nova janela de compromisso aberta. O usuário revisa e salva manualmente.");
    } catch (error) {
      reject(error);
    }
  });
}

async function executeTool(toolName, args) {
  try {
    let result;
    switch (toolName) {
      case "get_email_info":
        result = await getEmailInfo();
        break;
      case "set_email_body":
        result = await setEmailBody(args.html, args.mode);
        break;
      case "set_subject":
        result = await setSubject(args.subject);
        break;
      case "compose_new":
        result = await composeNew(args);
        break;
      case "compose_reply":
        result = await composeReply(args.htmlBody);
        break;
      case "compose_forward":
        result = await composeForward(args.htmlBody);
        break;
      case "compose_new_appointment":
        result = await composeNewAppointment(args);
        break;
      default:
        throw new Error("Ferramenta desconhecida: " + toolName);
    }
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

async function describeSelection() {
  try {
    const item = requireMailboxItem();
    const { ehComposicao, tipo } = obterModoDoItem(item);
    const assunto = await obterAssunto().catch(() => "");
    const remetente = (await getFromAsync(item)) || "?";
    const dests = (await getRecipientsAsync(item.to)).join(", ") || "?";
    const modo = ehComposicao ? "compose" : "read";
    const rotulo = tipo === "appointment" ? "compromisso" : "email";
    let texto =
      rotulo + " " + tipo + " em " + modo + ", assunto '" + (assunto || "(sem assunto)") + "'";
    if (tipo === "message") {
      texto += ", de " + remetente + " para " + dests;
    }
    return texto;
  } catch {
    return null;
  }
}

export { executeTool, describeSelection, clearOutlookCache };
