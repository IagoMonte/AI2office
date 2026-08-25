import { SYSTEM_PROMPT, TOOLS } from "./tools/schema.js";
import { executeTool, describeSelection } from "./tools/excel.js";
import { requestJson } from "./http.js";
import { getHistory, pushHistory, isBusy, setLoading } from "./state.js";
import { getEndpoint, getApiKey, setSavedModel } from "./config.js";
import {
  addMessage,
  addLoadingIndicator,
  removeLoadingIndicator,
  beginToolActivity,
  completeToolActivity,
  resetToolActivity,
  collapseToolActivity,
  updateSendButton,
} from "./chatView.js";

function getChatConfig() {
  return {
    apiEndpoint: getEndpoint(),
    apiKey: getApiKey(),
    model: document.getElementById("modelSelect").value,
  };
}

export function sendMessage() {
  const input = document.getElementById("chatInput");
  const text = input.value.trim();
  if (!text || isBusy()) return;

  const config = getChatConfig();
  if (!config.apiKey || !config.apiEndpoint) {
    addMessage("error", "Configure a API nas configurações antes de enviar mensagens.");
    return;
  }

  if (!config.model) {
    addMessage("error", "Selecione um modelo.");
    return;
  }

  setSavedModel(config.model);

  setLoading(true);
  input.value = "";
  updateSendButton();

  addMessage("user", text);
  pushHistory({ role: "user", content: text });

  resetToolActivity();
  addLoadingIndicator();

  const chatEndpoint = config.apiEndpoint.replace(/\/+$/, "") + "/chat/completions";
  startTurn(chatEndpoint, config);
}

async function startTurn(chatEndpoint, config) {
  const selection = await describeSelection();
  callApiWithTools(chatEndpoint, config, selection);
}

function buildMessages(selection) {
  let systemContent = SYSTEM_PROMPT;
  if (selection) {
    systemContent += "\n\nSeleção atual do usuário na planilha: " + selection;
  }
  return [{ role: "system", content: systemContent }].concat(getHistory());
}

function finishTurn(errorMessage) {
  removeLoadingIndicator();
  collapseToolActivity();
  setLoading(false);
  updateSendButton();
  if (errorMessage) {
    addMessage("error", errorMessage);
  }
}

async function handleResponse(chatEndpoint, config, selection, xhr) {
  if (xhr.status < 200 || xhr.status >= 300) {
    finishTurn("Erro: HTTP " + xhr.status + " - " + xhr.responseText.substring(0, 100));
    return;
  }

  try {
    const data = JSON.parse(xhr.responseText);
    const choice = data.choices && data.choices[0];
    if (!choice) {
      finishTurn("Sem resposta.");
      return;
    }

    const message = choice.message;

    if (message.tool_calls && message.tool_calls.length > 0) {
      pushHistory({
        role: "assistant",
        content: message.content || null,
        tool_calls: message.tool_calls,
      });

      for (const tc of message.tool_calls) {
        const fnName = tc.function.name;
        let fnArgs = {};
        try {
          fnArgs = JSON.parse(tc.function.arguments);
        } catch {
          fnArgs = {};
        }

        const activity = beginToolActivity(fnName, fnArgs);
        const result = await executeTool(fnName, fnArgs);
        completeToolActivity(activity, fnName, result);

        pushHistory({
          role: "tool",
          tool_call_id: tc.id,
          content: JSON.stringify(result),
        });
      }

      addLoadingIndicator();
      callApiWithTools(chatEndpoint, config, selection);
      return;
    }

    const reply = message.content || "Sem resposta.";
    pushHistory({ role: "assistant", content: reply });
    addMessage("assistant", reply);
    finishTurn();
  } catch {
    finishTurn("Erro ao processar resposta.");
  }
}

function callApiWithTools(chatEndpoint, config, selection) {
  requestJson("POST", chatEndpoint, {
    apiKey: config.apiKey,
    body: JSON.stringify({
      model: config.model,
      messages: buildMessages(selection),
      tools: TOOLS,
    }),
  }).then(
    (xhr) => handleResponse(chatEndpoint, config, selection, xhr),
    () => finishTurn("Erro de conexão.")
  );
}
