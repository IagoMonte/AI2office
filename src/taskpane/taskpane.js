import { sendMessage } from "./agent.js";
import { updateSendButton, addMessage, renderAttachmentChips } from "./chatView.js";
import {
  showSettings,
  hideSettings,
  toggleKeyVisibility,
  saveSettings,
  loadSettingsForm,
} from "./settingsView.js";
import { fetchModels } from "./modelSelect.js";
import { addAttachment, removeAttachment, onAttachmentsChanged } from "./attachments.js";
import { initPptDebug } from "./pptDebug.js";

function isPowerPointHost() {
  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.host) {
      return Office.context.host === Office.HostType.PowerPoint;
    }
  } catch {
    return false;
  }
  return false;
}

async function handleFiles(files) {
  for (const file of files) {
    try {
      await addAttachment(file);
    } catch (error) {
      addMessage("error", error.message);
    }
  }
}

function isFileDrag(event) {
  const types = (event.dataTransfer && event.dataTransfer.types) || [];
  return Array.prototype.indexOf.call(types, "Files") !== -1;
}

document.addEventListener("DOMContentLoaded", () => {
  // PowerPoint: abre a tela provisória de debug (um bloco por módulo),
  // não o chat — facilita depurar leitura/busca/escrita/polling.
  // Estratégia: inicia o chat por padrão (Excel/Word/navegador) e,
  // se o host for PowerPoint, troca para o debug (esconde o chat).
  initChat();

  const forcePpt =
    typeof window !== "undefined" && /[?&]host=ppt/i.test(window.location.search || "");
  if (forcePpt) {
    initPptDebug();
    mostrarBotaoDebugNoChat();
    return;
  }

  if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
    Office.onReady((info) => {
      if (info && info.host === Office.HostType.PowerPoint) {
        initPptDebug();
        mostrarBotaoDebugNoChat();
      } else if (isPowerPointHost()) {
        initPptDebug();
        mostrarBotaoDebugNoChat();
      }
    });
  } else if (isPowerPointHost()) {
    initPptDebug();
    mostrarBotaoDebugNoChat();
  }
});

// Exibe no chat o botão de volta ao debug (só faz sentido no host PowerPoint).
function mostrarBotaoDebugNoChat() {
  const btn = document.getElementById("btnChatGoPptDebug");
  if (btn) btn.classList.remove("hidden");
}

function initChat() {
  if (initChat.done) return;
  initChat.done = true;
  document.getElementById("sendBtn").addEventListener("click", sendMessage);
  document.getElementById("chatInput").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });
  document.getElementById("chatInput").addEventListener("input", updateSendButton);
  document.getElementById("settingsBtn").addEventListener("click", showSettings);
  document.getElementById("closeSettingsBtn").addEventListener("click", hideSettings);
  document.getElementById("settingsForm").addEventListener("submit", saveSettings);
  document.getElementById("toggleKeyBtn").addEventListener("click", toggleKeyVisibility);
  document.getElementById("refreshModelsBtn").addEventListener("click", fetchModels);
  const btnVoltarDebug = document.getElementById("btnChatGoPptDebug");
  if (btnVoltarDebug) btnVoltarDebug.addEventListener("click", initPptDebug);

  const fileInput = document.getElementById("fileInput");
  document.getElementById("attachBtn").addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => {
    const files = Array.from(fileInput.files || []);
    fileInput.value = "";
    handleFiles(files);
  });

  document.getElementById("attachmentChips").addEventListener("click", (e) => {
    const btn = e.target.closest(".chip-remove");
    if (btn) {
      removeAttachment(Number(btn.dataset.index));
    }
  });

  const appContainer = document.querySelector(".app-container");
  const dropOverlay = document.getElementById("dropOverlay");
  let dragDepth = 0;

  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => e.preventDefault());

  appContainer.addEventListener("dragenter", (e) => {
    if (!isFileDrag(e)) return;
    e.preventDefault();
    dragDepth++;
    dropOverlay.classList.remove("hidden");
  });
  appContainer.addEventListener("dragover", (e) => {
    if (!isFileDrag(e)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
  });
  appContainer.addEventListener("dragleave", (e) => {
    if (!isFileDrag(e)) return;
    dragDepth--;
    if (dragDepth <= 0) {
      dragDepth = 0;
      dropOverlay.classList.add("hidden");
    }
  });
  appContainer.addEventListener("drop", (e) => {
    if (!isFileDrag(e)) return;
    e.preventDefault();
    dragDepth = 0;
    dropOverlay.classList.add("hidden");
    const files = Array.from((e.dataTransfer && e.dataTransfer.files) || []);
    if (files.length) {
      handleFiles(files);
    }
  });

  onAttachmentsChanged(renderAttachmentChips);

  loadSettingsForm();
  fetchModels();
}
