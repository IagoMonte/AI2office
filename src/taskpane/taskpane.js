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
});
