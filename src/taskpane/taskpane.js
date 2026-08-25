import { sendMessage } from "./agent.js";
import { updateSendButton } from "./chatView.js";
import {
  showSettings,
  hideSettings,
  toggleKeyVisibility,
  saveSettings,
  loadSettingsForm,
} from "./settingsView.js";
import { fetchModels } from "./modelSelect.js";

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
  loadSettingsForm();
  fetchModels();
});
