import { getEndpoint, setEndpoint, getApiKey, setApiKey } from "./config.js";
import { fetchModels } from "./modelSelect.js";

export function showSettings() {
  document.getElementById("chatView").classList.add("hidden");
  document.getElementById("settingsView").classList.remove("hidden");
}

export function hideSettings() {
  document.getElementById("settingsView").classList.add("hidden");
  document.getElementById("chatView").classList.remove("hidden");
}

export function toggleKeyVisibility() {
  const input = document.getElementById("apiKey");
  input.type = input.type === "password" ? "text" : "password";
}

export function loadSettingsForm() {
  document.getElementById("apiEndpoint").value = getEndpoint();
  document.getElementById("apiKey").value = getApiKey();
}

export function saveSettings(e) {
  e.preventDefault();
  setEndpoint(document.getElementById("apiEndpoint").value);
  setApiKey(document.getElementById("apiKey").value);
  hideSettings();
  fetchModels();
}
