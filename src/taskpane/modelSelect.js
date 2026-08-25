import { requestJson } from "./http.js";
import { getEndpoint, getApiKey, getSavedModel } from "./config.js";

function getModelsUrl() {
  const endpoint = getEndpoint();
  if (!endpoint) return null;
  return endpoint.replace(/\/chat\/completions\/?$/i, "").replace(/\/+$/, "") + "/models";
}

function parseModelList(data) {
  if (Array.isArray(data)) {
    return data.map((m) => (typeof m === "string" ? m : m.id || m.name || "")).filter(Boolean);
  }
  if (data.data && Array.isArray(data.data)) {
    return data.data.map((m) => m.id || m.name || "").filter(Boolean);
  }
  if (data.models && Array.isArray(data.models)) {
    return data.models.map((m) => m.id || m.name || "").filter(Boolean);
  }
  return [];
}

function renderModels(select, models) {
  models.sort((a, b) => a.localeCompare(b));

  const savedModel = getSavedModel();
  select.innerHTML = "";

  if (models.length === 0) {
    select.innerHTML = '<option value="" disabled>Nenhum modelo encontrado</option>';
    return;
  }

  let hasSelected = false;
  models.forEach(function (id) {
    const opt = document.createElement("option");
    opt.value = id;
    opt.textContent = id;
    if (id === savedModel) {
      opt.selected = true;
      hasSelected = true;
    }
    select.appendChild(opt);
  });

  if (!hasSelected && select.options.length > 0) {
    select.options[0].selected = true;
  }
}

export async function fetchModels() {
  const select = document.getElementById("modelSelect");
  const btn = document.getElementById("refreshModelsBtn");
  const modelsUrl = getModelsUrl();
  const apiKey = getApiKey();

  if (!modelsUrl || !apiKey) {
    select.innerHTML = '<option value="" disabled selected>Configure a API primeiro</option>';
    return;
  }

  if (btn) btn.classList.add("spinning");
  select.innerHTML = '<option value="" disabled>Carregando...</option>';

  let xhr;
  try {
    xhr = await requestJson("GET", modelsUrl, { apiKey: apiKey });
  } catch {
    if (btn) btn.classList.remove("spinning");
    select.innerHTML = '<option value="" disabled>Erro de conexão</option>';
    return;
  }

  if (btn) btn.classList.remove("spinning");

  if (xhr.status < 200 || xhr.status >= 300) {
    select.innerHTML = '<option value="" disabled>Erro: ' + xhr.status + "</option>";
    return;
  }

  try {
    renderModels(select, parseModelList(JSON.parse(xhr.responseText)));
  } catch {
    select.innerHTML = '<option value="" disabled>Erro ao processar resposta</option>';
  }
}
