const KEYS = {
  endpoint: "ai2office_endpoint",
  apiKey: "ai2office_key",
  model: "ai2office_model",
};

export function getEndpoint() {
  return (localStorage.getItem(KEYS.endpoint) || "").trim();
}

export function setEndpoint(value) {
  localStorage.setItem(KEYS.endpoint, value.trim());
}

export function getApiKey() {
  return (localStorage.getItem(KEYS.apiKey) || "").trim();
}

export function setApiKey(value) {
  localStorage.setItem(KEYS.apiKey, value.trim());
}

export function getSavedModel() {
  return (localStorage.getItem(KEYS.model) || "").trim();
}

export function setSavedModel(value) {
  localStorage.setItem(KEYS.model, value.trim());
}
