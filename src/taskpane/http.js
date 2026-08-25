export function proxyUrl(url) {
  if (window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1") {
    return "/api-proxy?target=" + encodeURIComponent(url);
  }
  return url;
}

export function requestJson(method, url, options = {}) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open(method, proxyUrl(url), true);
    if (options.apiKey) {
      xhr.setRequestHeader("Authorization", "Bearer " + options.apiKey);
    }
    if (options.body !== undefined) {
      xhr.setRequestHeader("Content-Type", "application/json");
    }
    xhr.onload = () => resolve(xhr);
    xhr.onerror = () => reject(new Error("Erro de conexão"));
    xhr.send(options.body);
  });
}
