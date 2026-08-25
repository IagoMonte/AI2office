import { isBusy } from "./state.js";
import { renderMarkdown } from "./markdown.js";

export function addMessage(role, content) {
  const container = document.getElementById("chatMessages");
  const welcome = container.querySelector(".welcome-message");
  if (welcome) welcome.remove();

  const div = document.createElement("div");
  div.className = "message " + role;
  if (role === "assistant") {
    div.innerHTML = renderMarkdown(content);
  } else {
    div.textContent = content;
  }
  container.appendChild(div);
  container.scrollTop = container.scrollHeight;
  return div;
}

export function addLoadingIndicator() {
  removeLoadingIndicator();
  const container = document.getElementById("chatMessages");
  const div = document.createElement("div");
  div.className = "message loading";
  div.id = "loadingIndicator";
  div.innerHTML = '<div class="loading-dots"><span></span><span></span><span></span></div>';
  container.appendChild(div);
  container.scrollTop = container.scrollHeight;
}

export function removeLoadingIndicator() {
  document.querySelectorAll("#loadingIndicator").forEach((el) => el.remove());
}

const TOOL_GROUP_SVG =
  '<svg class="tool-chevron" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"></polyline></svg>' +
  '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"></path></svg>';

let currentToolGroup = null;

function buildToolGroup(container) {
  const group = document.createElement("div");
  group.className = "tool-group";

  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "tool-toggle";
  toggle.innerHTML =
    TOOL_GROUP_SVG + '<span>Uso de ferramentas <span class="tool-count">(0)</span></span>';
  toggle.addEventListener("click", () => {
    group.classList.toggle("open");
  });

  const body = document.createElement("div");
  body.className = "tool-body";

  group.appendChild(toggle);
  group.appendChild(body);

  const loader = document.getElementById("loadingIndicator");
  if (loader && loader.parentNode === container) {
    container.insertBefore(group, loader);
  } else {
    container.appendChild(group);
  }

  return group;
}

export function beginToolActivity(fnName, fnArgs) {
  const container = document.getElementById("chatMessages");

  if (!currentToolGroup || currentToolGroup.parentNode !== container) {
    currentToolGroup = buildToolGroup(container);
  }

  const entry = document.createElement("div");
  entry.className = "tool-entry running";

  const header = document.createElement("div");
  header.className = "tool-entry-header";
  header.textContent = "Executando: " + fnName + "(" + JSON.stringify(fnArgs) + ")";
  entry.appendChild(header);

  currentToolGroup.querySelector(".tool-body").appendChild(entry);
  container.scrollTop = container.scrollHeight;
  return entry;
}

export function completeToolActivity(entry, fnName, result) {
  if (!entry) return;

  entry.classList.remove("running");

  const resultEl = document.createElement("pre");
  resultEl.className = "tool-entry-result";
  resultEl.textContent = fnName + ": " + JSON.stringify(result);
  entry.appendChild(resultEl);

  const group = entry.closest(".tool-group");
  if (group) {
    const count = group.querySelector(".tool-count");
    if (count) count.textContent = "(" + group.querySelectorAll(".tool-entry").length + ")";
  }

  const container = document.getElementById("chatMessages");
  container.scrollTop = container.scrollHeight;
}

export function resetToolActivity() {
  currentToolGroup = null;
}

export function collapseToolActivity() {
  if (currentToolGroup) currentToolGroup.classList.remove("open");
}

export function updateSendButton() {
  const input = document.getElementById("chatInput");
  const btn = document.getElementById("sendBtn");
  btn.disabled = !input.value.trim() || isBusy();
}
