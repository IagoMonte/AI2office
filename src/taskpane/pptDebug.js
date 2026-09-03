/* Tela provisória de debug do PowerPoint.
 * Aberta no lugar do chat quando o host é PowerPoint.
 * Um bloco por módulo da arquitetura (leitura, busca, escrita, atualização).
 */

import {
  getCache,
  clearCache,
  getAllContent,
  isPowerPointApi18Supported,
  getLastReadErrors,
  getLastReadSkipped,
  buscarCombinado,
  buscarPorTabela,
  buscarPorIndice,
  substituirValor,
  medirTempoLeituraInicial,
  revalidarCacheIncremental,
  iniciarPolling,
  pararPolling,
  getPollingState,
} from "./tools/powerpoint/index.js";

function $(id) {
  return document.getElementById(id);
}

function now() {
  return new Date().toLocaleTimeString();
}

function log(msg, isError) {
  const el = $("pptLog");
  if (!el) return;
  const line = `[${now()}] ${msg}\n`;
  el.textContent += line;
  el.scrollTop = el.scrollHeight;
  if (isError) {
    // Erros também aparecem no bloco de status para visibilidade.
    const status = $("pptStatus");
    if (status) status.textContent = msg;
  }
}

function setOut(id, obj) {
  const el = $(id);
  if (!el) return;
  el.textContent = typeof obj === "string" ? obj : JSON.stringify(obj, null, 2);
}

function refreshStatus(extra) {
  const status = $("pptStatus");
  if (!status) return;
  const poll = getPollingState();
  const api18 = isPowerPointApi18Supported();
  status.textContent =
    `host=PowerPoint | api1.8=${api18 ? "sim" : "não"} | cache=${getCache().length} item(ns)` +
    ` | polling=${poll.ativo ? "ON (" + poll.intervaloMs + "ms)" : "OFF"}` +
    (extra ? ` | ${extra}` : "");
}

function previewItens(itens, limite) {
  return itens.slice(0, limite || 10).map((c) => ({
    id: c.id,
    slide: c.slideIndex,
    tipo: c.tipo,
    shape: c.shapeName,
    texto: String(c.texto || "").substring(0, 120),
  }));
}

function resumoPorSlide(itens) {
  const porSlide = {};
  itens.forEach((c) => {
    porSlide[c.slideIndex] = (porSlide[c.slideIndex] || 0) + 1;
  });
  return porSlide;
}

const MAX_ITENS_NA_SAIDA = 500;

function listaCompleta(itens) {
  const lista = itens.slice(0, MAX_ITENS_NA_SAIDA).map((c) => ({
    id: c.id,
    slide: c.slideIndex,
    tipo: c.tipo,
    shape: c.shapeName,
    texto: c.texto || "",
  }));
  if (itens.length > MAX_ITENS_NA_SAIDA) {
    lista.push({
      id: "...",
      slide: "...",
      tipo: "...",
      shape: `mostrando ${MAX_ITENS_NA_SAIDA} de ${itens.length} — use o Módulo 2 (Busca) para navegar`,
      texto: "",
    });
  }
  return lista;
}

async function onLerTudo() {
  log("M1: getAllContent() iniciado...");
  const t0 = performance.now();
  try {
    const itens = await getAllContent((info) => {
      if (info.fase === "slides") {
        log(`M1: ${info.totalSlides} slide(s) na apresentação.`);
      } else if (info.fase === "slide") {
        log(`M1: lendo slide ${info.slideIndex + 1}/${info.totalSlides}...`);
      }
    });
    const ms = Math.round(performance.now() - t0);
    const ignorados = getLastReadErrors();
    const pulados = getLastReadSkipped();
    log(
      `M1: OK — ${itens.length} item(ns) em ${ms}ms` +
        (ignorados.length ? ` (${ignorados.length} erro(s))` : "") +
        (pulados.length ? ` (${pulados.length} sem texto)` : "")
    );
    if (ignorados.length) {
      log(
        `M1 AVISO: shapes ignoradas: ${ignorados.map((g) => `slide ${g.slideIndex} "${g.shapeName}" [${g.tipo}]: ${g.erro}`).join(" | ")}`,
        true
      );
    }
    if (pulados.length) {
      log(
        `M1: puladas (tipo sem textFrame): ${pulados.map((p) => `"${p.shapeName}" [${p.tipo}]`).join(", ")}`
      );
    }
    setOut("pptReadOut", {
      totalItens: itens.length,
      tempoMs: ms,
      resumoPorSlide: resumoPorSlide(itens),
      ignorados: ignorados,
      puladosSemTexto: pulados,
      itens: listaCompleta(itens),
    });
  } catch (e) {
    const code = e && e.code ? ` [${e.code}]` : "";
    log(`M1 ERRO${code}: ` + (e && e.message ? e.message : e), true);
    // Detalhe completo vai para o console (stack + debugInfo do Office).
    try {
      console.log("[M1] erro completo:", e);
      if (e && e.debugInfo) console.log("[M1] debugInfo:", JSON.stringify(e.debugInfo));
      if (e && e.stack) console.log("[M1] stack:", String(e.stack).substring(0, 1000));
    } catch {
      // Console indisponível: segue com o painel.
    }
    setOut("pptReadOut", {
      erro: (e && e.message ? e.message : String(e)) + code,
      dica: "Verifique também o console (F12) para stack/debugInfo completos.",
    });
  }
  refreshStatus();
}

async function onBuscar() {
  const termo = ($("pptSearchTerm") || {}).value || "";
  const top = Number(($("pptSearchTop") || {}).value) || 5;
  const refId = (($("pptSearchRef") || {}).value || "").trim();
  if (!termo.trim()) {
    log("M2 ERRO: informe o termo de busca.", true);
    return;
  }
  try {
    const cache = getCache();
    if (!cache.length) {
      log("M2 AVISO: cache vazio — rode o Módulo 1 primeiro.", true);
      setOut("pptSearchOut", "Cache vazio. Rode Módulo 1 (Ler tudo).");
      return;
    }
    const ref = refId ? cache.find((c) => c.id === refId) || null : null;
    if (refId && !ref) log(`M2 AVISO: referenceId ${refId} não achado; busca sem localidade.`);
    const resultados = buscarCombinado(termo, cache, ref);
    log(`M2: "${termo}" → ${resultados.length} match(es), exibindo top ${top}`);
    setOut(
      "pptSearchOut",
      resultados.slice(0, top).map((r) => ({
        id: r.item.id,
        slideIndex: r.item.slideIndex,
        tipo: r.item.tipo,
        texto: String(r.item.texto || "").substring(0, 140),
        scoreSimilaridade: Number(r.scoreSimilaridade.toFixed(3)),
        scoreLocalidade: Number(r.scoreLocalidade.toFixed(3)),
        scoreFinal: Number(r.scoreFinal.toFixed(3)),
      }))
    );
  } catch (e) {
    log("M2 ERRO: " + (e && e.message ? e.message : e), true);
  }
}

function onBuscarTabela() {
  const prefixo = (($("pptTableId") || {}).value || "").trim();
  if (!prefixo) {
    log("M2-Tabela ERRO: informe o prefixo do id (ex.: s0_85).", true);
    return;
  }
  try {
    const cache = getCache();
    if (!cache.length) {
      log("M2-Tabela AVISO: cache vazio — rode o Módulo 1 primeiro.", true);
      setOut("pptSearchOut", "Cache vazio. Rode Módulo 1 (Ler tudo).");
      return;
    }
    const celulas = buscarPorTabela(prefixo, cache);
    log(`M2-Tabela: "${prefixo}" → ${celulas.length} célula(s)`);
    setOut(
      "pptSearchOut",
      celulas.map((item) => ({
        id: item.id,
        index: cache.indexOf(item),
        slideIndex: item.slideIndex,
        row: item.row,
        column: item.column,
        parentGroupId: item.parentGroupId,
        texto: item.texto || "",
      }))
    );
  } catch (e) {
    log("M2-Tabela ERRO: " + (e && e.message ? e.message : e), true);
  }
}

function onBuscarIndice() {
  const indice = Number(($("pptIndexNum") || {}).value);
  const raioRaw = ($("pptIndexRadius") || {}).value;
  const raio = raioRaw === "" || raioRaw === undefined ? 3 : Number(raioRaw);
  try {
    const cache = getCache();
    const resultados = buscarPorIndice(indice, cache, raio);
    log(
      `M2-Índice: ${indice} ± ${raio} → ${resultados.length} item(ns) ` +
        `[${resultados.map((r) => r.index).join(", ")}]`
    );
    setOut(
      "pptSearchOut",
      resultados.map((r) => ({
        index: r.index,
        id: r.item.id,
        slideIndex: r.item.slideIndex,
        tipo: r.item.tipo,
        texto: r.item.texto || "",
      }))
    );
  } catch (e) {
    log("M2-Índice ERRO: " + (e && e.message ? e.message : e), true);
    setOut("pptSearchOut", "ERRO: " + ((e && e.message) || e));
  }
}

async function onSubstituir() {
  const id = (($("pptReplaceId") || {}).value || "").trim();
  const texto = ($("pptReplaceText") || {}).value || "";
  if (!id) {
    log("M3 ERRO: informe o id do item (copie do M1/M2).", true);
    return;
  }
  try {
    const cache = getCache();
    const alvo = cache.find((c) => c.id === id);
    if (!alvo) {
      log(`M3 ERRO: id ${id} não está no cache. Revalide ou releia.`, true);
      setOut("pptReplaceOut", `Id não encontrado no cache: ${id}`);
      return;
    }
    log(`M3: substituir ${id} (slide ${alvo.slideIndex}, tipo ${alvo.tipo})...`);
    const res = await substituirValor(alvo, texto);
    log(`M3: OK — ${res.id} atualizado.`);
    setOut("pptReplaceOut", { ok: true, ...res });
  } catch (e) {
    log("M3 ERRO: " + (e && e.message ? e.message : e), true);
    setOut(
      "pptReplaceOut",
      "ERRO: " + ((e && e.message) || e) + " — tente M4 (Revalidar) e busque de novo."
    );
  }
  refreshStatus();
}

async function onMedir() {
  try {
    log("M4: medindo tempo de leitura inicial...");
    const { tempoMs, totalItens } = await medirTempoLeituraInicial(getAllContent);
    const intervalo = (await import("./tools/powerpoint/poller.js")).calcularIntervaloPolling(
      tempoMs
    );
    log(
      `M4: leitura=${Math.round(tempoMs)}ms, itens=${totalItens}, intervalo sugerido=${intervalo}ms`
    );
    setOut("pptPollOut", {
      tempoMs: Math.round(tempoMs),
      totalItens,
      intervaloSugeridoMs: intervalo,
    });
  } catch (e) {
    log("M4 ERRO: " + (e && e.message ? e.message : e), true);
  }
  refreshStatus();
}

async function onPollStart() {
  try {
    const intervalo = await iniciarPolling(getAllContent, (err, cache) => {
      if (err) {
        log("M4 polling tick ERRO: " + (err.message || err), true);
      } else {
        log(`M4 polling tick: cache com ${cache.length} item(ns)`);
      }
      refreshStatus();
    });
    log(`M4: polling iniciado a cada ${intervalo}ms`);
  } catch (e) {
    log("M4 ERRO ao iniciar polling: " + (e && e.message ? e.message : e), true);
  }
  refreshStatus();
}

function onPollStop() {
  pararPolling();
  log("M4: polling parado.");
  refreshStatus();
}

async function onRevalidar() {
  try {
    log("M4: revalidação incremental...");
    const cache = await revalidarCacheIncremental();
    log(`M4: OK — cache com ${cache.length} item(ns)`);
    setOut("pptPollOut", { totalItens: cache.length, preview: previewItens(cache, 10) });
  } catch (e) {
    log("M4 ERRO: " + (e && e.message ? e.message : e), true);
  }
  refreshStatus();
}

function initPptDebug() {
  const view = $("pptDebugView");
  const chat = $("chatView");
  if (view) view.classList.remove("hidden");
  if (chat) chat.classList.add("hidden");

  if ($("btnPptRead")) $("btnPptRead").addEventListener("click", onLerTudo);
  if ($("btnPptSearch")) $("btnPptSearch").addEventListener("click", onBuscar);
  if ($("btnPptSearchTable")) $("btnPptSearchTable").addEventListener("click", onBuscarTabela);
  if ($("btnPptSearchIndex")) $("btnPptSearchIndex").addEventListener("click", onBuscarIndice);
  if ($("btnPptReplace")) $("btnPptReplace").addEventListener("click", onSubstituir);
  if ($("btnPptMeasure")) $("btnPptMeasure").addEventListener("click", onMedir);
  if ($("btnPptPollStart")) $("btnPptPollStart").addEventListener("click", onPollStart);
  if ($("btnPptPollStop")) $("btnPptPollStop").addEventListener("click", onPollStop);
  if ($("btnPptRevalidate")) $("btnPptRevalidate").addEventListener("click", onRevalidar);
  if ($("btnPptClearLog")) {
    $("btnPptClearLog").addEventListener("click", () => {
      if ($("pptLog")) $("pptLog").textContent = "";
    });
  }
  if ($("btnPptClearCache")) {
    $("btnPptClearCache").addEventListener("click", () => {
      clearCache();
      log("Cache limpo.");
      refreshStatus();
    });
  }
  if ($("btnPptApiCheck")) {
    $("btnPptApiCheck").addEventListener("click", () => {
      const ok = isPowerPointApi18Supported();
      log(`Suporte PowerPoint API 1.8: ${ok ? "SIM" : "NÃO"} (grupos exigem 1.8+)`);
      console.log(
        "1.8 (grupos):",
        Office.context.requirements.isSetSupported("PowerPointApi", "1.8")
      );
      console.log(
        "1.9 (tabelas avançado):",
        Office.context.requirements.isSetSupported("PowerPointApi", "1.9")
      );
      console.log("1.10:", Office.context.requirements.isSetSupported("PowerPointApi", "1.10"));
      refreshStatus();
    });
  }

  // Enter na busca executa a busca.
  if ($("pptSearchTerm")) {
    $("pptSearchTerm").addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        onBuscar();
      }
    });
  }
  if ($("pptTableId")) {
    $("pptTableId").addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        onBuscarTabela();
      }
    });
  }
  if ($("pptIndexNum")) {
    $("pptIndexNum").addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        onBuscarIndice();
      }
    });
  }

  refreshStatus("tela provisória de debug");
  log("Debug PowerPoint pronto. Comece pelo Módulo 1 (Ler tudo).");
}

export { initPptDebug };
