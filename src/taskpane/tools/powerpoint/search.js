
function normalize(str) {
  return String(str || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function levenshtein(a, b) {
  const s = normalize(a);
  const t = normalize(b);
  if (s === t) return 0;
  if (!s.length) return t.length;
  if (!t.length) return s.length;
  if (Math.abs(s.length - t.length) > Math.max(s.length, t.length) * 0.6) {
    return Math.max(s.length, t.length);
  }
  const prev = new Array(t.length + 1);
  for (let j = 0; j <= t.length; j++) prev[j] = j;
  for (let i = 1; i <= s.length; i++) {
    let cur0 = i;
    let prevDiag = prev[0];
    prev[0] = i;
    for (let j = 1; j <= t.length; j++) {
      const tmp = prev[j];
      const cost = s[i - 1] === t[j - 1] ? 0 : 1;
      prev[j] = Math.min(prev[j] + 1, cur0 + 1, prevDiag + cost);
      prevDiag = tmp;
      cur0 = prev[j];
    }
  }
  return prev[t.length];
}

function scoreTexto(termo, texto) {
  const nTermo = normalize(termo).trim();
  const nTexto = normalize(texto);
  if (!nTermo) return 0;
  if (!nTexto) return 0;
  if (nTexto.includes(nTermo)) {
    return 0.7 + 0.3 * (nTermo.length / Math.max(nTexto.length, nTermo.length));
  }
  const tokensTermo = nTermo.split(/\s+/);
  const tokensTexto = nTexto.split(/\s+/);
  let soma = 0;
  for (const tok of tokensTermo) {
    let melhor = 0;
    for (const cand of tokensTexto) {
      const dist = levenshtein(tok, cand);
      const maior = Math.max(tok.length, cand.length) || 1;
      const sim = 1 - dist / maior;
      if (sim > melhor) melhor = sim;
    }
    soma += melhor;
  }
  const media = soma / tokensTermo.length;
  return media >= 0.4 ? media * 0.85 : media * 0.5;
}

function buscarPorSimilaridade(termo, cache) {
  return cache
    .map((item) => ({ item, scoreSimilaridade: scoreTexto(termo, item.texto) }))
    .filter((r) => r.scoreSimilaridade >= 0.2)
    .sort((a, b) => b.scoreSimilaridade - a.scoreSimilaridade);
}

function calcularDistanciaLocalidade(itemA, itemB) {
  if (itemA.shapeId === itemB.shapeId && itemA.slideIndex === itemB.slideIndex) return 0;
  if (itemA.slideIndex === itemB.slideIndex) return 1;
  if (Math.abs(itemA.slideIndex - itemB.slideIndex) === 1) return 2;
  return 4;
}

function buscarCombinado(termo, cache, shapeReferencia, pesos) {
  const p = pesos || { similaridade: 0.7, localidade: 0.3 };
  const resultadosSimilaridade = buscarPorSimilaridade(termo, cache || []);

  const comScoreFinal = resultadosSimilaridade.map((r) => {
    let scoreLocalidade = 0.5;
    if (shapeReferencia) {
      const distancia = calcularDistanciaLocalidade(shapeReferencia, r.item);
      scoreLocalidade = 1 / (1 + distancia);
    }

    const scoreFinal = r.scoreSimilaridade * p.similaridade + scoreLocalidade * p.localidade;

    return { ...r, scoreLocalidade, scoreFinal };
  });

  return comScoreFinal.sort((a, b) => b.scoreFinal - a.scoreFinal);
}

function buscarPorTabela(prefixoOuId, cache) {
  const entrada = String(prefixoOuId || "").trim();
  if (!entrada) return [];
  const prefixo = entrada.replace(/_r\d+c\d+$/, "");
  return (cache || [])
    .filter((item) => item.tipo === "Table" && String(item.id).startsWith(prefixo))
    .sort((a, b) => a.row - b.row || a.column - b.column);
}

function buscarPorIndice(indice, cache, raio) {
  const lista = cache || [];
  const idx = Math.floor(Number(indice));
  const r = raio === undefined || raio === null ? 3 : Math.max(0, Math.floor(Number(raio)));
  if (!lista.length) throw new Error("Cache vazio. Rode o Módulo 1 (Ler tudo).");
  if (!Number.isFinite(idx) || idx < 0 || idx >= lista.length) {
    throw new Error(`Índice ${indice} fora do intervalo [0, ${lista.length - 1}].`);
  }
  const inicio = Math.max(0, idx - r);
  const fim = Math.min(lista.length - 1, idx + r);
  const resultados = [];
  for (let i = inicio; i <= fim; i++) {
    resultados.push({ index: i, item: lista[i] });
  }
  return resultados;
}

export {
  buscarPorSimilaridade,
  calcularDistanciaLocalidade,
  buscarCombinado,
  buscarPorTabela,
  buscarPorIndice,
  scoreTexto,
};
