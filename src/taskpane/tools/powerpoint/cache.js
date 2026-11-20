
let cacheConteudo = [];

function getCache() {
  return cacheConteudo;
}

function setCache(novoCache) {
  cacheConteudo = novoCache || [];
  return cacheConteudo;
}

function updateCacheItem(id, novoTexto) {
  const idx = cacheConteudo.findIndex((c) => c.id === id);
  if (idx !== -1) {
    cacheConteudo[idx].texto = novoTexto;
  }
  return idx !== -1;
}

function findCacheItem(ref) {
  if (!ref) return null;
  if (ref.id) {
    const porId = cacheConteudo.find((c) => c.id === ref.id);
    if (porId) return porId;
  }
  if (ref.slideIndex !== undefined && ref.shapeId !== undefined) {
    return (
      cacheConteudo.find(
        (c) =>
          c.slideIndex === Number(ref.slideIndex) &&
          String(c.shapeId) === String(ref.shapeId) &&
          (ref.row === undefined || c.row === Number(ref.row)) &&
          (ref.column === undefined || c.column === Number(ref.column))
      ) || null
    );
  }
  return null;
}

function resolverRefShape(refShapeId, slideIndex) {
  const texto = String(refShapeId === undefined || refShapeId === null ? "" : refShapeId);
  const item = findCacheItem({ id: texto });
  if (item) {
    return { slideIndex: item.slideIndex, shapeId: String(item.shapeId) };
  }
  const idx = Math.floor(Number(slideIndex));
  if (!Number.isFinite(idx) || idx < 0) {
    throw new Error(
      `Referência inválida: "${texto}". Use o shapeId (ex: 85) + slideIndex, ou o id do cache (ex: s0_85).`
    );
  }
  if (!texto) throw new Error("shapeId vazio. Use o shapeId ou o id do cache.");
  return { slideIndex: idx, shapeId: texto };
}

function clearCache() {
  cacheConteudo = [];
  return cacheConteudo;
}

export { getCache, setCache, updateCacheItem, clearCache, findCacheItem, resolverRefShape };
