/* Cache compartilhado dos 4 módulos do PowerPoint.
 * Válido apenas durante a sessão atual do documento aberto
 * (shape.id não é garantidamente estável entre sessões).
 */

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

function clearCache() {
  cacheConteudo = [];
  return cacheConteudo;
}

export { getCache, setCache, updateCacheItem, clearCache };
