let conversationHistory = [];
let isLoading = false;

export function getHistory() {
  return conversationHistory;
}

export function pushHistory(entry) {
  conversationHistory.push(entry);
}

export function resetHistory() {
  conversationHistory = [];
}

export function isBusy() {
  return isLoading;
}

export function setLoading(value) {
  isLoading = value;
}
