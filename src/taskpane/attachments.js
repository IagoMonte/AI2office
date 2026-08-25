import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import * as pdfjsWorkerModule from "pdfjs-dist/build/pdf.worker.min.mjs";
import mammoth from "mammoth";

const MAX_ATTACHMENT_CHARS = 12000;
const MAX_TOTAL_CHARS = 40000;
const MAX_FILES = 8;

const pendingAttachments = [];

let changeHandler = null;

window.pdfjsWorker = pdfjsWorkerModule;
pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";

function notifyChange() {
  if (changeHandler) {
    changeHandler(pendingAttachments.slice());
  }
}

export function onAttachmentsChanged(handler) {
  changeHandler = handler;
}

function truncateText(text) {
  if (text.length <= MAX_ATTACHMENT_CHARS) {
    return { text: text, truncated: false };
  }
  return {
    text:
      text.substring(0, MAX_ATTACHMENT_CHARS) + "\n[...conteúdo truncado por limite de tamanho]",
    truncated: true,
  };
}

function getExtension(name) {
  const index = name.lastIndexOf(".");
  return index === -1 ? "" : name.substring(index + 1).toLowerCase();
}

async function readSpreadsheet(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  let text = "";
  workbook.SheetNames.forEach((sheetName) => {
    const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
    if (!csv.trim()) return;
    text += "--- Planilha: " + sheetName + " ---\n" + csv.trim() + "\n\n";
  });
  return text.trim();
}

async function readPdf(file) {
  const buffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(buffer) });
  // eslint-disable-next-line office-addins/call-sync-before-read, office-addins/load-object-before-read
  const doc = await loadingTask.promise;

  let text = "";
  for (let pageNum = 1; pageNum <= doc.numPages; pageNum++) {
    const page = await doc.getPage(pageNum);
    const content = await page.getTextContent();
    const pageText = content.items.map((item) => item.str).join(" ");
    text += "--- Página " + pageNum + " ---\n" + pageText.replace(/\s+/g, " ").trim() + "\n\n";
    page.cleanup();
  }
  await loadingTask.destroy();

  if (!text.trim()) {
    throw new Error(
      'O PDF "' + file.name + '" não contém texto extraível (provavelmente é digitalizado/imagem)'
    );
  }
  return text.trim();
}

async function readDocx(file) {
  const result = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
  return result.value.trim();
}

async function readPlainText(file) {
  return (await file.text()).trim();
}

async function extractContent(file) {
  switch (getExtension(file.name)) {
    case "xlsx":
    case "xls":
      return readSpreadsheet(file);
    case "pdf":
      return readPdf(file);
    case "docx":
      return readDocx(file);
    case "csv":
    case "txt":
      return readPlainText(file);
    default:
      throw new Error("Formato não suportado");
  }
}

export function getSupportedFormatsLabel() {
  return ".xlsx, .xls, .csv, .txt, .pdf, .docx";
}

export async function addAttachment(file) {
  if (pendingAttachments.length >= MAX_FILES) {
    throw new Error("Limite de " + MAX_FILES + " anexos por mensagem.");
  }
  if (
    pendingAttachments.some(
      (existing) => existing.name === file.name && existing.size === file.size
    )
  ) {
    throw new Error("Arquivo já anexado: " + file.name);
  }

  const content = await extractContent(file);
  if (!content) {
    throw new Error("Não foi possível extrair texto de: " + file.name);
  }

  const truncatedResult = truncateText(content);
  const total =
    pendingAttachments.reduce((sum, attachment) => sum + attachment.text.length, 0) +
    truncatedResult.text.length;
  if (total > MAX_TOTAL_CHARS) {
    throw new Error("Limite total de conteúdo dos anexos excedido. Envie menos material.");
  }

  pendingAttachments.push({
    name: file.name,
    size: file.size,
    text: truncatedResult.text,
    truncated: truncatedResult.truncated,
  });
  notifyChange();
  return pendingAttachments.slice();
}

export function removeAttachment(index) {
  pendingAttachments.splice(index, 1);
  notifyChange();
  return pendingAttachments.slice();
}

export function getPendingAttachments() {
  return pendingAttachments.slice();
}

export function clearPendingAttachments() {
  pendingAttachments.length = 0;
  notifyChange();
}

export function formatPendingAttachments() {
  if (!pendingAttachments.length) return null;

  return pendingAttachments
    .map((attachment) => {
      let block = "[Anexo: " + attachment.name + "]\n" + attachment.text;
      if (attachment.truncated) {
        block += "\n[Anexo truncado]";
      }
      return block;
    })
    .join("\n\n");
}
