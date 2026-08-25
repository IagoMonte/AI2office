const MAX_READ_CHARS = 8000;
const MAX_SELECTION_CHARS = 2000;
const MAX_PARAGRAPH_CHARS = 500;
const DEFAULT_PARAGRAPH_WINDOW = 15;
const MAX_PARAGRAPH_WINDOW = 50;
const MAX_MATCHES = 10;
const MATCH_CONTEXT_CHARS = 40;
const MAX_OUTLINE = 50;
const OUTLINE_SNIPPET_CHARS = 80;

async function executeTool(toolName, args) {
  try {
    let result;
    switch (toolName) {
      case "read_document":
        result = await readDocument();
        break;
      case "read_paragraphs":
        result = await readParagraphs(args.from, args.count);
        break;
      case "find_in_document":
        result = await findInDocument(args.query, args.matchCase);
        break;
      case "insert_text":
        result = await insertText(args.text, args.location);
        break;
      case "replace_selection":
        result = await replaceSelection(args.text);
        break;
      case "format_selection":
        result = await formatSelection(args);
        break;
      case "find_replace":
        result = await findReplace(args.find, args.replaceWith, args.matchCase);
        break;
      case "clear_content":
        result = await clearContent();
        break;
      case "get_doc_info":
        result = await getDocInfo();
        break;
      case "get_selection":
        result = await getSelection();
        break;
      default:
        throw new Error("Ferramenta desconhecida: " + toolName);
    }
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function normalizeLineBreaks(text) {
  return String(text).replace(/\r\n/g, "\r").replace(/\n/g, "\r");
}

function readDocument() {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();

      const text = body.text || "";
      const truncated = text.length > MAX_READ_CHARS;
      resolve({
        totalChars: text.length,
        truncated: truncated,
        text: truncated ? text.substring(0, MAX_READ_CHARS) : text,
      });
    }).catch(reject);
  });
}

function readParagraphs(from, count) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const requestedFrom = Math.floor(Number(from));
      if (!requestedFrom || requestedFrom < 1) {
        throw new Error('Parâmetro "from" inválido: informe o índice do parágrafo (1-based).');
      }

      const windowSize = Math.max(
        1,
        Math.min(Math.floor(Number(count)) || DEFAULT_PARAGRAPH_WINDOW, MAX_PARAGRAPH_WINDOW)
      );

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const total = paragraphs.items.length;
      if (!total) {
        resolve({ from: 0, to: 0, totalParagraphs: 0, paragraphs: [] });
        return;
      }

      const first = Math.min(requestedFrom, total);
      const last = Math.min(first + windowSize - 1, total);

      const slice = [];
      for (let i = first - 1; i < last; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text,style");
        slice.push(paragraph);
      }
      await context.sync();

      resolve({
        from: first,
        to: last,
        totalParagraphs: total,
        paragraphs: slice.map((paragraph, offset) => {
          const text = paragraph.text || "";
          const truncated = text.length > MAX_PARAGRAPH_CHARS;
          return {
            index: first + offset,
            style: paragraph.style || "",
            charCount: text.length,
            truncated: truncated,
            text: truncated ? text.substring(0, MAX_PARAGRAPH_CHARS) : text,
          };
        }),
      });
    }).catch(reject);
  });
}

function findInDocument(query, matchCase) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      if (!query || !String(query).trim()) {
        throw new Error("Texto de busca vazio.");
      }

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();
      paragraphs.items.forEach((paragraph) => paragraph.load("text"));
      await context.sync();

      const needle = String(matchCase ? query : String(query).toLowerCase());
      const matches = [];
      let totalMatches = 0;

      paragraphs.items.forEach((paragraph, index) => {
        const rawText = paragraph.text || "";
        const haystack = String(matchCase ? rawText : rawText.toLowerCase());
        let position = haystack.indexOf(needle);

        while (position !== -1) {
          totalMatches++;
          if (matches.length < MAX_MATCHES) {
            const contextStart = Math.max(0, position - MATCH_CONTEXT_CHARS);
            const contextEnd = Math.min(
              rawText.length,
              position + needle.length + MATCH_CONTEXT_CHARS
            );
            matches.push({
              paragraph: index + 1,
              snippet:
                (contextStart > 0 ? "..." : "") +
                rawText.substring(contextStart, contextEnd).trim() +
                (contextEnd < rawText.length ? "..." : ""),
            });
          }
          position = haystack.indexOf(needle, position + needle.length);
        }
      });

      resolve({
        query: query,
        totalMatches: totalMatches,
        shown: matches.length,
        matches: matches,
      });
    }).catch(reject);
  });
}

function insertText(text, location) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const content = normalizeLineBreaks(text);
      let where;

      if (location === "start") {
        context.document.body.insertText(content, "Start");
        where = "início do documento";
      } else if (location === "end") {
        context.document.body.insertText(content, "End");
        where = "final do documento";
      } else {
        context.document.getSelection().insertText(content, "End");
        where = "posição do cursor";
      }

      await context.sync();
      resolve("Texto inserido com sucesso no " + where + ".");
    }).catch(reject);
  });
}

function replaceSelection(text) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      context.document.getSelection().insertText(normalizeLineBreaks(text), "Replace");
      await context.sync();
      resolve("Seleção substituída com sucesso.");
    }).catch(reject);
  });
}

function formatSelection(format) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const font = context.document.getSelection().font;

      if (format.bold !== undefined) font.bold = !!format.bold;
      if (format.italic !== undefined) font.italic = !!format.italic;
      if (format.underline !== undefined) font.underline = format.underline ? "Single" : "None";
      if (format.strikethrough !== undefined) font.strikeThrough = !!format.strikethrough;
      if (format.fontSize !== undefined) font.size = format.fontSize;
      if (format.fontColor) font.color = format.fontColor;
      if (format.highlightColor) font.highlightColor = format.highlightColor;

      await context.sync();
      resolve("Formatação aplicada na seleção.");
    }).catch(reject);
  });
}

function findReplace(findText, replaceText, matchCase) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const results = context.document.body.search(findText, { matchCase: !!matchCase });
      context.load(results, "items");
      await context.sync();

      const count = results.items.length;
      for (let i = 0; i < count; i++) {
        results.items[i].insertText(replaceText, "Replace");
      }
      await context.sync();

      resolve(count + ' ocorrência(s) de "' + findText + '" substituída(s).');
    }).catch(reject);
  });
}

function clearContent() {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      context.document.body.clear();
      await context.sync();
      resolve("Conteúdo do documento limpo.");
    }).catch(reject);
  });
}

function getDocInfo() {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      body.load("text");
      paragraphs.load("items");
      await context.sync();

      const items = paragraphs.items;

      for (let i = 0; i < items.length; i++) {
        items[i].load("style");
      }
      await context.sync();

      const headingIndexes = [];
      for (let i = 0; i < items.length; i++) {
        const style = (items[i].style || "").trim();
        if (/^(heading|título|title)/i.test(style)) {
          if (headingIndexes.length < MAX_OUTLINE) {
            headingIndexes.push(i);
          }
        }
      }

      for (let h = 0; h < headingIndexes.length; h++) {
        items[headingIndexes[h]].load("text");
      }
      await context.sync();

      const text = body.text || "";
      resolve({
        paragraphCount: items.length,
        charCount: text.length,
        wordCount: text.trim() ? text.trim().split(/\s+/).length : 0,
        outline: headingIndexes.map((index) => ({
          paragraph: index + 1,
          style: items[index].style || "",
          text: (items[index].text || "").trim().substring(0, OUTLINE_SNIPPET_CHARS),
        })),
      });
    }).catch(reject);
  });
}

function getSelection() {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const text = selection.text || "";
      const truncated = text.length > MAX_SELECTION_CHARS;
      resolve({
        charCount: text.length,
        truncated: truncated,
        text: truncated ? text.substring(0, MAX_SELECTION_CHARS) : text,
      });
    }).catch(reject);
  });
}

async function describeSelection() {
  try {
    const info = await getSelection();
    if (!info.charCount) {
      return null;
    }
    let text = "trecho selecionado (" + info.charCount + " caractere(s))";
    if (info.truncated) {
      text += " [truncado]";
    }
    text += ": " + info.text;
    return text;
  } catch {
    return null;
  }
}

export { executeTool, describeSelection };
