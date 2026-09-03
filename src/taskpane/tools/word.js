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
        result = await findReplace(args.find, args.replaceWith, args.matchCase, args.occurrence);
        break;
      case "list_tables":
        result = await listTables();
        break;
      case "read_table":
        result = await readTable(args.table);
        break;
      case "insert_table":
        result = await insertTable(args.rows, args.columns, args.location, args.values);
        break;
      case "table_add_row":
        result = await tableAddRow(args.table, args.count);
        break;
      case "table_add_column":
        result = await tableAddColumn(args.table, args.count);
        break;
      case "write_table_cell":
        result = await writeTableCell(args.table, args.row, args.column, args.text);
        break;
      case "delete_table":
        result = await deleteTable(args.table);
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

function isWordApi13Supported() {
  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.requirements) {
      return Office.context.requirements.isSetSupported("WordApi", "1.3");
    }
  } catch {
    return false;
  }
  return false;
}

function requireTablesApi() {
  if (!isWordApi13Supported()) {
    throw new Error("Recursos de tabela/cabeçalho exigem Word API 1.3+.");
  }
}

async function resolveTableIndex(context, table) {
  const n = Math.floor(Number(table));
  if (!Number.isFinite(n) || n < 1) {
    throw new Error('Parâmetro "table" inválido: informe o número da tabela (1-based).');
  }
  const tables = context.document.body.tables;
  tables.load("items");
  await context.sync();
  if (n > tables.items.length) {
    throw new Error(`Tabela ${n} não existe: o documento tem ${tables.items.length} tabela(s).`);
  }
  return { tables: tables, item: tables.items[n - 1], index: n };
}

function listTables() {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      const items = tables.items;
      items.forEach((table) => table.load("rowCount,columnCount"));
      await context.sync();

      resolve({
        totalTables: items.length,
        tables: items.map((table, i) => ({
          table: i + 1,
          rows: table.rowCount,
          columns: table.columnCount,
        })),
      });
    }).catch(reject);
  });
}

function readTable(table) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const resolved = await resolveTableIndex(context, table);
      resolved.item.load("values,rowCount,columnCount");
      await context.sync();

      resolve({
        table: resolved.index,
        rows: resolved.item.rowCount,
        columns: resolved.item.columnCount,
        values: resolved.item.values,
      });
    }).catch(reject);
  });
}

function insertTable(rows, columns, location, values) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const rowCount = Math.floor(Number(rows));
      const columnCount = Math.floor(Number(columns));
      if (
        !Number.isFinite(rowCount) ||
        rowCount < 1 ||
        !Number.isFinite(columnCount) ||
        columnCount < 1
      ) {
        throw new Error('Parâmetros "rows"/"columns" inválidos: informe inteiros >= 1.');
      }
      if (values !== undefined && values !== null && !Array.isArray(values)) {
        throw new Error('Parâmetro "values" inválido: use matriz 2D (array de arrays) ou omita.');
      }

      const where = location === "start" ? "Start" : "End";
      context.document.body.insertTable(rowCount, columnCount, where, values || undefined);
      await context.sync();

      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      resolve({
        table: tables.items.length,
        rows: rowCount,
        columns: columnCount,
        location: location === "start" ? "início do documento" : "final do documento",
      });
    }).catch(reject);
  });
}

function tableAddRow(table, count) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const resolved = await resolveTableIndex(context, table);
      const qty = count === undefined || count === null ? 1 : Math.floor(Number(count));
      if (!Number.isFinite(qty) || qty < 1) {
        throw new Error('Parâmetro "count" inválido: informe inteiro >= 1.');
      }
      resolved.item.addRows(qty);
      await context.sync();
      resolved.item.load("rowCount,columnCount");
      await context.sync();

      resolve({
        table: resolved.index,
        addedRows: qty,
        rows: resolved.item.rowCount,
        columns: resolved.item.columnCount,
      });
    }).catch(reject);
  });
}

function tableAddColumn(table, count) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const resolved = await resolveTableIndex(context, table);
      const qty = count === undefined || count === null ? 1 : Math.floor(Number(count));
      if (!Number.isFinite(qty) || qty < 1) {
        throw new Error('Parâmetro "count" inválido: informe inteiro >= 1.');
      }
      resolved.item.addColumns(qty);
      await context.sync();
      resolved.item.load("rowCount,columnCount");
      await context.sync();

      resolve({
        table: resolved.index,
        addedColumns: qty,
        rows: resolved.item.rowCount,
        columns: resolved.item.columnCount,
      });
    }).catch(reject);
  });
}

function writeTableCell(table, row, column, text) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const resolved = await resolveTableIndex(context, table);
      resolved.item.load("values,rowCount,columnCount");
      await context.sync();

      const r = Math.floor(Number(row));
      const c = Math.floor(Number(column));
      if (
        !Number.isFinite(r) ||
        r < 1 ||
        r > resolved.item.rowCount ||
        !Number.isFinite(c) ||
        c < 1 ||
        c > resolved.item.columnCount
      ) {
        throw new Error(
          `Célula inválida: tabela ${resolved.index} tem ${resolved.item.rowCount}x${resolved.item.columnCount}.`
        );
      }

      const values = resolved.item.values.map((line) => line.slice());
      values[r - 1][c - 1] = text === undefined || text === null ? "" : String(text);
      resolved.item.values = values;
      await context.sync();

      resolve({ table: resolved.index, row: r, column: c });
    }).catch(reject);
  });
}

function deleteTable(table) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const resolved = await resolveTableIndex(context, table);
      resolved.item.delete();
      await context.sync();

      resolve({ deleted: resolved.index });
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
