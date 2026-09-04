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
      case "format_paragraph":
        result = await formatParagraph(args.paragraph, args.location, args);
        break;
      case "format_table_cell":
        result = await formatTableCell(args.table, args.row, args.column, args.location, args);
        break;
      case "format_table_row":
        result = await formatTableRow(args.table, args.row, args.location, args);
        break;
      case "format_table":
        result = await formatTable(args.table, args.location, args);
        break;
      case "find_replace":
        result = await findReplace(args.find, args.replaceWith, args.matchCase, args.occurrence);
        break;
      case "list_tables":
        result = await listTables(args.location);
        break;
      case "read_table":
        result = await readTable(args.table, args.location, args.fromRow, args.rowCount);
        break;
      case "insert_table":
        result = await insertTable(
          args.rows,
          args.columns,
          args.location,
          args.values,
          args.target
        );
        break;
      case "table_add_row":
        result = await tableAddRow(args.table, args.count, args.location);
        break;
      case "table_add_column":
        result = await tableAddColumn(args.table, args.count, args.location);
        break;
      case "write_table_cell":
        result = await writeTableCell(args.table, args.row, args.column, args.text, args.location);
        break;
      case "delete_table":
        result = await deleteTable(args.table, args.location);
        break;
      case "find_in_table":
        result = await findInTable(args.table, args.query, args.matchCase, args.location);
        break;
      case "read_header":
        result = await readHeaderFooter("header");
        break;
      case "write_header":
        result = await writeHeaderFooter("header", args.text);
        break;
      case "read_footer":
        result = await readHeaderFooter("footer");
        break;
      case "write_footer":
        result = await writeHeaderFooter("footer", args.text);
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

      function snippetAround(rawText, position, length) {
        const contextStart = Math.max(0, position - MATCH_CONTEXT_CHARS);
        const contextEnd = Math.min(rawText.length, position + length + MATCH_CONTEXT_CHARS);
        return (
          (contextStart > 0 ? "..." : "") +
          rawText.substring(contextStart, contextEnd).trim() +
          (contextEnd < rawText.length ? "..." : "")
        );
      }

      function scanText(rawText, onHit) {
        const haystack = String(matchCase ? rawText : rawText.toLowerCase());
        let position = haystack.indexOf(needle);
        while (position !== -1) {
          totalMatches++;
          if (matches.length < MAX_MATCHES) {
            onHit(position);
          }
          position = haystack.indexOf(needle, position + needle.length);
        }
      }

      if (isWordApi13Supported()) {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        const tableItems = tables.items;
        tableItems.forEach((table) => table.load("values"));
        await context.sync();

        tableItems.forEach((table, tableIndex) => {
          const values = table.values || [];
          values.forEach((line, rowIndex) => {
            (line || []).forEach((cell, columnIndex) => {
              scanText(String(cell || ""), (position) => {
                matches.push({
                  table: tableIndex + 1,
                  row: rowIndex + 1,
                  column: columnIndex + 1,
                  snippet: snippetAround(String(cell || ""), position, needle.length),
                });
              });
            });
          });
        });

        const section = context.document.sections.getFirst();
        const header = section.getHeader("Primary");
        const footer = section.getFooter("Primary");
        header.load("text");
        footer.load("text");
        await context.sync();

        scanText(header.text || "", (position) => {
          matches.push({
            location: "header",
            snippet: snippetAround(header.text || "", position, needle.length),
          });
        });
        scanText(footer.text || "", (position) => {
          matches.push({
            location: "footer",
            snippet: snippetAround(footer.text || "", position, needle.length),
          });
        });
      }

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
      if (format.fontName) font.name = format.fontName;
      if (format.fontColor) font.color = format.fontColor;
      if (format.highlightColor) font.highlightColor = format.highlightColor;
      if (format.subscript !== undefined) font.subscript = !!format.subscript;
      if (format.superscript !== undefined) font.superscript = !!format.superscript;

      if (format.alignment !== undefined && format.alignment !== null) {
        const paragraphs = context.document.getSelection().paragraphs;
        paragraphs.load("items");
        await context.sync();
        const mapped = mapAlignment(format.alignment, "seleção");
        paragraphs.items.forEach((paragraph) => {
          paragraph.alignment = mapped;
        });
      }

      await context.sync();
      resolve("Formatação aplicada na seleção.");
    }).catch(reject);
  });
}

function applyFontProps(font, format, applied) {
  if (format.bold !== undefined) {
    font.bold = !!format.bold;
    applied.push("bold");
  }
  if (format.italic !== undefined) {
    font.italic = !!format.italic;
    applied.push("italic");
  }
  if (format.underline !== undefined) {
    font.underline = format.underline ? "Single" : "None";
    applied.push("underline");
  }
  if (format.fontSize !== undefined) {
    font.size = format.fontSize;
    applied.push("fontSize");
  }
  if (format.fontName) {
    font.name = format.fontName;
    applied.push("fontName");
  }
  if (format.fontColor) {
    font.color = format.fontColor;
    applied.push("fontColor");
  }
}

function formatParagraph(paragraph, location, format) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const fmt = format || {};
      const target = resolveBody(context, location);
      const n = Math.floor(Number(paragraph));
      if (!Number.isFinite(n) || n < 1) {
        throw new Error('Parâmetro "paragraph" inválido: informe o índice (1-based).');
      }

      const paragraphs = target.body.paragraphs;
      paragraphs.load("items");
      await context.sync();
      if (n > paragraphs.items.length) {
        throw new Error(
          `Parágrafo ${n} não existe: há ${paragraphs.items.length} parágrafo(s) em ${target.location}.`
        );
      }

      const item = paragraphs.items[n - 1];
      const applied = [];
      if (fmt.alignment !== undefined && fmt.alignment !== null) {
        item.alignment = mapAlignment(fmt.alignment, "parágrafo");
        applied.push("alignment");
      }
      if (fmt.styleBuiltIn) {
        item.styleBuiltIn = fmt.styleBuiltIn;
        applied.push("styleBuiltIn");
      }
      if (fmt.spaceBefore !== undefined) {
        item.spaceBefore = Number(fmt.spaceBefore);
        applied.push("spaceBefore");
      }
      if (fmt.spaceAfter !== undefined) {
        item.spaceAfter = Number(fmt.spaceAfter);
        applied.push("spaceAfter");
      }
      if (fmt.lineSpacing !== undefined) {
        item.lineSpacing = Number(fmt.lineSpacing);
        applied.push("lineSpacing");
      }
      if (fmt.leftIndent !== undefined) {
        item.leftIndent = Number(fmt.leftIndent);
        applied.push("leftIndent");
      }
      if (fmt.rightIndent !== undefined) {
        item.rightIndent = Number(fmt.rightIndent);
        applied.push("rightIndent");
      }
      if (fmt.firstLineIndent !== undefined) {
        item.firstLineIndent = Number(fmt.firstLineIndent);
        applied.push("firstLineIndent");
      }
      applyFontProps(item.font, fmt, applied);

      if (!applied.length) {
        throw new Error("Nada para formatar: informe ao menos um parâmetro de formatação.");
      }
      await context.sync();

      resolve({ location: target.location, paragraph: n, applied: applied });
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

function resolveBody(context, location) {
  const loc = location === undefined || location === null ? "body" : String(location).toLowerCase();
  if (loc === "header" || loc === "cabecalho" || loc === "cabeçalho") {
    return {
      body: context.document.sections.getFirst().getHeader("Primary"),
      location: "header",
    };
  }
  if (loc === "footer" || loc === "rodape" || loc === "rodapé") {
    return {
      body: context.document.sections.getFirst().getFooter("Primary"),
      location: "footer",
    };
  }
  if (loc !== "body") {
    throw new Error('Parâmetro "location" inválido: use "body", "header" ou "footer".');
  }
  return { body: context.document.body, location: "body" };
}

function mapAlignment(value, kind) {
  const v = String(value).toLowerCase();
  const table = {
    left: "Left",
    center: "Centered",
    centered: "Centered",
    right: "Right",
    justify: "Justified",
    justified: "Justified",
  };
  if (!table[v]) {
    throw new Error(
      `Alinhamento inválido para ${kind}: use "left", "center", "right" ou "justify".`
    );
  }
  return table[v];
}

function mapVerticalAlignment(value) {
  const v = String(value).toLowerCase();
  const table = { top: "Top", center: "Center", middle: "Center", bottom: "Bottom" };
  if (!table[v]) {
    throw new Error('Alinhamento vertical inválido: use "top", "center" ou "bottom".');
  }
  return table[v];
}

async function resolveTableIndex(context, table, body) {
  const n = Math.floor(Number(table));
  if (!Number.isFinite(n) || n < 1) {
    throw new Error('Parâmetro "table" inválido: informe o número da tabela (1-based).');
  }
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  if (n > tables.items.length) {
    throw new Error(`Tabela ${n} não existe: há ${tables.items.length} tabela(s) neste local.`);
  }
  return { tables: tables, item: tables.items[n - 1], index: n };
}

async function listTablesIn(body, context) {
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  const items = tables.items;
  items.forEach((table) => table.load("rowCount,columnCount"));
  await context.sync();
  return items.map((table, i) => ({
    table: i + 1,
    rows: table.rowCount,
    columns: table.columnCount,
  }));
}

function listTables(location) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const target = resolveBody(context, location);
      const tables = await listTablesIn(target.body, context);

      resolve({ location: target.location, totalTables: tables.length, tables: tables });
    }).catch(reject);
  });
}

function readTable(table, location, fromRow, rowCount) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const target = resolveBody(context, location);
      const resolved = await resolveTableIndex(context, table, target.body);
      resolved.item.load("rowCount,columnCount");
      await context.sync();

      const total = resolved.item.rowCount;

      if (fromRow === undefined || fromRow === null) {
        resolved.item.load("values");
        await context.sync();
        resolve({
          table: resolved.index,
          location: target.location,
          rows: total,
          columns: resolved.item.columnCount,
          values: resolved.item.values,
        });
        return;
      }

      const first = Math.floor(Number(fromRow));
      const qty =
        rowCount === undefined || rowCount === null ? total : Math.floor(Number(rowCount));
      if (!Number.isFinite(first) || first < 1 || first > total) {
        throw new Error(`"fromRow" inválido: a tabela ${resolved.index} tem ${total} linha(s).`);
      }
      if (!Number.isFinite(qty) || qty < 1) {
        throw new Error('"rowCount" inválido: informe inteiro >= 1.');
      }

      const rows = resolved.item.rows;
      rows.load("items");
      await context.sync();

      const last = Math.min(first + qty - 1, total);
      const slice = [];
      for (let i = first - 1; i < last; i++) {
        rows.items[i].load("values");
        slice.push(rows.items[i]);
      }
      await context.sync();

      resolve({
        table: resolved.index,
        location: target.location,
        rows: total,
        columns: resolved.item.columnCount,
        fromRow: first,
        rowCount: slice.length,
        values: slice.map((row) => (row.values && row.values[0]) || []),
      });
    }).catch(reject);
  });
}

function insertTable(rows, columns, location, values, target) {
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

      const dest = resolveBody(context, target);
      const where = location === "start" ? "Start" : "End";
      dest.body.insertTable(rowCount, columnCount, where, values || undefined);
      await context.sync();

      const tables = await listTablesIn(dest.body, context);

      resolve({
        table: tables.length,
        location: dest.location,
        rows: rowCount,
        columns: columnCount,
        position: location === "start" ? "início" : "final",
      });
    }).catch(reject);
  });
}

function tableAddRow(table, count, location) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const target = resolveBody(context, location);
      const resolved = await resolveTableIndex(context, table, target.body);
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
        location: target.location,
        addedRows: qty,
        rows: resolved.item.rowCount,
        columns: resolved.item.columnCount,
      });
    }).catch(reject);
  });
}

function tableAddColumn(table, count, location) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const target = resolveBody(context, location);
      const resolved = await resolveTableIndex(context, table, target.body);
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
        location: target.location,
        addedColumns: qty,
        rows: resolved.item.rowCount,
        columns: resolved.item.columnCount,
      });
    }).catch(reject);
  });
}

function writeTableCell(table, row, column, text, location) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const target = resolveBody(context, location);
      const resolved = await resolveTableIndex(context, table, target.body);
      resolved.item.load("rowCount,columnCount");
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

      // Escrita cirúrgica na célula (getCell é 0-based): não reescreve a
      // matriz inteira, preservando a formatação das demais células.
      const cell = resolved.item.getCell(r - 1, c - 1);
      cell.value = text === undefined || text === null ? "" : String(text);
      await context.sync();

      resolve({ table: resolved.index, location: target.location, row: r, column: c });
    }).catch(reject);
  });
}

function deleteTable(table, location) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const target = resolveBody(context, location);
      const resolved = await resolveTableIndex(context, table, target.body);
      resolved.item.delete();
      await context.sync();

      resolve({ location: target.location, deleted: resolved.index });
    }).catch(reject);
  });
}

function findInTable(table, query, matchCase, location) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      if (!query || !String(query).trim()) {
        throw new Error("Texto de busca vazio.");
      }
      const target = resolveBody(context, location);
      const resolved = await resolveTableIndex(context, table, target.body);
      resolved.item.load("values,rowCount,columnCount");
      await context.sync();

      const needle = String(matchCase ? query : String(query).toLowerCase());
      const matches = [];
      let totalMatches = 0;
      const values = resolved.item.values || [];

      values.forEach((line, rowIndex) => {
        (line || []).forEach((cell, columnIndex) => {
          const rawText = String(cell || "");
          const haystack = String(matchCase ? rawText : rawText.toLowerCase());
          let position = haystack.indexOf(needle);
          while (position !== -1) {
            totalMatches++;
            if (matches.length < MAX_MATCHES) {
              matches.push({ row: rowIndex + 1, column: columnIndex + 1, text: rawText });
            }
            position = haystack.indexOf(needle, position + needle.length);
          }
        });
      });

      resolve({
        table: resolved.index,
        location: target.location,
        query: query,
        totalMatches: totalMatches,
        shown: matches.length,
        matches: matches,
      });
    }).catch(reject);
  });
}

function readHeaderFooter(kind) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const section = context.document.sections.getFirst();
      const body = kind === "header" ? section.getHeader("Primary") : section.getFooter("Primary");
      body.load("text");
      await context.sync();

      const tables = await listTablesIn(body, context);
      const text = body.text || "";
      resolve({
        location: kind,
        charCount: text.length,
        text: text,
        totalTables: tables.length,
        tables: tables,
      });
    }).catch(reject);
  });
}

function writeHeaderFooter(kind, text) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      requireTablesApi();
      const section = context.document.sections.getFirst();
      const body = kind === "header" ? section.getHeader("Primary") : section.getFooter("Primary");
      const tables = await listTablesIn(body, context);
      if (tables.length) {
        throw new Error(
          `Este ${kind === "header" ? "cabeçalho" : "rodapé"} contém ${tables.length} tabela(s): a escrita integral as destruiria. Edite com as table tools usando location: "${kind}".`
        );
      }
      body.clear();
      body.insertText(
        normalizeLineBreaks(text === undefined || text === null ? "" : String(text)),
        "End"
      );
      await context.sync();

      resolve((kind === "header" ? "Cabeçalho" : "Rodapé") + " atualizado com sucesso.");
    }).catch(reject);
  });
}

function formatTableCell(table, row, column, location, format) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const fmt = format || {};
      requireTablesApi();
      const target = resolveBody(context, location);
      const resolved = await resolveTableIndex(context, table, target.body);
      resolved.item.load("rowCount,columnCount");
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

      const cell = resolved.item.getCell(r - 1, c - 1);
      const applied = [];
      if (fmt.shadingColor) {
        cell.shadingColor = fmt.shadingColor;
        applied.push("shadingColor");
      }
      if (fmt.horizontalAlignment !== undefined && fmt.horizontalAlignment !== null) {
        cell.horizontalAlignment = mapAlignment(fmt.horizontalAlignment, "célula");
        applied.push("horizontalAlignment");
      }
      if (fmt.verticalAlignment !== undefined && fmt.verticalAlignment !== null) {
        cell.verticalAlignment = mapVerticalAlignment(fmt.verticalAlignment);
        applied.push("verticalAlignment");
      }
      if (
        fmt.bold !== undefined ||
        fmt.italic !== undefined ||
        fmt.underline !== undefined ||
        fmt.fontSize !== undefined ||
        fmt.fontName ||
        fmt.fontColor
      ) {
        cell.load("body/paragraphs");
        await context.sync();
        const paragraphs = cell.body.paragraphs;
        paragraphs.load("items");
        await context.sync();
        if (paragraphs.items.length) {
          applyFontProps(paragraphs.items[0].font, fmt, applied);
        }
      }

      if (!applied.length) {
        throw new Error("Nada para formatar: informe ao menos um parâmetro de formatação.");
      }
      await context.sync();

      resolve({
        table: resolved.index,
        location: target.location,
        row: r,
        column: c,
        applied: applied,
      });
    }).catch(reject);
  });
}

function formatTableRow(table, row, location, format) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const fmt = format || {};
      requireTablesApi();
      const target = resolveBody(context, location);
      const resolved = await resolveTableIndex(context, table, target.body);
      const n = Math.floor(Number(row));

      const rows = resolved.item.rows;
      rows.load("items");
      await context.sync();
      if (!Number.isFinite(n) || n < 1 || n > rows.items.length) {
        throw new Error(
          `Linha inválida: tabela ${resolved.index} tem ${rows.items.length} linha(s).`
        );
      }

      const item = rows.items[n - 1];
      const applied = [];
      if (fmt.shadingColor) {
        item.shadingColor = fmt.shadingColor;
        applied.push("shadingColor");
      }
      if (fmt.horizontalAlignment !== undefined && fmt.horizontalAlignment !== null) {
        item.horizontalAlignment = mapAlignment(fmt.horizontalAlignment, "linha");
        applied.push("horizontalAlignment");
      }
      if (fmt.verticalAlignment !== undefined && fmt.verticalAlignment !== null) {
        item.verticalAlignment = mapVerticalAlignment(fmt.verticalAlignment);
        applied.push("verticalAlignment");
      }
      applyFontProps(item.font, fmt, applied);

      if (!applied.length) {
        throw new Error("Nada para formatar: informe ao menos um parâmetro de formatação.");
      }
      await context.sync();

      resolve({ table: resolved.index, location: target.location, row: n, applied: applied });
    }).catch(reject);
  });
}

function formatTable(table, location, format) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const fmt = format || {};
      requireTablesApi();
      const target = resolveBody(context, location);
      const resolved = await resolveTableIndex(context, table, target.body);

      const applied = [];
      if (fmt.styleBuiltIn) {
        resolved.item.styleBuiltIn = fmt.styleBuiltIn;
        applied.push("styleBuiltIn");
      }
      if (fmt.shadingColor) {
        resolved.item.shadingColor = fmt.shadingColor;
        applied.push("shadingColor");
      }
      if (fmt.horizontalAlignment !== undefined && fmt.horizontalAlignment !== null) {
        resolved.item.horizontalAlignment = mapAlignment(fmt.horizontalAlignment, "tabela");
        applied.push("horizontalAlignment");
      }
      if (fmt.bandedRows !== undefined) {
        resolved.item.styleBandedRows = !!fmt.bandedRows;
        applied.push("bandedRows");
      }
      if (fmt.headerRows !== undefined) {
        const headerCount = Math.floor(Number(fmt.headerRows));
        if (!Number.isFinite(headerCount) || headerCount < 0) {
          throw new Error('"headerRows" inválido: informe inteiro >= 0.');
        }
        resolved.item.headerRowCount = headerCount;
        applied.push("headerRows");
      }

      if (!applied.length) {
        throw new Error("Nada para formatar: informe ao menos um parâmetro de formatação.");
      }
      await context.sync();

      resolve({ table: resolved.index, location: target.location, applied: applied });
    }).catch(reject);
  });
}

function findReplace(findText, replaceText, matchCase, occurrence) {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      if (!findText || !String(findText).trim()) {
        throw new Error("Texto de busca vazio.");
      }

      const results = context.document.body.search(findText, { matchCase: !!matchCase });
      context.load(results, "items");
      await context.sync();

      const total = results.items.length;
      let targets = results.items;
      let occurrenceLabel = "all";

      if (
        occurrence !== undefined &&
        occurrence !== null &&
        String(occurrence).toLowerCase() !== "all"
      ) {
        const n = Math.floor(Number(occurrence));
        if (!Number.isFinite(n) || n < 1) {
          throw new Error(
            'Parâmetro "occurrence" inválido: use "all" ou o número da ocorrência (1-based).'
          );
        }
        if (n > total) {
          throw new Error(`Ocorrência ${n} não existe: "${findText}" tem ${total} ocorrência(s).`);
        }
        targets = [results.items[n - 1]];
        occurrenceLabel = n;
      }

      for (let i = 0; i < targets.length; i++) {
        targets[i].insertText(replaceText, "Replace");
      }
      await context.sync();

      resolve({
        find: findText,
        occurrence: occurrenceLabel,
        totalFound: total,
        replaced: targets.length,
      });
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
      const tables = await listTablesIn(body, context);
      resolve({
        paragraphCount: items.length,
        charCount: text.length,
        wordCount: text.trim() ? text.trim().split(/\s+/).length : 0,
        totalTables: tables.length,
        tables: tables,
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
