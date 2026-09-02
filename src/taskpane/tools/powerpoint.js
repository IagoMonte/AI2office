const MAX_READ_CHARS = 8000;
const MAX_SELECTION_CHARS = 2000;
const DEFAULT_SLIDE_WINDOW = 10;
const MAX_SLIDE_WINDOW = 30;

async function executeTool(toolName, args) {
  try {
    let result;
    switch (toolName) {
      case "read_presentation":
        result = await readPresentation();
        break;
      case "read_slides":
        result = await readSlides(args.from, args.count);
        break;
      case "get_presentation_info":
        result = await getPresentationInfo();
        break;
      case "get_selection":
        result = await getSelection();
        break;
      case "replace_selection":
        result = await replaceSelection(args.text);
        break;
      case "read_content":
        result = await readContent(args.slideIndex);
        break;
      case "find_in_presentation":
        result = await findInPresentation(args.query);
        break;
      case "add_slide":
        result = await addSlide(args.index);
        break;
      case "delete_slide":
        result = await deleteSlide(args.index);
        break;
      default:
        return { error: "Ferramenta desconhecida: " + toolName };
    }
    return result;
  } catch (error) {
    return { error: "Erro ao executar " + toolName + ": " + error.message };
  }
}

async function getSlideTexts(context, slide) {
  slide.load("shapes");
  await context.sync();

  var texts = [];
  for (var i = 0; i < slide.shapes.items.length; i++) {
    var shape = slide.shapes.items[i];
    shape.load("name");
    await context.sync();

    var hasText = false;
    var text = "";
    try {
      shape.load("textFrame");
      await context.sync();
      if (shape.textFrame && shape.textFrame.textRange) {
        shape.textFrame.textRange.load("text");
        await context.sync();
        text = shape.textFrame.textRange.text || "";
        hasText = text.length > 0;
      }
    } catch (e) {
      hasText = false;
    }

    if (hasText) {
      texts.push({
        shapeIndex: i,
        shapeName: shape.name || "",
        text: text,
      });
    }
  }
  return texts;
}

function readPresentation() {
  return new Promise(function (resolve, reject) {
    PowerPoint.run(async function (context) {
      var slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      var allText = "";
      for (var i = 0; i < slides.items.length; i++) {
        var texts = await getSlideTexts(context, slides.items[i]);
        var slideText = texts.map(function (t) { return t.text; }).join(" ");
        allText += "\n--- Slide " + (i + 1) + " ---\n" + slideText;

        if (allText.length >= MAX_READ_CHARS) {
          allText = allText.substring(0, MAX_READ_CHARS) + "\n[Conteudo truncado]";
          break;
        }
      }

      resolve({ content: allText || "Apresentacao vazia" });
    }).catch(reject);
  });
}

function readSlides(from, count) {
  return new Promise(function (resolve, reject) {
    var startSlide = Math.max(1, from || 1);
    var slideCount = Math.min(count || DEFAULT_SLIDE_WINDOW, MAX_SLIDE_WINDOW);

    PowerPoint.run(async function (context) {
      var slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      var allText = "";
      var endIndex = Math.min(startSlide + slideCount - 1, slides.items.length);

      for (var i = startSlide - 1; i < endIndex; i++) {
        var texts = await getSlideTexts(context, slides.items[i]);
        var slideText = texts.map(function (t) { return t.text; }).join(" ");
        allText += "\n--- Slide " + (i + 1) + " ---\n" + slideText;

        if (allText.length >= MAX_READ_CHARS) {
          allText = allText.substring(0, MAX_READ_CHARS) + "\n[Conteudo truncado]";
          break;
        }
      }

      resolve({ content: allText || "Nenhum conteudo encontrado" });
    }).catch(reject);
  });
}

function getPresentationInfo() {
  return new Promise(function (resolve, reject) {
    PowerPoint.run(async function (context) {
      var slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      var slideInfos = [];
      for (var i = 0; i < slides.items.length; i++) {
        var texts = await getSlideTexts(context, slides.items[i]);
        var title = texts.length > 0 ? texts[0].text.substring(0, 50) : "Sem titulo";

        slideInfos.push({
          slideNumber: i + 1,
          title: title,
        });
      }

      resolve({
        slideCount: slides.items.length,
        slides: slideInfos,
      });
    }).catch(reject);
  });
}

function getSelection() {
  return new Promise(function (resolve, reject) {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          var text = result.value || "";
          resolve({
            text: text.substring(0, MAX_SELECTION_CHARS),
            length: text.length,
          });
        } else {
          reject(result.error);
        }
      }
    );
  });
}

function replaceSelection(text) {
  return new Promise(function (resolve, reject) {
    Office.context.document.setSelectedDataAsync(
      text,
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve({ success: true, message: "Texto substituido com sucesso" });
        } else {
          reject(result.error);
        }
      }
    );
  });
}

function readContent(slideIndex) {
  return new Promise(function (resolve, reject) {
    PowerPoint.run(async function (context) {
      var slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      if (slideIndex < 1 || slideIndex > slides.items.length) {
        resolve({ error: "Indice de slide invalido: " + slideIndex + ". Total de slides: " + slides.items.length });
        return;
      }

      var texts = await getSlideTexts(context, slides.items[slideIndex - 1]);

      resolve({
        slideIndex: slideIndex,
        shapes: texts,
      });
    }).catch(reject);
  });
}

function findInPresentation(query) {
  return new Promise(function (resolve, reject) {
    if (!query || query.trim() === "") {
      resolve({ error: "Query de busca nao fornecida" });
      return;
    }

    PowerPoint.run(async function (context) {
      var slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      var matches = [];
      var lowerQuery = query.toLowerCase();

      for (var i = 0; i < slides.items.length; i++) {
        var texts = await getSlideTexts(context, slides.items[i]);

        for (var j = 0; j < texts.length; j++) {
          var shapeData = texts[j];
          if (shapeData.text.toLowerCase().indexOf(lowerQuery) !== -1) {
            matches.push({
              slideNumber: i + 1,
              shapeIndex: shapeData.shapeIndex,
              shapeName: shapeData.shapeName,
              snippet: shapeData.text.substring(0, 100),
            });
          }
        }
      }

      resolve({
        query: query,
        matchCount: matches.length,
        matches: matches,
      });
    }).catch(reject);
  });
}

function addSlide(index) {
  return new Promise(function (resolve, reject) {
    PowerPoint.run(async function (context) {
      var slides = context.presentation.slides;

      if (index !== undefined && index !== null) {
        slides.add(index);
      } else {
        slides.add();
      }

      await context.sync();
      resolve({ success: true, message: "Slide adicionado com sucesso" });
    }).catch(reject);
  });
}

function deleteSlide(index) {
  return new Promise(function (resolve, reject) {
    if (index === undefined || index === null) {
      resolve({ error: "Indice do slide nao fornecido" });
      return;
    }

    PowerPoint.run(async function (context) {
      var slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      if (index < 1 || index > slides.items.length) {
        resolve({ error: "Indice de slide invalido: " + index + ". Total de slides: " + slides.items.length });
        return;
      }

      var slide = slides.items[index - 1];
      slide.delete();
      await context.sync();

      resolve({ success: true, message: "Slide " + index + " excluido com sucesso" });
    }).catch(reject);
  });
}

async function describeSelection() {
  try {
    var info = await getSelection();
    if (!info.text || info.text.length === 0) {
      return null;
    }
    var text = "trecho selecionado (" + info.length + " caractere(s))";
    text += ": " + info.text;
    return text;
  } catch (e) {
    return null;
  }
}

export { executeTool, describeSelection };
