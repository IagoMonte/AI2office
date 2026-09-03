const POWERPOINT_SYSTEM_PROMPT = `# Papel
Você é um especialista em edição de apresentações Microsoft PowerPoint.
Seu objetivo é localizar e substituir textos em slides de forma segura.

# Regras de Ouro
1. **Leitura primeiro:** antes de buscar ou substituir, use \`ppt_list_content\` para popular o cache.
2. **Um item por vez:** \`ppt_replace\` substitui um único item do cache; repita a chamada para vários itens.
3. **Grupos e tabelas:** itens do tipo GroupChild exigem desagrupar/editar/reagrupar; itens Table usam linha/coluna.
4. **Cache desatualizado:** se a escrita falhar com "não existe mais", rode \`ppt_revalidate\` e peça o id novamente.
5. **Dois formatos de id:** \`id\` do cache (ex: "s0_85", "s0_85_r1c2") identifica o ITEM; \`shapeId\` (ex: "85") identifica a SHAPE no slide. As tools de mover/rotacionar/reordenar/duplicar/deletar/estrutura aceitam qualquer um dos dois.
6. **Verifique destruição:** \`ppt_delete\` retorna \`removida\` (único) ou \`detalhes[]\` (lote) — se false, a shape NÃO foi deletada (id errado ou já removida). Confirme com busca antes de declarar sucesso.

# Posicionamento (geometria)
- Todo item do cache carrega \`left, top, width, height, rotation\` em **pontos** (origem no canto superior esquerdo do slide). Valores \`null\` = indisponível.
- Itens Table (células) carregam a **bounding box da tabela toda** + sua posição \`row/column\`.
- Use a geometria para: evitar sobreposição ao inserir (\`ppt_insert\`), alinhar/empilhar, e escolher \`left/top\` em \`ppt_move_resize\`.
- O cache acompanha mover/rotacionar/inserir/deletar e mudanças em tabelas na hora; arrastar shapes no slide é capturado pelo \`ppt_revalidate\`.

# Fluxo recomendado
- \`ppt_list_content\` → \`ppt_search\` (termo) → mostrar top 5 com id/slide/texto → \`ppt_replace\` no id confirmado.
- Para varrer uma tabela inteira use \`ppt_search_table\` (prefixo do id, ex: "s0_85"); para ver a vizinhança de um item no cache use \`ppt_search_index\` (índice + raio).
- Inserção: \`ppt_insert\` (kind textbox/shape/line) já entra no cache; use o shapeId retornado para mover/rotacionar/deletar.
- Tabelas: \`ppt_create_table\` (+ values 2D opcional); estrutura via \`ppt_table_structure\` (dimension rows/columns, action add/remove/removeMany) — o cache da tabela é relido sozinho.
- Reordenar (z-order) e criar tabela exigem API 1.8; linhas/colunas exigem 1.9 — confira com \`ppt_api_support\`.`;

const POWERPOINT_TOOLS = [
  {
    type: "function",
    function: {
      name: "ppt_list_content",
      description: "Lê todos os textos da apresentação (shapes, tabelas, grupos) e popula o cache",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_search",
      description:
        "Busca no cache por similaridade textual + localidade; retorna top resultados com id",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string", description: "Termo a buscar" },
          top: { type: "number", description: "Quantos resultados retornar (padrão 5)" },
          referenceId: {
            type: "string",
            description: "Id do item de referência para priorizar proximidade",
          },
        },
        required: ["query"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_search_table",
      description:
        "Retorna todas as células de uma tabela do cache pelo prefixo do id (ex: 's0_85'), ordenadas por linha/coluna",
      parameters: {
        type: "object",
        properties: {
          tableId: {
            type: "string",
            description:
              "Prefixo do id da tabela (ex: 's0_85') ou id de uma célula (puxa a tabela toda)",
          },
        },
        required: ["tableId"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_search_index",
      description:
        "Retorna a vizinhança de um item no cache pela posição (índice ± raio, padrão 3)",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Posição do item no cache (0-based)" },
          radius: { type: "number", description: "Raio da vizinhança (padrão 3)" },
        },
        required: ["index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_replace",
      description: "Substitui o texto de um único item do cache pelo id",
      parameters: {
        type: "object",
        properties: {
          id: { type: "string", description: "Id do item (ex: s1_shape42 ou s1_shape42_r0c1)" },
          text: { type: "string", description: "Novo texto" },
        },
        required: ["id", "text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_revalidate",
      description:
        "Revalida o cache de forma incremental (remove deletadas, lê novas, atualiza textos)",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_api_support",
      description:
        "Verifica suporte a PowerPoint API 1.8 (grupos, z-order, addTable) e 1.9 (linhas/colunas)",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_insert",
      description:
        "Insere conteúdo novo no slide: caixa de texto (textbox), forma geométrica (shape) ou linha (line)",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Índice 0-based do slide" },
          kind: { type: "string", description: "textbox, shape ou line" },
          text: { type: "string", description: "Texto (textbox, ou opcional na shape)" },
          shapeType: {
            type: "string",
            description: "Tipo geométrico (ex: Rectangle, Oval, RightArrow)",
          },
          connectorType: { type: "string", description: "Straight, Elbow ou Curve" },
          left: { type: "number", description: "Posição X em pontos (padrão 50)" },
          top: { type: "number", description: "Posição Y em pontos (padrão 50)" },
          width: {
            type: "number",
            description: "Largura (em linhas = deslocamento X do ponto final)",
          },
          height: {
            type: "number",
            description: "Altura (em linhas = deslocamento Y do ponto final)",
          },
        },
        required: ["slideIndex", "kind"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_move_resize",
      description: "Move e/ou redimensiona uma shape (informe só os campos a alterar)",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Índice 0-based do slide" },
          shapeId: {
            type: "string",
            description: "Id da shape (numérico, ex: 85) ou id do item no cache (ex: s0_85)",
          },
          left: { type: "number", description: "Nova posição X" },
          top: { type: "number", description: "Nova posição Y" },
          width: { type: "number", description: "Nova largura" },
          height: { type: "number", description: "Nova altura" },
        },
        required: ["slideIndex", "shapeId"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_rotate",
      description: "Rotaciona uma shape (graus, positivo = sentido horário)",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Índice 0-based do slide" },
          shapeId: {
            type: "string",
            description: "Id da shape (numérico, ex: 85) ou id do item no cache (ex: s0_85)",
          },
          degrees: { type: "number", description: "Graus de rotação" },
        },
        required: ["slideIndex", "shapeId", "degrees"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_reorder",
      description: "Reordena empilhamento (z-order) de uma shape; exige API 1.8",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Índice 0-based do slide" },
          shapeId: {
            type: "string",
            description: "Id da shape (numérico, ex: 85) ou id do item no cache (ex: s0_85)",
          },
          order: {
            type: "string",
            description: "bringToFront, sendToBack, bringForward ou sendBackward",
          },
        },
        required: ["slideIndex", "shapeId"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_duplicate",
      description:
        "Duplica uma shape com deslocamento (recria como TextBox com posição, tamanho, rotação e texto; não preserva preenchimento/contorno)",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Índice 0-based do slide" },
          shapeId: {
            type: "string",
            description: "Id da shape original (numérico) ou id do item no cache",
          },
          offsetLeft: { type: "number", description: "Deslocamento X (padrão 20)" },
          offsetTop: { type: "number", description: "Deslocamento Y (padrão 20)" },
        },
        required: ["slideIndex", "shapeId"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_delete",
      description: "Deleta uma ou várias shapes (idempotente); atualiza o cache",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Índice 0-based do slide (deleção única)" },
          shapeId: {
            type: "string",
            description: "Id da shape ou id do item no cache (deleção única)",
          },
          items: {
            type: "array",
            description: "Deleção em lote",
            items: {
              type: "object",
              properties: {
                slideIndex: { type: "number" },
                shapeId: { type: "string" },
              },
            },
          },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_create_table",
      description: "Cria tabela nova no slide (exige API 1.8); values é array 2D opcional",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Índice 0-based do slide" },
          rows: { type: "number", description: "Número de linhas" },
          columns: { type: "number", description: "Número de colunas" },
          left: { type: "number", description: "Posição X (padrão 50)" },
          top: { type: "number", description: "Posição Y (padrão 50)" },
          values: {
            type: "array",
            description: "Conteúdo inicial (array 2D de strings)",
            items: { type: "array", items: { type: "string" } },
          },
        },
        required: ["slideIndex", "rows", "columns"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "ppt_table_structure",
      description:
        "Altera estrutura de tabela existente: add/remove/removeMany em rows ou columns (exige API 1.9); o cache da tabela é relido sozinho",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Índice 0-based do slide" },
          shapeId: {
            type: "string",
            description:
              "Id da shape da tabela (numérico) ou id de qualquer célula no cache (ex: s0_85_r1c2)",
          },
          dimension: { type: "string", description: "rows ou columns" },
          action: { type: "string", description: "add, remove ou removeMany" },
          count: { type: "number", description: "Quantidade (add, padrão 1)" },
          index: {
            type: "number",
            description: "Posição (add: omitido = fim; remove: índice alvo)",
          },
          indices: {
            type: "array",
            items: { type: "number" },
            description: "Índices (removeMany)",
          },
        },
        required: ["slideIndex", "shapeId", "dimension", "action"],
      },
    },
  },
];

export { POWERPOINT_SYSTEM_PROMPT, POWERPOINT_TOOLS };
