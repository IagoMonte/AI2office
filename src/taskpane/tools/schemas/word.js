const WORD_SYSTEM_PROMPT = `# Papel
Você é um especialista em edição e formatação de documentos Microsoft Word. Seu objetivo é manipular texto, estrutura e formatação de forma coerente e segura.

# Regras de Ouro
1. **Ações Destrutivas:** Antes de executar \`clear_content\` ou \`delete_table\`, você **DEVE** perguntar ao usuário para confirmar.
2. **Região de tabela = só tools de tabela:** NUNCA use \`insert_text\`, \`replace_selection\` ou blocos multi-linha do \`find_replace\` com cursor/seleção dentro ou atravessando tabela — isso destrói a estrutura (vira tabs sem linhas). Edite células com \`write_table_cell\`, estrutura com \`table_add_row\`/\`table_add_column\`, formatação com \`format_table_cell\`/\`format_table_row\`/\`format_table\`.
3. **Cursor vs. Seleção:** Entenda a diferença entre a posição do cursor (onde o texto será inserido) e a seleção atual (trecho destacado que será substituído ou formatado).

# Estratégias de Substituição Robusta (find_replace)
Documentos reais possuem formatações ocultas e espaçamentos inconsistentes. Para evitar falhas silenciosas:
- **NUNCA** use blocos grandes de texto com quebras de linha (\`\\r\`, \`\\n\`), tabulações (\`\\t\`) ou múltiplas linhas no parâmetro \`find\`.
- **SEMPRE** isole o menor trecho de texto único possível para substituir.
  - *ERRADO:* Usar linhas inteiras com formatação invisível (falha se houver um espaço extra ou diferença de formatação).
  - *CERTO:* Usar apenas a parte essencial e única do texto que precisa ser alterado.
- Se o trecho que você quer substituir não for único no documento, use um contexto um pouco maior, mas **evite incluir quebras de linha e tabs**.
- Quando precisar substituir múltiplas linhas ou blocos complexos, considere fazer múltiplas chamadas menores em vez de uma chamada grande.

# Substituição por Ocorrência (find_replace + occurrence)
O parâmetro \`occurrence\` controla QUAIS ocorrências trocar:
- \`"all"\` (padrão, pode omitir): troca todas — ex: \`find_replace {find: "Contratante", replaceWith: "Contratada"}\` renomeia no documento inteiro.
- Número como string (1-based, ordem de leitura do documento): troca SÓ aquela — ex: \`find_replace {find: "Contratante", replaceWith: "Contratada", occurrence: "2"}\` troca apenas a 2ª ocorrência.
- O retorno traz \`totalFound\` (quantas existem) e \`replaced\` (quantas foram trocadas) — confira antes de declarar sucesso.
- Fluxo dinâmico típico: \`find_in_document {query: "Contratante"}\` → vê \`totalMatches: 3\` → decide "trocar só a 2ª" → \`find_replace\` com \`occurrence: "2"\`.
- Se pedir ocorrência além do total, a ferramenta retorna erro — reduza o número ou use \`"all"\`.

# Inserções e Substituições Grandes (replace_selection / insert_text)
- **NUNCA** use tabulações (\`\\t\`) no texto a inserir ou substituir: o Word pode rejeitar a operação com erro genérico. Use espaços no lugar.
- **EVITE** blocos muito grandes em uma única chamada de \`replace_selection\`: prefira várias substituições menores (parágrafo a parágrafo) ou \`find_replace\` com trechos pequenos e únicos.
- Atenção ao escopo da seleção: se ela atravessar tabelas, quebras de página ou de seção, a substituição integral tende a falhar. Nesse caso, trabalhe por partes menores dentro da região.
- Se \`replace_selection\` falhar, NÃO repita a mesma chamada idêntica: divida o texto em partes menores e tente novamente por etapas.

# Fluxo de Trabalho e Contexto
- **Seleção do Usuário:** Se o usuário disser "aqui", "este trecho", "seleção", use \`get_selection\` para identificar o escopo antes de formatar ou substituir.
  - Ex: "coloque isto em negrito" → \`get_selection\` → confirma o trecho → \`format_selection {bold: true}\`.
- **Leitura barata primeiro (fluxo padrão):** nunca comece lendo o documento inteiro. Fluxo recomendado: \`get_doc_info\` para obter contagens e o **outline de títulos** (cada título traz o índice do seu parágrafo), depois \`read_paragraphs\` abrindo janelas nas regiões relevantes e \`find_in_document\` para localizar termos específicos. Use \`read_document\` apenas em documentos curtos; se o campo \`truncated\` vier \`true\`, NÃO repita a leitura integral: migre obrigatoriamente para \`get_doc_info\`/\`read_paragraphs\`/\`find_in_document\`.
  - Ex: "atualize a seção de preços" → \`get_doc_info\` → outline mostra "Preços" no parágrafo 40 → \`read_paragraphs {from: 38, count: 10}\` → edita só ali.
- **Conferência barata:** após editar, NÃO releia o documento inteiro para conferir. Confirme no ponto: \`find_in_document\` com o texto novo ou \`read_paragraphs\` na janela editada.
  - Ex: trocou o e-mail no parágrafo 12 → \`read_paragraphs {from: 12, count: 2}\` e pronto.
- **Posições de Parágrafo:** Os índices usados em \`read_paragraphs\` são 1-based (o parágrafo 1 é o primeiro do documento). O campo \`paragraph\` do \`find_in_document\` e os índices do \`outline\` de \`get_doc_info\` usam essa mesma numeração, então você pode navegar direto para eles.
- **Tabelas (ciclo padrão):** conteúdo de tabela NÃO aparece nas leituras de parágrafos. Ciclo: \`list_tables\` → \`read_table {table: 1}\` → escreve com \`write_table_cell {table: 1, row: 2, column: 3, text: "novo"}\` (escrita cirúrgica: só a célula, sem tocar na formatação das vizinhas).
  - Ex: "corrija o total na primeira tabela" → \`list_tables\` → \`read_table {table: 1}\` → localiza a célula → \`write_table_cell {table: 1, row: 4, column: 2, text: "R$ 1.200,00"}\`.
  - \`find_in_document\` também enxerga tabelas: resultados com \`table\`/\`row\`/\`column\` indicam a célula exata.
  - **Tabelas grandes (20+ linhas):** NUNCA leia a matriz inteira para localizar algo — use \`find_in_table {table: 1, query: "total"}\` e leia por janela: \`read_table {table: 1, fromRow: 40, rowCount: 10}\`.
  - Para criar: \`insert_table {rows: 2, columns: 3}\` (fim do doc) + \`write_table_cell\` para preencher; \`table_add_row\`/\`table_add_column\` crescem a tabela; \`delete_table\` remove (confirme antes, como ação destrutiva).
  - **Corpo, cabeçalho e rodapé:** quase todas as table tools aceitam \`location: "body"|"header"|"footer"\` (padrão: body).
  - Ex: tabela no cabeçalho → \`list_tables {location: "header"}\` → \`write_table_cell {table: 1, row: 1, column: 2, text: "Novo", location: "header"}\`.
  - \`write_header\`/\`write_footer\` RECUSAM se houver tabela no local (para não destruir) — edite nesses casos com as table tools + \`location\`.
- **Cabeçalho e rodapé (texto):** sem tabelas no local, use o ciclo ler→escrever: \`read_header\` → \`write_header {text: "..."}\` (idem para \`read_footer\`/\`write_footer\`). Referem-se ao cabeçalho/rodapé primário da primeira seção; texto vazio limpa.
  - Ex: "atualize o cabeçalho com o novo nome da empresa" → \`read_header\` → \`write_header {text: "Empresa Nova LTDA"}\`.
  - \`find_in_document\` também varre cabeçalho/rodapé: resultados com \`location: "header"\` ou \`"footer"\`.
- **Formatação (personalizar):** três níveis, do específico ao geral.
  - Célula: \`format_table_cell {table: 1, row: 1, column: 2, shadingColor: "#1F4E79", fontColor: "#FFFFFF", bold: true, horizontalAlignment: "center"}\` — ideal para cabeçalho de tabela.
  - Linha: \`format_table_row {table: 1, row: 3, shadingColor: "#D9E1F2"}\` — ex: zebrar ou destacar total.
  - Tabela: \`format_table {table: 1, styleBuiltIn: "GridTable4_Accent1", bandedRows: true}\` — estilos prontos resolvem 90% do visual.
  - Parágrafo: \`format_paragraph {paragraph: 5, styleBuiltIn: "Heading1"}\` ou manual \`format_paragraph {paragraph: 5, alignment: "center", spaceAfter: 12, fontSize: 14, bold: true}\`; \`location\` também vale aqui (body/header/footer).
  - Seleção: \`format_selection {bold: true}\` / \`format_selection {alignment: "justify", fontName: "Arial"}\` — use após \`get_selection\` confirmar o trecho.
- **Anexos:** Mensagens podem conter blocos "[Anexo: ...]". Use esse conteúdo como base para inserções ou substituições, mas lembre-se: anexos são **somente leitura**. Alterações só ocorrem no documento aberto.

# Diretrizes de Ferramentas
- Use \`find_replace\` para alterações em massa no documento inteiro.
- Use \`insert_text\` com \`location: "end"\` para adicionar conteúdo ao final sem perder o cursor atual.`;

const WORD_TOOLS = [
  {
    type: "function",
    function: {
      name: "read_document",
      description:
        "Lê o texto completo do corpo do documento ativo (limitado; o campo 'truncated' indica leitura parcial). Em documentos longos prefira get_doc_info/read_paragraphs/find_in_document",
      parameters: {
        type: "object",
        properties: {},
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "read_paragraphs",
      description:
        "Lê uma janela de parágrafos do documento por posição (índices 1-based), incluindo estilo e tamanho de cada um. Use para navegar em documentos longos",
      parameters: {
        type: "object",
        properties: {
          from: {
            type: "number",
            description: "Índice (1-based) do primeiro parágrafo a ser lido",
          },
          count: {
            type: "number",
            description: "Quantidade de parágrafos a ler (padrão: 15, máximo: 50)",
          },
        },
        required: ["from"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_in_document",
      description:
        "Localiza ocorrências de um texto nos parágrafos, tabelas, cabeçalho e rodapé; retorna posição e contexto de cada ocorrência",
      parameters: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description: "Texto a localizar",
          },
          matchCase: {
            type: "boolean",
            description: "Diferenciar maiúsculas de minúsculas (padrão: false)",
          },
        },
        required: ["query"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_text",
      description:
        "Insere texto na posição do cursor ou em ponto específico do documento. Quebras de linha criam novos parágrafos",
      parameters: {
        type: "object",
        properties: {
          text: {
            type: "string",
            description: "Texto a inserir (use \\n para separar parágrafos)",
          },
          location: {
            type: "string",
            enum: ["cursor", "start", "end"],
            description: "Onde inserir: posição do cursor (padrão), início ou final do documento",
          },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "replace_selection",
      description: "Substitui o trecho atualmente selecionado pelo texto informado",
      parameters: {
        type: "object",
        properties: {
          text: {
            type: "string",
            description: "Novo texto que substituirá a seleção",
          },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "format_selection",
      description: "Aplica formatação ao texto atualmente selecionado",
      parameters: {
        type: "object",
        properties: {
          bold: {
            type: "boolean",
            description: "Aplicar negrito",
          },
          italic: {
            type: "boolean",
            description: "Aplicar itálico",
          },
          underline: {
            type: "boolean",
            description: "Aplicar sublinhado",
          },
          strikethrough: {
            type: "boolean",
            description: "Aplicar tachado",
          },
          fontSize: {
            type: "number",
            description: "Tamanho da fonte em pontos (ex: 12, 14)",
          },
          fontName: {
            type: "string",
            description: "Nome da fonte (ex: 'Arial', 'Calibri')",
          },
          fontColor: {
            type: "string",
            description: "Cor da fonte (hex, ex: '#FF0000' ou nome como 'red')",
          },
          highlightColor: {
            type: "string",
            description: "Cor de destaque/fundo do texto (hex ou nome como 'yellow')",
          },
          subscript: {
            type: "boolean",
            description: "Formatar como subscrito",
          },
          superscript: {
            type: "boolean",
            description: "Formatar como sobrescrito",
          },
          alignment: {
            type: "string",
            enum: ["left", "center", "right", "justify"],
            description: "Alinhamento dos parágrafos da seleção",
          },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "format_paragraph",
      description:
        "Aplica formatação a um parágrafo por índice: estilo pronto, alinhamento, espaçamentos e fonte",
      parameters: {
        type: "object",
        properties: {
          paragraph: {
            type: "number",
            description: "Índice do parágrafo (1-based, como em read_paragraphs)",
          },
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde está o parágrafo (padrão: body)",
          },
          styleBuiltIn: {
            type: "string",
            description: "Estilo pronto (ex: 'Heading1', 'Title', 'Quote', 'Normal')",
          },
          alignment: {
            type: "string",
            enum: ["left", "center", "right", "justify"],
            description: "Alinhamento do parágrafo",
          },
          spaceBefore: {
            type: "number",
            description: "Espaçamento antes, em pontos",
          },
          spaceAfter: {
            type: "number",
            description: "Espaçamento depois, em pontos",
          },
          lineSpacing: {
            type: "number",
            description: "Espaçamento entre linhas, em pontos",
          },
          leftIndent: {
            type: "number",
            description: "Recuo à esquerda, em pontos",
          },
          rightIndent: {
            type: "number",
            description: "Recuo à direita, em pontos",
          },
          firstLineIndent: {
            type: "number",
            description: "Recuo da primeira linha, em pontos",
          },
          fontName: {
            type: "string",
            description: "Nome da fonte (ex: 'Arial')",
          },
          fontSize: {
            type: "number",
            description: "Tamanho da fonte em pontos",
          },
          fontColor: {
            type: "string",
            description: "Cor da fonte (hex)",
          },
          bold: {
            type: "boolean",
            description: "Aplicar negrito",
          },
          italic: {
            type: "boolean",
            description: "Aplicar itálico",
          },
          underline: {
            type: "boolean",
            description: "Aplicar sublinhado",
          },
        },
        required: ["paragraph"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "format_table_cell",
      description: "Aplica formatação a uma célula de tabela: fundo, alinhamentos e fonte",
      parameters: {
        type: "object",
        properties: {
          table: {
            type: "number",
            description: "Número da tabela (1-based)",
          },
          row: {
            type: "number",
            description: "Número da linha (1-based)",
          },
          column: {
            type: "number",
            description: "Número da coluna (1-based)",
          },
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde está a tabela (padrão: body)",
          },
          shadingColor: {
            type: "string",
            description: "Cor de fundo da célula (hex, ex: '#D9E1F2')",
          },
          horizontalAlignment: {
            type: "string",
            enum: ["left", "center", "right", "justify"],
            description: "Alinhamento horizontal do conteúdo",
          },
          verticalAlignment: {
            type: "string",
            enum: ["top", "center", "bottom"],
            description: "Alinhamento vertical do conteúdo",
          },
          bold: {
            type: "boolean",
            description: "Aplicar negrito",
          },
          italic: {
            type: "boolean",
            description: "Aplicar itálico",
          },
          underline: {
            type: "boolean",
            description: "Aplicar sublinhado",
          },
          fontSize: {
            type: "number",
            description: "Tamanho da fonte em pontos",
          },
          fontName: {
            type: "string",
            description: "Nome da fonte",
          },
          fontColor: {
            type: "string",
            description: "Cor da fonte (hex)",
          },
        },
        required: ["table", "row", "column"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "format_table_row",
      description: "Aplica formatação a uma linha inteira de tabela: fundo, alinhamentos e fonte",
      parameters: {
        type: "object",
        properties: {
          table: {
            type: "number",
            description: "Número da tabela (1-based)",
          },
          row: {
            type: "number",
            description: "Número da linha (1-based)",
          },
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde está a tabela (padrão: body)",
          },
          shadingColor: {
            type: "string",
            description: "Cor de fundo da linha (hex)",
          },
          horizontalAlignment: {
            type: "string",
            enum: ["left", "center", "right", "justify"],
            description: "Alinhamento horizontal do conteúdo",
          },
          verticalAlignment: {
            type: "string",
            enum: ["top", "center", "bottom"],
            description: "Alinhamento vertical do conteúdo",
          },
          bold: {
            type: "boolean",
            description: "Aplicar negrito",
          },
          italic: {
            type: "boolean",
            description: "Aplicar itálico",
          },
          underline: {
            type: "boolean",
            description: "Aplicar sublinhado",
          },
          fontSize: {
            type: "number",
            description: "Tamanho da fonte em pontos",
          },
          fontName: {
            type: "string",
            description: "Nome da fonte",
          },
          fontColor: {
            type: "string",
            description: "Cor da fonte (hex)",
          },
        },
        required: ["table", "row"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "format_table",
      description: "Aplica formatação à tabela inteira: estilo pronto, fundo, alinhamento e faixas",
      parameters: {
        type: "object",
        properties: {
          table: {
            type: "number",
            description: "Número da tabela (1-based)",
          },
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde está a tabela (padrão: body)",
          },
          styleBuiltIn: {
            type: "string",
            description: "Estilo pronto de tabela (ex: 'TableGrid', 'GridTable4_Accent1')",
          },
          shadingColor: {
            type: "string",
            description: "Cor de fundo da tabela (hex)",
          },
          horizontalAlignment: {
            type: "string",
            enum: ["left", "center", "right", "justify"],
            description: "Alinhamento horizontal das células",
          },
          bandedRows: {
            type: "boolean",
            description: "Linhas em faixas alternadas",
          },
          headerRows: {
            type: "number",
            description: "Quantidade de linhas de cabeçalho",
          },
        },
        required: ["table"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_replace",
      description:
        "Substitui ocorrências de um texto no documento por outro texto (todas por padrão, ou uma específica via 'occurrence')",
      parameters: {
        type: "object",
        properties: {
          find: {
            type: "string",
            description: "Texto a ser localizado",
          },
          replaceWith: {
            type: "string",
            description: "Novo texto que substituirá cada ocorrência encontrada",
          },
          matchCase: {
            type: "boolean",
            description: "Diferenciar maiúsculas de minúsculas (padrão: false)",
          },
          occurrence: {
            type: "string",
            description:
              'Qual ocorrência substituir: "all" (padrão) ou o número da ocorrência (1-based, ex: "2")',
          },
        },
        required: ["find", "replaceWith"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_tables",
      description:
        "Lista as tabelas com número (1-based), linhas e colunas de cada uma (corpo, cabeçalho ou rodapé)",
      parameters: {
        type: "object",
        properties: {
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde listar: corpo, cabeçalho ou rodapé (padrão: body)",
          },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "read_table",
      description:
        "Lê o conteúdo de uma tabela como matriz 2D de valores; em tabelas grandes use fromRow/rowCount para ler por janela",
      parameters: {
        type: "object",
        properties: {
          table: {
            type: "number",
            description: "Número da tabela (1-based, conforme list_tables)",
          },
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde está a tabela (padrão: body)",
          },
          fromRow: {
            type: "number",
            description: "Primeira linha da janela (1-based; omitido = tabela inteira)",
          },
          rowCount: {
            type: "number",
            description: "Quantidade de linhas da janela (padrão: até o fim)",
          },
        },
        required: ["table"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_in_table",
      description:
        "Localiza ocorrências de um texto dentro de uma tabela e retorna linha/coluna de cada uma (barato para tabelas grandes)",
      parameters: {
        type: "object",
        properties: {
          table: {
            type: "number",
            description: "Número da tabela (1-based)",
          },
          query: {
            type: "string",
            description: "Texto a localizar",
          },
          matchCase: {
            type: "boolean",
            description: "Diferenciar maiúsculas de minúsculas (padrão: false)",
          },
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde está a tabela (padrão: body)",
          },
        },
        required: ["table", "query"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_table",
      description: "Insere uma tabela nova no início ou no final do documento",
      parameters: {
        type: "object",
        properties: {
          rows: {
            type: "number",
            description: "Quantidade de linhas (>= 1)",
          },
          columns: {
            type: "number",
            description: "Quantidade de colunas (>= 1)",
          },
          location: {
            type: "string",
            enum: ["start", "end"],
            description: "Onde inserir: início ou final do destino (padrão: end)",
          },
          target: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Destino: corpo, cabeçalho ou rodapé (padrão: body)",
          },
          values: {
            type: "array",
            description: "Matriz 2D opcional com o conteúdo inicial das células",
            items: {
              type: "array",
              items: {},
            },
          },
        },
        required: ["rows", "columns"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "table_add_row",
      description: "Adiciona linhas ao final de uma tabela",
      parameters: {
        type: "object",
        properties: {
          table: {
            type: "number",
            description: "Número da tabela (1-based)",
          },
          count: {
            type: "number",
            description: "Quantidade de linhas a adicionar (padrão: 1)",
          },
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde está a tabela (padrão: body)",
          },
        },
        required: ["table"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "table_add_column",
      description: "Adiciona colunas ao final de uma tabela",
      parameters: {
        type: "object",
        properties: {
          table: {
            type: "number",
            description: "Número da tabela (1-based)",
          },
          count: {
            type: "number",
            description: "Quantidade de colunas a adicionar (padrão: 1)",
          },
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde está a tabela (padrão: body)",
          },
        },
        required: ["table"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "write_table_cell",
      description: "Escreve texto em uma célula específica de uma tabela",
      parameters: {
        type: "object",
        properties: {
          table: {
            type: "number",
            description: "Número da tabela (1-based)",
          },
          row: {
            type: "number",
            description: "Número da linha (1-based)",
          },
          column: {
            type: "number",
            description: "Número da coluna (1-based)",
          },
          text: {
            type: "string",
            description: "Texto a escrever na célula",
          },
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde está a tabela (padrão: body)",
          },
        },
        required: ["table", "row", "column", "text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "delete_table",
      description: "Remove uma tabela do documento",
      parameters: {
        type: "object",
        properties: {
          table: {
            type: "number",
            description: "Número da tabela (1-based)",
          },
          location: {
            type: "string",
            enum: ["body", "header", "footer"],
            description: "Onde está a tabela (padrão: body)",
          },
        },
        required: ["table"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "read_header",
      description:
        "Lê o texto do cabeçalho primário da primeira seção (inclui resumo das tabelas, se houver)",
      parameters: {
        type: "object",
        properties: {},
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "write_header",
      description:
        "Substitui o texto do cabeçalho primário (recusa se houver tabelas: edite com table tools + location header)",
      parameters: {
        type: "object",
        properties: {
          text: {
            type: "string",
            description: "Novo texto do cabeçalho (vazio limpa)",
          },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "read_footer",
      description:
        "Lê o texto do rodapé primário da primeira seção (inclui resumo das tabelas, se houver)",
      parameters: {
        type: "object",
        properties: {},
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "write_footer",
      description:
        "Substitui o texto do rodapé primário (recusa se houver tabelas: edite com table tools + location footer)",
      parameters: {
        type: "object",
        properties: {
          text: {
            type: "string",
            description: "Novo texto do rodapé (vazio limpa)",
          },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "clear_content",
      description: "Limpa todo o conteúdo do documento ativo",
      parameters: {
        type: "object",
        properties: {},
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_doc_info",
      description:
        "Obtém estatísticas do documento (parágrafos, caracteres, palavras), outline de títulos e resumo das tabelas",
      parameters: {
        type: "object",
        properties: {},
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_selection",
      description:
        "Obtém o texto atualmente selecionado no documento (com contagem de caracteres; limitado em tamanho)",
      parameters: {
        type: "object",
        properties: {},
        required: [],
      },
    },
  },
];

export { WORD_SYSTEM_PROMPT, WORD_TOOLS };
