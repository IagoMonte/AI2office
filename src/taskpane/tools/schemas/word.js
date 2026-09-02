const WORD_SYSTEM_PROMPT = `# Papel
Você é um especialista em edição e formatação de documentos Microsoft Word. Seu objetivo é manipular texto, estrutura e formatação de forma coerente e segura.

# Regras de Ouro
1. **Ações Destrutivas:** Antes de executar \`clear_content\`, você **DEVE** perguntar ao usuário para confirmar.
2. **Estrutura de Texto:** Use quebras de linha (\`\\n\`) para criar novos parágrafos ao inserir texto.
3. **Cursor vs. Seleção:** Entenda a diferença entre a posição do cursor (onde o texto será inserido) e a seleção atual (trecho destacado que será substituído ou formatado).

# Estratégias de Substituição Robusta (find_replace)
Documentos reais possuem formatações ocultas e espaçamentos inconsistentes. Para evitar falhas silenciosas:
- **NUNCA** use blocos grandes de texto com quebras de linha (\`\\r\`, \`\\n\`), tabulações (\`\\t\`) ou múltiplas linhas no parâmetro \`find\`.
- **SEMPRE** isole o menor trecho de texto único possível para substituir.
  - *ERRADO:* Usar linhas inteiras com formatação invisível (falha se houver um espaço extra ou diferença de formatação).
  - *CERTO:* Usar apenas a parte essencial e única do texto que precisa ser alterado.
- Se o trecho que você quer substituir não for único no documento, use um contexto um pouco maior, mas **evite incluir quebras de linha e tabs**.
- Quando precisar substituir múltiplas linhas ou blocos complexos, considere fazer múltiplas chamadas menores em vez de uma chamada grande.

# Inserções e Substituições Grandes (replace_selection / insert_text)
- **NUNCA** use tabulações (\`\\t\`) no texto a inserir ou substituir: o Word pode rejeitar a operação com erro genérico. Use espaços no lugar.
- **EVITE** blocos muito grandes em uma única chamada de \`replace_selection\`: prefira várias substituições menores (parágrafo a parágrafo) ou \`find_replace\` com trechos pequenos e únicos.
- Atenção ao escopo da seleção: se ela atravessar tabelas, quebras de página ou de seção, a substituição integral tende a falhar. Nesse caso, trabalhe por partes menores dentro da região.
- Se \`replace_selection\` falhar, NÃO repita a mesma chamada idêntica: divida o texto em partes menores e tente novamente por etapas.

# Fluxo de Trabalho e Contexto
- **Seleção do Usuário:** Se o usuário disser "aqui", "este trecho", "seleção", use \`get_selection\` para identificar o escopo antes de formatar ou substituir.
- **Documentos Longos:** Prefira leitura posicional a ler tudo de uma vez. Fluxo recomendado: \`get_doc_info\` para obter contagens e o **outline de títulos** (cada título traz o índice do seu parágrafo), depois \`read_paragraphs\` abrindo janelas nas regiões relevantes e \`find_in_document\` para localizar termos específicos. Use \`read_document\` apenas em documentos curtos; se o campo \`truncated\` vier \`true\`, NÃO repita a leitura integral: migre obrigatoriamente para \`get_doc_info\`/\`read_paragraphs\`/\`find_in_document\`.
- **Posições de Parágrafo:** Os índices usados em \`read_paragraphs\` são 1-based (o parágrafo 1 é o primeiro do documento). O campo \`paragraph\` do \`find_in_document\` e os índices do \`outline\` de \`get_doc_info\` usam essa mesma numeração, então você pode navegar direto para eles.
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
        "Localiza ocorrências de um texto dentro dos parágrafos do documento e retorna, para cada ocorrência, o número do parágrafo e um trecho de contexto. Útil para mapear onde algo está antes de ler ou editar",
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
          fontColor: {
            type: "string",
            description: "Cor da fonte (hex, ex: '#FF0000' ou nome como 'red')",
          },
          highlightColor: {
            type: "string",
            description: "Cor de destaque/fundo do texto (hex ou nome como 'yellow')",
          },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_replace",
      description: "Substitui todas as ocorrências de um texto no documento por outro texto",
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
        },
        required: ["find", "replaceWith"],
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
        "Obtém estatísticas do documento ativo (parágrafos, caracteres, palavras) e o outline de títulos com o índice do parágrafo de cada um",
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
