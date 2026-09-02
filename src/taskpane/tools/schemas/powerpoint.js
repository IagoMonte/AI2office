const POWERPOINT_SYSTEM_PROMPT = `# Papel
Você é um especialista em manipulação de apresentações Microsoft PowerPoint. Seu objetivo é ler, organizar e editar conteúdo de slides com precisão.

# Regras de Ouro
1. **Ações Destrutivas:** Antes de executar \`delete_slide\`, você **DEVE** perguntar ao usuário para confirmar.
2. **Leitura antes de escrita:** Sempre use \`read_content\` ou \`find_in_presentation\` para localizar o texto exato antes de tentar substituir.
3. **Substituição:** Use \`replace_selection\` para substituir o texto atualmente selecionado. O usuário precisa ter o shape selecionado no PowerPoint.

# Fluxo de Trabalho para Edição de Texto
1. Use \`find_in_presentation\` para localizar onde o texto está (retorna slide e shape).
2. Informe ao usuário qual shape selecionar (ex: "Selecione o título do slide 2").
3. Use \`replace_selection\` para substituir o texto na seleção atual.

# Fluxo de Trabalho e Contexto
- **Apresentações Longas:** Prefira leitura posicional. Use \`get_presentation_info\` para visão geral, \`read_content\` para ler um slide específico, \`read_slides\` para janelas de slides.
- **Busca de Texto:** Use \`find_in_presentation\` para localizar termos em toda a apresentação antes de editar.
- **Anexos:** Mensagens podem conter blocos "[Anexo: ...]". Use esse conteúdo como base para substituições, mas lembre-se: anexos são **somente leitura**.

# Diretrizes de Ferramentas
- \`read_content(slideIndex)\` retorna cada shape do slide com índice, nome e texto — use para entender a estrutura de um slide antes de editar.
- \`find_in_presentation(query)\` busca em todos os slides e retorna onde o texto foi encontrado.
- \`replace_selection(text)\` substitui o texto selecionado no PowerPoint — o usuário precisa selecionar o shape antes.
- \`add_slide\` e \`delete_slide\` para gestão de estrutura.`;

const POWERPOINT_TOOLS = [
  {
    type: "function",
    function: {
      name: "read_presentation",
      description: "Lê o conteúdo de todos os slides da apresentação (limitado; o campo 'truncated' indica leitura parcial). Em apresentações longas prefira get_presentation_info/read_slides/read_content",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "read_slides",
      description: "Lê uma janela de slides por posição (índices 1-based), incluindo texto de cada um",
      parameters: {
        type: "object",
        properties: {
          from: { type: "number", description: "Índice (1-based) do primeiro slide a ser lido" },
          count: { type: "number", description: "Quantidade de slides a ler (padrão: 10, máximo: 30)" },
        },
        required: ["from"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_presentation_info",
      description: "Obtém número de slides e títulos de cada slide",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "get_selection",
      description: "Obtém o texto atualmente selecionado na apresentação",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "replace_selection",
      description: "Substitui o texto atualmente selecionado pelo texto informado. O usuário precisa ter um shape com texto selecionado no PowerPoint.",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string", description: "Novo texto que substituirá a seleção" },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "read_content",
      description: "Lê todo o conteúdo de texto de um slide específico, retornando cada shape com seu índice, nome e texto. Use para entender a estrutura de um slide antes de editar.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Índice (1-based) do slide a ser lido" },
        },
        required: ["slideIndex"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_in_presentation",
      description: "Busca um texto em todos os slides da apresentação e retorna as ocorrências com número do slide, índice do shape e trecho de contexto",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string", description: "Texto a localizar" },
        },
        required: ["query"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "add_slide",
      description: "Adiciona um novo slide na posição especificada (padrão: final)",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Posição (1-based) onde inserir o slide (padrão: final da apresentação)" },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "delete_slide",
      description: "Remove um slide da apresentação",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Índice (1-based) do slide a ser removido" },
        },
        required: ["index"],
      },
    },
  },
];

export { POWERPOINT_SYSTEM_PROMPT, POWERPOINT_TOOLS };
