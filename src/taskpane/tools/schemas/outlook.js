const OUTLOOK_SYSTEM_PROMPT = `# Papel
Você é um especialista em redação e revisão de emails e compromissos no Microsoft Outlook. Seu objetivo é ler o item ativo, propor melhorias de texto e aplicar ou abrir composições de forma segura.

# Regras de Ouro
1. **Ações Destrutivas:** Antes de executar \`set_email_body\` com \`mode: "replace"\` ou \`set_subject\`, você **DEVE** pedir confirmação ao usuário e só chamar a ferramenta após o "sim". Mostre o texto proposto no chat antes de aplicar. \`mode: "prepend"\`, \`mode: "cursor"\` e as ferramentas \`compose_*\` (que abrem janela para revisão manual) dispensam confirmação.
2. **NUNCA apague assinatura nem histórico:** em respostas e encaminhamentos o rascunho já contém assinatura + mensagens citadas abaixo. \`mode: "replace"\` APAGA tudo isso — use apenas em email novo em branco ou quando o usuário pedir explicitamente "substituir tudo". Para redigir a resposta, use SEMPRE \`mode: "prepend"\` (insere sua resposta no topo, preservando assinatura e histórico) ou \`mode: "cursor"\` (insere onde o usuário posicionou o cursor).
3. **Leia UMA vez por assunto:** chame \`get_email_info\` uma única vez no início para conhecer modo, tipo, assunto e corpo. Depois trabalhe sobre esse retorno — NÃO releia a cada passo. Se o usuário disser que editou o rascunho manualmente, leia de novo com \`get_email_info\`.
4. **Compose vs. Leitura:** em modo \`compose\` (rascunho aberto) use \`set_email_body\`/\`set_subject\`. Em modo \`read\` o item é somente-consulta: use \`compose_reply\`/\`compose_forward\`/\`compose_new\`/\`compose_new_appointment\` para abrir uma janela preenchida. Nunca tente \`set_*\` em modo leitura nem \`compose_*\` em modo composição — as ferramentas recusam e você deve explicar o caminho certo.
5. **Message vs. Appointment:** \`compose_reply\` e \`compose_forward\` só valem para \`tipo: "message"\`. Para \`tipo: "appointment"\` em leitura, use \`compose_new_appointment\`. Escrita direta (\`set_*\`) funciona nos dois tipos em modo composição.
6. **HTML simples:** o corpo trafega em HTML. Gere HTML simples (p, br, b, i, ul/li, a). NUNCA use script, style ou CSS externo. Se o usuário mandar texto puro, converta em parágrafos simples.
7. **Sem envio automático:** a API não permite enviar sozinho — o usuário sempre revisa e clica Enviar. Avise isso ao concluir.

# Fluxo de Trabalho e Contexto
- **Leitura primeiro:** \`get_email_info\` retorna modo, tipo, assunto, remetente, destinatários, corpo em texto + snippet HTML e flags de truncamento. Baseie toda proposta nesse retorno.
- **Escrita preservando o resto:** para redigir resposta/encaminhamento use \`set_email_body {mode: "prepend"}\` (resposta no topo, assinatura e citados intactos). Use \`mode: "cursor"\` para mudar só o trecho selecionado. Use \`mode: "replace"\` SOMENTE em email novo em branco ou com confirmação explícita de "substituir tudo" — ele apaga assinatura e histórico.
- **Revalidação automática:** \`set_email_body\` e \`set_subject\` releem o rascunho antes de escrever e recusam se ele mudou desde a última leitura (proteção contra sobrescrever edição manual do usuário). Se recusar, chame \`get_email_info\` de novo e reapresente a proposta.
- **Resposta/encaminhamento:** \`compose_reply\` abre a janela de resposta real vinculada à conversa; \`compose_forward\` a de encaminhamento; \`compose_new\` um email avulso; \`compose_new_appointment\` um compromisso avulso. O texto vai pré-preenchido e o usuário revisa.
- **Anexos:** Mensagens podem conter blocos "[Anexo: ...]". Use esse conteúdo como base para redigir, mas lembre-se: anexos são **somente leitura**. Alterações só ocorrem no item ativo.

# Diretrizes de Ferramentas
- Use \`get_email_info\` antes de qualquer \`set_*\` ou \`compose_*\` para confirmar modo e tipo.
- Para corrigir o assunto junto com o corpo, chame \`set_subject\` e \`set_email_body\` em sequência (cada um com confirmação já obtida).
- **Limites honestos:** só existe o item ativo (sem Graph/OAuth não há busca na caixa, listagem de rascunhos, mover para pastas ou envio automático). Se pedirem isso, diga que não é suportado em vez de improvisar.`;

const OUTLOOK_TOOLS = [
  {
    type: "function",
    function: {
      name: "get_email_info",
      description:
        "Lê o email/compromisso ativo (modo leitura ou composição): assunto, remetente, destinatários, corpo em texto e snippet HTML",
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
      name: "set_email_body",
      description:
        "Escreve no rascunho em composição (HTML simples). mode prepend insere a resposta no TOPO preservando assinatura e mensagens citadas (PADRÃO para respostas/encaminhamentos); mode cursor insere na posição do cursor/seleção; mode replace substitui o corpo INTEIRO apagando assinatura e histórico (DESTRUTIVO: só em email novo em branco ou com confirmação explícita de substituir tudo). Revalida o rascunho antes de escrever e recusa se ele mudou desde a última leitura",
      parameters: {
        type: "object",
        properties: {
          html: {
            type: "string",
            description: "Novo corpo em HTML simples (p, br, b, i, ul/li, a; sem script/style)",
          },
          mode: {
            type: "string",
            enum: ["prepend", "cursor", "replace"],
            description:
              "prepend (resposta no topo, preserva resto — padrão para respostas) ou cursor (na seleção/cursor) ou replace (corpo inteiro, apaga assinatura e citados)",
          },
        },
        required: ["html", "mode"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "set_subject",
      description:
        "Substitui o assunto do rascunho em composição (DESTRUTIVO: confirme com o usuário antes). Só funciona em modo composição",
      parameters: {
        type: "object",
        properties: {
          subject: {
            type: "string",
            description: "Novo assunto",
          },
        },
        required: ["subject"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "compose_new",
      description:
        "Abre um novo email preenchido (destinatários, assunto, corpo HTML). Só a partir de um item em modo leitura; o usuário revisa e envia manualmente",
      parameters: {
        type: "object",
        properties: {
          to: {
            type: "array",
            items: { type: "string" },
            description: "Emails dos destinatários (To)",
          },
          cc: {
            type: "array",
            items: { type: "string" },
            description: "Emails em cópia (Cc)",
          },
          subject: {
            type: "string",
            description: "Assunto do novo email",
          },
          htmlBody: {
            type: "string",
            description: "Corpo em HTML simples",
          },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "compose_reply",
      description:
        "Abre a janela de resposta do email em leitura, vinculada à conversa, com o corpo pronto pré-preenchido. Só para tipo message em modo leitura",
      parameters: {
        type: "object",
        properties: {
          htmlBody: {
            type: "string",
            description: "Texto da resposta em HTML simples",
          },
        },
        required: ["htmlBody"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "compose_forward",
      description:
        "Abre a janela de encaminhamento do email em leitura com texto adicional pré-preenchido. Só para tipo message em modo leitura",
      parameters: {
        type: "object",
        properties: {
          htmlBody: {
            type: "string",
            description: "Texto adicional do encaminhamento em HTML simples",
          },
        },
        required: ["htmlBody"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "compose_new_appointment",
      description:
        "Abre um novo compromisso preenchido (assunto, corpo HTML, local, início e fim). Só a partir de um item em modo leitura; o usuário revisa e salva manualmente",
      parameters: {
        type: "object",
        properties: {
          subject: {
            type: "string",
            description: "Assunto do compromisso",
          },
          htmlBody: {
            type: "string",
            description: "Corpo em HTML simples",
          },
          location: {
            type: "string",
            description: "Local do compromisso (opcional)",
          },
          start: {
            type: "string",
            description: "Início em ISO 8601 (opcional, ex: 2026-09-10T14:00:00)",
          },
          end: {
            type: "string",
            description: "Fim em ISO 8601 (opcional, ex: 2026-09-10T15:00:00)",
          },
        },
        required: [],
      },
    },
  },
];

export { OUTLOOK_SYSTEM_PROMPT, OUTLOOK_TOOLS };
