# AI2office

**Assistente de IA para o Microsoft Office — manipule seus documentos conversando.**

O AI2office adiciona um painel de chat (task pane) aos aplicativos do Microsoft Office, permitindo ler, escrever e formatar conteúdo por meio de linguagem natural. O projeto começa pelo **Excel** e foi desenhado desde o início para se expandir ao **Word**, **PowerPoint** e **Outlook**.

## Visão geral

- Conecta-se a **qualquer API compatível com OpenAI** (`/chat/completions`) — OpenAI, Azure OpenAI, LM Studio, Ollama, proxies locais etc.
- Loop agêntico de *tool calling*: o modelo decide quais ferramentas chamar, o add-in executa via Office.js e devolve os resultados até a resposta final.
- Configuração flexível: endpoint, chave de API e modelo definidos na própria interface; a lista de modelos é carregada diretamente da API configurada.
- Sem back-end: tudo roda no add-in, e as credenciais ficam apenas no `localStorage` do painel.

## Status por aplicativo

| Aplicativo | Status | Ferramentas |
|---|---|---|
| Excel | ✅ Implementado | Leitura, escrita, formatação, limpeza, info da planilha, seleção atual |
| Word | 🔜 Planejado | — |
| PowerPoint | 🔜 Planejado | — |
| Outlook | 🔜 Planejado | — |

## Funcionalidades atuais (Excel)

- Chat com histórico de conversa mantido durante a sessão
- Loop de *tool calling* com múltiplas chamadas por turno
- **Consciência de seleção**: a seleção atual do usuário é capturada a cada envio e injetada no prompt do sistema — o assistente entende "isto", "aqui", "minha seleção" sem perguntas extras
- Tool dedicada `get_selection` para buscar endereço, dimensões e valores da seleção (valores limitados a 200 células para proteger o contexto da API)
- Painel de configurações (endpoint, chave, modelo) com validação e carregamento dinâmico de modelos

## Arquitetura

Fluxo de um turno:

1. Usuário envia mensagem → `agent.js` captura a seleção atual (`describeSelection`) e monta as mensagens (system prompt + histórico + contexto)
2. Requisição POST para `/chat/completions` com as ferramentas registradas
3. Se o modelo responde com `tool_calls`, o add-in executa via Office.js e devolve os resultados como novas mensagens
4. O ciclo se repete até a resposta final, exibida no chat

Estrutura em `src/taskpane/`:

```
src/taskpane/
├── taskpane.js        # entry point: wiring de eventos
├── config.js          # persistência de configurações (localStorage)
├── http.js            # proxyUrl + requestJson (XHR promisificado)
├── state.js           # histórico da conversa e estado de carregamento
├── chatView.js        # renderização de mensagens e indicador de loading
├── settingsView.js    # formulário de configurações
├── modelSelect.js     # fetch/render do seletor de modelos
├── agent.js           # sendMessage + loop de tool calling
└── tools/
    ├── schema.js      # SYSTEM_PROMPT + definição JSON das ferramentas
    └── excel.js       # execução das ferramentas via Office.js
```

A camada `tools/` foi projetada para crescer por aplicativo: cada novo app do Office receberá seu próprio módulo de execução (ex.: `tools/word.js`, `tools/outlook.js`) e entradas correspondentes em `schema.js`, reaproveitando o mesmo loop do `agent.js`.

## Ferramentas disponíveis (Excel)

| Ferramenta | Descrição |
|---|---|
| `read_cells` | Lê valores de um intervalo (ex.: `"B2:D10"`) |
| `write_cells` | Escreve valores em células a partir de um range inicial |
| `format_cells` | Aplica negrito, itálico, cor de fundo, cor da fonte, tamanho |
| `clear_cells` | Limpa o conteúdo de um intervalo |
| `get_sheet_info` | Retorna nome da aba ativa e intervalo usado |
| `get_selection` | Retorna endereço, dimensões e valores da seleção atual |

## Como executar

Pré-requisitos: [Node.js](https://nodejs.org) (LTS) e npm.

```bash
# 1. Instalar dependências
npm install

# 2. Confiar nos certificados de desenvolvimento (uma vez só)
npm run dev-certs

# 3. Iniciar o add-in (build + servidor dev + sideload no Excel)
npm start
```

Ao finalizar a depuração:

```bash
npm run stop
```

Outros comandos úteis:

```bash
npm run build:dev   # build de desenvolvimento
npm run build       # build de produção
npm run lint        # ESLint + Prettier
npm run lint:fix    # correções automáticas de lint
npm run validate    # valida o manifest.xml
```

## Configuração da API

1. Abra o painel do AI2office no Excel
2. Clique em ⚙️ e preencha:
   - **Endpoint**: URL base compatível com OpenAI (ex.: `https://api.openai.com/v1`, `http://localhost:1234/v1`)
   - **Chave de API**: token de acesso (pode ficar vazio em servidores locais)
   - **Modelo**: selecionado automaticamente a partir dos modelos expostos pela API
3. Salve — as credenciais ficam apenas no `localStorage` do painel

> Se você usa um servidor local (LM Studio, Ollama), pode ser necessário habilitar CORS ou usar HTTPS no endpoint, pois o add-in roda em `https://localhost`.

## Stack técnica

- [Office.js](https://learn.microsoft.com/office/dev/add-ins/) — integração com os aplicativos Office
- Webpack 5 + Babel — build e transpilação (ES5 para compatibilidade)
- ESLint (flat config) + Prettier — qualidade e formatação de código
- JavaScript puro, sem frameworks — bundle leve e controle total do DOM

## Roadmap

- [ ] Validação completa em runtime das ferramentas do Excel
- [ ] Novas ferramentas do Excel (gráficos, ordenar/filtrar, tabelas)
- [ ] Streaming de respostas do modelo
- [ ] Módulo de ferramentas para **Word**
- [ ] Módulo de ferramentas para **PowerPoint**
- [ ] Módulo de ferramentas para **Outlook**
