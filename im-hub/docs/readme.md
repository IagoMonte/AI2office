# AI2Office

Assistente de IA para o Microsoft Office: um painel de chat (task pane) que lê, escreve e formata documentos por linguagem natural.

## Problema

Tarefas repetitivas em planilhas, documentos, apresentações e e-mails exigem navegar menus e decorar fórmulas. O AI2Office deixa o usuário descrever o que quer e o modelo executa a edição direto no arquivo aberto.

## O que foi construído

- Chat com loop agêntico de *tool calling*: o modelo escolhe as ferramentas, o add-in executa via Office.js e devolve o resultado até a resposta final.
- Conexão com **qualquer API compatível com OpenAI** (`/chat/completions`): OpenAI, Azure OpenAI, LM Studio, Ollama, proxies locais. Endpoint, chave e modelo são configurados no painel; a lista de modelos vem da própria API.
- Módulos por aplicativo, escolhidos pelo host ativo: **Excel**, **Word**, **PowerPoint** e **Outlook**.
- Consciência de seleção: o que o usuário selecionou é injetado no contexto ("isto", "aqui" funcionam sem perguntas extras).
- Anexos (`.xlsx`, `.xls`, `.csv`, `.txt`, `.pdf`, `.docx`) por botão ou arrastar e soltar, extraídos no próprio add-in.
- Sem back-end: as credenciais ficam só no `localStorage` do painel.

## Como rodar

Requisitos: Node.js (LTS) e npm.

```bash
npm install
npm run dev-certs   # uma vez, confia nos certificados de desenvolvimento
npm start           # build, servidor dev e sideload no Office
npm run stop        # encerra a depuração
```

Depois, abra o painel, clique em ⚙️ e informe endpoint, chave e modelo.

## Próximos passos

- Validação completa em runtime das ferramentas de cada aplicativo.
- Streaming das respostas do modelo.
- Mais ferramentas (gráficos, ordenar/filtrar e tabelas no Excel; títulos, estilos e listas no Word).
