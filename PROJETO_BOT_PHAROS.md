# Bot Telegram Pharos — Especificação do Projeto

> Documento de referência para desenvolvimento no Claude Code. Contém arquitetura, prompts finais, specs visuais extraídas dos templates reais e roadmap de implementação.

---

## Status atual (atualizado em 13/08/2026)

### ✅ Implementado

**Infra**
- Bot Telegraf — webhook em produção (Render) / polling em desenvolvimento.
- Cadastro por senha (`/start`) com dados salvos no Postgres; whitelist efetiva é "usuário cadastrado e ativo" (o middleware `whitelist.ts` por `ALLOWED_CHAT_ID` existe no código mas não está mais em uso — foi substituído pelo fluxo de cadastro).
- Sessões de conversa persistidas em banco (`sessions`), healthcheck em `/health`.
- `/perfil` — visualizar e editar nome, e-mail e login/senha do banco de horas (senha criptografada com AES-256-GCM).

**`/ata`** — transcrição de reunião (texto, `.docx` ou `.txt`) → pergunta projeto/assunto/data, modalidade e nível de detalhe → Gemini estrutura em JSON (prompt seção 4.3) → gera `.docx` no papel timbrado real da Pharos (margens, cores, fontes da marca) → entrega o arquivo + lista de pontos a confirmar.

**`/opr`** — relatório semanal completo:
1. Texto livre → Gemini extrai atividades realizadas / próximas / reuniões (prompt seção 7.3).
2. Pergunta cliente e período da semana, se a IA não identificar no texto.
3. Confirma o dia da semana de cada reunião — calculado deterministicamente a partir da data numérica (nunca confia no "dia" que a IA supôs).
4. Preview com botões Confirmar/Corrigir antes de gerar.
5. Entrega mensagem de follow-up (saudação calculada por horário, `America/Fortaleza`) + `.pptx` — logo real da Pharos (farol dourado, extraído do material da marca), fontes Playfair Display/Helvetica Now, cards com altura/espaçamento dinâmicos (nunca estouram o slide, mesmo com várias reuniões na semana).
6. **Sem PDF** — foi removido; hoje o `/opr` entrega só `.pptx` (ver "Pendente" abaixo).

**`/followup`** — mesmo fluxo de extração e confirmação de dias do `/opr`, mas sem perguntar cliente/semana e sem gerar PPT: entrega só a mensagem de follow-up semanal, pronta pra copiar.

**`/turno`** — pergunta cliente + data do turno → relato livre do que foi feito/está pendente/vai fazer → Gemini estrutura em três parágrafos seguindo o guia oficial de follow-up de turno da Pharos ("O que foi feito no turno" / "O que ficou pendente" / "O que farei no próximo turno"), preservando números e nomes citados, sem inventar conteúdo.

Todos os fluxos de IA (ata, opr, followup, turno) seguem o mesmo princípio: a IA nunca completa informação que não foi dita, marca "(a confirmar)" quando há ambiguidade, e o usuário sempre revisa antes do resultado final ser entregue.

### ⏳ Pendente

- **Fase 2 — Outlook (Microsoft Graph)**: nada implementado. Sem OAuth2/MSAL, sem comando `/evento` (criar/editar/listar). A tabela `oauth_tokens` já existe no schema (seção 3), esperando por isso.
- **Fase 4 — Banco de horas (RPA)**: nada implementado. Sem Playwright, sem comando `/banco_horas`. O cadastro já coleta e criptografa login/senha do sistema de banco de horas, esperando por isso.
- **PDF do relatório semanal**: removido a pedido do usuário. Para reativar no futuro, precisa de LibreOffice (`soffice --headless`) no ambiente de deploy — o Render nativo (`runtime: node`) não tem isso disponível; exigiria migrar pra `runtime: docker` com um Dockerfile próprio, ou usar um serviço externo de conversão.
- **Decisões em aberto da seção 10**: revisar antes de avançar a Fase 2 (fonte das reuniões do relatório semanal — só texto livre ou puxar do Outlook como complemento).

---

## 1. Visão geral

Bot pessoal no Telegram (uso exclusivo, acesso por cadastro com senha) que automatiza fluxos de trabalho na Pharos Consultoria:

1. **Atas de reunião** ✅ — transcrição (texto ou `.docx`) → `.docx` formatado no papel timbrado Pharos. (`/ata`)
2. **Relatório semanal (OPR)** ✅ — texto livre → PPT de uma página + mensagem de follow-up. (`/opr`)
3. **Follow-up semanal avulso** ✅ — mesmo fluxo do OPR, só a mensagem, sem PPT. (`/followup`, não estava no plano original)
4. **Follow-up de turno** ✅ — relato livre → mensagem em 3 seções (feito/pendente/próximo). (`/turno`, não estava no plano original)
5. **Outlook** ⏳ — criar/editar/listar eventos via Microsoft Graph. Não iniciado.
6. **Banco de horas** ⏳ — lançamento automatizado no site interno da empresa (sem API — via automação de navegador). Não iniciado.

Rodando 24h num VPS (Render), IA de backend: **Gemini API**.

---

## 2. Stack técnica (decisão: unificar em Node.js)

Motivo: o gerador de PPT (`pptxgenjs`) e o parser de Excel (`xlsx`/SheetJS) já existem prontos em JavaScript (ver `opr_pharos_gera_ppt_local.html` do usuário) — reaproveitar em vez de reescrever em Python.

| Módulo | Biblioteca |
|---|---|
| Bot / Telegram | `telegraf` (webhook em produção, polling em dev) |
| IA | `@google/genai` (Gemini — migrado do `@google/generative-ai` original) |
| Geração de Word (ata) | `docx` (npm) |
| Leitura de `.docx` enviado | `mammoth` (extrai texto preservando estrutura) |
| Geração de PPT (OPR) | `pptxgenjs` — **reaproveitar script existente quase 1:1** |
| PPT → PDF | LibreOffice headless (`soffice --headless --convert-to pdf`) |
| Outlook | `@microsoft/microsoft-graph-client` + `@azure/msal-node` (OAuth2) |
| Banco de horas (RPA) | `playwright` (Node) |
| Banco de dados | Postgres (Render) — tokens, estado de conversas, sessões |
| Scheduler | `node-cron` ou cron job nativo do Render |
| Deploy | Render Web Service (plano pago — free hiberna e quebra o "24h") |

---

## 3. Estrutura de pastas sugerida

> Esta é a estrutura planejada originalmente. A árvore real do repositório evoluiu de forma um pouco diferente (ex.: `commands/relatorio_semanal.ts` cobre `/opr` e `/followup`; `commands/turno.ts` e `prompts/turno.prompt.ts` foram adicionados; `modules/pdf/` foi criado e depois removido junto com o PDF do OPR; `evento.ts`, `banco_horas.ts`, `outlook/` e `banco-horas/` ainda não existem, por serem Fases 2 e 4). Consulte a árvore atual do repositório para a estrutura exata.

```
pharos-bot/
├── src/
│   ├── bot/
│   │   ├── index.ts            # inicialização do telegraf + webhook
│   │   ├── middleware/
│   │   │   └── whitelist.ts    # só responde ao chat_id autorizado
│   │   └── commands/
│   │       ├── ata.ts
│   │       ├── evento.ts
│   │       ├── banco_horas.ts
│   │       └── relatorio_semanal.ts
│   ├── modules/
│   │   ├── gemini/
│   │   │   ├── client.ts
│   │   │   ├── prompts/
│   │   │   │   ├── ata.prompt.ts
│   │   │   │   └── relatorio-semanal.prompt.ts
│   │   ├── docx/
│   │   │   ├── ata-template.ts     # margens/estilos do papel timbrado
│   │   │   └── assets/
│   │   │       └── header1.xml     # extraído do timbrado original
│   │   ├── pptx/
│   │   │   └── opr-generator.ts    # port do opr_pharos_gera_ppt_local.html
│   │   ├── outlook/
│   │   │   └── graph-client.ts
│   │   ├── banco-horas/
│   │   │   └── playwright-flow.ts
│   │   └── pdf/
│   │       └── convert.ts          # wrapper do soffice headless
│   └── db/
│       └── schema.sql
├── assets/
│   ├── papel-timbrado.png          # imagem de fundo extraída do docx
│   └── pharos-fonts/               # Playfair Display, Helvetica Now
├── .env.example
├── package.json
└── README.md
```

---

## 4. Módulo 1 — Ata de reunião

### 4.1 Especificação visual (extraída do `Papel_Timbrado_-_Pharos.docx` real)

- Página A4: `w=11906` `h=16838` (twips)
- Margens: `top=2268` `right=1417` `bottom=1531` `left=1417` (twips) — **essencial**, é o que evita o texto invadir a moldura do timbrado
- Imagem de fundo ancorada no header (`word/header1.xml`), full-page
- Títulos de seção: Playfair Display, verde `#14663D`, linha inferior dourada `#AF7932`
- Título principal: Playfair Display, verde escuro `#072E22`
- Corpo: Helvetica Now 11pt `#111111`; destaques inline dourado `#AF7932` negrito
- Subtítulos de tópico: itálico dourado `#AF7932`
- Tabela de encaminhamentos: cabeçalho fundo `#072E22` texto branco; linhas alternadas `#FFFFFF`/`#F0F4F0`; números `#AF7932` negrito
- Rodapé fixo: "Ata elaborada por Pharos Consultoria." + "Você define o destino, nós mostramos o caminho."
- Nome do arquivo: `Ata [Projeto] - [DD-MM-AAAA] - [Assunto].docx`

### 4.2 Fluxo

1. Usuário manda transcrição (texto direto ou arquivo `.docx`/`.txt`).
2. Se `.docx`: extrair texto via `mammoth`.
3. Bot pergunta (Telegram, resposta rápida): **modalidade/local** da reunião e **nível de detalhe** (com números / sem números) — só se não estiver claro no texto.
4. Bot monta o prompt (seção 4.3) e chama o Gemini.
5. Gemini retorna JSON estruturado (nunca o arquivo — API só devolve texto).
6. Bot preenche o template `.docx` (script `ata-template.ts`) com o JSON, aplicando os estilos da seção 4.1.
7. Bot envia o `.docx` final + lista de itens `(a confirmar)` em uma mensagem separada.

### 4.3 Prompt final (ata)

```
SYSTEM (fixo no bot):

Você é o assistente de atas da Pharos Consultoria. Vai receber a
transcrição completa de uma ou mais reuniões/entrevistas de um mesmo
projeto e deve devolver APENAS um JSON válido (sem markdown, sem texto
fora do JSON), no schema abaixo.

═══ LEITURA COMPLETA — REGRA ABSOLUTA ═══
Leia e processe o texto INTEIRO antes de responder, independentemente do
tamanho. Nenhum trecho relevante pode ficar de fora por limitação de
espaço. Se a transcrição contiver várias reuniões/entrevistas juntas
(identificadas por marcações de data/hora, mudança de participantes,
cabeçalhos como "Entrevista 2" etc.), trate cada uma como uma fonte de
informação para a MESMA ata: agrupe por assunto — não crie uma seção por
reunião. Se duas fontes falarem do mesmo tema em momentos diferentes,
junte e conecte as informações no mesmo subtema (2.1, 2.2...). Se duas
fontes se contradisserem sobre o mesmo fato, registre as duas versões e
marque "(a confirmar)", citando que houve divergência entre falas —
nunca escolha uma versão por conta própria.

═══ FILTRO DE RELEVÂNCIA ═══
Nem tudo que foi dito vira ata. Identifique e EXCLUA do conteúdo final:
- Conversa social, piadas, comentários pessoais, fofoca, histórias que
  não têm relação com o projeto/processo em discussão.
- Digressões sem consequência prática para o diagnóstico, decisão ou
  encaminhamento.
Mantenha, mesmo que soe informal: analogias, exemplos ou anedotas que a
pessoa usou para ILUSTRAR um ponto de negócio real — o critério é
relevância para o projeto, não o tom da fala.

═══ RUÍDO DE TRANSCRIÇÃO x INVENÇÃO — SÃO COISAS DIFERENTES ═══
Transcrição automática erra: palavras trocadas por termos parecidos,
trechos em inglês mal captados, nomes fonéticos. Você PODE e DEVE
corrigir isso quando o contexto deixa claro qual era a palavra/termo
pretendido (ex.: um termo técnico do setor do cliente, mal grafado, mas
óbvio pelo resto da frase). Isso é correção de forma, não invenção de
conteúdo. O que você NUNCA pode fazer: completar informação que não foi
dita, supor motivo/causa não explicitado, arredondar ou estimar número,
presumir responsável não citado, ou inferir uma decisão a partir de uma
discussão que não chegou a ser fechada.
Se depois de tentar inferir pelo contexto a palavra/nome/número
continuar ambíguo, NÃO adivinhe — marque "(a confirmar)" e, se fizer
diferença, registre as opções ouvidas (ex.: "Jorge/George — grafia a
confirmar").

═══ NÚMEROS E FATOS ═══
Reproduza valores, datas, percentuais e prazos exatamente como foram
ditos. Nunca arredonde, nunca estime, nunca crie um número plausível
para preencher lacuna. Se o número foi dito com ressalva ("acho que",
"salvo engano", "por volta de"), transcreva o valor E marque
"(a confirmar)".

═══ ESTRUTURA DE SAÍDA ═══
"secoes": organizadas por subtema (2.1, 2.2...), em prosa, cobrindo TODO
o conteúdo relevante identificado — não resuma a ponto de perder
informação, o objetivo é completude com linguagem profissional, não
brevidade. Último item de "secoes" = "Decisões e alinhamentos", em
tópicos.

Antes de finalizar, faça uma checagem interna: todo tema, decisão ou
encaminhamento relevante mencionado em qualquer trecho da transcrição
está representado em "secoes" ou "encaminhamentos"? Se não, adicione.

Schema JSON de saída:
{
  "projeto": string, "assunto": string, "data": string,
  "duracao": string, "modalidade": string,
  "presentes": [{"nome": string, "empresa": string, "papel": string}],
  "observacao_participantes": string | null,
  "secoes": [{"titulo": string, "corpo": string}],
  "decisoes_alinhamentos": [string],
  "encaminhamentos": [{"numero": int, "encaminhamento": string,
                        "responsavel": string, "situacao": string}],
  "pontos_a_confirmar": [string]
}

USER (montado pelo bot a cada chamada):
Projeto/Cliente: {{projeto}}
Modalidade e local: {{modalidade}}
Nível de detalhe: {{detalhe}}
Transcrição (pode conter mais de uma reunião/entrevista):
{{transcricao_completa}}
```

---

## 5. Módulo 2 — Outlook (Microsoft Graph)

- Registro de app no Azure AD (gratuito), fluxo OAuth2 uma vez, refresh token salvo no Postgres.
- Comandos: `/evento criar`, `/evento editar`, `/evento listar`.
- Usado também como fonte de dados para o relatório semanal (reuniões da semana), se o usuário optar por puxar do calendário em vez de digitar manualmente.

---

## 6. Módulo 3 — Banco de horas (RPA)

- Login apenas usuário/senha (sem 2FA confirmado pelo usuário).
- Playwright headless: login → preencher formulário → confirmar → screenshot de confirmação enviado de volta no Telegram.
- **Tratamento de erro obrigatório**: se o site mudar de layout, o bot deve detectar falha no seletor e avisar no Telegram em vez de falhar silenciosamente ou lançar dado errado.
- Fase mais frágil do projeto — implementar por último, depois da infra estar validada nos outros módulos.

---

## 7. Módulo 4 — Relatório semanal (OPR) + follow-up

> **Status**: implementado como `/opr` (não `/relatorio_semanal`, nome original). Ver "Status atual" no topo do documento para o que mudou de escopo (PDF removido, logo real + fontes da marca no PPT, comandos extras `/followup` e `/turno`).

### 7.1 Especificação visual do PPT (extraída de `opr_pharos_gera_ppt_local.html`)

Layout `LAYOUT_16x9` (10" × 5.625"). Paleta:

```js
const C = {
  dg: '072E22',    // verde escuro (dark green)
  mg: '14663D',    // verde médio (medium green)
  gold: 'AF7932',
  off: 'F2F5F2',
  white: 'FFFFFF',
  border: 'C2D4C5',
  muted: '5A7A60',
  light: 'E8F0EA'
};
```

Estrutura: barra superior com logo "PHAROS" + pill "OPR – One Page Report" + "CONFIDENCIAL", 3 colunas (Atividades Realizadas / Próximas Entregas / Reuniões Agendadas) com cards por item, rodapé com cliente/semana. **O script completo já existe e deve ser portado quase 1:1** — está em `opr_pharos_gera_ppt_local.html`, função `gerarPPT()`. No backend, a única mudança é trocar `pres.writeFile({fileName})` (que baixa no navegador) por salvar em disco/buffer para anexar no Telegram.

**Ajustes feitos em relação ao script original** (`src/modules/pptx/opr-generator.ts`):
- Logo "PHAROS" trocado por imagem real do farol dourado da marca (`assets/pharos-logo.png`), com o texto "PHAROS" ao lado.
- Fontes trocadas de Georgia/Trebuchet MS para **Playfair Display** (títulos) e **Helvetica Now** (corpo) — mesma dupla usada no template de ata.
- Layout dos cards (atividades e reuniões) ficou responsivo: altura e espaçamento se ajustam dinamicamente para nunca ultrapassar o limite do slide, mesmo com várias reuniões na mesma semana.
- Nome do arquivo segue a convenção `OPR_{cliente com "_" no lugar de espaço}_{semana com "-" no lugar de "/" e espaços}.pptx`.

### 7.2 Fluxo de entrada — texto livre (decisão final, substituiu Excel e formato rígido)

1. Bot mostra um **exemplo ilustrativo** (não é template a seguir à risca) e pede o texto da semana:

```
📋 Me conta como foi a semana, do seu jeito — pode escrever corrido.

Exemplo do tipo de informação que preciso encontrar no seu texto
(não precisa seguir essa ordem nem esse formato):

"Essa semana avançamos na análise dos dados pra RMR com Lucas e
Hudson, e seguimos as entrevistas de processo na oficina com Lucas
e Ítalo. Pra semana que vem: reunião mensal de resultados com o
Rogério, e apresentação dos gargalos da oficina. Reuniões marcadas:
segunda 10/08 às 9h, apresentação da RMR, in loco na Paupina, com
Lucas, Hudson e Ítalo; quarta 12/08 às 14h, apresentação da análise
de processos, mesma equipe, in loco na Paupina."

Pode mandar sua mensagem 👇
```

2. Usuário manda texto livre.
3. Gemini extrai estrutura (prompt 7.3).
4. **[implementado, não estava no plano original]** Bot pergunta o cliente e o período da semana, se a IA não identificou no texto (o período vira o nome do arquivo).
5. **[implementado, não estava no plano original]** Bot mostra os dias da semana de cada reunião, **calculados deterministicamente a partir da data numérica** (nunca a partir do que a IA "achou" que era o dia) — usuário confirma ou corrige antes de seguir.
6. Bot mostra **preview formatado** (não o JSON cru) + botões `✅ Confirmar` / `✏️ Corrigir`.
7. Se `Corrigir`: usuário digita a correção em texto livre; bot reenvia JSON atual + correção ao Gemini, pedindo só o JSON atualizado.
8. Confirmado → gera follow-up (template determinístico, sem IA) + PPT (`pptxgenjs`).
9. Entrega: mensagem de follow-up (texto, pronta pra copiar) + `.pptx`. **Sem PDF** — foi removido do escopo (ver "Status atual").

### 7.3 Prompt de extração (relatório semanal)

```
SYSTEM (fixo no bot):

Você recebe um texto livre escrito por um consultor da Pharos contando
como foi a semana de um projeto. Extraia as informações em JSON
estruturado. Data de hoje: {{data_hoje}} — use isso para resolver
referências relativas ("semana que vem", "quinta-feira", etc.).

═══ REGRA ABSOLUTA PARA REUNIÕES ═══
Data, horário e local de reunião só entram no JSON se estiverem
EXPLICITAMENTE ditos no texto. Nunca calcule, nunca estime, nunca
presuma horário "padrão". Se a pessoa disser um dia da semana sem data
numérica, calcule a data a partir de {{data_hoje}} — mas se o
dia/horário for vago ("ainda essa semana", "de manhã"), marque o campo
correspondente como "(a confirmar)" em vez de adivinhar.

═══ REGRA GERAL — NÃO INVENTAR ═══
Nunca crie responsável, cliente, objetivo ou atividade que não esteja
no texto. Campo não mencionado = "-". Nunca invente números.

═══ CLASSIFICAÇÃO REALIZADA x PRÓXIMA ═══
Separe pelo tempo verbal e contexto: o que já aconteceu (realizada) x
o que está planejado (próxima). Na dúvida entre as duas, use o contexto
temporal explícito no texto, nunca suponha.

═══ RESPONSÁVEIS ═══
Se o texto citar nomes junto da atividade/reunião (ex.: "com o Lucas e
o Hudson"), associe-os ao campo correto (Pharos ou Cliente) só se for
identificável pelo contexto quem é de qual lado. Se não for claro de
qual lado a pessoa é, coloque em "ph" (Pharos) apenas se for nome de
consultor já conhecido do projeto; senão registre em "cli".

Gere um "obj" curto pra cada item, baseado só no que foi dito — sem
elaborar além do texto.

Schema de saída (APENAS JSON, sem texto fora dele):
{
  "cliente": string|null, "semana": string|null,
  "real": [{"ativ":string,"obj":string,"cli":string,"ph":string}],
  "prox": [{"ativ":string,"obj":string,"cli":string,"ph":string}],
  "meet": [{"dia":string,"data":string,"horario":string,
            "local":string,"obj":string,"cli":string,"ph":string}]
}

USER:
{{texto_livre_do_usuario}}
```

### 7.4 Template da mensagem de follow-up (determinístico, sem IA)

```
{saudação}

Nessa semana avançamos com os seguintes pontos:
{para cada item de "real"}: - {ativ} ({obj});

Próximos passos:
{para cada item de "prox"}: - {ativ};

Reuniões Agendadas:
{para cada item de "meet"}: - {dia} ({data}) às {horario} - {obj} - {local}
```

- Saudação: **decidido — calculada por horário** (Bom dia / Boa tarde / Boa noite, `America/Fortaleza`), seguida de "amigos! Todos bem?".
- Bloco "Próximas" não leva parênteses de objetivo (assimetria intencional, conforme exemplo original do usuário).
- Bloco "Realizadas" leva `obj` já vindo enxuto do próprio prompt de extração (7.3) — não precisa de segunda chamada de IA.

### 7.5 Preview de confirmação (formato usado no passo 4 da seção 7.2)

```
Entendi assim 👇

✅ Realizadas (N)
1. {ativ} — Pharos: {ph}

➡️ Próximas (N)
1. {ativ} — Cliente: {cli}

📅 Reuniões (N)
1. {dia} {data} às {horario} — {obj} — {local}

[✅ Confirmar]   [✏️ Corrigir]
```

### 7.6 `/followup` — só a mensagem semanal, sem PPT [não estava no plano original]

Pedido do usuário depois do `/opr` estar pronto: em alguns casos só a mensagem de follow-up é necessária, sem gerar o PPT.

Reaproveita **exatamente** o mesmo fluxo do `/opr` (extração, confirmação de dias das reuniões, preview, correção) — a única diferença é que **não pergunta cliente nem semana** (campos que não aparecem na mensagem de follow-up) e o passo final só envia a mensagem de texto, sem chamar `generateOprPptx`.

Implementação: `src/bot/commands/relatorio_semanal.ts` compartilha todo o código entre `/opr` e `/followup` através de um parâmetro `mode: 'opr' | 'followup'` guardado na sessão, que controla se as perguntas de cliente/semana aparecem e se o PPT é gerado no final.

### 7.7 `/turno` — follow-up de turno [não estava no plano original]

Pedido do usuário: mensagem curta de passagem de turno, formato diferente do relatório semanal — baseado no guia oficial "Follow-up de turno" da Pharos.

**Fluxo:**
1. Bot pergunta cliente e data do turno (formato `Cliente - Data`, ex.: `Auto Center - 31/07`).
2. Usuário manda relato livre do que fez, o que ficou pendente e o que pretende fazer no próximo turno.
3. Gemini estrutura em três parágrafos objetivos, em primeira pessoa (prompt em `src/modules/gemini/prompts/turno.prompt.ts`), preservando números e nomes citados — nunca inventa pendência ou atividade não mencionada.
4. Preview com `✅ Está bom` / `✏️ Corrigir` antes de finalizar.

**Formato final (template determinístico, montado pelo bot):**
```
Follow-up de turno — {Cliente} ({Data})

O que foi feito no turno
{parágrafo gerado pela IA — pretérito perfeito, resultados concretos, menciona quem participou}

O que ficou pendente
{parágrafo gerado pela IA — específico, com responsável quando citado; nunca genérico tipo "falta terminar"}

O que farei no próximo turno
{parágrafo gerado pela IA — futuro/presente, o que foi combinado com cliente/gestor quando citado}
```

Schema JSON do Gemini: `{ "feito": string, "pendente": string, "proximo": string }`.

---

## 8. Segurança e infraestrutura

- Bot só responde ao `chat_id` do usuário (whitelist no middleware).
- Webhook (não polling) — mais estável para uptime 24h.
- Todas as credenciais (Gemini API key, tokens Graph, login do banco de horas) em variáveis de ambiente no painel do Render — nunca no código.
- Plano pago no Render (free hiberna e quebra o funcionamento contínuo).
- Healthcheck endpoint simples para monitoramento externo (ex.: UptimeRobot apontando pro webhook).

---

## 9. Roadmap de implementação (ordem definida)

**Fase 0 — Infra** ✅ concluída
Bot no webhook do Render, cadastro por senha (substituiu a whitelist estática original), Postgres conectado.
Critério de pronto: bot responde depois de 24h sem intervenção. ✅

**Fase 1 — Ata** ✅ concluída
Placeholders do template `.docx` + prompt da seção 4.3 + preenchimento via `docx` (npm).
Critério de pronto: manda texto de reunião, recebe `.docx` formatado. ✅

**Fase 2 — Outlook** ⏳ não iniciada
OAuth2 do Graph + criar/editar/listar eventos.
Critério de pronto: cria evento pelo Telegram e aparece no Outlook.

**Fase 3 — Relatório semanal + follow-up** ✅ concluída (com mudanças de escopo)
Fluxo da seção 7 completo (exemplo → texto livre → extração → confirmação de cliente/semana/dia da semana → preview → confirmação → PPT + follow-up).
Critério de pronto: `/opr` entrega PPT + follow-up. ✅ (PDF foi removido do escopo, ver "Status atual")
Além do planejado originalmente, dois comandos extras foram adicionados: `/followup` (só a mensagem semanal, sem PPT) e `/turno` (follow-up de turno em 3 seções, formato próprio — não fazia parte da especificação original).

**Fase 4 — Banco de horas (RPA)** ⏳ não iniciada
Playwright headless com tratamento de erro robusto.
Critério de pronto: `/banco_horas` lança e retorna print de confirmação.

---

## 10. Decisões em aberto (revisar antes ou durante a Fase 2)

- [x] Saudação da mensagem de follow-up: **variável por horário** (Bom dia / Boa tarde / Boa noite, calculado em `America/Fortaleza`).
- [ ] Fonte das reuniões no relatório semanal: só texto livre, ou puxar automaticamente do Outlook (Fase 2) como complemento/checagem?
- [x] Nome exato dos placeholders no template `.docx` da ata: definido em `src/modules/docx/ata-generator.ts`.
- [ ] PDF do relatório semanal: reativar ou manter só PPT? Se reativar, decidir entre migrar o deploy pra `runtime: docker` (LibreOffice) ou usar um serviço externo de conversão.

---

## 11. Arquivos de referência usados nesta especificação

- `Papel_Timbrado_-_Pharos.docx` — fonte das margens/estilos da seção 4.1
- `Ata_Truckão_-_28-07-2026_-_Alinhamento_Interno.docx` — gabarito de validação de conteúdo/estrutura da ata
- `opr_pharos_gera_ppt_local.html` — script `pptxgenjs` a portar (seção 7.1)
- `Modelo_Agenda_Semanal_-_Pharos.xlsx` — estrutura de campos original (substituída pelo fluxo de texto livre, mas mantida aqui como referência de nomenclatura)
- `OPR_Tuckão_Autopeças_*.pptx` — exemplos de output esperado do módulo de relatório semanal
