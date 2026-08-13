# Bot Telegram Pharos — Especificação do Projeto

> Documento de referência para desenvolvimento no Claude Code. Contém arquitetura, prompts finais, specs visuais extraídas dos templates reais e roadmap de implementação.

---

## 1. Visão geral

Bot pessoal no Telegram (uso exclusivo, whitelist de 1 `chat_id`) que automatiza 4 fluxos de trabalho na Pharos Consultoria:

1. **Atas de reunião** — transcrição (texto ou `.docx`) → `.docx` formatado no papel timbrado Pharos.
2. **Outlook** — criar/editar/listar eventos via Microsoft Graph.
3. **Banco de horas** — lançamento automatizado no site interno da empresa (sem API — via automação de navegador).
4. **Relatório semanal (OPR)** — texto livre → PPT de uma página + PDF + mensagem de follow-up.

Rodando 24h num VPS (Render), IA de backend: **Gemini API**.

---

## 2. Stack técnica (decisão: unificar em Node.js)

Motivo: o gerador de PPT (`pptxgenjs`) e o parser de Excel (`xlsx`/SheetJS) já existem prontos em JavaScript (ver `opr_pharos_gera_ppt_local.html` do usuário) — reaproveitar em vez de reescrever em Python.

| Módulo | Biblioteca |
|---|---|
| Bot / Telegram | `telegraf` (webhook, não polling) |
| IA | `@google/generative-ai` (Gemini) |
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
4. Bot mostra **preview formatado** (não o JSON cru) + botões `✅ Confirmar` / `✏️ Corrigir`.
5. Se `Corrigir`: usuário digita a correção em texto livre; bot reenvia JSON atual + correção ao Gemini, pedindo só o JSON atualizado.
6. Confirmado → gera follow-up (template determinístico, sem IA) + PPT (`pptxgenjs`) + PDF (LibreOffice).
7. Entrega: mensagem de follow-up (texto, pronta pra copiar) + `.pptx` + `.pdf`.

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

- Saudação: decidir se é fixa ("Bom dia, amigos! Todos bem?") ou calculada por horário — **pendente de decisão do usuário**.
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

---

## 8. Segurança e infraestrutura

- Bot só responde ao `chat_id` do usuário (whitelist no middleware).
- Webhook (não polling) — mais estável para uptime 24h.
- Todas as credenciais (Gemini API key, tokens Graph, login do banco de horas) em variáveis de ambiente no painel do Render — nunca no código.
- Plano pago no Render (free hiberna e quebra o funcionamento contínuo).
- Healthcheck endpoint simples para monitoramento externo (ex.: UptimeRobot apontando pro webhook).

---

## 9. Roadmap de implementação (ordem definida)

**Fase 0 — Infra**
Bot básico (`/start`, `/ping`) no webhook do Render, whitelist, Postgres conectado.
Critério de pronto: bot responde depois de 24h sem intervenção.

**Fase 1 — Ata**
Placeholders do template `.docx` + prompt da seção 4.3 + preenchimento via `docx` (npm).
Critério de pronto: manda texto de reunião, recebe `.docx` formatado.

**Fase 2 — Outlook**
OAuth2 do Graph + criar/editar/listar eventos.
Critério de pronto: cria evento pelo Telegram e aparece no Outlook.

**Fase 3 — Relatório semanal + follow-up**
Fluxo da seção 7 completo (exemplo → texto livre → extração → preview → confirmação → PPT + PDF + follow-up).
Critério de pronto: `/relatorio_semanal` entrega os 3 outputs.

**Fase 4 — Banco de horas (RPA)**
Playwright headless com tratamento de erro robusto.
Critério de pronto: `/banco_horas` lança e retorna print de confirmação.

---

## 10. Decisões em aberto (revisar antes ou durante a Fase 3)

- [ ] Saudação da mensagem de follow-up: fixa ou variável por horário/humor?
- [ ] Fonte das reuniões no relatório semanal: só texto livre, ou puxar automaticamente do Outlook (Fase 2) como complemento/checagem?
- [ ] Nome exato dos placeholders no template `.docx` da ata (a definir durante a Fase 1, com o arquivo real em mãos no Claude Code).

---

## 11. Arquivos de referência usados nesta especificação

- `Papel_Timbrado_-_Pharos.docx` — fonte das margens/estilos da seção 4.1
- `Ata_Truckão_-_28-07-2026_-_Alinhamento_Interno.docx` — gabarito de validação de conteúdo/estrutura da ata
- `opr_pharos_gera_ppt_local.html` — script `pptxgenjs` a portar (seção 7.1)
- `Modelo_Agenda_Semanal_-_Pharos.xlsx` — estrutura de campos original (substituída pelo fluxo de texto livre, mas mantida aqui como referência de nomenclatura)
- `OPR_Tuckão_Autopeças_*.pptx` — exemplos de output esperado do módulo de relatório semanal
