import { Context, Telegraf } from 'telegraf';
import { getSession, setSession, clearSession } from '../../db/sessions';
import { callGemini, extractJSON } from '../../modules/gemini/client';
import {
  buildRelatorioSystem,
  buildRelatorioUserMessage,
  buildRelatorioCorrectionMessage,
} from '../../modules/gemini/prompts/relatorio-semanal.prompt';
import { generateOprPptx, buildOprFileName, OprData } from '../../modules/pptx/opr-generator';

// 'opr'      → /opr: gera a mensagem de follow-up + o PPT.
// 'followup' → /followup: gera só a mensagem de follow-up semanal (sem cliente/semana/PPT).
type RelatorioMode = 'opr' | 'followup';

const EXEMPLO_BASE = `Exemplo do tipo de informação que preciso encontrar no seu texto
(não precisa seguir essa ordem nem esse formato):

"Essa semana avançamos na análise dos dados pra RMR com Lucas e
Hudson, e seguimos as entrevistas de processo na oficina com Lucas
e Ítalo. Pra semana que vem: reunião mensal de resultados com o
Rogério, e apresentação dos gargalos da oficina. Reuniões marcadas:
segunda 10/08 às 9h, apresentação da RMR, in loco na Paupina, com
Lucas, Hudson e Ítalo; quarta 12/08 às 14h, apresentação da análise
de processos, mesma equipe, in loco na Paupina."

Pode mandar sua mensagem 👇

_Para cancelar: /cancelar_`;

const EXEMPLO_OPR = `📋 Me conta como foi a semana, do seu jeito — pode escrever corrido.

${EXEMPLO_BASE}`;

const EXEMPLO_FOLLOWUP = `📋 Me conta como foi a semana, do seu jeito — pode escrever corrido.
_Essa aqui gera só a mensagem de follow-up semanal, sem o PPT._

${EXEMPLO_BASE}`;

// ─── /opr e /followup — pontos de entrada ────────────────────────────────────

export async function relatorioCommand(ctx: Context) {
  const chatId = ctx.from!.id;
  await setSession(chatId, { flow: 'relatorio', step: 'waiting_text', mode: 'opr' satisfies RelatorioMode });
  await ctx.reply(EXEMPLO_OPR, { parse_mode: 'Markdown' });
}

export async function followupCommand(ctx: Context) {
  const chatId = ctx.from!.id;
  await setSession(chatId, { flow: 'relatorio', step: 'waiting_text', mode: 'followup' satisfies RelatorioMode });
  await ctx.reply(EXEMPLO_FOLLOWUP, { parse_mode: 'Markdown' });
}

// ─── Helpers de data/hora ─────────────────────────────────────────────────

function dataHojeExtenso(): string {
  return new Intl.DateTimeFormat('pt-BR', {
    timeZone: 'America/Fortaleza',
    weekday: 'long',
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }).format(new Date());
}

function greeting(): string {
  const hour = Number(
    new Intl.DateTimeFormat('pt-BR', { timeZone: 'America/Fortaleza', hour: '2-digit', hour12: false }).format(new Date()),
  );
  if (hour < 12) return 'Bom dia';
  if (hour < 18) return 'Boa tarde';
  return 'Boa noite';
}

function capitalize(s: string): string {
  return s.length ? s[0].toUpperCase() + s.slice(1) : s;
}

// Calcula o dia da semana de forma determinística a partir de "DD/MM" ou
// "DD/MM/AAAA" — nunca confia no "dia" que a IA extraiu, só na data numérica.
function parseDataBR(dataStr: string, refDate: Date): Date | null {
  const m = dataStr.trim().match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (!m) return null;

  const day = Number(m[1]);
  const month = Number(m[2]) - 1;
  let year = m[3] ? Number(m[3]) : refDate.getFullYear();
  if (m[3] && m[3].length === 2) year += 2000;

  let d = new Date(Date.UTC(year, month, day, 12));
  if (!m[3]) {
    const diffDays = (refDate.getTime() - d.getTime()) / 86_400_000;
    if (diffDays > 30) d = new Date(Date.UTC(year + 1, month, day, 12));
  }
  return Number.isNaN(d.getTime()) ? null : d;
}

const WEEKDAYS_PT = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];

function weekdayPt(date: Date): string {
  return WEEKDAYS_PT[date.getUTCDay()];
}

// Recalcula e sobrescreve data.meet[].dia com o dia da semana real (calculado
// a partir da data numérica), e devolve as linhas de texto para conferência.
function recomputeWeekdays(data: OprData): string[] {
  const hoje = new Date();
  return data.meet.map((r) => {
    const parsed = parseDataBR(r.data, hoje);
    if (!parsed) {
      return `• ${r.data || '(sem data)'} — não deu pra calcular automaticamente; confira: "${r.dia}"`;
    }
    r.dia = capitalize(weekdayPt(parsed));
    return `• ${r.data} → *${r.dia}*`;
  });
}

// ─── Normalização dos dados vindos do Gemini ─────────────────────────────────

function normalizeOprData(raw: OprData): OprData {
  return {
    cliente: raw.cliente || '-',
    semana: raw.semana || '-',
    real: raw.real ?? [],
    prox: raw.prox ?? [],
    meet: raw.meet ?? [],
  };
}

// ─── Mensagens ────────────────────────────────────────────────────────────

function buildPreview(data: OprData): string {
  const real = data.real.map((it, i) => `${i + 1}. ${it.ativ} — Pharos: ${it.ph}`).join('\n') || '_nenhuma_';
  const prox = data.prox.map((it, i) => `${i + 1}. ${it.ativ} — Cliente: ${it.cli}`).join('\n') || '_nenhuma_';
  const meet = data.meet.map((r, i) => `${i + 1}. ${r.dia} ${r.data} às ${r.horario} — ${r.obj} — ${r.local}`).join('\n') || '_nenhuma_';

  return `Entendi assim 👇\n\n` +
    `✅ *Realizadas (${data.real.length})*\n${real}\n\n` +
    `➡️ *Próximas (${data.prox.length})*\n${prox}\n\n` +
    `📅 *Reuniões (${data.meet.length})*\n${meet}`;
}

function buildFollowUpMessage(data: OprData): string {
  const real = data.real.map((it) => `- ${it.ativ} (${it.obj});`).join('\n');
  const prox = data.prox.map((it) => `- ${it.ativ};`).join('\n');
  const meet = data.meet.map((r) => `- ${r.dia} (${r.data}) às ${r.horario} - ${r.obj} - ${r.local}`).join('\n');

  return `${greeting()}, amigos! Todos bem?\n\n` +
    `Nessa semana avançamos com os seguintes pontos:\n${real}\n\n` +
    `Próximos passos:\n${prox}\n\n` +
    `Reuniões Agendadas:\n${meet}`;
}

// ─── Etapa 1 (só no modo 'opr'): nome do cliente e período da semana ─────────

async function askCliente(ctx: Context, chatId: number, data: OprData, mode: RelatorioMode) {
  await setSession(chatId, { flow: 'relatorio', step: 'entering_cliente', mode, oprData: JSON.stringify(data) });
  await ctx.reply('Qual é o nome do cliente?');
}

async function askSemana(ctx: Context, chatId: number, data: OprData, mode: RelatorioMode) {
  await setSession(chatId, { flow: 'relatorio', step: 'entering_semana', mode, oprData: JSON.stringify(data) });
  await ctx.reply(
    'Qual é o período da semana? (ex.: 10/08 a 14/08)\n\n' +
    '_Esse texto também vira o nome do arquivo do PPT._',
    { parse_mode: 'Markdown' },
  );
}

// ─── Etapa 2: confirmação de dias das reuniões (+ cliente/semana no modo 'opr') ─

async function sendDetailsConfirmation(ctx: Context, chatId: number, data: OprData, mode: RelatorioMode) {
  const diasLines = recomputeWeekdays(data);
  const diasText = data.meet.length > 0 ? diasLines.join('\n') : '_nenhuma reunião com data para conferir_';

  const clienteSemanaBlock = mode === 'opr'
    ? `👤 *Cliente:* ${data.cliente}\n\n🗓️ *Semana* (vai no nome do arquivo): ${data.semana}\n\n`
    : '';

  const editButtonsRow = mode === 'opr'
    ? [
        { text: '✏️ Corrigir cliente', callback_data: 'relatorio_details_edit_cliente' },
        { text: '✏️ Corrigir semana', callback_data: 'relatorio_details_edit_semana' },
      ]
    : null;

  await setSession(chatId, { flow: 'relatorio', step: 'confirming_details', mode, oprData: JSON.stringify(data) });
  await ctx.reply(
    `Só confirmando antes de montar o relatório 👇\n\n` +
    clienteSemanaBlock +
    `📅 *Dias das reuniões* (calculados a partir da data):\n${diasText}\n\n` +
    `Tudo certo?`,
    {
      parse_mode: 'Markdown',
      reply_markup: {
        inline_keyboard: [
          [{ text: '✅ Confirmar e continuar', callback_data: 'relatorio_details_ok' }],
          ...(editButtonsRow ? [editButtonsRow] : []),
          [{ text: '✏️ Corrigir datas', callback_data: 'relatorio_details_edit_dias' }],
        ],
      },
    },
  );
}

// ─── Etapa 3: preview final (realizadas / próximas / reuniões) ──────────────

async function sendPreview(ctx: Context, chatId: number, data: OprData, mode: RelatorioMode) {
  await setSession(chatId, { flow: 'relatorio', step: 'confirming', mode, oprData: JSON.stringify(data) });
  await ctx.reply(buildPreview(data), {
    parse_mode: 'Markdown',
    reply_markup: {
      inline_keyboard: [[
        { text: '✅ Confirmar', callback_data: 'relatorio_confirm' },
        { text: '✏️ Corrigir', callback_data: 'relatorio_correct' },
      ]],
    },
  });
}

// Depois da extração (inicial ou por correção): no modo 'opr' pergunta cliente
// e semana se a IA não identificou; no modo 'followup' pula direto pra
// confirmação de dias, já que cliente/semana não entram na mensagem final.
async function proceedAfterExtraction(ctx: Context, chatId: number, data: OprData, mode: RelatorioMode) {
  if (mode === 'opr') {
    if (!data.cliente || data.cliente === '-') {
      await askCliente(ctx, chatId, data, mode);
      return;
    }
    if (!data.semana || data.semana === '-') {
      await askSemana(ctx, chatId, data, mode);
      return;
    }
  }
  await sendDetailsConfirmation(ctx, chatId, data, mode);
}

// ─── Handler de mensagens enquanto flow = 'relatorio' ────────────────────────

export async function handleRelatorioMessage(ctx: Context, session: Record<string, unknown>) {
  const chatId = ctx.from!.id;
  const step = session.step as string;
  const mode: RelatorioMode = session.mode === 'followup' ? 'followup' : 'opr';
  const msg = ctx.message;
  const text = msg && 'text' in msg ? msg.text?.trim() : '';
  if (!text || text.startsWith('/')) return;

  if (step === 'waiting_text') {
    if (text.length < 20) {
      await ctx.reply('Texto muito curto. Conta um pouco mais sobre a semana.');
      return;
    }

    const processingMsg = await ctx.reply('⏳ Analisando...');

    try {
      const system = buildRelatorioSystem(dataHojeExtenso());
      const raw = await callGemini(system, buildRelatorioUserMessage(text));
      const data = normalizeOprData(JSON.parse(extractJSON(raw)));

      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      await proceedAfterExtraction(ctx, chatId, data, mode);
    } catch (err) {
      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      console.error('[relatorio:error]', err);
      await ctx.reply(`❌ Erro ao interpretar o relatório. Tente novamente com /${mode === 'opr' ? 'opr' : 'followup'}.`);
      await clearSession(chatId);
    }
    return;
  }

  if (step === 'entering_cliente') {
    const data: OprData = JSON.parse(session.oprData as string);
    data.cliente = text;
    await proceedAfterExtraction(ctx, chatId, data, mode);
    return;
  }

  if (step === 'entering_semana') {
    const data: OprData = JSON.parse(session.oprData as string);
    data.semana = text;
    await sendDetailsConfirmation(ctx, chatId, data, mode);
    return;
  }

  if (step === 'correcting_dias') {
    const data: OprData = JSON.parse(session.oprData as string);
    const processingMsg = await ctx.reply('⏳ Ajustando as datas...');

    try {
      const system = buildRelatorioSystem(dataHojeExtenso());
      const userMsg = buildRelatorioCorrectionMessage({ jsonAtual: JSON.stringify(data), correcao: text });
      const raw = await callGemini(system, userMsg);
      const corrected = normalizeOprData(JSON.parse(extractJSON(raw)));

      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      await sendDetailsConfirmation(ctx, chatId, corrected, mode);
    } catch (err) {
      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      console.error('[relatorio:dias:error]', err);
      await ctx.reply('❌ Erro ao aplicar a correção. Tente descrever de outro jeito.');
    }
    return;
  }

  if (step === 'waiting_correction') {
    const jsonAtual = session.oprData as string;
    const processingMsg = await ctx.reply('⏳ Aplicando correção...');

    try {
      const system = buildRelatorioSystem(dataHojeExtenso());
      const userMsg = buildRelatorioCorrectionMessage({ jsonAtual, correcao: text });
      const raw = await callGemini(system, userMsg);
      const data = normalizeOprData(JSON.parse(extractJSON(raw)));

      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      await sendPreview(ctx, chatId, data, mode);
    } catch (err) {
      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      console.error('[relatorio:correction:error]', err);
      await ctx.reply('❌ Erro ao aplicar a correção. Tente descrever de outro jeito.');
    }
    return;
  }
}

// ─── Callbacks dos botões inline ─────────────────────────────────────────────

export function registerRelatorioActions(bot: Telegraf) {
  bot.action('relatorio_details_ok', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'relatorio') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery();
    await ctx.editMessageReplyMarkup(undefined);

    const mode: RelatorioMode = session.mode === 'followup' ? 'followup' : 'opr';
    const data: OprData = JSON.parse(session.oprData as string);
    await sendPreview(ctx, chatId, data, mode);
  });

  bot.action('relatorio_details_edit_cliente', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'relatorio') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery();
    await ctx.editMessageReplyMarkup(undefined);
    await setSession(chatId, { ...session, step: 'entering_cliente' });
    await ctx.reply('Qual é o nome correto do cliente?');
  });

  bot.action('relatorio_details_edit_semana', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'relatorio') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery();
    await ctx.editMessageReplyMarkup(undefined);
    await setSession(chatId, { ...session, step: 'entering_semana' });
    await ctx.reply(
      'Qual é o período correto da semana? (ex.: 10/08 a 14/08)\n\n' +
      '_Esse texto também vira o nome do arquivo do PPT._',
      { parse_mode: 'Markdown' },
    );
  });

  bot.action('relatorio_details_edit_dias', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'relatorio') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery();
    await ctx.editMessageReplyMarkup(undefined);
    await setSession(chatId, { ...session, step: 'correcting_dias' });
    await ctx.reply('Qual reunião está com a data errada? Me diga qual e a data certa.');
  });

  bot.action('relatorio_correct', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'relatorio') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery();
    await ctx.editMessageReplyMarkup(undefined);
    await setSession(chatId, { ...session, step: 'waiting_correction' });
    await ctx.reply('O que você quer corrigir? Descreva em texto livre.');
  });

  bot.action('relatorio_confirm', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'relatorio') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery('⏳ Gerando...');
    await ctx.editMessageReplyMarkup(undefined);

    const mode: RelatorioMode = session.mode === 'followup' ? 'followup' : 'opr';
    const data: OprData = JSON.parse(session.oprData as string);
    await clearSession(chatId);

    if (mode === 'followup') {
      try {
        await ctx.reply(buildFollowUpMessage(data));
      } catch (err) {
        console.error('[relatorio:confirm:followup:error]', err);
        await ctx.reply('❌ Erro ao gerar a mensagem. Tente novamente com /followup.');
      }
      return;
    }

    const processingMsg = await ctx.reply('⏳ Gerando follow-up e PPT...');

    try {
      const pptxBuffer = await generateOprPptx(data);
      const fileName = buildOprFileName(data.cliente, data.semana);

      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);

      await ctx.reply(buildFollowUpMessage(data));
      await ctx.replyWithDocument(
        { source: pptxBuffer, filename: fileName },
        { caption: `📊 ${fileName}` },
      );
    } catch (err) {
      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      console.error('[relatorio:confirm:error]', err);
      await ctx.reply('❌ Erro ao gerar os arquivos. Tente novamente com /opr.');
    }
  });
}
