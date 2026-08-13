import { Context, Telegraf } from 'telegraf';
import { getSession, setSession, clearSession } from '../../db/sessions';
import { callGemini, extractJSON } from '../../modules/gemini/client';
import {
  buildTurnoSystem,
  buildTurnoUserMessage,
  buildTurnoCorrectionMessage,
} from '../../modules/gemini/prompts/turno.prompt';

interface TurnoJSON {
  feito: string;
  pendente: string;
  proximo: string;
}

interface TurnoState {
  cliente: string;
  data: string;
  turno?: TurnoJSON;
}

// ─── /turno — ponto de entrada ───────────────────────────────────────────────

export async function turnoCommand(ctx: Context) {
  const chatId = ctx.from!.id;
  await setSession(chatId, { flow: 'turno', step: 'waiting_cliente_data' });
  await ctx.reply(
    '📋 *Follow-up de turno*\n\n' +
    'Qual o cliente e a data do turno?\n\n' +
    'Responda no formato: *Cliente - Data*\n\n' +
    `_Exemplo: Auto Center - ${dataHojeCurta()}_\n\n` +
    '_Para cancelar: /cancelar_',
    { parse_mode: 'Markdown' },
  );
}

// ─── Helpers ──────────────────────────────────────────────────────────────

function dataHojeCurta(): string {
  return new Intl.DateTimeFormat('pt-BR', {
    timeZone: 'America/Fortaleza',
    day: '2-digit',
    month: '2-digit',
  }).format(new Date());
}

function parseClienteData(text: string): { cliente: string; data: string } {
  const parts = text.split(/\s*[-–]\s*/);
  const cliente = (parts[0] ?? '').trim();
  const data = (parts[1] ?? '').trim() || dataHojeCurta();
  return { cliente, data };
}

function buildTurnoMessage(state: TurnoState): string {
  const turno = state.turno!;
  return `Follow-up de turno — ${state.cliente} (${state.data})\n\n` +
    `O que foi feito no turno\n${turno.feito}\n\n` +
    `O que ficou pendente\n${turno.pendente}\n\n` +
    `O que farei no próximo turno\n${turno.proximo}`;
}

async function sendPreview(ctx: Context, chatId: number, state: TurnoState) {
  await setSession(chatId, { flow: 'turno', step: 'confirming', state: JSON.stringify(state) });
  await ctx.reply(buildTurnoMessage(state), {
    reply_markup: {
      inline_keyboard: [[
        { text: '✅ Está bom', callback_data: 'turno_ok' },
        { text: '✏️ Corrigir', callback_data: 'turno_correct' },
      ]],
    },
  });
}

// ─── Handler de mensagens enquanto flow = 'turno' ────────────────────────────

export async function handleTurnoMessage(ctx: Context, session: Record<string, unknown>) {
  const chatId = ctx.from!.id;
  const step = session.step as string;
  const msg = ctx.message;
  const text = msg && 'text' in msg ? msg.text?.trim() : '';
  if (!text || text.startsWith('/')) return;

  if (step === 'waiting_cliente_data') {
    const { cliente, data } = parseClienteData(text);
    if (!cliente) {
      await ctx.reply('Não consegui identificar o cliente. Responda no formato: *Cliente - Data*', { parse_mode: 'Markdown' });
      return;
    }

    await setSession(chatId, { flow: 'turno', step: 'waiting_texto', state: JSON.stringify({ cliente, data }) });
    await ctx.reply(
      'Me conta o que você fez nesse turno, o que ficou pendente e o que pretende fazer no próximo, do seu jeito — pode escrever corrido.',
    );
    return;
  }

  if (step === 'waiting_texto') {
    if (text.length < 20) {
      await ctx.reply('Texto muito curto. Conta um pouco mais sobre o turno.');
      return;
    }

    const state: TurnoState = JSON.parse(session.state as string);
    const processingMsg = await ctx.reply('⏳ Analisando...');

    try {
      const raw = await callGemini(buildTurnoSystem(), buildTurnoUserMessage(text));
      state.turno = JSON.parse(extractJSON(raw));

      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      await sendPreview(ctx, chatId, state);
    } catch (err) {
      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      console.error('[turno:error]', err);
      await ctx.reply('❌ Erro ao interpretar o relato. Tente novamente com /turno.');
      await clearSession(chatId);
    }
    return;
  }

  if (step === 'correcting') {
    const state: TurnoState = JSON.parse(session.state as string);
    const processingMsg = await ctx.reply('⏳ Ajustando...');

    try {
      const userMsg = buildTurnoCorrectionMessage({ jsonAtual: JSON.stringify(state.turno), correcao: text });
      const raw = await callGemini(buildTurnoSystem(), userMsg);
      state.turno = JSON.parse(extractJSON(raw));

      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      await sendPreview(ctx, chatId, state);
    } catch (err) {
      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      console.error('[turno:correction:error]', err);
      await ctx.reply('❌ Erro ao aplicar a correção. Tente descrever de outro jeito.');
    }
    return;
  }
}

// ─── Callbacks dos botões inline ─────────────────────────────────────────────

export function registerTurnoActions(bot: Telegraf) {
  bot.action('turno_correct', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'turno') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery();
    await ctx.editMessageReplyMarkup(undefined);
    await setSession(chatId, { ...session, step: 'correcting' });
    await ctx.reply('O que você quer corrigir? Descreva em texto livre.');
  });

  bot.action('turno_ok', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'turno') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery('✅');
    await ctx.editMessageReplyMarkup(undefined);
    await clearSession(chatId);
    await ctx.reply('✅ Prontinho — é só copiar a mensagem acima.');
  });
}
