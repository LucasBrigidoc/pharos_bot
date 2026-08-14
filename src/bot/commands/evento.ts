import { Context, Telegraf } from 'telegraf';
import { getSession, setSession, clearSession } from '../../db/sessions';
import { getOAuthToken } from '../../db/oauth';
import { callGemini, extractJSON } from '../../modules/gemini/client';
import {
  buildEventoSystem,
  buildEventoUserMessage,
  buildEventoCorrectionMessage,
} from '../../modules/gemini/prompts/evento.prompt';
import { dataHojeExtenso } from '../../modules/date/br-date';
import {
  buildAuthorizationUrl,
  withAutoRefresh,
  createEvent,
  listUpcomingEvents,
  OutlookNotConnectedError,
  GraphApiError,
  EventoInput,
} from '../../modules/outlook/graph-client';

interface EventoData {
  titulo: string;
  data: string;
  hora_inicio: string;
  hora_fim: string;
  local: string;
  participantes: string[];
  descricao: string;
}

const A_CONFIRMAR = '(a confirmar)';
const isPending = (v: string | undefined) => !v || v === A_CONFIRMAR;

// ─── /evento — ponto de entrada ───────────────────────────────────────────────

export async function eventoCommand(ctx: Context) {
  const chatId = ctx.from!.id;
  await clearSession(chatId);

  const connected = await getOAuthToken(chatId, 'outlook');
  if (!connected) {
    let url: string;
    try {
      url = buildAuthorizationUrl(chatId);
    } catch (err) {
      console.error('[evento:config:error]', err);
      await ctx.reply('⚙️ O módulo de Outlook ainda não foi configurado (credenciais do Azure AD pendentes). Fale com o administrador.');
      return;
    }
    await ctx.reply(
      '📅 *Módulo de Eventos (Outlook)*\n\n' +
      'Você ainda não conectou sua conta do Outlook da Pharos.',
      {
        parse_mode: 'Markdown',
        reply_markup: { inline_keyboard: [[{ text: '🔗 Conectar Outlook', url }]] },
      },
    );
    return;
  }

  await setSession(chatId, { flow: 'evento', step: 'menu' });
  await ctx.reply('📅 O que você quer fazer?', {
    reply_markup: {
      inline_keyboard: [[
        { text: '📅 Criar evento', callback_data: 'evento_criar' },
        { text: '📋 Listar eventos', callback_data: 'evento_listar' },
      ]],
    },
  });
}

// ─── Helpers ──────────────────────────────────────────────────────────────

function normalizeEventoData(raw: Partial<EventoData>): EventoData {
  return {
    titulo: raw.titulo ?? A_CONFIRMAR,
    data: raw.data ?? A_CONFIRMAR,
    hora_inicio: raw.hora_inicio ?? A_CONFIRMAR,
    hora_fim: raw.hora_fim ?? A_CONFIRMAR,
    local: raw.local || '-',
    participantes: raw.participantes ?? [],
    descricao: raw.descricao || '-',
  };
}

function buildPreview(data: EventoData): string {
  const participantes = data.participantes.length > 0 ? data.participantes.join(', ') : '_nenhum citado_';
  return `Entendi assim 👇\n\n` +
    `📌 *Título:* ${data.titulo}\n` +
    `📆 *Data:* ${data.data}\n` +
    `🕐 *Horário:* ${data.hora_inicio} às ${data.hora_fim}\n` +
    `📍 *Local:* ${data.local}\n` +
    `👥 *Participantes citados:* ${participantes}\n` +
    `📝 *Descrição:* ${data.descricao}`;
}

function formatGraphDateTime(iso: string): string {
  const [datePart, timePart] = iso.split('T');
  const [year, month, day] = datePart.split('-');
  const time = (timePart ?? '').slice(0, 5);
  return `${day}/${month} ${time}`;
}

async function handleOutlookError(ctx: Context, chatId: number, err: unknown) {
  if (err instanceof OutlookNotConnectedError) {
    const url = buildAuthorizationUrl(chatId);
    await ctx.reply(
      '🔌 Sua conexão com o Outlook expirou ou foi desfeita.',
      { reply_markup: { inline_keyboard: [[{ text: '🔗 Reconectar Outlook', url }]] } },
    );
    return;
  }
  if (err instanceof GraphApiError) {
    if (err.status === 403) {
      await ctx.reply('🚫 Sem permissão para acessar o calendário. Peça pro administrador da Pharos conceder o consentimento no Azure AD (API permissions → Grant admin consent).');
      return;
    }
    if (err.status === 429) {
      await ctx.reply('⏳ Muitas requisições ao Outlook agora. Tente de novo em instantes.');
      return;
    }
  }
  console.error('[evento:error]', err);
  await ctx.reply('❌ Erro ao falar com o Outlook. Tente novamente.');
}

async function askField(ctx: Context, chatId: number, data: EventoData, field: keyof EventoData, question: string) {
  await setSession(chatId, { flow: 'evento', step: 'entering_field', field, eventoData: JSON.stringify(data) });
  await ctx.reply(question);
}

async function sendPreview(ctx: Context, chatId: number, data: EventoData) {
  await setSession(chatId, { flow: 'evento', step: 'confirming', eventoData: JSON.stringify(data) });
  await ctx.reply(buildPreview(data), {
    parse_mode: 'Markdown',
    reply_markup: {
      inline_keyboard: [[
        { text: '✅ Confirmar', callback_data: 'evento_confirm' },
        { text: '✏️ Corrigir', callback_data: 'evento_correct' },
      ]],
    },
  });
}

// Depois da extração (inicial ou por correção): pergunta os campos
// obrigatórios que faltarem, um de cada vez, antes de mostrar o preview.
async function proceedAfterExtraction(ctx: Context, chatId: number, data: EventoData) {
  if (isPending(data.titulo)) return askField(ctx, chatId, data, 'titulo', 'Qual é o título do evento?');
  if (isPending(data.data)) return askField(ctx, chatId, data, 'data', 'Qual é a data do evento? (ex.: 14/08)');
  if (isPending(data.hora_inicio)) return askField(ctx, chatId, data, 'hora_inicio', 'Qual o horário de início? (ex.: 14:00)');
  if (isPending(data.hora_fim)) return askField(ctx, chatId, data, 'hora_fim', 'Qual o horário de término? (ex.: 15:00)');
  await sendPreview(ctx, chatId, data);
}

// ─── Handler de mensagens enquanto flow = 'evento' ────────────────────────────

export async function handleEventoMessage(ctx: Context, session: Record<string, unknown>) {
  const chatId = ctx.from!.id;
  const step = session.step as string;
  const msg = ctx.message;
  const text = msg && 'text' in msg ? msg.text?.trim() : '';
  if (!text || text.startsWith('/')) return;

  if (step === 'waiting_texto') {
    if (text.length < 10) {
      await ctx.reply('Descreve um pouco mais o evento (data, horário, com quem).');
      return;
    }

    const processingMsg = await ctx.reply('⏳ Analisando...');
    try {
      const raw = await callGemini(buildEventoSystem(dataHojeExtenso()), buildEventoUserMessage(text));
      const data = normalizeEventoData(JSON.parse(extractJSON(raw)));

      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      await proceedAfterExtraction(ctx, chatId, data);
    } catch (err) {
      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      console.error('[evento:extract:error]', err);
      await ctx.reply('❌ Erro ao interpretar o evento. Tente novamente com /evento.');
      await clearSession(chatId);
    }
    return;
  }

  if (step === 'entering_field') {
    const field = session.field as keyof EventoData;
    const data: EventoData = JSON.parse(session.eventoData as string);
    (data[field] as string) = text;
    await proceedAfterExtraction(ctx, chatId, data);
    return;
  }

  if (step === 'waiting_correction') {
    const data: EventoData = JSON.parse(session.eventoData as string);
    const processingMsg = await ctx.reply('⏳ Ajustando...');

    try {
      const userMsg = buildEventoCorrectionMessage({ jsonAtual: JSON.stringify(data), correcao: text });
      const raw = await callGemini(buildEventoSystem(dataHojeExtenso()), userMsg);
      const corrected = normalizeEventoData(JSON.parse(extractJSON(raw)));

      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      await proceedAfterExtraction(ctx, chatId, corrected);
    } catch (err) {
      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      console.error('[evento:correction:error]', err);
      await ctx.reply('❌ Erro ao aplicar a correção. Tente descrever de outro jeito.');
    }
    return;
  }
}

// ─── Callbacks dos botões inline ─────────────────────────────────────────────

export function registerEventoActions(bot: Telegraf) {
  bot.action('evento_criar', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'evento') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery();
    await ctx.editMessageReplyMarkup(undefined);
    await setSession(chatId, { flow: 'evento', step: 'waiting_texto' });
    await ctx.reply(
      '📝 Descreve o evento, do seu jeito — pode escrever corrido.\n\n' +
      '_Exemplo: "reunião com o Rogério amanhã às 14h até 15h pra falar do fechamento do mês, na sede da Truckão"_',
      { parse_mode: 'Markdown' },
    );
  });

  bot.action('evento_listar', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'evento') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery('⏳ Buscando...');
    await ctx.editMessageReplyMarkup(undefined);
    await clearSession(chatId);

    try {
      const events = await withAutoRefresh(chatId, listUpcomingEvents);
      if (events.length === 0) {
        await ctx.reply('📋 Nenhum evento nos próximos 7 dias.');
        return;
      }
      const lista = events.map((ev, i) =>
        `${i + 1}. *${ev.subject}*\n   ${formatGraphDateTime(ev.start)} – ${formatGraphDateTime(ev.end).split(' ')[1]}  ·  ${ev.location}`,
      ).join('\n\n');
      await ctx.reply(`📋 *Próximos 7 dias:*\n\n${lista}`, { parse_mode: 'Markdown' });
    } catch (err) {
      await handleOutlookError(ctx, chatId, err);
    }
  });

  bot.action('evento_correct', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'evento') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery();
    await ctx.editMessageReplyMarkup(undefined);
    await setSession(chatId, { ...session, step: 'waiting_correction' });
    await ctx.reply('O que você quer corrigir? Descreva em texto livre.');
  });

  bot.action('evento_confirm', async (ctx) => {
    const chatId = ctx.from!.id;
    const session = await getSession(chatId);
    if ((session.flow as string) !== 'evento') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery('⏳ Criando...');
    await ctx.editMessageReplyMarkup(undefined);

    const data: EventoData = JSON.parse(session.eventoData as string);
    await clearSession(chatId);

    const input: EventoInput = {
      titulo: data.titulo,
      data: data.data,
      horaInicio: data.hora_inicio,
      horaFim: data.hora_fim,
      local: data.local,
      descricao: data.participantes.length > 0
        ? `${data.descricao !== '-' ? data.descricao + '\n\n' : ''}Participantes: ${data.participantes.join(', ')}`
        : data.descricao,
    };

    try {
      const { webLink } = await withAutoRefresh(chatId, (token) => createEvent(token, input));
      await ctx.reply('✅ Evento criado no Outlook!', {
        reply_markup: { inline_keyboard: [[{ text: '🔗 Abrir no Outlook', url: webLink }]] },
      });
    } catch (err) {
      await handleOutlookError(ctx, chatId, err);
    }
  });
}
