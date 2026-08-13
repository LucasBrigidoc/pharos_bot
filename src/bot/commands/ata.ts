import { Context, Telegraf } from 'telegraf';
import { getSession, setSession, clearSession } from '../../db/sessions';
import { callGemini, extractJSON } from '../../modules/gemini/client';
import { ATA_SYSTEM, buildAtaUserMessage } from '../../modules/gemini/prompts/ata.prompt';
import { generateAtaDocx, buildAtaFileName, AtaJSON } from '../../modules/docx/ata-generator';

// ─── Opções dos botões inline ────────────────────────────────────────────────

const MODALIDADE_OPTS = [
  { text: '🏢 Presencial', callback_data: 'ata_mod:Presencial' },
  { text: '💻 Remoto (online)', callback_data: 'ata_mod:Remoto (online)' },
  { text: '🔀 Híbrido', callback_data: 'ata_mod:Híbrido' },
];

const DETALHE_OPTS = [
  { text: '📊 Com números e dados', callback_data: 'ata_det:com números e dados explícitos' },
  { text: '📝 Sem números', callback_data: 'ata_det:sem mencionar números ou percentuais' },
];

// ─── /ata — ponto de entrada ─────────────────────────────────────────────────

export async function ataCommand(ctx: Context) {
  const chatId = ctx.from!.id;
  await setSession(chatId, { flow: 'ata', step: 'waiting_transcription' });
  await ctx.reply(
    '📋 *Módulo de Ata de Reunião*\n\n' +
    'Mande a transcrição da reunião. Pode ser:\n' +
    '• Texto digitado diretamente\n' +
    '• Arquivo *.docx*\n' +
    '• Arquivo *.txt*\n\n' +
    '_Para cancelar: /cancelar_',
    { parse_mode: 'Markdown' },
  );
}

// ─── Extração de texto de mensagens ─────────────────────────────────────────

async function extractText(ctx: Context): Promise<string | null> {
  const msg = ctx.message;
  if (!msg) return null;

  // Texto direto
  if ('text' in msg && msg.text && !msg.text.startsWith('/')) {
    return msg.text;
  }

  // Arquivo
  if ('document' in msg && msg.document) {
    const { file_id, file_name = '', mime_type = '' } = msg.document;
    const isDocx = file_name.endsWith('.docx') || mime_type.includes('wordprocessingml');
    const isTxt  = file_name.endsWith('.txt')  || mime_type === 'text/plain';

    if (!isDocx && !isTxt) {
      await ctx.reply('Formato não suportado. Envie texto, *.docx* ou *.txt*.', { parse_mode: 'Markdown' });
      return null;
    }

    const link = await ctx.telegram.getFileLink(file_id);
    const res  = await fetch(link.href);
    const buf  = Buffer.from(await res.arrayBuffer());

    if (isDocx) {
      const mammoth = await import('mammoth');
      const { value } = await mammoth.extractRawText({ buffer: buf });
      return value;
    }
    return buf.toString('utf8');
  }

  return null;
}

// ─── Parse da resposta "Projeto - Nome - Data" ───────────────────────────────

function parseMeetingInfo(text: string): { projeto: string; assunto: string; data: string } {
  const parts = text.split(/\s*[-–]\s*/);
  const projeto = (parts[0] ?? '').trim();
  const assunto = (parts[1] ?? '').trim();
  const data    = (parts[2] ?? '').trim();
  return { projeto, assunto, data };
}

// ─── Handler de mensagens enquanto flow = 'ata' ──────────────────────────────

export async function handleAtaMessage(ctx: Context, session: Record<string, unknown>) {
  const chatId = ctx.from!.id;
  const step = session.step as string;

  if (step === 'waiting_transcription') {
    const text = await extractText(ctx);
    if (!text) return;

    if (text.trim().length < 50) {
      await ctx.reply('Transcrição muito curta. Envie o texto completo da reunião.');
      return;
    }

    await setSession(chatId, { ...session, step: 'waiting_meeting_info', transcricao: text });
    await ctx.reply(
      '📁 Qual o projeto, nome e data da reunião?\n\n' +
      'Responda no formato: *Projeto - Nome da Reunião - Data*\n\n' +
      '_Exemplo: Truckão - Alinhamento Interno - 28/07/2026_',
      { parse_mode: 'Markdown' },
    );
    return;
  }

  if (step === 'waiting_meeting_info') {
    const msg = ctx.message;
    const rawText = msg && 'text' in msg ? msg.text?.trim() : '';
    if (!rawText || rawText.startsWith('/')) return;

    const { projeto, assunto, data } = parseMeetingInfo(rawText);

    await setSession(chatId, { ...session, step: 'waiting_modalidade', projeto, assunto, data });
    await ctx.reply(
      'Qual foi a modalidade da reunião?',
      { reply_markup: { inline_keyboard: [MODALIDADE_OPTS] } },
    );
    return;
  }
}

// ─── Callbacks dos botões inline ─────────────────────────────────────────────

export function registerAtaActions(bot: Telegraf) {
  // Seleção de modalidade
  bot.action(/^ata_mod:(.+)$/, async (ctx) => {
    const chatId = ctx.from!.id;
    const modalidade = ctx.match[1];
    const session = await getSession(chatId);

    if ((session.flow as string) !== 'ata') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery(`✅ ${modalidade}`);
    await ctx.editMessageReplyMarkup(undefined);
    await setSession(chatId, { ...session, step: 'waiting_detalhe', modalidade });
    await ctx.reply(
      'Nível de detalhe para números e dados?',
      { reply_markup: { inline_keyboard: [DETALHE_OPTS] } },
    );
  });

  // Seleção de detalhe → dispara o processamento
  bot.action(/^ata_det:(.+)$/, async (ctx) => {
    const chatId = ctx.from!.id;
    const detalhe = ctx.match[1];
    const session = await getSession(chatId);

    if ((session.flow as string) !== 'ata') { await ctx.answerCbQuery(); return; }

    await ctx.answerCbQuery('⏳ Processando...');
    await ctx.editMessageReplyMarkup(undefined);

    const transcricao = session.transcricao as string;
    const modalidade  = session.modalidade  as string;
    const projeto     = (session.projeto as string | undefined) ?? '';
    const assunto     = (session.assunto as string | undefined) ?? '';
    const data        = (session.data    as string | undefined) ?? '';

    await clearSession(chatId);

    const processingMsg = await ctx.reply('⏳ Analisando transcrição com IA...\n\nIsso pode levar até 1 minuto para transcrições longas.');

    try {
      const userMsg = buildAtaUserMessage({ modalidade, detalhe, transcricao, projeto, assunto, data });
      const raw = await callGemini(ATA_SYSTEM, userMsg);
      const ata: AtaJSON = JSON.parse(extractJSON(raw));

      // Usar os campos fornecidos pelo usuário se a IA não preencheu
      if (!ata.projeto && projeto) ata.projeto = projeto;
      if (!ata.assunto && assunto) ata.assunto = assunto;
      if (!ata.data    && data)    ata.data    = data;

      const docBuffer = await generateAtaDocx(ata);
      const fileName  = buildAtaFileName(
        projeto || ata.projeto,
        data    || ata.data,
        assunto || ata.assunto,
      );

      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);

      await ctx.replyWithDocument(
        { source: docBuffer, filename: fileName },
        { caption: `✅ *Ata gerada!*\n📁 ${fileName}`, parse_mode: 'Markdown' },
      );

      if (ata.pontos_a_confirmar.length > 0) {
        const lista = ata.pontos_a_confirmar.map((p, i) => `${i + 1}. ${p}`).join('\n');
        await ctx.reply(
          `⚠️ *Pontos a confirmar (${ata.pontos_a_confirmar.length}):*\n\n${lista}`,
          { parse_mode: 'Markdown' },
        );
      }

    } catch (err) {
      await ctx.telegram.deleteMessage(chatId, processingMsg.message_id).catch(() => undefined);
      console.error('[ata:error]', err);

      if (err instanceof SyntaxError) {
        await ctx.reply('❌ A IA retornou um formato inesperado. Tente novamente com /ata.');
      } else {
        await ctx.reply('❌ Erro ao gerar a ata. Tente novamente com /ata.');
      }
    }
  });
}
