import { Context, Telegraf } from 'telegraf';
import { findUser, updateUserField } from '../../db/users';
import { getSession, setSession, clearSession } from '../../db/sessions';

// ─── Exibe o perfil do usuário com botões de edição ──────────────────────────

export async function perfilCommand(ctx: Context) {
  const chatId = ctx.from!.id;
  await clearSession(chatId);

  const user = await findUser(chatId);
  if (!user) {
    await ctx.reply('Cadastro não encontrado. Use /start para se cadastrar.');
    return;
  }

  const senhaOculta = user.banco_senha ? '••••••••' : '_(não informado)_';

  await ctx.reply(
    `👤 *Seu perfil*\n\n` +
    `*Nome:* ${user.nome}\n` +
    `*E-mail:* ${user.email ?? '_(não informado)_'}\n` +
    `*Login banco de horas:* ${user.banco_login ?? '_(não informado)_'}\n` +
    `*Senha banco de horas:* ${senhaOculta}\n\n` +
    `O que deseja editar?`,
    {
      parse_mode: 'Markdown',
      reply_markup: {
        inline_keyboard: [
          [
            { text: '✏️ Nome',          callback_data: 'perfil_edit:nome' },
            { text: '✏️ E-mail',         callback_data: 'perfil_edit:email' },
          ],
          [
            { text: '✏️ Login banco',    callback_data: 'perfil_edit:banco_login' },
            { text: '✏️ Senha banco',    callback_data: 'perfil_edit:banco_senha' },
          ],
        ],
      },
    },
  );
}

// ─── Labels legíveis por campo ────────────────────────────────────────────────

const FIELD_LABEL: Record<string, string> = {
  nome:        'nome completo',
  email:       'e-mail corporativo',
  banco_login: 'login do banco de horas',
  banco_senha: 'senha do banco de horas',
};

// ─── Handler de mensagens enquanto flow = 'perfil' ───────────────────────────

export async function handlePerfilMessage(ctx: Context, session: Record<string, unknown>) {
  const chatId = ctx.from!.id;
  const msg = ctx.message;
  const text = msg && 'text' in msg ? msg.text?.trim() : '';
  if (!text || text.startsWith('/')) return;

  const field = session.editing_field as string;

  if (field === 'email' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text)) {
    await ctx.reply('E-mail inválido. Informe um e-mail corporativo válido:');
    return;
  }

  const dbField = field === 'banco_senha'
    ? 'banco_senha_enc'
    : field as 'nome' | 'email' | 'banco_login' | 'banco_senha_enc';

  await updateUserField(chatId, dbField, text);
  await clearSession(chatId);

  const display = field === 'banco_senha' ? '••••••••' : text;
  await ctx.reply(
    `✅ *${FIELD_LABEL[field] ?? field}* atualizado com sucesso!\n\n_Valor salvo: ${display}_\n\nUse /perfil para ver seu cadastro completo.`,
    { parse_mode: 'Markdown' },
  );
}

// ─── Callbacks dos botões inline ─────────────────────────────────────────────

export function registerPerfilActions(bot: Telegraf) {
  bot.action(/^perfil_edit:(.+)$/, async (ctx) => {
    const chatId = ctx.from!.id;
    const field  = ctx.match[1];

    await ctx.answerCbQuery();
    await ctx.editMessageReplyMarkup(undefined);
    await setSession(chatId, { flow: 'perfil', step: 'editing', editing_field: field });

    const prompts: Record<string, string> = {
      nome:        'Qual o seu novo nome completo?',
      email:       'Qual o seu novo e-mail corporativo?',
      banco_login: 'Qual o novo login do sistema de banco de horas?',
      banco_senha: '🔒 Qual a nova senha do sistema de banco de horas?\n\n_Será salva de forma criptografada._',
    };

    await ctx.reply(prompts[field] ?? 'Informe o novo valor:', { parse_mode: 'Markdown' });
  });
}
