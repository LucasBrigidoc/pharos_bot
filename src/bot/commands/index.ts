import { Context, Telegraf } from 'telegraf';
import { startCommand } from './start';
import { pingCommand } from './ping';
import { ataCommand, handleAtaMessage, registerAtaActions } from './ata';
import { getSession, clearSession } from '../../db/sessions';

export function registerCommands(bot: Telegraf) {
  // ─── Comandos ──────────────────────────────────────────────────────────────
  bot.command('start', startCommand);
  bot.command('ping', pingCommand);
  bot.command('ata', ataCommand);

  bot.command('cancelar', async (ctx: Context) => {
    await clearSession(ctx.from!.id);
    await ctx.reply('✅ Operação cancelada.');
  });

  // ─── Ações de botões inline ───────────────────────────────────────────────
  registerAtaActions(bot);

  // ─── Handler de mensagens para fluxos ativos ──────────────────────────────
  // Captura mensagens que não são comandos e as roteia para o fluxo correto.
  bot.on('message', async (ctx: Context) => {
    const msg = ctx.message;
    if (!msg || ('text' in msg && msg.text?.startsWith('/'))) return;

    const session = await getSession(ctx.from!.id);
    const flow = session.flow as string | undefined;

    if (flow === 'ata') {
      await handleAtaMessage(ctx, session);
      return;
    }

    // Nenhum fluxo ativo — orienta o usuário
    await ctx.reply('Use /start para ver os comandos disponíveis.');
  });
}
