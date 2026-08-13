import { Context, Telegraf } from 'telegraf';
import { startCommand } from './start';
import { ataCommand, handleAtaMessage, registerAtaActions } from './ata';
import { perfilCommand, handlePerfilMessage, registerPerfilActions } from './perfil';
import { relatorioCommand, followupCommand, handleRelatorioMessage, registerRelatorioActions } from './relatorio_semanal';
import { turnoCommand, handleTurnoMessage, registerTurnoActions } from './turno';
import { getSession, clearSession } from '../../db/sessions';

export function registerCommands(bot: Telegraf) {
  // ─── Comandos ──────────────────────────────────────────────────────────────
  bot.command('start',              startCommand);
  bot.command('ata',                ataCommand);
  bot.command('perfil',             perfilCommand);
  bot.command('opr',                relatorioCommand);
  bot.command('followup',           followupCommand);
  bot.command('turno',              turnoCommand);

  bot.command('cancelar', async (ctx: Context) => {
    await clearSession(ctx.from!.id);
    await ctx.reply('✅ Operação cancelada. Use /start para ver os comandos.');
  });

  // ─── Ações de botões inline ───────────────────────────────────────────────
  registerAtaActions(bot);
  registerPerfilActions(bot);
  registerRelatorioActions(bot);
  registerTurnoActions(bot);

  // ─── Handler de mensagens para fluxos ativos ──────────────────────────────
  bot.on('message', async (ctx: Context) => {
    const msg = ctx.message;
    if (!msg || ('text' in msg && msg.text?.startsWith('/'))) return;

    const session = await getSession(ctx.from!.id);
    const flow = session.flow as string | undefined;

    if (flow === 'ata') {
      await handleAtaMessage(ctx, session);
      return;
    }

    if (flow === 'perfil') {
      await handlePerfilMessage(ctx, session);
      return;
    }

    if (flow === 'relatorio') {
      await handleRelatorioMessage(ctx, session);
      return;
    }

    if (flow === 'turno') {
      await handleTurnoMessage(ctx, session);
      return;
    }

    await ctx.reply('Use /start para ver os comandos disponíveis.');
  });
}
