import { Context } from 'telegraf';
import { clearSession } from '../../db/sessions';

export async function startCommand(ctx: Context) {
  await clearSession(ctx.from!.id);
  await ctx.reply(
    '👋 Olá! Sou o assistente interno da *Pharos Consultoria*.\n\n' +
    '*Comandos disponíveis:*\n' +
    '/ata — gera ata de reunião em .docx\n' +
    '/perfil — veja e edite seus dados de cadastro\n' +
    '/cancelar — cancela a operação em andamento\n\n' +
    '_Em breve:_\n' +
    '/evento — gerencia seu calendário Outlook\n' +
    '/relatorio\\_semanal — gera OPR semanal\n' +
    '/banco\\_horas — lança horas no sistema',
    { parse_mode: 'Markdown' },
  );
}
