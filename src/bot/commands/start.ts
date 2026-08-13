import { Context } from 'telegraf';

export async function startCommand(ctx: Context) {
  await ctx.reply(
    'Olá! Sou o bot interno da Pharos Consultoria.\n\n' +
    '*Comandos disponíveis:*\n' +
    '/ata — gera ata de reunião em .docx\n' +
    '/ping — verifica se estou online\n' +
    '/cancelar — cancela a operação em andamento\n\n' +
    '_Em breve:_\n' +
    '/evento — gerencia seu calendário Outlook\n' +
    '/relatorio\\_semanal — gera OPR semanal\n' +
    '/banco\\_horas — lança horas no sistema',
    { parse_mode: 'Markdown' },
  );
}
