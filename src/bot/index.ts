import { Telegraf } from 'telegraf';
import { registrationFlow } from './registration';
import { registerCommands } from './commands';

const BOT_TOKEN = process.env.BOT_TOKEN;
if (!BOT_TOKEN) throw new Error('BOT_TOKEN not set');

export const bot = new Telegraf(BOT_TOKEN);

// Toda mensagem passa pelo fluxo de registro antes de qualquer comando.
// Usuários não cadastrados ficam nesse fluxo até concluir o cadastro.
// Usuários cadastrados passam direto para os comandos.
bot.use(registrationFlow);

registerCommands(bot);

bot.catch((err, ctx) => {
  console.error(`[bot:error] update ${ctx.updateType}`, err);
  ctx.reply('Ocorreu um erro interno. Tente novamente.').catch(() => undefined);
});
