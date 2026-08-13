import { Context, MiddlewareFn } from 'telegraf';
import { findUser, upsertUser } from '../db/users';
import { getSession, setSession, clearSession } from '../db/sessions';

const REG_PASSWORD = process.env.REGISTRATION_PASSWORD ?? 'consultor';

function getText(ctx: Context): string | null {
  return ctx.message && 'text' in ctx.message ? ctx.message.text?.trim() ?? null : null;
}

function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

export const registrationFlow: MiddlewareFn<Context> = async (ctx, next) => {
  const chatId = ctx.from?.id;
  if (!chatId) return;

  const user = await findUser(chatId);

  // Usuário ativo → fluxo normal
  if (user?.ativo) return next();

  // Usuário desativado → bloqueia
  if (user && !user.ativo) {
    await ctx.reply('Seu acesso foi desativado. Entre em contato com o administrador.');
    return;
  }

  // Usuário não cadastrado → fluxo de registro
  const session = await getSession(chatId);
  const step = (session.reg_step as string) ?? null;
  const text = getText(ctx);

  if (!step) {
    await setSession(chatId, { reg_step: 'password' });
    await ctx.reply(
      'Olá! Este é o bot interno da Pharos Consultoria.\n\n' +
      '🔐 Para ter acesso, informe a senha de cadastro:'
    );
    return;
  }

  if (step === 'password') {
    if (text !== REG_PASSWORD) {
      await ctx.reply('❌ Senha incorreta. Tente novamente:');
      return;
    }
    await setSession(chatId, { reg_step: 'name' });
    await ctx.reply('✅ Senha correta!\n\nQual é o seu nome completo?');
    return;
  }

  if (step === 'name') {
    if (!text || text.length < 2) {
      await ctx.reply('Por favor, informe seu nome completo:');
      return;
    }
    await setSession(chatId, { reg_step: 'email', nome: text });
    await ctx.reply(
      `Olá, ${text}! 👋\n\n` +
      'Qual é o seu *e-mail corporativo*?\n' +
      '_(Será usado para conectar seu Outlook ao bot)_',
      { parse_mode: 'Markdown' }
    );
    return;
  }

  if (step === 'email') {
    if (!text || !isValidEmail(text)) {
      await ctx.reply('E-mail inválido. Por favor, informe um e-mail corporativo válido:');
      return;
    }
    await setSession(chatId, { ...session, reg_step: 'banco_login', email: text });
    await ctx.reply(
      'Agora informe seu *login* do sistema de banco de horas\n' +
      '_(o mesmo usuário que você usa para acessar o site da empresa)_:',
      { parse_mode: 'Markdown' }
    );
    return;
  }

  if (step === 'banco_login') {
    if (!text) {
      await ctx.reply('Por favor, informe seu login:');
      return;
    }
    await setSession(chatId, { ...session, reg_step: 'banco_senha', banco_login: text });
    await ctx.reply(
      'Por último, informe sua *senha* do sistema de banco de horas:\n\n' +
      '🔒 Ela será salva de forma criptografada e usada somente para lançar suas horas automaticamente.',
      { parse_mode: 'Markdown' }
    );
    return;
  }

  if (step === 'banco_senha') {
    if (!text) {
      await ctx.reply('Por favor, informe sua senha:');
      return;
    }
    const nome = session.nome as string;
    const email = session.email as string;
    const bancoLogin = session.banco_login as string;
    await upsertUser(chatId, nome, email, bancoLogin, text);
    await clearSession(chatId);
    await ctx.reply(
      `✅ *Cadastro concluído!* Bem-vindo(a), ${nome}.\n\n` +
      `📧 E-mail registrado: ${email}\n\n` +
      'Use /start para ver os comandos disponíveis.\n\n' +
      '_Quando o módulo de Outlook estiver disponível, você receberá um link para conectar seu calendário._',
      { parse_mode: 'Markdown' }
    );
    return;
  }
};
