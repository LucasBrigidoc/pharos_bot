import { Context, MiddlewareFn } from 'telegraf';

const ALLOWED_ID = Number(process.env.ALLOWED_CHAT_ID);

if (!ALLOWED_ID) throw new Error('ALLOWED_CHAT_ID not set or invalid');

export const whitelist: MiddlewareFn<Context> = (ctx, next) => {
  const id = ctx.from?.id ?? ctx.chat?.id;
  if (id !== ALLOWED_ID) {
    console.warn(`[whitelist] blocked ${id}`);
    return;
  }
  return next();
};
