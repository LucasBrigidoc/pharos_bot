import { Context } from 'telegraf';
import { pool } from '../../db';

export async function pingCommand(ctx: Context) {
  const t0 = Date.now();
  await pool.query('SELECT 1');
  const dbMs = Date.now() - t0;
  await ctx.reply(`pong — db: ${dbMs}ms — uptime: ${Math.floor(process.uptime())}s`);
}
