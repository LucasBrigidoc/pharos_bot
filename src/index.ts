import 'dotenv/config';
import express from 'express';
import { bot } from './bot';
import { pool } from './db';
import { handleOutlookCallback } from './modules/outlook/graph-client';

const PORT = parseInt(process.env.PORT ?? '3000', 10);
const BOT_TOKEN = process.env.BOT_TOKEN!;
const WEBHOOK_DOMAIN = process.env.WEBHOOK_DOMAIN;

async function main() {
  await pool.query('SELECT 1');
  console.log('[db] connected');

  const app = express();
  app.use(express.json());

  // Mantido sem path no app.use: o token do bot tem ":" (formato id:hash),
  // e o Express interpreta ":" como início de parâmetro de rota quando o
  // path é passado pro roteador — isso corrompe o casamento e a rota nunca
  // bate (404). O telegraf já faz a checagem exata da URL internamente.
  const webhookPath = `/webhook/${BOT_TOKEN}`;
  app.use(bot.webhookCallback(webhookPath));

  app.get('/health', (_req, res) => {
    res.json({ ok: true, uptime: Math.floor(process.uptime()) });
  });

  app.get('/auth/outlook/callback', async (req, res) => {
    const { code, state, error } = req.query as { code?: string; state?: string; error?: string };
    const { status, html } = await handleOutlookCallback({ code, state, error }, bot);
    res.status(status).send(html);
  });

  app.listen(PORT, async () => {
    console.log(`[server] port ${PORT}`);
    if (WEBHOOK_DOMAIN) {
      const url = `${WEBHOOK_DOMAIN}${webhookPath}`;
      await bot.telegram.setWebhook(url);
      console.log(`[webhook] ${url}`);
    } else {
      await bot.launch();
      console.log('[bot] polling (dev)');
    }
  });

  process.once('SIGINT', () => bot.stop('SIGINT'));
  process.once('SIGTERM', () => bot.stop('SIGTERM'));
}

main().catch(err => {
  console.error('[fatal]', err);
  process.exit(1);
});
