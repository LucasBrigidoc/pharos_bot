import 'dotenv/config';
import express from 'express';
import { bot } from './bot';
import { pool } from './db';

const PORT = parseInt(process.env.PORT ?? '3000', 10);
const BOT_TOKEN = process.env.BOT_TOKEN!;
const WEBHOOK_DOMAIN = process.env.WEBHOOK_DOMAIN;

async function main() {
  await pool.query('SELECT 1');
  console.log('[db] connected');

  const app = express();
  app.use(express.json());

  const webhookPath = `/webhook/${BOT_TOKEN}`;
  app.use(webhookPath, bot.webhookCallback(webhookPath));

  app.get('/health', (_req, res) => {
    res.json({ ok: true, uptime: Math.floor(process.uptime()) });
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
