import { pool } from './index';

export async function getSession(chatId: number): Promise<Record<string, unknown>> {
  const { rows } = await pool.query(
    'SELECT state FROM sessions WHERE chat_id = $1',
    [chatId]
  );
  return (rows[0]?.state as Record<string, unknown>) ?? {};
}

export async function setSession(
  chatId: number,
  state: Record<string, unknown>
): Promise<void> {
  await pool.query(
    `INSERT INTO sessions (chat_id, state, updated_at)
     VALUES ($1, $2, NOW())
     ON CONFLICT (chat_id) DO UPDATE SET state = $2, updated_at = NOW()`,
    [chatId, JSON.stringify(state)]
  );
}

export async function clearSession(chatId: number): Promise<void> {
  await pool.query('DELETE FROM sessions WHERE chat_id = $1', [chatId]);
}
