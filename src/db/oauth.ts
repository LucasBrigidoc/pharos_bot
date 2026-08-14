import { pool } from './index';
import { encrypt, decrypt } from '../modules/crypto';

export interface OAuthToken {
  accessToken: string;
  refreshToken: string;
  expiresAt: Date | null;
}

export async function getOAuthToken(chatId: number, service: string): Promise<OAuthToken | null> {
  const { rows } = await pool.query(
    'SELECT access_token, refresh_token, expires_at FROM oauth_tokens WHERE chat_id = $1 AND service = $2',
    [chatId, service]
  );
  const r = rows[0];
  if (!r || !r.access_token || !r.refresh_token) return null;
  return {
    accessToken: decrypt(r.access_token),
    refreshToken: decrypt(r.refresh_token),
    expiresAt: r.expires_at ? new Date(r.expires_at) : null,
  };
}

export async function upsertOAuthToken(
  chatId: number,
  service: string,
  token: OAuthToken
): Promise<void> {
  await pool.query(
    `INSERT INTO oauth_tokens (chat_id, service, access_token, refresh_token, expires_at, updated_at)
     VALUES ($1, $2, $3, $4, $5, NOW())
     ON CONFLICT (chat_id, service) DO UPDATE
       SET access_token = $3, refresh_token = $4, expires_at = $5, updated_at = NOW()`,
    [chatId, service, encrypt(token.accessToken), encrypt(token.refreshToken), token.expiresAt]
  );
}

export async function deleteOAuthToken(chatId: number, service: string): Promise<void> {
  await pool.query('DELETE FROM oauth_tokens WHERE chat_id = $1 AND service = $2', [chatId, service]);
}
