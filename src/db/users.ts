import { pool } from './index';
import { encrypt, decrypt } from '../modules/crypto';

export interface User {
  chat_id: number;
  nome: string;
  email: string | null;
  banco_login: string | null;
  banco_senha: string | null;
  ativo: boolean;
}

export async function findUser(chatId: number): Promise<User | null> {
  const { rows } = await pool.query(
    'SELECT * FROM users WHERE chat_id = $1',
    [chatId]
  );
  if (!rows[0]) return null;
  const r = rows[0];
  return {
    chat_id: Number(r.chat_id),
    nome: r.nome,
    email: r.email ?? null,
    banco_login: r.banco_login ?? null,
    banco_senha: r.banco_senha_enc ? decrypt(r.banco_senha_enc) : null,
    ativo: r.ativo,
  };
}

export async function upsertUser(
  chatId: number,
  nome: string,
  email: string,
  bancoLogin: string,
  bancoPwd: string
): Promise<void> {
  await pool.query(
    `INSERT INTO users (chat_id, nome, email, banco_login, banco_senha_enc)
     VALUES ($1, $2, $3, $4, $5)
     ON CONFLICT (chat_id) DO UPDATE
       SET nome = $2, email = $3, banco_login = $4, banco_senha_enc = $5, ativo = true`,
    [chatId, nome, email, bancoLogin, encrypt(bancoPwd)]
  );
}
