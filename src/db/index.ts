import { Pool } from 'pg';

const DATABASE_URL = process.env.DATABASE_URL;
if (!DATABASE_URL) throw new Error('DATABASE_URL not set');

export const pool = new Pool({
  connectionString: DATABASE_URL,
  // SSL sempre ativo — banco é remoto (Supabase) em todos os ambientes
  ssl: DATABASE_URL.includes('localhost') ? false : { rejectUnauthorized: false },
});
