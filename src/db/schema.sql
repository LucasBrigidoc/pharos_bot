-- Tabela de usuários cadastrados
CREATE TABLE IF NOT EXISTS users (
  chat_id         BIGINT       PRIMARY KEY,
  nome            VARCHAR(100) NOT NULL,
  email           VARCHAR(150),                  -- email corporativo (usado para conectar Outlook)
  banco_login     VARCHAR(100),
  banco_senha_enc TEXT,                          -- senha criptografada (AES-256-GCM)
  ativo           BOOLEAN      NOT NULL DEFAULT true,
  criado_em       TIMESTAMPTZ  NOT NULL DEFAULT NOW()
);

-- Estado de conversas em andamento (registro, fluxo de ata, etc.)
CREATE TABLE IF NOT EXISTS sessions (
  chat_id    BIGINT      PRIMARY KEY,
  state      JSONB       NOT NULL DEFAULT '{}',
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Tokens OAuth por usuário (Outlook / Microsoft Graph — Fase 2)
CREATE TABLE IF NOT EXISTS oauth_tokens (
  chat_id       BIGINT      NOT NULL,
  service       VARCHAR(50) NOT NULL,
  access_token  TEXT,
  refresh_token TEXT,
  expires_at    TIMESTAMPTZ,
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  PRIMARY KEY (chat_id, service)
);
