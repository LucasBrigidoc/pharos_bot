import { Telegraf } from 'telegraf';
import { encrypt, decrypt } from '../crypto';
import { getOAuthToken, upsertOAuthToken, deleteOAuthToken } from '../../db/oauth';
import { parseDataBR } from '../date/br-date';

const SERVICE = 'outlook';
const SCOPES = 'openid profile offline_access Calendars.ReadWrite User.Read';
const STATE_MAX_AGE_MS = 10 * 60_000;
const REFRESH_MARGIN_MS = 2 * 60_000;

// As credenciais do Azure AD são checadas sob demanda (não no import do
// módulo) — assim, enquanto o usuário não termina o cadastro do app no
// Azure, o resto do bot (ata, opr, turno) continua funcionando normalmente
// em vez de derrubar o processo inteiro por uma env var ausente.
function getAzureConfig() {
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  const tenantId = process.env.AZURE_TENANT_ID;
  const webhookDomain = process.env.WEBHOOK_DOMAIN;
  if (!clientId || !clientSecret || !tenantId) {
    throw new Error('AZURE_CLIENT_ID/AZURE_CLIENT_SECRET/AZURE_TENANT_ID não configurados');
  }
  if (!webhookDomain) {
    throw new Error('WEBHOOK_DOMAIN não configurado (necessário para o redirect URI do Outlook)');
  }
  return {
    clientId,
    clientSecret,
    tenantId,
    redirectUri: `${webhookDomain}/auth/outlook/callback`,
    authBase: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0`,
  };
}

export class OutlookNotConnectedError extends Error {
  constructor() {
    super('Outlook não conectado');
    this.name = 'OutlookNotConnectedError';
  }
}

export class GraphApiError extends Error {
  constructor(public status: number, message: string) {
    super(message);
    this.name = 'GraphApiError';
  }
}

// ─── Autorização ──────────────────────────────────────────────────────────

export function buildAuthorizationUrl(chatId: number): string {
  const cfg = getAzureConfig();
  // URLSearchParams já faz o percent-encoding ao montar a query string —
  // não codificar manualmente aqui, senão o valor fica com dupla-codificação.
  const state = encrypt(JSON.stringify({ chatId, ts: Date.now() }));
  const params = new URLSearchParams({
    client_id: cfg.clientId,
    response_type: 'code',
    redirect_uri: cfg.redirectUri,
    response_mode: 'query',
    scope: SCOPES,
    state,
  });
  return `${cfg.authBase}/authorize?${params.toString()}`;
}

function parseAndValidateState(state: string | undefined): number {
  if (!state) throw new Error('state ausente');
  // O Express já faz o URL-decode do query param antes de chegar aqui —
  // decodificar de novo corromperia o valor (dupla-decodificação).
  const decoded = decrypt(state);
  const { chatId, ts } = JSON.parse(decoded) as { chatId: number; ts: number };
  if (Date.now() - ts > STATE_MAX_AGE_MS) throw new Error('state expirado');
  return chatId;
}

interface TokenResponse {
  access_token: string;
  refresh_token: string;
  expires_in: number;
}

async function postToTokenEndpoint(params: URLSearchParams): Promise<TokenResponse> {
  const cfg = getAzureConfig();
  const res = await fetch(`${cfg.authBase}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });
  const json = await res.json() as any;
  if (!res.ok) {
    throw new Error(`token endpoint: ${json.error} — ${json.error_description ?? ''}`);
  }
  return json as TokenResponse;
}

async function exchangeCodeForTokens(code: string): Promise<TokenResponse> {
  const cfg = getAzureConfig();
  return postToTokenEndpoint(new URLSearchParams({
    client_id: cfg.clientId,
    client_secret: cfg.clientSecret,
    grant_type: 'authorization_code',
    code,
    redirect_uri: cfg.redirectUri,
    scope: SCOPES,
  }));
}

async function refreshAccessToken(refreshToken: string): Promise<TokenResponse> {
  const cfg = getAzureConfig();
  return postToTokenEndpoint(new URLSearchParams({
    client_id: cfg.clientId,
    client_secret: cfg.clientSecret,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope: SCOPES,
  }));
}

// ─── Token válido (com refresh automático) ───────────────────────────────────

export async function getValidAccessToken(chatId: number, forceRefresh = false): Promise<string> {
  const stored = await getOAuthToken(chatId, SERVICE);
  if (!stored) throw new OutlookNotConnectedError();

  const needsRefresh = forceRefresh || !stored.expiresAt || stored.expiresAt.getTime() <= Date.now() + REFRESH_MARGIN_MS;
  if (!needsRefresh) return stored.accessToken;

  try {
    const fresh = await refreshAccessToken(stored.refreshToken);
    await upsertOAuthToken(chatId, SERVICE, {
      accessToken: fresh.access_token,
      refreshToken: fresh.refresh_token ?? stored.refreshToken,
      expiresAt: new Date(Date.now() + fresh.expires_in * 1000),
    });
    return fresh.access_token;
  } catch (err) {
    // Refresh token revogado/expirado (senha trocada, 90 dias sem uso, etc.)
    // — apaga a linha morta e força reconexão limpa em vez de insistir.
    await deleteOAuthToken(chatId, SERVICE);
    throw new OutlookNotConnectedError();
  }
}

// Encapsula o padrão "pega token válido, chama a Graph API, se vier 401
// tenta renovar à força uma vez e repete a chamada" — usado por createEvent
// e listUpcomingEvents através do comando /evento.
export async function withAutoRefresh<T>(chatId: number, fn: (accessToken: string) => Promise<T>): Promise<T> {
  const token = await getValidAccessToken(chatId);
  try {
    return await fn(token);
  } catch (err) {
    if (err instanceof GraphApiError && err.status === 401) {
      const freshToken = await getValidAccessToken(chatId, true);
      return await fn(freshToken);
    }
    throw err;
  }
}

// ─── Callback HTTP (chamado pela rota Express) ───────────────────────────────

const HTML_ERRO = (msg: string) => `<html><body><p>❌ ${msg}</p></body></html>`;
const HTML_SUCESSO = `<html><body><p>✅ Outlook conectado! Pode fechar esta aba e voltar pro Telegram.</p></body></html>`;
const HTML_CANCELADO = `<html><body><p>Conexão cancelada. Volte ao Telegram e use /evento se quiser tentar de novo.</p></body></html>`;

export async function handleOutlookCallback(
  params: { code?: string; state?: string; error?: string },
  bot: Telegraf
): Promise<{ status: number; html: string }> {
  if (params.error) {
    return { status: 200, html: HTML_CANCELADO };
  }

  let chatId: number;
  try {
    chatId = parseAndValidateState(params.state);
  } catch {
    return { status: 400, html: HTML_ERRO('Link inválido ou expirado. Peça um novo com /evento no bot.') };
  }

  if (!params.code) {
    return { status: 400, html: HTML_ERRO('Código de autorização ausente.') };
  }

  try {
    const tokens = await exchangeCodeForTokens(params.code);
    await upsertOAuthToken(chatId, SERVICE, {
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token,
      expiresAt: new Date(Date.now() + tokens.expires_in * 1000),
    });
    await bot.telegram.sendMessage(chatId, '✅ Outlook conectado! Use /evento para criar ou listar eventos.').catch(() => undefined);
    return { status: 200, html: HTML_SUCESSO };
  } catch (err) {
    console.error('[outlook:callback:error]', err);
    return { status: 500, html: HTML_ERRO('Erro ao conectar com o Outlook. Tente de novo com /evento no bot.') };
  }
}

// ─── Microsoft Graph API ─────────────────────────────────────────────────

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

async function graphFetch(accessToken: string, path: string, init?: RequestInit): Promise<any> {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...(init?.headers ?? {}),
    },
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({})) as any;
    const message = body?.error?.message ?? `HTTP ${res.status}`;
    throw new GraphApiError(res.status, message);
  }
  if (res.status === 204) return null;
  return res.json();
}

function pad2(n: number): string {
  return String(n).padStart(2, '0');
}

export interface EventoInput {
  titulo: string;
  data: string;       // "DD/MM" ou "DD/MM/AAAA"
  horaInicio: string; // "HH:mm"
  horaFim: string;    // "HH:mm"
  local: string;
  descricao: string;
}

export async function createEvent(accessToken: string, input: EventoInput): Promise<{ webLink: string }> {
  const dataObj = parseDataBR(input.data, new Date());
  if (!dataObj) throw new Error(`Data inválida: ${input.data}`);
  const dateStr = `${dataObj.getUTCFullYear()}-${pad2(dataObj.getUTCMonth() + 1)}-${pad2(dataObj.getUTCDate())}`;

  const json = await graphFetch(accessToken, '/me/events', {
    method: 'POST',
    body: JSON.stringify({
      subject: input.titulo,
      body: { contentType: 'text', content: input.descricao },
      start: { dateTime: `${dateStr}T${input.horaInicio}:00`, timeZone: 'America/Fortaleza' },
      end: { dateTime: `${dateStr}T${input.horaFim}:00`, timeZone: 'America/Fortaleza' },
      location: { displayName: input.local },
    }),
  });
  return { webLink: json.webLink };
}

export interface GraphEventSummary {
  subject: string;
  start: string;
  end: string;
  location: string;
  webLink: string;
}

export async function listUpcomingEvents(accessToken: string): Promise<GraphEventSummary[]> {
  const now = new Date();
  const in7days = new Date(now.getTime() + 7 * 86_400_000);
  const params = new URLSearchParams({
    startDateTime: now.toISOString(),
    endDateTime: in7days.toISOString(),
    $orderby: 'start/dateTime',
    $top: '25',
    $select: 'subject,start,end,location,webLink',
  });

  const json = await graphFetch(accessToken, `/me/calendarView?${params.toString()}`, {
    headers: { Prefer: 'outlook.timezone="America/Fortaleza"' },
  });

  return (json.value ?? []).map((ev: any) => ({
    subject: ev.subject || '(sem título)',
    start: ev.start?.dateTime ?? '',
    end: ev.end?.dateTime ?? '',
    location: ev.location?.displayName || '-',
    webLink: ev.webLink,
  }));
}
