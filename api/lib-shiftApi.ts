import { requestRouteWebToken, getTokenPreview } from './lib-routeWebServer.js';
import fs from 'node:fs';
import path from 'node:path';

// ─── Env helpers ──────────────────────────────────────────────────────────

const readEnvFromDotEnvFiles = (name: string): string => {
  const candidates = ['.env.local', '.env'];
  for (const file of candidates) {
    const fullPath = path.join(process.cwd(), file);
    if (!fs.existsSync(fullPath)) continue;
    const content = fs.readFileSync(fullPath, 'utf-8');
    const lines = content.split(/\r?\n/);
    for (const rawLine of lines) {
      const line = String(rawLine || '').trim();
      if (!line || line.startsWith('#')) continue;
      const eqIndex = line.indexOf('=');
      if (eqIndex <= 0) continue;
      const key = line.slice(0, eqIndex).trim().replace(/^export\s+/i, '');
      if (key !== name) continue;
      const rawValue = line.slice(eqIndex + 1).trim();
      return rawValue.replace(/^['"]|['"]$/g, '').trim();
    }
  }
  return '';
};

const readRequiredEnv = (name: string): string => {
  const fromProcess = String(process.env[name] || '').trim();
  const fromDotEnv = readEnvFromDotEnvFiles(name);
  const value = fromDotEnv || fromProcess;
  if (!value) throw new Error(`${name} não configurada`);
  return value;
};

const readOptionalEnv = (name: string): string => {
  const fromDotEnv = readEnvFromDotEnvFiles(name);
  const fromProcess = String(process.env[name] || '').trim();
  return fromDotEnv || fromProcess;
};

// ─── Shift API base URL ──────────────────────────────────────────────────

export const getShiftApiBaseUrl = (): string =>
  readRequiredEnv('SHIFT_API_URL').replace(/\/+$/, '');

// ─── Token: usa credenciais próprias (SHIFT_*) ou fallback para Route Web ─

const TOKEN_PATH = '/api/oauth/token';

const parseTokenResponse = async (response: Response): Promise<{ contentType: string; raw: string; data: any }> => {
  const contentType = String(response.headers.get('content-type') || '');
  const raw = await response.text();
  if (!raw) return { contentType, raw: '', data: null };
  try { return { contentType, raw, data: JSON.parse(raw) }; } catch { return { contentType, raw, data: raw }; }
};

const getTokenFromCandidate = (candidate: any): { token: string; tokenField: string } | null => {
  if (!candidate || typeof candidate !== 'object') return null;
  for (const field of ['access_token', 'token', 'TOKEN', 'bearer', 'Bearer', 'accessToken']) {
    const value = candidate[field];
    if (typeof value === 'string' && value.trim()) {
      return { token: value.trim(), tokenField: field };
    }
  }
  return null;
};

const extractBearerToken = (payload: any): { token: string; tokenField: string } | null => {
  const directString = typeof payload === 'string' ? payload.trim() : '';
  if (directString) return { token: directString, tokenField: 'raw' };
  for (const candidate of [payload, payload?.data, payload?.result, payload?.response]) {
    const tokenMatch = getTokenFromCandidate(candidate);
    if (tokenMatch) return tokenMatch;
  }
  return null;
};

const buildTokenAttempts = (clientId: string, clientSecret: string, username: string, password: string) => {
  const jsonBody = JSON.stringify({
    client_id: clientId,
    client_secret: clientSecret,
    username,
    password,
    grant_type: 'password'
  });
  const formBody = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    username,
    password,
    grant_type: 'password'
  }).toString();
  return [
    { format: 'json' as const, headers: { 'Content-Type': 'application/json', Accept: 'application/json, text/plain, */*' }, body: jsonBody },
    { format: 'form' as const, headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json, text/plain, */*' }, body: formBody }
  ];
};

const hasShiftOwnCredentials = (): boolean => {
  return Boolean(
    readOptionalEnv('SHIFT_URL') &&
    readOptionalEnv('SHIFT_CLIENT_ID') &&
    readOptionalEnv('SHIFT_CLIENT_SECRET')
  );
};

async function requestShiftOwnToken(): Promise<string> {
  const baseUrl = readOptionalEnv('SHIFT_URL').replace(/\/+$/, '');
  const clientId = readOptionalEnv('SHIFT_CLIENT_ID');
  const clientSecret = readOptionalEnv('SHIFT_CLIENT_SECRET');
  const username = readOptionalEnv('SHIFT_USERNAME') || readOptionalEnv('ROUTE_WEB_USERNAME');
  const password = readOptionalEnv('SHIFT_PASSWORD') || readOptionalEnv('ROUTE_WEB_PASSWORD');
  const tokenUrl = `${baseUrl}${TOKEN_PATH}`;

  console.log(`[SHIFT_TOKEN] Usando credenciais próprias: ${baseUrl}${TOKEN_PATH}`);

  let lastErrorMessage = 'Falha desconhecida ao obter token Shift';

  for (const attempt of buildTokenAttempts(clientId, clientSecret, username, password)) {
    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: attempt.headers,
      body: attempt.body
    });

    const parsed = await parseTokenResponse(response);

    if (response.ok) {
      const tokenData = extractBearerToken(parsed.data);
      if (tokenData) {
        console.log(`[SHIFT_TOKEN] Token próprio obtido: preview=${getTokenPreview(tokenData.token)}, field=${tokenData.tokenField}, format=${attempt.format}`);
        return tokenData.token;
      }
      lastErrorMessage = `Token não encontrado na resposta (${attempt.format})`;
      continue;
    }

    const snippet = String(parsed.raw || '').slice(0, 400);
    lastErrorMessage = `Status ${response.status} ao obter token Shift (${attempt.format})${snippet ? `: ${snippet}` : ''}`;
  }

  throw new Error(lastErrorMessage);
}

// ─── Exports ──────────────────────────────────────────────────────────────

export { getTokenPreview };

export async function requestShiftToken(): Promise<string> {
  if (hasShiftOwnCredentials()) {
    return requestShiftOwnToken();
  }
  console.log('[SHIFT_TOKEN] Sem credenciais próprias, usando token Route Web');
  const result = await requestRouteWebToken();
  console.log(`[SHIFT_TOKEN] Token Route Web obtido: preview=${getTokenPreview(result.token)}`);
  return result.token;
}
