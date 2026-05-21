import fs from 'node:fs';
import path from 'node:path';

export type RouteWebTokenResult = {
  token: string;
  tokenField: string;
  format: 'json' | 'form';
  url: string;
  status: number;
  contentType: string;
  data: any;
  raw: string;
};

const TOKEN_PATH = '/api/oauth/token';
const ROUTES_PATH_FALLBACK = '/api/routes';

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
      const unquoted = rawValue.replace(/^['"]|['"]$/g, '').trim();
      return unquoted;
    }
  }

  return '';
};

const readRequiredEnv = (name: string): string => {
  const fromProcess = String(process.env[name] || '').trim();
  const fromDotEnv = readEnvFromDotEnvFiles(name);
  const value = fromDotEnv || fromProcess;
  if (!value) {
    throw new Error(`${name} não configurada`);
  }
  return value;
};

const normalizeBaseUrl = (value: string): string => value.replace(/\/+$/, '');
const isAbsoluteHttpUrl = (value: string): boolean => /^https?:\/\//i.test(value);

export const getRouteWebBaseUrl = (): string =>
  normalizeBaseUrl(readRequiredEnv('ROUTE_WEB_URL'));

export const getRouteWebTokenUrl = (): string =>
  `${getRouteWebBaseUrl()}${TOKEN_PATH}`;

export const getRouteWebUpstreamUrl = (path: string): string => {
  const rawPath = String(path || '').trim();
  if (!rawPath) {
    throw new Error('Informe um path relativo para o endpoint do RP');
  }

  if (/^https?:\/\//i.test(rawPath)) {
    throw new Error('Use apenas path relativo do RP, sem http/https');
  }

  const normalizedPath = rawPath.startsWith('/') ? rawPath : `/${rawPath}`;
  return `${getRouteWebBaseUrl()}${normalizedPath}`;
};

type RouteWebRoutesUrlSource = '.env' | 'process.env' | 'missing';

const getRouteWebRoutesUrlConfig = (): {
  source: RouteWebRoutesUrlSource;
  value: string;
  valueFromDotEnv: string;
  valueFromProcessEnv: string;
} => {
  const customFromProcess = String(process.env.ROUTE_WEB_ROUTES_URL || '').trim();
  const customFromDotEnv = readEnvFromDotEnvFiles('ROUTE_WEB_ROUTES_URL');
  const custom = customFromDotEnv || customFromProcess;
  return {
    source: customFromDotEnv ? '.env' : customFromProcess ? 'process.env' : 'missing',
    value: custom,
    valueFromDotEnv: customFromDotEnv,
    valueFromProcessEnv: customFromProcess
  };
};

export const getRouteWebRoutesEnvDebug = () => {
  const config = getRouteWebRoutesUrlConfig();
  return {
    cwd: process.cwd(),
    source: config.source,
    configured: Boolean(config.value),
    value: config.value,
    fromDotEnv: config.valueFromDotEnv,
    fromProcessEnv: config.valueFromProcessEnv
  };
};

const resolveRouteWebRoutesUrl = (): { parsed: URL; source: RouteWebRoutesUrlSource; raw: string } => {
  const config = getRouteWebRoutesUrlConfig();

  if (!config.value) {
    throw new Error('ROUTE_WEB_ROUTES_URL não configurada');
  }

  if (!isAbsoluteHttpUrl(config.value)) {
    throw new Error('ROUTE_WEB_ROUTES_URL deve ser uma URL absoluta iniciando com http/https');
  }

  const parsed = new URL(config.value);
  parsed.hash = '';
  parsed.search = '';

  const normalizedPath = String(parsed.pathname || '').replace(/\/+$/, '');
  if (!normalizedPath || normalizedPath === '/') {
    parsed.pathname = ROUTES_PATH_FALLBACK;
  }

  return { parsed, source: config.source, raw: config.value };
};

export const getRouteWebRoutesEndpointUrl = (): string => {
  const resolved = resolveRouteWebRoutesUrl();
  const url = resolved.parsed.toString();

  console.log('[ROUTE_WEB_SERVER] ROUTE_WEB_ROUTES_URL resolvida:', {
    source: resolved.source,
    raw: resolved.raw,
    resolved: url
  });

  return url;
};

export const getRouteWebRouteEventsEndpointUrl = (routeId: number | string): string => {
  const routeIdText = String(routeId ?? '').trim();
  if (!routeIdText || !/^\d+$/.test(routeIdText)) {
    throw new Error('routeId inválido para consulta de eventos');
  }

  const resolved = resolveRouteWebRoutesUrl();
  const routePath = String(resolved.parsed.pathname || '').replace(/\/+$/, '') || ROUTES_PATH_FALLBACK;
  resolved.parsed.pathname = `${routePath}/${routeIdText}/events`;

  const url = resolved.parsed.toString();
  console.log('[ROUTE_WEB_SERVER] ROUTE_WEB_EVENTS_URL resolvida:', {
    routeId: routeIdText,
    source: resolved.source,
    raw: resolved.raw,
    resolved: url
  });

  return url;
};

export const appendQueryToUrl = (url: string, query: URLSearchParams): string => {
  const queryString = query.toString();
  if (!queryString) return url;
  return `${url}${url.includes('?') ? '&' : '?'}${queryString}`;
};

const parseResponse = async (response: Response): Promise<{ contentType: string; raw: string; data: any }> => {
  const contentType = String(response.headers.get('content-type') || '');
  const raw = await response.text();

  if (!raw) {
    return { contentType, raw: '', data: null };
  }

  try {
    return { contentType, raw, data: JSON.parse(raw) };
  } catch {
    return { contentType, raw, data: raw };
  }
};

const getTokenFromCandidate = (candidate: any): { token: string; tokenField: string } | null => {
  if (!candidate || typeof candidate !== 'object') return null;

  const possibleFields = [
    'access_token',
    'token',
    'TOKEN',
    'bearer',
    'Bearer',
    'accessToken'
  ];

  for (const field of possibleFields) {
    const value = candidate[field];
    if (typeof value === 'string' && value.trim()) {
      return { token: value.trim(), tokenField: field };
    }
  }

  return null;
};

const extractBearerToken = (payload: any): { token: string; tokenField: string } | null => {
  const directString = typeof payload === 'string' ? payload.trim() : '';
  if (directString) {
    return { token: directString, tokenField: 'raw' };
  }

  const candidates = [
    payload,
    payload?.data,
    payload?.result,
    payload?.response
  ];

  for (const candidate of candidates) {
    const tokenMatch = getTokenFromCandidate(candidate);
    if (tokenMatch) return tokenMatch;
  }

  return null;
};

const buildTokenAttempts = (
  clientId: string,
  clientSecret: string,
  username: string,
  password: string
) => {
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
    {
      format: 'json' as const,
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json, text/plain, */*'
      },
      body: jsonBody
    },
    {
      format: 'form' as const,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        Accept: 'application/json, text/plain, */*'
      },
      body: formBody
    }
  ];
};

export const getTokenPreview = (token: string): string => {
  const normalized = String(token || '').trim();
  if (!normalized) return '';
  if (normalized.length <= 24) return normalized;
  return `${normalized.slice(0, 18)}...${normalized.slice(-6)}`;
};

export async function requestRouteWebToken(): Promise<RouteWebTokenResult> {
  const url = getRouteWebTokenUrl();
  const clientId = readRequiredEnv('ROUTE_WEB_CLIENT_ID');
  const clientSecret = readRequiredEnv('ROUTE_WEB_CLIENT_SECRET');
  const username = readRequiredEnv('ROUTE_WEB_USERNAME');
  const password = readRequiredEnv('ROUTE_WEB_PASSWORD');

  let lastErrorMessage = 'Falha desconhecida ao obter token';

  for (const attempt of buildTokenAttempts(clientId, clientSecret, username, password)) {
    const response = await fetch(url, {
      method: 'POST',
      headers: attempt.headers,
      body: attempt.body
    });

    const parsed = await parseResponse(response);

    if (response.ok) {
      const tokenData = extractBearerToken(parsed.data);
      if (tokenData) {
        return {
          token: tokenData.token,
          tokenField: tokenData.tokenField,
          format: attempt.format,
          url,
          status: response.status,
          contentType: parsed.contentType,
          data: parsed.data,
          raw: parsed.raw
        };
      }

      lastErrorMessage = `Token não encontrado na resposta do endpoint (${attempt.format})`;
      continue;
    }

    const snippet = String(parsed.raw || '').slice(0, 400);
    lastErrorMessage = `Status ${response.status} ao obter token (${attempt.format})${snippet ? `: ${snippet}` : ''}`;
  }

  throw new Error(lastErrorMessage);
}
