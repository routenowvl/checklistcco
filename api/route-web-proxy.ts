import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getRouteWebUpstreamUrl, getTokenPreview, requestRouteWebToken } from '../utils/routeWebServer.js';

type ProxyRequestBody = {
  path?: string;
  method?: string;
  bearerToken?: string;
  body?: string;
  contentType?: string;
  headers?: Record<string, string>;
};

const METHODS_WITHOUT_BODY = new Set(['GET', 'HEAD']);

const normalizeMethod = (value: string | undefined): string => {
  const candidate = String(value || 'GET').trim().toUpperCase();
  const allowed = ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'HEAD'];
  return allowed.includes(candidate) ? candidate : 'GET';
};

const sanitizeHeaders = (headers: Record<string, string> | undefined): Record<string, string> => {
  const result: Record<string, string> = {};
  if (!headers || typeof headers !== 'object') return result;

  for (const [key, value] of Object.entries(headers)) {
    const normalizedKey = String(key || '').trim();
    const normalizedValue = String(value || '').trim();
    if (!normalizedKey || !normalizedValue) continue;
    if (/^(authorization|host|content-length)$/i.test(normalizedKey)) continue;
    result[normalizedKey] = normalizedValue;
  }

  return result;
};

const parseUpstreamResponse = async (response: Response): Promise<{ contentType: string; raw: string; data: any }> => {
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

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const body = (req.body || {}) as ProxyRequestBody;
    const method = normalizeMethod(body.method);
    const upstreamUrl = getRouteWebUpstreamUrl(body.path || '');
    const manualToken = String(body.bearerToken || '').trim();

    let bearerToken = manualToken;
    let tokenSource: 'manual' | 'oauth' = 'manual';

    if (!bearerToken) {
      const tokenResult = await requestRouteWebToken();
      bearerToken = tokenResult.token;
      tokenSource = 'oauth';
    }

    const headers: Record<string, string> = {
      Accept: 'application/json, text/plain, */*',
      Authorization: `Bearer ${bearerToken}`,
      ...sanitizeHeaders(body.headers)
    };

    const rawBody = typeof body.body === 'string' ? body.body : '';
    if (!METHODS_WITHOUT_BODY.has(method)) {
      headers['Content-Type'] = String(body.contentType || 'application/json').trim() || 'application/json';
    }

    const response = await fetch(upstreamUrl, {
      method,
      headers,
      body: METHODS_WITHOUT_BODY.has(method) ? undefined : rawBody
    });

    const upstreamPayload = await parseUpstreamResponse(response);

    return res.status(200).json({
      success: response.ok,
      upstreamStatus: response.status,
      upstreamStatusText: response.statusText,
      upstreamUrl,
      tokenSource,
      tokenPreview: getTokenPreview(bearerToken),
      contentType: upstreamPayload.contentType,
      response: upstreamPayload.data,
      raw: upstreamPayload.raw
    });
  } catch (error: any) {
    console.error('[ROUTE_WEB_PROXY] Erro ao chamar endpoint do Route Web:', error?.message || error);
    return res.status(500).json({
      success: false,
      error: error?.message || 'Erro ao chamar endpoint do Route Web'
    });
  }
}
