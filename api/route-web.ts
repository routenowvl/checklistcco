import type { VercelRequest, VercelResponse } from '@vercel/node';
import {
  getRouteWebUpstreamUrl,
  getRouteWebRoutesEndpointUrl,
  getRouteWebRouteEventsEndpointUrl,
  getRouteWebRoutesEnvDebug,
  getTokenPreview,
  requestRouteWebToken,
  appendQueryToUrl,
  getRouteWebTokenUrl
} from './lib/lib-routeWebServer.js';

/**
 * Endpoint consolidado Route Web.
 * Uso: POST /api/route-web
 * Body: { "resource": "token"|"plants"|"routes"|"route-events"|"proxy", ... }
 */

// ─── Shared helpers ──────────────────────────────────────────────────────

const parseUpstreamResponse = async (response: Response): Promise<{ contentType: string; raw: string; data: any }> => {
  const contentType = String(response.headers.get('content-type') || '');
  const raw = await response.text();
  if (!raw) return { contentType, raw: '', data: null };
  try { return { contentType, raw, data: JSON.parse(raw) }; } catch { return { contentType, raw, data: raw }; }
};

const toOptionalInt = (value: unknown): number | null => {
  if (value == null) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? Math.trunc(parsed) : null;
};

const toBoolean = (value: unknown, fallback: boolean): boolean => {
  if (typeof value === 'boolean') return value;
  if (value == null) return fallback;
  const raw = String(value).trim().toLowerCase();
  if (!raw) return fallback;
  return raw === '1' || raw === 'true' || raw === 'yes' || raw === 'sim';
};

const clampInt = (value: number | null, fallback: number, min: number, max: number): number => {
  const base = value == null || !Number.isFinite(value) ? fallback : value;
  return Math.max(min, Math.min(max, Math.trunc(base)));
};

const pickArray = (payload: any, keys: string[]): any[] => {
  for (const key of keys) {
    const arr = key.includes('.') ? key.split('.').reduce((o, k) => o?.[k], payload) : payload?.[key];
    if (Array.isArray(arr)) return arr;
  }
  return Array.isArray(payload) ? payload : [];
};

// ─── Token ───────────────────────────────────────────────────────────────

const handleToken = async (_body: any, res: VercelResponse) => {
  const tokenResult = await requestRouteWebToken();
  return res.status(200).json({
    success: true,
    url: getRouteWebTokenUrl(),
    status: tokenResult.status,
    format: tokenResult.format,
    tokenField: tokenResult.tokenField,
    token: tokenResult.token,
    tokenPreview: getTokenPreview(tokenResult.token),
    response: tokenResult.data,
    raw: tokenResult.raw
  });
};

// ─── Plants ──────────────────────────────────────────────────────────────

const handlePlants = async (_body: any, res: VercelResponse) => {
  const tokenResult = await requestRouteWebToken();
  const upstreamUrl = getRouteWebUpstreamUrl('/api/plants');

  const response = await fetch(upstreamUrl, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${tokenResult.token}`,
      'Content-Type': 'application/json',
      'X-Requested-With': 'XMLHttpRequest',
      'x-requested_with': 'XLMHttpRequest'
    }
  });

  const upstreamPayload = await parseUpstreamResponse(response);
  const plants = pickArray(upstreamPayload.data, ['data', 'plants', 'items', 'results', 'data.plants']);

  return res.status(200).json({
    success: response.ok,
    upstreamStatus: response.status,
    upstreamStatusText: response.statusText,
    upstreamUrl,
    tokenPreview: getTokenPreview(tokenResult.token),
    tokenField: tokenResult.tokenField,
    tokenFormat: tokenResult.format,
    contentType: upstreamPayload.contentType,
    count: plants.length,
    plants,
    response: upstreamPayload.data,
    raw: upstreamPayload.raw
  });
};

// ─── Routes (with cache) ─────────────────────────────────────────────────

const compactRoute = (route: any): any => ({
  id: route?.id,
  roadmap_code: route?.roadmap_code,
  route_code: route?.route_code,
  code: route?.code,
  route: route?.route,
  route_plan_id: route?.route_plan_id,
  unloading_plate: route?.unloading_plate,
  plant_name: route?.plant_name,
  driver_name: route?.driver_name,
  last_driver_name: route?.last_driver_name,
  driver: route?.driver ? { name: route.driver.name } : undefined,
  last_driver: route?.last_driver ? { name: route.last_driver.name } : undefined,
  plant: route?.plant ? { id: route.plant.id, code: route.plant.code, display_name: route.plant.display_name, name: route.plant.name } : undefined
});

type CachedEntry = {
  ok: boolean; status: number; statusText: string; contentType: string;
  raw: string; data: any; routes: any[]; upstreamSnippet: string; fetchedAt: number;
};

const ROUTES_SUCCESS_TTL = 2 * 60 * 1000;
const ROUTES_ERROR_TTL = 30 * 1000;
const routesCache = new Map<string, CachedEntry>();
const routesInFlight = new Map<string, Promise<CachedEntry>>();

const handleRoutes = async (body: any, res: VercelResponse) => {
  const plantId = toOptionalInt(body.plantId);
  if (plantId == null) return res.status(400).json({ success: false, error: 'plantId é obrigatório' });

  const perPage = toOptionalInt(body.perPage) ?? 60;
  const strictDate = toOptionalInt(body.strictDate) ?? 1;
  const compact = toBoolean(body.compact, false);
  const initialExpectedStartDate = String(body.initialExpectedStartDate || '').trim();
  const finalExpectedStartDate = String(body.finalExpectedStartDate || '').trim();
  if (!initialExpectedStartDate || !finalExpectedStartDate) {
    return res.status(400).json({ success: false, error: 'initialExpectedStartDate e finalExpectedStartDate são obrigatórios' });
  }

  const query = new URLSearchParams({
    plant_id: String(plantId), per_page: String(perPage),
    initial_expected_start_date: initialExpectedStartDate,
    final_expected_start_date: finalExpectedStartDate, strict_date: String(strictDate)
  });

  const manualToken = String(body.bearerToken || '').trim();
  const tokenResult = manualToken ? null : await requestRouteWebToken();
  const bearerToken = manualToken || String(tokenResult?.token || '').trim();
  const tokenSource = manualToken ? 'manual' : 'oauth';
  if (!bearerToken) throw new Error('Token de acesso não disponível');

  const routesEndpointUrl = getRouteWebRoutesEndpointUrl();
  const upstreamUrl = appendQueryToUrl(routesEndpointUrl, query);
  const routesEnvDebug = getRouteWebRoutesEnvDebug();

  const cacheKey = `${upstreamUrl}|${compact ? 1 : 0}`;
  const now = Date.now();
  let fromCache = false;
  let cachedEntry: CachedEntry | undefined;

  const cached = routesCache.get(cacheKey);
  if (cached) {
    const maxAge = cached.ok ? ROUTES_SUCCESS_TTL : ROUTES_ERROR_TTL;
    if (now - cached.fetchedAt <= maxAge) { cachedEntry = cached; fromCache = true; }
  }

  if (!cachedEntry) {
    const inFlight = routesInFlight.get(cacheKey);
    if (inFlight) {
      cachedEntry = await inFlight;
      fromCache = true;
    } else {
      const promise = (async () => {
        const response = await fetch(upstreamUrl, {
          method: 'GET',
          headers: { Authorization: `Bearer ${bearerToken}`, 'Content-Type': 'application/json', 'X-Requested-With': 'XMLHttpRequest', 'x-requested_with': 'XLMHttpRequest' }
        });
        const upstreamPayload = await parseUpstreamResponse(response);
        const routes = pickArray(upstreamPayload.data, ['data', 'routes', 'items', 'results', 'data.routes']);
        const upstreamSnippet = String(upstreamPayload.raw || '').slice(0, 600);
        const entry: CachedEntry = { ok: response.ok, status: response.status, statusText: response.statusText, contentType: upstreamPayload.contentType, raw: upstreamPayload.raw, data: upstreamPayload.data, routes, upstreamSnippet, fetchedAt: Date.now() };
        routesCache.set(cacheKey, entry);
        return entry;
      })().finally(() => { routesInFlight.delete(cacheKey); });
      routesInFlight.set(cacheKey, promise);
      cachedEntry = await promise;
    }
  }

  const routes = cachedEntry.routes;
  const errorMessage = cachedEntry.ok ? undefined : `Upstream ${cachedEntry.status} ${cachedEntry.statusText}`;
  const payload: any = {
    success: cachedEntry.ok, error: errorMessage, upstreamStatus: cachedEntry.status, upstreamStatusText: cachedEntry.statusText,
    routesEndpointUrl, routesEnvDebug, requestQuery: Object.fromEntries(query.entries()), upstreamUrl, upstreamSnippet: cachedEntry.upstreamSnippet,
    fromCache, tokenSource, tokenPreview: getTokenPreview(bearerToken),
    tokenField: tokenResult?.tokenField || 'manual', tokenFormat: tokenResult?.format || 'manual',
    contentType: cachedEntry.contentType, count: routes.length, routes: compact ? routes.map(compactRoute) : routes
  };
  if (!compact) { payload.response = cachedEntry.data; payload.raw = cachedEntry.raw; }

  return res.status(200).json(payload);
};

// ─── Route Events (with cache) ───────────────────────────────────────────

const EVENTS_SUCCESS_TTL = 2 * 60 * 1000;
const EVENTS_ERROR_TTL = 30 * 1000;
const DEFAULT_TIMEOUT = 9000;
const DEFAULT_RETRY_TIMEOUT = 14000;
const DEFAULT_RETRY_DELAY = 120;

type CachedRouteEvents = {
  ok: boolean; routeId: number; upstreamStatus: number; upstreamStatusText: string;
  upstreamUrl: string; upstreamSnippet: string; contentType: string;
  requestQuery: Record<string, string>; events: any[]; error?: string; fetchedAt: number;
};

const routeEventsCache = new Map<string, CachedRouteEvents>();
const routeEventsInflight = new Map<string, Promise<CachedRouteEvents>>();

const sleep = (ms: number): Promise<void> => new Promise(resolve => setTimeout(resolve, ms));

const normalizeText = (value: unknown): string =>
  String(value ?? '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

const isScraperSmartQuestionOccurrence = (occ: any): boolean =>
  normalizeText(occ?.inserted_by) === 'scrapersmartquestion';

const NON_COLLECTION_TECHNICAL_OCCURRENCE_IDS = new Set<number>([2]);
const NON_COLLECTION_EXCLUDED_REASON_PATTERNS = [
  'troca de caminhao', 'troca de caminhão', 'evento extra', 'evento_extra',
  'tanque comunitario', 'tanque_comunitario', 'alteracao de horario', 'troca de reboque', 'falta de sinal do rastreador'
];

const getOccurrenceDescription = (occ: any): string =>
  String(occ?.occurrence_type?.description || occ?.occurrence_type_description || occ?.description || occ?.type_name || occ?.name || occ?.title || '').trim();

const isTechnicalOccurrence = (occ: any): boolean => {
  const occurrenceTypeId = toOptionalInt(occ?.occurrence_type_id ?? occ?.occurrence_type?.id);
  if (occurrenceTypeId != null && NON_COLLECTION_TECHNICAL_OCCURRENCE_IDS.has(occurrenceTypeId)) return true;
  const desc = normalizeText(getOccurrenceDescription(occ));
  return desc.includes('atualizacao de posicao pelo rastreador') || desc.includes('evento fora de ordem');
};

const isNonCollectionOccurrence = (occ: any): boolean => {
  if (!occ || typeof occ !== 'object') return false;
  const desc = normalizeText(getOccurrenceDescription(occ));
  if (NON_COLLECTION_EXCLUDED_REASON_PATTERNS.some(p => desc.includes(normalizeText(p)))) return false;
  if (isScraperSmartQuestionOccurrence(occ)) return true;
  if (isTechnicalOccurrence(occ)) return false;
  const occurrenceTypeId = toOptionalInt(occ?.occurrence_type_id ?? occ?.occurrence_type?.id);
  return occurrenceTypeId != null || Boolean(desc);
};

const filterNonCollectionEvents = (events: any[]): any[] =>
  events.map(e => {
    if (normalizeText(e?.type_name) !== 'coleta') return null;
    const matching = (Array.isArray(e?.occurrences) ? e.occurrences : []).filter(isNonCollectionOccurrence);
    return matching.length === 0 ? null : { ...e, occurrences: matching };
  }).filter(Boolean) as any[];

const compactOccurrence = (occ: any): any => ({
  id: occ?.id, inserted_by: occ?.inserted_by, inserted_at: occ?.inserted_at,
  occurrence_type_id: occ?.occurrence_type_id ?? occ?.occurrence_type?.id,
  occurrence_type_description: occ?.occurrence_type?.description || occ?.occurrence_type_description || ''
});

const compactEvent = (event: any): any => ({
  id: event?.id, route_id: event?.route_id, event_id: event?.event_id, type_name: event?.type_name,
  reference: event?.reference, reference_code: event?.reference_code, status: event?.status, executed: event?.executed,
  expected_arrival: event?.expected_arrival, actual_arrival: event?.actual_arrival,
  expected_departure: event?.expected_departure, actual_departure: event?.actual_departure,
  created_at: event?.created_at, updated_at: event?.updated_at,
  occurrences: Array.isArray(event?.occurrences) ? event.occurrences.map(compactOccurrence) : []
});

const fetchRouteEventsFromUpstream = async (
  routeId: number, options: { withOccurrences: boolean; nonCollectionOnly: boolean; bearerToken: string; timeoutMs: number; retryTimeoutMs: number; retryDelayMs: number }
): Promise<CachedRouteEvents> => {
  const query = new URLSearchParams({ with_occurrences: options.withOccurrences ? 'true' : 'false' });
  const requestQuery = Object.fromEntries(query.entries());
  const upstreamUrl = appendQueryToUrl(getRouteWebRouteEventsEndpointUrl(routeId), query);

  const executeCall = async (timeoutMs: number): Promise<Omit<CachedRouteEvents, 'fetchedAt'>> => {
    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      let response: Response;
      try {
        response = await fetch(upstreamUrl, {
          method: 'GET', signal: controller.signal,
          headers: { Authorization: `Bearer ${options.bearerToken}`, 'Content-Type': 'application/json', 'X-Requested-With': 'XMLHttpRequest', 'x-requested_with': 'XLMHttpRequest' }
        });
      } finally { clearTimeout(timer); }
      const upstreamPayload = await parseUpstreamResponse(response);
      const sourceEvents = pickArray(upstreamPayload.data, ['data', 'events', 'items', 'results', 'data.events']);
      const filteredEvents = options.nonCollectionOnly
        ? filterNonCollectionEvents(sourceEvents).map((e: any) => ({ ...e, occurrences: Array.isArray(e?.occurrences) ? e.occurrences.filter(isNonCollectionOccurrence) : [] }))
        : sourceEvents;
      const upstreamSnippet = String(upstreamPayload.raw || '').slice(0, 600);
      return {
        ok: response.ok, routeId, upstreamStatus: response.status, upstreamStatusText: response.statusText,
        upstreamUrl, upstreamSnippet, contentType: upstreamPayload.contentType, requestQuery, events: filteredEvents,
        error: response.ok ? undefined : `Upstream ${response.status} ${response.statusText}`
      };
    } catch (error: any) {
      const isTimeout = error?.name === 'AbortError';
      return {
        ok: false, routeId, upstreamStatus: 0, upstreamStatusText: isTimeout ? 'Timeout' : 'NetworkError',
        upstreamUrl, upstreamSnippet: '', contentType: '', requestQuery, events: [],
        error: isTimeout ? `Timeout ao consultar eventos da rota ${routeId}` : error?.message || `Erro inesperado`
      };
    }
  };

  let result = await executeCall(options.timeoutMs);
  if (!result.ok && /timeout/i.test(String(result.error || ''))) {
    await sleep(options.retryDelayMs);
    const retry = await executeCall(options.retryTimeoutMs);
    if (retry.ok) result = retry;
    else result = { ...retry, error: `${retry.error} (retry 1/1 também falhou)` };
  }

  return { ...result, fetchedAt: Date.now() };
};

const fetchRouteEventsWithCache = async (routeId: number, options: Parameters<typeof fetchRouteEventsFromUpstream>[1]): Promise<{ entry: CachedRouteEvents; fromCache: boolean }> => {
  const cacheKey = `${routeId}|${options.withOccurrences ? 1 : 0}|${options.nonCollectionOnly ? 1 : 0}`;
  const now = Date.now();
  const cached = routeEventsCache.get(cacheKey);
  if (cached) {
    const maxAge = cached.ok ? EVENTS_SUCCESS_TTL : EVENTS_ERROR_TTL;
    if (now - cached.fetchedAt <= maxAge) return { entry: cached, fromCache: true };
  }
  const inFlight = routeEventsInflight.get(cacheKey);
  if (inFlight) return { entry: await inFlight, fromCache: true };

  const promise = fetchRouteEventsFromUpstream(routeId, options).then(entry => {
    routeEventsCache.set(cacheKey, entry);
    routeEventsInflight.delete(cacheKey);
    return entry;
  }).catch(error => { routeEventsInflight.delete(cacheKey); throw error; });
  routeEventsInflight.set(cacheKey, promise);
  return { entry: await promise, fromCache: false };
};

const runWithConcurrency = async <T, R>(items: T[], concurrency: number, mapper: (item: T, i: number) => Promise<R>): Promise<R[]> => {
  if (items.length === 0) return [];
  const safeConcurrency = Math.max(1, Math.min(concurrency, items.length));
  const results = new Array<R>(items.length);
  let cursor = 0;
  await Promise.all(Array.from({ length: safeConcurrency }, async () => {
    while (true) {
      const index = cursor++;
      if (index >= items.length) break;
      results[index] = await mapper(items[index], index);
    }
  }));
  return results;
};

const handleRouteEvents = async (body: any, res: VercelResponse) => {
  const routeIdList = (Array.isArray(body.routeIds) ? body.routeIds : []).map((v: any) => toOptionalInt(v)).filter((id: number | null): id is number => id != null);
  const single = toOptionalInt(body.routeId);
  if (single != null) routeIdList.push(single);
  const routeIds = Array.from(new Set(routeIdList));
  if (routeIds.length === 0) return res.status(400).json({ success: false, error: 'routeId/routeIds é obrigatório' });

  const withOccurrences = toBoolean(body.withOccurrences, true);
  const compact = toBoolean(body.compact, false);
  const nonCollectionOnly = toBoolean(body.nonCollectionOnly, false);
  const timeoutMs = clampInt(toOptionalInt(body.timeoutMs), DEFAULT_TIMEOUT, 4000, 30000);
  const retryTimeoutMs = clampInt(toOptionalInt(body.retryTimeoutMs), DEFAULT_RETRY_TIMEOUT, 5000, 40000);
  const retryDelayMs = clampInt(toOptionalInt(body.retryDelayMs), DEFAULT_RETRY_DELAY, 0, 2000);
  const concurrency = clampInt(toOptionalInt(body.concurrency), 2, 1, 6);

  const manualToken = String(body.bearerToken || '').trim();
  const tokenResult = manualToken ? null : await requestRouteWebToken();
  const bearerToken = manualToken || String(tokenResult?.token || '').trim();
  const tokenSource = manualToken ? 'manual' : 'oauth';
  if (!bearerToken) throw new Error('Token de acesso não disponível');

  const routesEnvDebug = getRouteWebRoutesEnvDebug();
  const opts = { withOccurrences, nonCollectionOnly, bearerToken, timeoutMs, retryTimeoutMs, retryDelayMs };

  if (routeIds.length === 1) {
    const routeId = routeIds[0];
    const { entry, fromCache } = await fetchRouteEventsWithCache(routeId, opts);
    return res.status(200).json({
      success: entry.ok, error: entry.error, routeId, upstreamStatus: entry.upstreamStatus, upstreamStatusText: entry.upstreamStatusText,
      eventsEndpointUrl: getRouteWebRouteEventsEndpointUrl(routeId), routesEnvDebug, requestQuery: entry.requestQuery,
      upstreamUrl: entry.upstreamUrl, upstreamSnippet: entry.upstreamSnippet, tokenSource, tokenPreview: getTokenPreview(bearerToken),
      tokenField: tokenResult?.tokenField || 'manual', tokenFormat: tokenResult?.format || 'manual',
      contentType: entry.contentType, nonCollectionOnly, count: entry.events.length, fromCache,
      events: compact ? entry.events.map(compactEvent) : entry.events
    });
  }

  const startedAt = Date.now();
  const entries = await runWithConcurrency(routeIds, concurrency, async (routeId) => {
    const { entry, fromCache } = await fetchRouteEventsWithCache(routeId, opts);
    return { routeId, fromCache, entry };
  });

  const results = entries.map(item => ({
    routeId: item.routeId, success: item.entry.ok, error: item.entry.error,
    upstreamStatus: item.entry.upstreamStatus, upstreamStatusText: item.entry.upstreamStatusText,
    requestQuery: item.entry.requestQuery, upstreamUrl: item.entry.upstreamUrl,
    upstreamSnippet: item.entry.upstreamSnippet, contentType: item.entry.contentType,
    count: item.entry.events.length, fromCache: item.fromCache,
    events: compact ? item.entry.events.map(compactEvent) : item.entry.events
  }));

  const successCount = results.filter(r => r.success).length;
  const failedCount = results.length - successCount;

  return res.status(200).json({
    success: failedCount === 0, batch: true, totalRoutes: routeIds.length,
    successCount, failedCount, tokenSource, tokenPreview: getTokenPreview(bearerToken),
    tokenField: tokenResult?.tokenField || 'manual', tokenFormat: tokenResult?.format || 'manual',
    nonCollectionOnly, withOccurrences, compact, concurrency, timeoutMs, retryTimeoutMs, retryDelayMs,
    routesEnvDebug, durationMs: Date.now() - startedAt, results
  });
};

// ─── Proxy ───────────────────────────────────────────────────────────────

const METHODS_WITHOUT_BODY = new Set(['GET', 'HEAD']);

const sanitizeHeaders = (headers: Record<string, string> | undefined): Record<string, string> => {
  const result: Record<string, string> = {};
  if (!headers || typeof headers !== 'object') return result;
  for (const [key, value] of Object.entries(headers)) {
    const k = String(key || '').trim();
    const v = String(value || '').trim();
    if (!k || !v || /^(authorization|host|content-length)$/i.test(k)) continue;
    result[k] = v;
  }
  return result;
};

const handleProxy = async (body: any, res: VercelResponse) => {
  const method = (() => {
    const candidate = String(body.method || 'GET').trim().toUpperCase();
    return ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'HEAD'].includes(candidate) ? candidate : 'GET';
  })();

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
    method, headers, body: METHODS_WITHOUT_BODY.has(method) ? undefined : rawBody
  });

  const upstreamPayload = await parseUpstreamResponse(response);

  return res.status(200).json({
    success: response.ok, upstreamStatus: response.status, upstreamStatusText: response.statusText,
    upstreamUrl, tokenSource, tokenPreview: getTokenPreview(bearerToken),
    contentType: upstreamPayload.contentType, response: upstreamPayload.data, raw: upstreamPayload.raw
  });
};

// ─── Handler ─────────────────────────────────────────────────────────────

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const { resource } = req.body || {};
    if (!resource) return res.status(400).json({ success: false, error: 'resource é obrigatório (token, plants, routes, route-events, proxy)' });

    switch (resource) {
      case 'token': return await handleToken(req.body, res);
      case 'plants': return await handlePlants(req.body, res);
      case 'routes': return await handleRoutes(req.body, res);
      case 'route-events': return await handleRouteEvents(req.body, res);
      case 'proxy': return await handleProxy(req.body, res);
      default: return res.status(400).json({ success: false, error: `Resource desconhecido: ${resource}` });
    }
  } catch (error: any) {
    console.error('[ROUTE_WEB] Erro:', error?.message || error);
    return res.status(500).json({ success: false, error: error?.message || 'Erro no Route Web' });
  }
}
