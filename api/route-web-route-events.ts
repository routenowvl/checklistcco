import type { VercelRequest, VercelResponse } from '@vercel/node';
import {
  appendQueryToUrl,
  getRouteWebRouteEventsEndpointUrl,
  getRouteWebRoutesEnvDebug,
  getTokenPreview,
  requestRouteWebToken
} from '../utils/routeWebServer.js';

type RouteWebRouteEventsBody = {
  routeId?: number | string;
  routeIds?: Array<number | string>;
  withOccurrences?: boolean | number | string;
  bearerToken?: string;
  compact?: boolean | number | string;
  nonCollectionOnly?: boolean | number | string;
  timeoutMs?: number | string;
  retryTimeoutMs?: number | string;
  retryDelayMs?: number | string;
  concurrency?: number | string;
};

type RouteEventsFetchResult = {
  ok: boolean;
  routeId: number;
  upstreamStatus: number;
  upstreamStatusText: string;
  upstreamUrl: string;
  upstreamSnippet: string;
  contentType: string;
  requestQuery: Record<string, string>;
  events: any[];
  error?: string;
};

type CachedRouteEventsFetchResult = RouteEventsFetchResult & {
  fetchedAt: number;
};

const EVENTS_SUCCESS_CACHE_TTL_MS = 2 * 60 * 1000;
const EVENTS_ERROR_CACHE_TTL_MS = 30 * 1000;
const DEFAULT_TIMEOUT_MS = 9000;
const DEFAULT_RETRY_TIMEOUT_MS = 14000;
const DEFAULT_RETRY_DELAY_MS = 120;

const routeEventsCache = new Map<string, CachedRouteEventsFetchResult>();
const routeEventsInflight = new Map<string, Promise<CachedRouteEventsFetchResult>>();

const toOptionalInt = (value: unknown): number | null => {
  if (value == null) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  const parsed = Number(raw);
  if (!Number.isFinite(parsed)) return null;
  return Math.trunc(parsed);
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

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

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

const pickEventsArray = (payload: any): any[] => {
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload?.events)) return payload.events;
  if (Array.isArray(payload?.items)) return payload.items;
  if (Array.isArray(payload?.results)) return payload.results;
  if (Array.isArray(payload?.data?.events)) return payload.data.events;
  return [];
};

const normalizeText = (value: unknown): string =>
  String(value ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

const isScraperSmartQuestionOccurrence = (occ: any): boolean =>
  normalizeText(occ?.inserted_by) === 'scrapersmartquestion';

const NON_COLLECTION_TECHNICAL_OCCURRENCE_IDS = new Set<number>([2]);
const NON_COLLECTION_EXCLUDED_REASON_PATTERNS = [
  'troca de caminhao',
  'troca de caminhão',
  'evento extra',
  'evento_extra',
  'tanque comunitario',
  'tanque_comunitario',
  'alteracao de horario',
  'troca de reboque',
  'falta de sinal do rastreador'
];

const getOccurrenceDescription = (occ: any): string =>
  String(
    occ?.occurrence_type?.description ||
      occ?.occurrence_type_description ||
      occ?.description ||
      occ?.type_name ||
      occ?.name ||
      occ?.title ||
      ''
  ).trim();

const isTechnicalOccurrence = (occ: any): boolean => {
  const occurrenceTypeId = toOptionalInt(occ?.occurrence_type_id ?? occ?.occurrence_type?.id);
  if (occurrenceTypeId != null && NON_COLLECTION_TECHNICAL_OCCURRENCE_IDS.has(occurrenceTypeId)) {
    return true;
  }

  const descriptionNormalized = normalizeText(getOccurrenceDescription(occ));
  if (!descriptionNormalized) return false;

  return (
    descriptionNormalized.includes('atualizacao de posicao pelo rastreador') ||
    descriptionNormalized.includes('evento fora de ordem')
  );
};

const isNonCollectionOccurrence = (occ: any): boolean => {
  if (!occ || typeof occ !== 'object') return false;
  const description = getOccurrenceDescription(occ);
  const descriptionNormalized = normalizeText(description);
  if (NON_COLLECTION_EXCLUDED_REASON_PATTERNS.some((pattern) => descriptionNormalized.includes(normalizeText(pattern)))) {
    return false;
  }

  if (isScraperSmartQuestionOccurrence(occ)) return true;
  if (isTechnicalOccurrence(occ)) return false;

  const occurrenceTypeId = toOptionalInt(occ?.occurrence_type_id ?? occ?.occurrence_type?.id);
  return occurrenceTypeId != null || Boolean(description);
};

const filterNonCollectionEvents = (events: any[]): any[] =>
  events
    .map((event) => {
      const typeNameNormalized = normalizeText(event?.type_name);
      if (typeNameNormalized !== 'coleta') return null;

      const occurrences = Array.isArray(event?.occurrences) ? event.occurrences : [];
      const matching = occurrences.filter(isNonCollectionOccurrence);
      if (matching.length === 0) return null;
      return { ...event, occurrences: matching };
    })
    .filter(Boolean) as any[];

const keepOnlyNonCollectionOccurrences = (event: any): any => ({
  ...event,
  occurrences: Array.isArray(event?.occurrences) ? event.occurrences.filter(isNonCollectionOccurrence) : []
});

const compactOccurrence = (occ: any): any => ({
  id: occ?.id,
  inserted_by: occ?.inserted_by,
  inserted_at: occ?.inserted_at,
  occurrence_type_id: occ?.occurrence_type_id ?? occ?.occurrence_type?.id,
  occurrence_type_description: occ?.occurrence_type?.description || occ?.occurrence_type_description || ''
});

const compactEvent = (event: any): any => ({
  id: event?.id,
  route_id: event?.route_id,
  event_id: event?.event_id,
  type_name: event?.type_name,
  reference: event?.reference,
  reference_code: event?.reference_code,
  status: event?.status,
  executed: event?.executed,
  expected_arrival: event?.expected_arrival,
  actual_arrival: event?.actual_arrival,
  expected_departure: event?.expected_departure,
  actual_departure: event?.actual_departure,
  created_at: event?.created_at,
  updated_at: event?.updated_at,
  occurrences: Array.isArray(event?.occurrences) ? event.occurrences.map(compactOccurrence) : []
});

const buildRouteCacheKey = (routeId: number, withOccurrences: boolean, nonCollectionOnly: boolean): string =>
  `${routeId}|${withOccurrences ? 1 : 0}|${nonCollectionOnly ? 1 : 0}`;

const withTimeout = async (upstreamUrl: string, bearerToken: string, timeoutMs: number): Promise<Response> => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(upstreamUrl, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${bearerToken}`,
        'Content-Type': 'application/json',
        'X-Requested-With': 'XMLHttpRequest',
        'x-requested_with': 'XLMHttpRequest'
      },
      signal: controller.signal
    });
  } finally {
    clearTimeout(timer);
  }
};

const fetchRouteEventsFromUpstream = async (
  routeId: number,
  options: {
    withOccurrences: boolean;
    nonCollectionOnly: boolean;
    bearerToken: string;
    timeoutMs: number;
    retryTimeoutMs: number;
    retryDelayMs: number;
  }
): Promise<CachedRouteEventsFetchResult> => {
  const query = new URLSearchParams({
    with_occurrences: options.withOccurrences ? 'true' : 'false'
  });
  const requestQuery = Object.fromEntries(query.entries());
  const eventsEndpointUrl = getRouteWebRouteEventsEndpointUrl(routeId);
  const upstreamUrl = appendQueryToUrl(eventsEndpointUrl, query);

  const executeCall = async (timeoutMs: number): Promise<RouteEventsFetchResult> => {
    try {
      const response = await withTimeout(upstreamUrl, options.bearerToken, timeoutMs);
      const upstreamPayload = await parseUpstreamResponse(response);
      const sourceEvents = pickEventsArray(upstreamPayload.data);
      const filteredEvents = options.nonCollectionOnly
        ? filterNonCollectionEvents(sourceEvents).map(keepOnlyNonCollectionOccurrences)
        : sourceEvents;
      const upstreamSnippet = String(upstreamPayload.raw || '').slice(0, 600);

      const errorMessage = response.ok
        ? undefined
        : `Upstream ${response.status} ${response.statusText} em ${upstreamUrl}${upstreamSnippet ? ` | ${upstreamSnippet}` : ''}`;

      return {
        ok: response.ok,
        routeId,
        upstreamStatus: response.status,
        upstreamStatusText: response.statusText,
        upstreamUrl,
        upstreamSnippet,
        contentType: upstreamPayload.contentType,
        requestQuery,
        events: filteredEvents,
        error: errorMessage
      };
    } catch (error: any) {
      const isTimeout = error?.name === 'AbortError';
      return {
        ok: false,
        routeId,
        upstreamStatus: 0,
        upstreamStatusText: isTimeout ? 'Timeout' : 'NetworkError',
        upstreamUrl,
        upstreamSnippet: '',
        contentType: '',
        requestQuery,
        events: [],
        error: isTimeout
          ? `Timeout ao consultar eventos da rota ${routeId}`
          : error?.message || `Erro inesperado ao consultar eventos da rota ${routeId}`
      };
    }
  };

  let result = await executeCall(options.timeoutMs);
  if (!result.ok && /timeout/i.test(String(result.error || ''))) {
    await sleep(options.retryDelayMs);
    const retry = await executeCall(options.retryTimeoutMs);
    if (retry.ok) {
      result = retry;
    } else {
      result = {
        ...retry,
        error: `${retry.error} (retry 1/1 também falhou)`
      };
    }
  }

  return {
    ...result,
    fetchedAt: Date.now()
  };
};

const fetchRouteEventsWithCache = async (
  routeId: number,
  options: {
    withOccurrences: boolean;
    nonCollectionOnly: boolean;
    bearerToken: string;
    timeoutMs: number;
    retryTimeoutMs: number;
    retryDelayMs: number;
  }
): Promise<{ entry: CachedRouteEventsFetchResult; fromCache: boolean }> => {
  const cacheKey = buildRouteCacheKey(routeId, options.withOccurrences, options.nonCollectionOnly);
  const now = Date.now();
  const cached = routeEventsCache.get(cacheKey);
  if (cached) {
    const maxAge = cached.ok ? EVENTS_SUCCESS_CACHE_TTL_MS : EVENTS_ERROR_CACHE_TTL_MS;
    if (now - cached.fetchedAt <= maxAge) {
      return { entry: cached, fromCache: true };
    }
  }

  const inFlight = routeEventsInflight.get(cacheKey);
  if (inFlight) {
    return { entry: await inFlight, fromCache: true };
  }

  const promise = fetchRouteEventsFromUpstream(routeId, options)
    .then((entry) => {
      routeEventsCache.set(cacheKey, entry);
      routeEventsInflight.delete(cacheKey);
      return entry;
    })
    .catch((error) => {
      routeEventsInflight.delete(cacheKey);
      throw error;
    });

  routeEventsInflight.set(cacheKey, promise);
  const entry = await promise;
  return { entry, fromCache: false };
};

const runWithConcurrency = async <TItem, TResult>(
  items: TItem[],
  concurrency: number,
  mapper: (item: TItem, index: number) => Promise<TResult>
): Promise<TResult[]> => {
  if (items.length === 0) return [];
  const safeConcurrency = Math.max(1, Math.min(concurrency, items.length));
  const results = new Array<TResult>(items.length);
  let cursor = 0;

  const workers = Array.from({ length: safeConcurrency }, async () => {
    while (true) {
      const index = cursor;
      cursor += 1;
      if (index >= items.length) break;
      results[index] = await mapper(items[index], index);
    }
  });

  await Promise.all(workers);
  return results;
};

const toUniqueRouteIds = (body: RouteWebRouteEventsBody): number[] => {
  const list = Array.isArray(body.routeIds) ? body.routeIds : [];
  const candidate = list
    .map((value) => toOptionalInt(value))
    .filter((id): id is number => id != null);
  const single = toOptionalInt(body.routeId);
  if (single != null) candidate.push(single);
  return Array.from(new Set(candidate));
};

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const body = (req.body || {}) as RouteWebRouteEventsBody;
    const routeIds = toUniqueRouteIds(body);
    if (routeIds.length === 0) {
      return res.status(400).json({ success: false, error: 'routeId/routeIds é obrigatório' });
    }

    const withOccurrences = toBoolean(body.withOccurrences, true);
    const compact = toBoolean(body.compact, false);
    const nonCollectionOnly = toBoolean(body.nonCollectionOnly, false);
    const timeoutMs = clampInt(toOptionalInt(body.timeoutMs), DEFAULT_TIMEOUT_MS, 4000, 30000);
    const retryTimeoutMs = clampInt(toOptionalInt(body.retryTimeoutMs), DEFAULT_RETRY_TIMEOUT_MS, 5000, 40000);
    const retryDelayMs = clampInt(toOptionalInt(body.retryDelayMs), DEFAULT_RETRY_DELAY_MS, 0, 2000);
    const concurrency = clampInt(toOptionalInt(body.concurrency), 2, 1, 6);

    const manualToken = String(body.bearerToken || '').trim();
    const tokenResult = manualToken ? null : await requestRouteWebToken();
    const bearerToken = manualToken || String(tokenResult?.token || '').trim();
    const tokenSource = manualToken ? 'manual' : 'oauth';

    if (!bearerToken) {
      throw new Error('Token de acesso não disponível para consulta de eventos');
    }

    const routesEnvDebug = getRouteWebRoutesEnvDebug();

    if (routeIds.length === 1) {
      const routeId = routeIds[0];
      const { entry, fromCache } = await fetchRouteEventsWithCache(routeId, {
        withOccurrences,
        nonCollectionOnly,
        bearerToken,
        timeoutMs,
        retryTimeoutMs,
        retryDelayMs
      });

      const payload: any = {
        success: entry.ok,
        error: entry.error,
        routeId,
        upstreamStatus: entry.upstreamStatus,
        upstreamStatusText: entry.upstreamStatusText,
        eventsEndpointUrl: getRouteWebRouteEventsEndpointUrl(routeId),
        routesEnvDebug,
        requestQuery: entry.requestQuery,
        upstreamUrl: entry.upstreamUrl,
        upstreamSnippet: entry.upstreamSnippet,
        tokenSource,
        tokenPreview: getTokenPreview(bearerToken),
        tokenField: tokenResult?.tokenField || 'manual',
        tokenFormat: tokenResult?.format || 'manual',
        contentType: entry.contentType,
        nonCollectionOnly,
        count: entry.events.length,
        fromCache,
        events: compact ? entry.events.map(compactEvent) : entry.events
      };

      return res.status(200).json(payload);
    }

    const startedAt = Date.now();
    const entries = await runWithConcurrency(routeIds, concurrency, async (routeId) => {
      const { entry, fromCache } = await fetchRouteEventsWithCache(routeId, {
        withOccurrences,
        nonCollectionOnly,
        bearerToken,
        timeoutMs,
        retryTimeoutMs,
        retryDelayMs
      });
      return {
        routeId,
        fromCache,
        entry
      };
    });

    const results = entries.map((item) => ({
      routeId: item.routeId,
      success: item.entry.ok,
      error: item.entry.error,
      upstreamStatus: item.entry.upstreamStatus,
      upstreamStatusText: item.entry.upstreamStatusText,
      requestQuery: item.entry.requestQuery,
      upstreamUrl: item.entry.upstreamUrl,
      upstreamSnippet: item.entry.upstreamSnippet,
      contentType: item.entry.contentType,
      count: item.entry.events.length,
      fromCache: item.fromCache,
      events: compact ? item.entry.events.map(compactEvent) : item.entry.events
    }));

    const successCount = results.filter((item) => item.success).length;
    const failedCount = results.length - successCount;

    console.log('[ROUTE_WEB_ROUTE_EVENTS] Batch summary:', {
      routeIds: routeIds.length,
      successCount,
      failedCount,
      fromCache: results.filter((item) => item.fromCache).length,
      durationMs: Date.now() - startedAt,
      nonCollectionOnly,
      concurrency
    });

    return res.status(200).json({
      success: failedCount === 0,
      batch: true,
      totalRoutes: routeIds.length,
      successCount,
      failedCount,
      tokenSource,
      tokenPreview: getTokenPreview(bearerToken),
      tokenField: tokenResult?.tokenField || 'manual',
      tokenFormat: tokenResult?.format || 'manual',
      nonCollectionOnly,
      withOccurrences,
      compact,
      concurrency,
      timeoutMs,
      retryTimeoutMs,
      retryDelayMs,
      routesEnvDebug,
      durationMs: Date.now() - startedAt,
      results
    });
  } catch (error: any) {
    console.error('[ROUTE_WEB_ROUTE_EVENTS] Erro ao consultar events:', error?.message || error);
    return res.status(500).json({
      success: false,
      error: error?.message || 'Erro ao consultar events da rota'
    });
  }
}

