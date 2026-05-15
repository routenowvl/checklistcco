import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';
import { VitePWA } from 'vite-plugin-pwa';
import {
  appendQueryToUrl,
  getRouteWebRouteEventsEndpointUrl,
  getRouteWebRoutesEndpointUrl,
  getRouteWebRoutesEnvDebug,
  getRouteWebTokenUrl,
  getRouteWebUpstreamUrl,
  getTokenPreview,
  requestRouteWebToken
} from './utils/routeWebServer';

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

const pickPlantsArray = (payload: any): any[] => {
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload?.plants)) return payload.plants;
  if (Array.isArray(payload?.items)) return payload.items;
  if (Array.isArray(payload?.results)) return payload.results;
  if (Array.isArray(payload?.data?.plants)) return payload.data.plants;
  return [];
};

const pickRoutesArray = (payload: any): any[] => {
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload?.routes)) return payload.routes;
  if (Array.isArray(payload?.items)) return payload.items;
  if (Array.isArray(payload?.results)) return payload.results;
  if (Array.isArray(payload?.data?.routes)) return payload.data.routes;
  return [];
};

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
  plant: route?.plant
    ? {
        id: route.plant.id,
        code: route.plant.code,
        display_name: route.plant.display_name,
        name: route.plant.name
      }
    : undefined
});

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

const keepOnlyScraperOccurrences = (event: any): any => ({
  ...event,
  occurrences: Array.isArray(event?.occurrences) ? event.occurrences.filter(isNonCollectionOccurrence) : []
});

const compactOccurrence = (occ: any): any => ({
  id: occ?.id,
  inserted_by: occ?.inserted_by,
  inserted_at: occ?.inserted_at,
  type: occ?.type,
  type_name: occ?.type_name,
  name: occ?.name,
  title: occ?.title,
  occurrence_type_description: occ?.occurrence_type?.description || occ?.occurrence_type_description || '',
  occurrence_type: occ?.occurrence_type
    ? {
        id: occ.occurrence_type.id,
        description: occ.occurrence_type.description
      }
    : undefined
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

const readJsonBody = async (req: any): Promise<any> => {
  const chunks: Buffer[] = [];
  for await (const chunk of req) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(String(chunk)));
  }
  if (chunks.length === 0) return {};
  const raw = Buffer.concat(chunks).toString('utf-8').trim();
  if (!raw) return {};
  try {
    return JSON.parse(raw);
  } catch {
    return {};
  }
};

const writeJson = (res: any, status: number, payload: any) => {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(JSON.stringify(payload));
};

const DEV_EVENTS_SUCCESS_CACHE_TTL_MS = 2 * 60 * 1000;
const DEV_EVENTS_ERROR_CACHE_TTL_MS = 30 * 1000;
const DEV_ROUTES_SUCCESS_CACHE_TTL_MS = 2 * 60 * 1000;
const DEV_ROUTES_ERROR_CACHE_TTL_MS = 30 * 1000;
const devRoutesCache = new Map<string, {
  ok: boolean;
  status: number;
  statusText: string;
  contentType: string;
  raw: string;
  data: any;
  routes: any[];
  upstreamSnippet: string;
  fetchedAt: number;
}>();
const devRoutesInflight = new Map<string, Promise<{
  ok: boolean;
  status: number;
  statusText: string;
  contentType: string;
  raw: string;
  data: any;
  routes: any[];
  upstreamSnippet: string;
  fetchedAt: number;
}>>();
const devRouteEventsCache = new Map<string, {
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
  fetchedAt: number;
}>();
const devRouteEventsInflight = new Map<string, Promise<{
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
  fetchedAt: number;
}>>();

const clampInt = (value: number | null, fallback: number, min: number, max: number): number => {
  const base = value == null || !Number.isFinite(value) ? fallback : value;
  return Math.max(min, Math.min(max, Math.trunc(base)));
};

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

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

const routeWebDevPlugin = (mode: string) => ({
  name: 'route-web-dev-api',
  configureServer(server: any) {
    server.middlewares.use(async (req: any, res: any, next: any) => {
      // Recarrega .env em dev para refletir alterações sem reiniciar o servidor.
      Object.assign(process.env, loadEnv(mode, '.', ''));

      const pathname = String(req.url || '').split('?')[0];

      if (req.method === 'POST' && pathname === '/api/route-web-token') {
        try {
          const tokenResult = await requestRouteWebToken();
          return writeJson(res, 200, {
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
        } catch (error: any) {
          return writeJson(res, 500, {
            success: false,
            error: error?.message || 'Erro ao obter token do Route Web'
          });
        }
      }

      if (req.method === 'POST' && pathname === '/api/route-web-plants') {
        try {
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
          const plants = pickPlantsArray(upstreamPayload.data);

          return writeJson(res, 200, {
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
        } catch (error: any) {
          return writeJson(res, 500, {
            success: false,
            error: error?.message || 'Erro ao consultar plants do Route Web'
          });
        }
      }

      if (req.method === 'POST' && pathname === '/api/route-web-routes') {
        try {
          const body = await readJsonBody(req);

          const plantId = toOptionalInt(body?.plantId);
          if (plantId == null) {
            return writeJson(res, 400, { success: false, error: 'plantId é obrigatório' });
          }

          const perPage = toOptionalInt(body?.perPage) ?? 60;
          const strictDate = toOptionalInt(body?.strictDate) ?? 1;
          const compact = toBoolean(body?.compact, false);
          const initialExpectedStartDate = String(body?.initialExpectedStartDate || '').trim();
          const finalExpectedStartDate = String(body?.finalExpectedStartDate || '').trim();

          if (!initialExpectedStartDate || !finalExpectedStartDate) {
            return writeJson(res, 400, {
              success: false,
              error: 'initialExpectedStartDate e finalExpectedStartDate são obrigatórios'
            });
          }

          const query = new URLSearchParams({
            plant_id: String(plantId),
            per_page: String(perPage),
            initial_expected_start_date: initialExpectedStartDate,
            final_expected_start_date: finalExpectedStartDate,
            strict_date: String(strictDate)
          });

          const manualToken = String(body?.bearerToken || '').trim();
          const tokenResult = manualToken ? null : await requestRouteWebToken();
          const bearerToken = manualToken || String(tokenResult?.token || '').trim();
          const tokenSource = manualToken ? 'manual' : 'oauth';

          if (!bearerToken) {
            return writeJson(res, 500, { success: false, error: 'Token de acesso não disponível para consulta de rotas' });
          }

          const routesEndpointUrl = getRouteWebRoutesEndpointUrl();
          const upstreamUrl = appendQueryToUrl(routesEndpointUrl, query);
          const requestQuery = Object.fromEntries(query.entries());
          const routesEnvDebug = getRouteWebRoutesEnvDebug();

          console.log('[ROUTE_WEB_ROUTES][DEV] Request:', {
            routesEndpointUrl,
            upstreamUrl,
            requestQuery,
            routesEnvDebug
          });

          const cacheKey = `${upstreamUrl}|${compact ? 1 : 0}`;
          const now = Date.now();
          let fromCache = false;
          let entry:
            | {
                ok: boolean;
                status: number;
                statusText: string;
                contentType: string;
                raw: string;
                data: any;
                routes: any[];
                upstreamSnippet: string;
                fetchedAt: number;
              }
            | undefined;

          const cached = devRoutesCache.get(cacheKey);
          if (cached) {
            const maxAge = cached.ok ? DEV_ROUTES_SUCCESS_CACHE_TTL_MS : DEV_ROUTES_ERROR_CACHE_TTL_MS;
            if (now - cached.fetchedAt <= maxAge) {
              entry = cached;
              fromCache = true;
            }
          }

          if (!entry) {
            const inFlight = devRoutesInflight.get(cacheKey);
            if (inFlight) {
              entry = await inFlight;
              fromCache = true;
            } else {
              const promise = (async () => {
                const response = await fetch(upstreamUrl, {
                  method: 'GET',
                  headers: {
                    Authorization: `Bearer ${bearerToken}`,
                    'Content-Type': 'application/json',
                    'X-Requested-With': 'XMLHttpRequest',
                    'x-requested_with': 'XLMHttpRequest'
                  }
                });

                const upstreamPayload = await parseUpstreamResponse(response);
                const routes = pickRoutesArray(upstreamPayload.data);
                const upstreamSnippet = String(upstreamPayload.raw || '').slice(0, 600);

                const built = {
                  ok: response.ok,
                  status: response.status,
                  statusText: response.statusText,
                  contentType: upstreamPayload.contentType,
                  raw: upstreamPayload.raw,
                  data: upstreamPayload.data,
                  routes,
                  upstreamSnippet,
                  fetchedAt: Date.now()
                };
                devRoutesCache.set(cacheKey, built);
                return built;
              })().finally(() => {
                devRoutesInflight.delete(cacheKey);
              });

              devRoutesInflight.set(cacheKey, promise);
              entry = await promise;
            }
          }

          const routes = entry.routes;
          const upstreamSnippet = entry.upstreamSnippet;

          console.log('[ROUTE_WEB_ROUTES][DEV] Response:', {
            upstreamUrl,
            status: entry.status,
            statusText: entry.statusText,
            contentType: entry.contentType,
            routesCount: routes.length,
            fromCache,
            snippet: upstreamSnippet
          });

          const errorMessage = entry.ok
            ? undefined
            : `Upstream ${entry.status} ${entry.statusText} em ${upstreamUrl}${upstreamSnippet ? ` | ${upstreamSnippet}` : ''}`;

          const payload: any = {
            success: entry.ok,
            error: errorMessage,
            upstreamStatus: entry.status,
            upstreamStatusText: entry.statusText,
            routesEndpointUrl,
            routesEnvDebug,
            requestQuery,
            upstreamUrl,
            upstreamSnippet,
            fromCache,
            tokenSource,
            tokenPreview: getTokenPreview(bearerToken),
            tokenField: tokenResult?.tokenField || 'manual',
            tokenFormat: tokenResult?.format || 'manual',
            contentType: entry.contentType,
            count: routes.length,
            routes: compact ? routes.map(compactRoute) : routes
          };

          if (!compact) {
            payload.response = entry.data;
            payload.raw = entry.raw;
          }

          return writeJson(res, 200, payload);
        } catch (error: any) {
          return writeJson(res, 500, {
            success: false,
            error: error?.message || 'Erro ao consultar routes do Route Web'
          });
        }
      }

      if (req.method === 'POST' && pathname === '/api/route-web-route-events') {
        try {
          const body = await readJsonBody(req);
          const routeIdsRaw = Array.isArray(body?.routeIds) ? body.routeIds : [];
          const singleRouteId = toOptionalInt(body?.routeId);
          const routeIds = Array.from(
            new Set(
              [...routeIdsRaw, singleRouteId]
                .map((value) => toOptionalInt(value))
                .filter((id): id is number => id != null)
            )
          );

          if (routeIds.length === 0) {
            return writeJson(res, 400, { success: false, error: 'routeId/routeIds é obrigatório' });
          }

          const withOccurrences = toBoolean(body?.withOccurrences, true);
          const compact = toBoolean(body?.compact, false);
          const nonCollectionOnly = toBoolean(body?.nonCollectionOnly, false);
          const timeoutMs = clampInt(toOptionalInt(body?.timeoutMs), 9000, 4000, 30000);
          const retryTimeoutMs = clampInt(toOptionalInt(body?.retryTimeoutMs), 14000, 5000, 40000);
          const retryDelayMs = clampInt(toOptionalInt(body?.retryDelayMs), 120, 0, 2000);
          const concurrency = clampInt(toOptionalInt(body?.concurrency), 2, 1, 6);

          const manualToken = String(body?.bearerToken || '').trim();
          const tokenResult = manualToken ? null : await requestRouteWebToken();
          const bearerToken = manualToken || String(tokenResult?.token || '').trim();
          const tokenSource = manualToken ? 'manual' : 'oauth';
          const routesEnvDebug = getRouteWebRoutesEnvDebug();

          if (!bearerToken) {
            return writeJson(res, 500, { success: false, error: 'Token de acesso não disponível para consulta de eventos' });
          }

          const fetchRouteEvents = async (routeId: number): Promise<{ entry: any; fromCache: boolean }> => {
            const cacheKey = `${routeId}|${withOccurrences ? 1 : 0}|${nonCollectionOnly ? 1 : 0}`;
            const now = Date.now();
            const cached = devRouteEventsCache.get(cacheKey);
            if (cached) {
              const maxAge = cached.ok ? DEV_EVENTS_SUCCESS_CACHE_TTL_MS : DEV_EVENTS_ERROR_CACHE_TTL_MS;
              if (now - cached.fetchedAt <= maxAge) {
                return { entry: cached, fromCache: true };
              }
            }

            const inFlight = devRouteEventsInflight.get(cacheKey);
            if (inFlight) {
              return { entry: await inFlight, fromCache: true };
            }

            const query = new URLSearchParams({
              with_occurrences: withOccurrences ? 'true' : 'false'
            });
            const requestQuery = Object.fromEntries(query.entries());
            const eventsEndpointUrl = getRouteWebRouteEventsEndpointUrl(routeId);
            const upstreamUrl = appendQueryToUrl(eventsEndpointUrl, query);

            const callUpstream = async (timeout: number): Promise<any> => {
              const controller = new AbortController();
              const timer = setTimeout(() => controller.abort(), timeout);
              try {
                const response = await fetch(upstreamUrl, {
                  method: 'GET',
                  headers: {
                    Authorization: `Bearer ${bearerToken}`,
                    'Content-Type': 'application/json',
                    'X-Requested-With': 'XMLHttpRequest',
                    'x-requested_with': 'XLMHttpRequest'
                  },
                  signal: controller.signal
                });

                const upstreamPayload = await parseUpstreamResponse(response);
                const events = pickEventsArray(upstreamPayload.data);
                const filteredEvents = nonCollectionOnly ? filterNonCollectionEvents(events).map(keepOnlyScraperOccurrences) : events;
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
                  error: isTimeout ? `Timeout ao consultar eventos da rota ${routeId}` : error?.message || `Erro ao consultar eventos da rota ${routeId}`
                };
              } finally {
                clearTimeout(timer);
              }
            };

            const promise = (async () => {
              let entry = await callUpstream(timeoutMs);
              if (!entry.ok && /timeout/i.test(String(entry.error || ''))) {
                await sleep(retryDelayMs);
                const retry = await callUpstream(retryTimeoutMs);
                entry = retry.ok ? retry : { ...retry, error: `${retry.error} (retry 1/1 também falhou)` };
              }
              const withTimestamp = { ...entry, fetchedAt: Date.now() };
              devRouteEventsCache.set(cacheKey, withTimestamp);
              return withTimestamp;
            })()
              .finally(() => {
                devRouteEventsInflight.delete(cacheKey);
              });

            devRouteEventsInflight.set(cacheKey, promise);
            return { entry: await promise, fromCache: false };
          };

          if (routeIds.length === 1) {
            const routeId = routeIds[0];
            const { entry, fromCache } = await fetchRouteEvents(routeId);
            return writeJson(res, 200, {
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
            });
          }

          const startedAt = Date.now();
          const batchEntries = await runWithConcurrency(routeIds, concurrency, async (routeId) => {
            const { entry, fromCache } = await fetchRouteEvents(routeId);
            return { routeId, entry, fromCache };
          });

          const results = batchEntries.map((item) => ({
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

          console.log('[ROUTE_WEB_ROUTE_EVENTS][DEV] Batch summary:', {
            routeIds: routeIds.length,
            successCount,
            failedCount,
            fromCache: results.filter((item) => item.fromCache).length,
            durationMs: Date.now() - startedAt,
            nonCollectionOnly,
            concurrency
          });

          return writeJson(res, 200, {
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
          return writeJson(res, 500, {
            success: false,
            error: error?.message || 'Erro ao consultar events da rota'
          });
        }
      }

      next();
    });
  }
});

export default defineConfig(({ mode }) => {
    const env = loadEnv(mode, '.', '');
    Object.assign(process.env, env);
    return {
      server: {
        port: Number(env.VITE_SERVER_PORT) || 3000,
        host: '0.0.0.0',
      },
      plugins: [
        react(),
        routeWebDevPlugin(mode),
        VitePWA({
          registerType: 'autoUpdate',
          includeAssets: ['favicon.ico', 'apple-touch-icon.png', 'masked-icon.svg'],
          manifest: {
            name: 'Checklist Web',
            short_name: 'Checklist',
            description: 'Gestão de Operações em Tempo Real',
            theme_color: '#020617',
            background_color: '#020617',
            display: 'standalone',
            orientation: 'landscape',
            scope: '/',
            start_url: '/',
            icons: [
              {
                src: 'pwa-192x192.png',
                sizes: '192x192',
                type: 'image/png'
              },
              {
                src: 'pwa-512x512.png',
                sizes: '512x512',
                type: 'image/png'
              },
              {
                src: 'pwa-512x512.png',
                sizes: '512x512',
                type: 'image/png',
                purpose: 'any maskable'
              }
            ]
          },
          workbox: {
            clientsClaim: true,
            skipWaiting: true,
            cleanupOutdatedCaches: true,
            globPatterns: ['**/*.{js,css,html,ico,png,svg,woff2}'],
            runtimeCaching: [
              {
                urlPattern: /^https:\/\/fonts\.googleapis\.com\/.*/i,
                handler: 'CacheFirst',
                options: {
                  cacheName: 'google-fonts-cache',
                  expiration: {
                    maxEntries: 10,
                    maxAgeSeconds: 60 * 60 * 24 * 365 // 1 ano
                  },
                  cacheableResponse: {
                    statuses: [0, 200]
                  }
                }
              },
              {
                urlPattern: /^https:\/\/fonts\.gstatic\.com\/.*/i,
                handler: 'CacheFirst',
                options: {
                  cacheName: 'gstatic-fonts-cache',
                  expiration: {
                    maxEntries: 10,
                    maxAgeSeconds: 60 * 60 * 24 * 365 // 1 ano
                  },
                  cacheableResponse: {
                    statuses: [0, 200]
                  }
                }
              },
              {
                urlPattern: /^https:\/\/cdn\.tailwindcss\.com/i,
                handler: 'NetworkFirst',
                options: {
                  cacheName: 'tailwind-cache',
                  expiration: {
                    maxEntries: 3,
                    maxAgeSeconds: 60 * 60 * 24 // 1 dia
                  },
                  cacheableResponse: {
                    statuses: [0, 200]
                  }
                }
              }
            ]
          }
        })
      ],
      define: {
        'process.env.LIMITAR_RETRY_LOGIN': JSON.stringify(env.LIMITAR_RETRY_LOGIN || '5'),
        'process.env.LOGIN_LOCKOUT_MINUTES': JSON.stringify(env.LOGIN_LOCKOUT_MINUTES || '15'),
        'process.env.SITE_KEY': JSON.stringify(env.SITE_KEY || '')
      }
    };
});

