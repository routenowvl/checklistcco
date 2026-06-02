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
import { getGraphAppToken } from './utils/graphAppAuth';
import { getShiftApiBaseUrl, getShiftToken, getTokenPreview as getShiftTokenPreview } from './utils/shiftApi';
import {
  getRouteWebEventsByDateAndPlants,
  getRouteWebRoutesByDateAndPlants,
  type RouteWebEventDbRow,
  type RouteWebRouteDbRow
} from './utils/rweDb';
import {
  getAllConfigs, getConfigByOperacao, updateConfigField, updateConfigFields,
  updateConteudoIfChanged, updateConteudoNcoletasIfChanged,
  getLockStatus, acquireLock, releaseLock,
  getDepartures, upsertDeparture, deleteDeparture,
  getNonCollections, insertNonCollection, updateNonCollection, deleteNonCollection,
  insertConfig, fixNonCollectionsRoutes
} from './utils/checklistDb';
import { queryMaintenanceEvents } from './utils/maintenanceDb';

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

      if (req.method === 'POST' && pathname === '/api/route-web') {
        const _rwBody = await readJsonBody(req);
        const _rwResource = String(_rwBody?.resource || '').trim();
        if (_rwResource === 'token') {
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

      if (_rwResource === 'plants') {
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

      if (_rwResource === 'routes') {
        try {
          const body = _rwBody;

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

      if (_rwResource === 'route-events') {
        try {
          const body = _rwBody;
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
      } // end /api/route-web

      if ((req.method === 'POST' || req.method === 'GET') && pathname === '/api/route-web-db') {
        const _dbBody = req.method === 'POST' ? await readJsonBody(req) : Object.fromEntries(new URL(String(req.url || ''), 'http://localhost').searchParams);
        const _dbEntity = String(_dbBody?.entity || '').trim();
        if (_dbEntity === 'events') {
        try {
          const body = _dbBody;
          const dataReferencia = String(body?.dataReferencia || '').trim();

          if (!dataReferencia || !/^\d{4}-\d{2}-\d{2}$/.test(dataReferencia)) {
            return writeJson(res, 400, { success: false, error: 'dataReferencia inválida (YYYY-MM-DD)' });
          }

          let plantIds: number[] = [];
          const plantIdsRaw = body?.plantIds;
          if (Array.isArray(plantIdsRaw)) {
            plantIds = plantIdsRaw.map(Number).filter(Number.isFinite);
          } else if (plantIdsRaw != null) {
            const parsed = Number(plantIdsRaw);
            if (Number.isFinite(parsed)) plantIds = [parsed];
          }

          const rows: RouteWebEventDbRow[] = await getRouteWebEventsByDateAndPlants(dataReferencia, plantIds);

          return writeJson(res, 200, {
            success: true,
            dataReferencia,
            plantIds,
            count: rows.length,
            events: rows
          });
        } catch (error: any) {
          console.error('[ROUTE_WEB_EVENTS][DEV] Erro:', error?.message || error);
          return writeJson(res, 500, {
            success: false,
            error: error?.message || 'Erro ao consultar eventos'
          });
        }
      }

      if (_dbEntity === 'routes') {
        try {
          const body = _dbBody;
          const dataReferencia = String(body?.dataReferencia || '').trim();

          if (!dataReferencia || !/^\d{4}-\d{2}-\d{2}$/.test(dataReferencia)) {
            return writeJson(res, 400, { success: false, error: 'dataReferencia inválida (YYYY-MM-DD)' });
          }

          let plantIds: number[] = [];
          const plantIdsRaw = body?.plantIds;
          if (Array.isArray(plantIdsRaw)) {
            plantIds = plantIdsRaw.map(Number).filter(Number.isFinite);
          } else if (plantIdsRaw != null) {
            const parsed = Number(plantIdsRaw);
            if (Number.isFinite(parsed)) plantIds = [parsed];
          }

          const rows: RouteWebRouteDbRow[] = await getRouteWebRoutesByDateAndPlants(dataReferencia, plantIds);

          return writeJson(res, 200, {
            success: true,
            dataReferencia,
            plantIds,
            count: rows.length,
            routes: rows
          });
        } catch (error: any) {
          console.error('[ROUTE_WEB_ROUTES_DB][DEV] Erro:', error?.message || error);
          return writeJson(res, 500, {
            success: false,
            error: error?.message || 'Erro ao consultar rotas'
          });
        }
      }
      } // end /api/route-web-db

      // Checklist API
      if (req.method === 'POST' && pathname === '/api/checklist') {
        const _clBody = await readJsonBody(req);
        const _clDomain = String(_clBody?.domain || '').trim();
        if (_clDomain === 'config') {
        try {
          const body = _clBody;
          const { action } = body;

          switch (action) {
            case 'getAll': {
              const configs = await getAllConfigs();
              return writeJson(res, 200, { success: true, configs });
            }
            case 'getByOperacao': {
              const config = await getConfigByOperacao(String(body.operacao || ''));
              return writeJson(res, 200, { success: true, config });
            }
            case 'updateField': {
              await updateConfigField(String(body.operacao || ''), String(body.field || ''), body.value);
              return writeJson(res, 200, { success: true });
            }
            case 'updateFields': {
              await updateConfigFields(String(body.operacao || ''), body.fields || {});
              return writeJson(res, 200, { success: true });
            }
            case 'updateConteudoIfChanged': {
              const changed = await updateConteudoIfChanged(String(body.operacao || ''), String(body.conteudo || ''));
              return writeJson(res, 200, { success: true, changed });
            }
            case 'updateConteudoNcoletasIfChanged': {
              const changed = await updateConteudoNcoletasIfChanged(String(body.operacao || ''), String(body.conteudoNcoletas || ''));
              return writeJson(res, 200, { success: true, changed });
            }
            case 'getLockStatus': {
              const lockStatus = await getLockStatus(String(body.operacao || ''));
              return writeJson(res, 200, { success: true, lock: lockStatus });
            }
            case 'acquireLock': {
              await acquireLock(String(body.operacao || ''), String(body.userEmail || ''), String(body.timestamp || ''));
              return writeJson(res, 200, { success: true });
            }
            case 'releaseLock': {
              await releaseLock(String(body.operacao || ''));
              return writeJson(res, 200, { success: true });
            }
            default:
              return writeJson(res, 400, { success: false, error: `Ação desconhecida: ${action}` });
          }
        } catch (error: any) {
          console.error('[CHECKLIST_CONFIG][DEV] Erro:', error?.message || error);
          return writeJson(res, 500, { success: false, error: 'Erro ao processar config' });
        }
      }

      // Checklist API — departures
      if (_clDomain === 'departures') {
        try {
          const body = _clBody;
          const { action } = body;

          switch (action) {
            case 'getAll': {
              const departures = await getDepartures();
              return writeJson(res, 200, { success: true, departures });
            }
            case 'upsert': {
              const id = await upsertDeparture(body.departure || {});
              return writeJson(res, 200, { success: true, id });
            }
            case 'delete': {
              await deleteDeparture(Number(body.id));
              return writeJson(res, 200, { success: true });
            }
            default:
              return writeJson(res, 400, { success: false, error: `Ação desconhecida: ${action}` });
          }
        } catch (error: any) {
          const detail = error?.detail || '';
          const hint = error?.hint || '';
          const code = error?.code || '';
          const schema = error?.schema || '';
          const table = error?.table || '';
          const column = error?.column || '';
          console.error('[CHECKLIST_DEPARTURES][DEV] Erro:', error?.message || error, { detail, hint, code, schema, table, column });
          return writeJson(res, 500, {
            success: false,
            error: error?.message || 'Erro ao processar departures',
            detail,
            hint,
            code,
            schema,
            table,
            column
          });
        }
      }

      // Checklist API — non-collections
      if (_clDomain === 'non-collections') {
        try {
          const body = _clBody;
          const { action } = body;

          switch (action) {
            case 'getAll': {
              const nonCollections = await getNonCollections();
              return writeJson(res, 200, { success: true, nonCollections });
            }
            case 'insert': {
              const id = await insertNonCollection(body.nonCollection || {});
              return writeJson(res, 200, { success: true, id });
            }
            case 'update': {
              await updateNonCollection(body.nonCollection || {});
              return writeJson(res, 200, { success: true });
            }
            case 'delete': {
              await deleteNonCollection(Number(body.id));
              return writeJson(res, 200, { success: true });
            }
            default:
              return writeJson(res, 400, { success: false, error: `Ação desconhecida: ${action}` });
          }
        } catch (error: any) {
          console.error('[CHECKLIST_NON_COLLECTIONS][DEV] Erro:', error?.message || error, error?.detail || '', error?.hint || '');
          return writeJson(res, 500, { success: false, error: `Erro ao processar non-collections: ${error?.message || error}` });
        }
      }

      // Migration: SharePoint → PostgreSQL (TEMPORÁRIO)
      const SITE_PATH = process.env.VITE_SHAREPOINT_SITE_PATH || '';
      const graphFetch = async (endpoint: string, token: string): Promise<any> => {
        const url = endpoint.startsWith('https://') ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`;
        const r = await fetch(url, { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } });
        if (!r.ok) { const t = await r.text(); throw new Error(`Graph API ${r.status}: ${t.slice(0, 400)}`); }
        return r.status === 204 ? null : r.json();
      };
      const normalizeStr = (str: string): string => str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '').trim();
      const resolveField = (mapping: Record<string, string>, target: string): string => mapping[normalizeStr(target)] || target;
      const formatISOtoBR = (iso: any): string => { if (!iso) return ''; const s = String(iso).trim(); const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : s; };
      const brDatetimeToISO = (v: any): string | null => { if (!v) return null; const s = String(v).trim(); const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}:\d{2}:\d{2})$/); if (m) return `${m[3]}-${m[2]}-${m[1]}T${m[4]}`; if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s; return null; };
      const getColumnMapping = async (siteId: string, listId: string, token: string): Promise<Record<string, string>> => {
        const columns = await graphFetch(`/sites/${siteId}/lists/${listId}/columns`, token);
        const mapping: Record<string, string> = {};
        for (const col of columns.value || []) { mapping[normalizeStr(col.name)] = col.name; mapping[normalizeStr(col.displayName)] = col.name; }
        return mapping;
      };
      const formatTime = (v: any): string => { if (!v) return ''; const s = String(v).trim(); if (s === '-') return ''; const brMatch = s.match(/(\d{2}:\d{2}):\d{2}$/); if (brMatch) return brMatch[1] + ':00'; const dtMatch = s.match(/T(\d{2}:\d{2})/); if (dtMatch) return dtMatch[1] + ':00'; const tMatch = s.match(/^(\d{2}:\d{2})/); return tMatch ? tMatch[1] + ':00' : ''; };
      const parseNumericId = (value: unknown): number | null => { if (value == null) return null; if (typeof value === 'number' && Number.isFinite(value)) return Math.trunc(value); const raw = String(value).trim(); if (!raw) return null; const match = raw.match(/-?\d+(?:[.,]\d+)?/); if (!match) return null; const parsed = Number(match[0].replace(',', '.')); return Number.isFinite(parsed) ? Math.trunc(parsed) : null; };
      const extractPlantId = (fields: Record<string, any>, mapping: Record<string, string>): any => {
        for (const c of ['Plant_id', 'Plant Id', 'PlantId', 'plant_id', 'IdPlant']) { const r = resolveField(mapping, c); if (fields?.[r] != null && String(fields[r]).trim() !== '') return fields[r]; }
        for (const [key, value] of Object.entries(fields || {})) { const nk = normalizeStr(key); if ((nk.includes('plantid') || nk.includes('idplant')) && value != null && String(value).trim() !== '') return value; }
        return null;
      };
      const parseUltimoEnvioNcoletas = (raw: any): { datetime: string | null; quantidade: number } => {
        if (!raw) return { datetime: null, quantidade: 0 }; const s = String(raw).trim(); const match = s.match(/^(.+?\d{2}:\d{2}:\d{2})\s+(\d+)$/);
        if (match) return { datetime: match[1].trim(), quantidade: parseInt(match[2], 10) || 0 }; return { datetime: s, quantidade: 0 };
      };

      if (_clDomain === 'migrate-config') {
        try {
          const appToken = await getGraphAppToken();
          const siteData = await graphFetch(`/sites/${SITE_PATH}`, appToken);
          const siteId = siteData.id;
          let list: any;
          try { list = await graphFetch(`/sites/${siteId}/lists/CONFIG_OPERACAO_SAIDA_DE_ROTAS`, appToken); } catch {
            const listsData = await graphFetch(`/sites/${siteId}/lists`, appToken);
            list = (listsData.value || []).find((l: any) => l.name?.toLowerCase() === 'config_operacao_saida_de_rotas');
            if (!list) throw new Error('Lista CONFIG_OPERACAO_SAIDA_DE_ROTAS não encontrada');
          }
          const mapping = await getColumnMapping(siteId, list.id, appToken);
          let allItems: any[] = [], nextUrl: string | null = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=100`;
          while (nextUrl) { const d = await graphFetch(nextUrl, appToken); allItems = allItems.concat(d.value || []); nextUrl = d['@odata.nextLink'] || null; }
          let upserted = 0, skipped = 0; const errors: string[] = [];
          for (const item of allItems) {
            const f = item.fields || {}; const operacao = String(f[resolveField(mapping, 'OPERACAO')] || f.Title || '').trim();
            if (!operacao) { skipped++; continue; }
            try {
              const ncoleta = parseUltimoEnvioNcoletas(f[resolveField(mapping, 'UltimoEnvioNcoleta')]);
              await insertConfig({ operacao, email: String(f[resolveField(mapping, 'EMAIL')] || '').toLowerCase().trim(), tolerancia: String(f[resolveField(mapping, 'TOLERANCIA')] || '00:00:00'), nome_exibicao: String(f[resolveField(mapping, 'NomeExibicao')] || operacao), plant_id: parseNumericId(extractPlantId(f, mapping)), ultimo_envio_saida: brDatetimeToISO(f[resolveField(mapping, 'UltimoEnvioSaida')]), status: String(f[resolveField(mapping, 'Status')] || ''), envio: String(f[resolveField(mapping, 'Envio')] || ''), copia: String(f[resolveField(mapping, 'Copia')] || ''), ultimo_envio_resumo_saida: brDatetimeToISO(f[resolveField(mapping, 'UltimoEnvioResumoSaida')]), status_resumo_saida: String(f[resolveField(mapping, 'StatusResumoSaida')] || ''), ultimo_envio_ncoleta: brDatetimeToISO(ncoleta.datetime), quantidade_ncoletas_registrada: ncoleta.quantidade, conteudo: String(f[resolveField(mapping, 'Conteudo')] || ''), conteudo_ncoletas: String(f[resolveField(mapping, 'ConteudoNcoletas')] || ''), lock_envio: f[resolveField(mapping, 'LockEnvio')] || null, lock_user: String(f[resolveField(mapping, 'LockUser')] || ''), lock_timestamp: brDatetimeToISO(f[resolveField(mapping, 'LockTimestamp')]) });
              upserted++;
            } catch (err: any) { errors.push(`${operacao}: ${err.message}`); }
          }
          return writeJson(res, 200, { success: true, total: allItems.length, upserted, skipped, errors: errors.length > 0 ? errors.slice(0, 20) : undefined });
        } catch (error: any) { return writeJson(res, 500, { success: false, error: error?.message || 'Erro migrate-config' }); }
      }

      if (_clDomain === 'migrate-departures') {
        try {
          const appToken = await getGraphAppToken();
          const siteData = await graphFetch(`/sites/${SITE_PATH}`, appToken);
          const siteId = siteData.id;
          let list: any;
          try { list = await graphFetch(`/sites/${siteId}/lists/Dados_Saida_de_rotas`, appToken); } catch {
            const listsData = await graphFetch(`/sites/${siteId}/lists`, appToken);
            list = (listsData.value || []).find((l: any) => l.name?.toLowerCase() === 'dados_saida_de_rotas');
            if (!list) throw new Error('Lista Dados_Saida_de_rotas não encontrada');
          }
          const mapping = await getColumnMapping(siteId, list.id, appToken);
          let allItems: any[] = [], nextUrl: string | null = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=100`;
          while (nextUrl) { const d = await graphFetch(nextUrl, appToken); allItems = allItems.concat(d.value || []); nextUrl = d['@odata.nextLink'] || null; }
          let upserted = 0; const errors: string[] = [];
          for (const item of allItems) {
            const f = item.fields || {};
            try {
              const dataBR = formatISOtoBR(f[resolveField(mapping, 'DataOperacao')]);
              if (!dataBR) { errors.push(`Item ${item.id}: sem data`); continue; }
              await upsertDeparture({ operacao: String(f[resolveField(mapping, 'Operacao')] || '').trim(), rota: String(f.Title || '').trim(), motorista: String(f[resolveField(mapping, 'Motorista')] || '').trim(), placa: String(f[resolveField(mapping, 'Placa')] || '').trim(), contato: String(f[resolveField(mapping, 'Contato')] || '').replace(/\D/g, ''), inicio: formatTime(f[resolveField(mapping, 'HorarioInicio')]), saida: formatTime(f[resolveField(mapping, 'HorarioSaida')]), statusGeral: String(f[resolveField(mapping, 'StatusGeral')] || '').trim(), motivo: String(f[resolveField(mapping, 'MotivoAtraso')] || '').trim(), observacao: String(f[resolveField(mapping, 'Observacao')] || '').trim(), data: dataBR, statusOp: String(f[resolveField(mapping, 'StatusOp')] || 'Previsto').trim(), checklistMotorista: String(f[resolveField(mapping, 'ChecklistMotorista')] || '').trim(), retornoMotorista: String(f[resolveField(mapping, 'RetornoMotorista')] || '').trim(), causaRaiz: String(f[resolveField(mapping, 'CausaRaiz')] || '').trim(), tempoResposta: String(f[resolveField(mapping, 'TempoResposta')] || '').trim(), logTempoResposta: String(f[resolveField(mapping, 'LogTempoResposta')] || '').trim() });
              upserted++;
            } catch (err: any) { errors.push(`Item ${item.id}: ${err.message}`); }
          }
          return writeJson(res, 200, { success: true, total: allItems.length, upserted, errors: errors.length > 0 ? errors.slice(0, 20) : undefined });
        } catch (error: any) { return writeJson(res, 500, { success: false, error: error?.message || 'Erro migrate-departures' }); }
      }

      if (_clDomain === 'migrate-non-collections') {
        try {
          const LIST_ID = '83e8cfb9-1982-47ae-b515-3fec112da457';
          const appToken = await getGraphAppToken();
          const siteData = await graphFetch(`/sites/${SITE_PATH}`, appToken);
          const siteId = siteData.id;
          const mapping = await getColumnMapping(siteId, LIST_ID, appToken);
          let allItems: any[] = [], nextUrl: string | null = `/sites/${siteId}/lists/${LIST_ID}/items?expand=fields&$top=100`;
          while (nextUrl) { const d = await graphFetch(nextUrl, appToken); allItems = allItems.concat(d.value || []); nextUrl = d['@odata.nextLink'] || null; }
          let upserted = 0; const errors: string[] = [];
          for (const item of allItems) {
            const f = item.fields || {};
            try {
              await insertNonCollection({ operacao: String(f[resolveField(mapping, 'Operacao')] || '').trim(), data: formatISOtoBR(f[resolveField(mapping, 'DataOperacao')]), rota: String(f.Title || '').trim(), observacao: String(f[resolveField(mapping, 'Observacao')] || '').trim(), semana: String(f[resolveField(mapping, 'Semana')] || '').trim(), codigo: String(f[resolveField(mapping, 'Codigo')] || '').trim(), produtor: String(f[resolveField(mapping, 'Produtor')] || '').trim(), motivo: String(f[resolveField(mapping, 'Motivo')] || '').trim(), acao: String(f[resolveField(mapping, 'Acao')] || '').trim(), dataAcao: String(f[resolveField(mapping, 'DataAcao')] || '').trim(), ultimaColeta: String(f[resolveField(mapping, 'UltimaColeta')] || '').trim(), Culpabilidade: String(f[resolveField(mapping, 'Culpabilidade')] || '').trim(), causaRaiz: String(f[resolveField(mapping, 'CausaRaiz')] || '').trim() });
              upserted++;
            } catch (err: any) { errors.push(`Item ${item.id}: ${err.message}`); }
          }
          return writeJson(res, 200, { success: true, total: allItems.length, upserted, errors: errors.length > 0 ? errors.slice(0, 20) : undefined });
        } catch (error: any) { return writeJson(res, 500, { success: false, error: error?.message || 'Erro migrate-non-collections' }); }
      }

      if (_clDomain === 'fix-nc-routes') {
        try {
          const LIST_ID = '83e8cfb9-1982-47ae-b515-3fec112da457';
          const appToken = await getGraphAppToken();
          const siteData = await graphFetch(`/sites/${SITE_PATH}`, appToken);
          const siteId = siteData.id;
          const mapping = await getColumnMapping(siteId, LIST_ID, appToken);
          let allItems: any[] = [], nextUrl: string | null = `/sites/${siteId}/lists/${LIST_ID}/items?expand=fields&$top=100`;
          while (nextUrl) { const d = await graphFetch(nextUrl, appToken); allItems = allItems.concat(d.value || []); nextUrl = d['@odata.nextLink'] || null; }
          const spItems: { operacao: string; codigo: string; rota: string }[] = [];
          for (const item of allItems) {
            const f = item.fields || {};
            const operacao = String(f[resolveField(mapping, 'Operacao')] || '').trim();
            const codigo = String(f[resolveField(mapping, 'Codigo')] || '').trim();
            const rota = String(f[resolveField(mapping, 'Rota')] || '').trim();
            if (operacao && codigo && rota) { spItems.push({ operacao, codigo, rota }); }
          }
          console.log(`[FIX-NC-ROUTES][DEV] ${spItems.length} itens com rota válida (total SP: ${allItems.length})`);
          const result = await fixNonCollectionsRoutes(spItems);
          return writeJson(res, 200, { success: true, spTotal: allItems.length, spWithRota: spItems.length, updated: result.updated, skipped: result.skipped, details: result.details.slice(0, 50) });
        } catch (error: any) { return writeJson(res, 500, { success: false, error: error?.message || 'Erro fix-nc-routes' }); }
      }
      } // end /api/checklist

      // Shift API — Escala de Motoristas (dev proxy)
      if (req.method === 'POST' && pathname === '/api/shift') {
        try {
          const _shiftBody = await readJsonBody(req);
          const _shiftAction = String(_shiftBody?.action || '').trim();
          const _shiftBaseUrl = getShiftApiBaseUrl();

          if (_shiftAction === 'schedules') {
            const _plantId = Number(_shiftBody?.plant_id);
            const _code = String(_shiftBody?.code || '').trim();
            const _perPage = Number(_shiftBody?.per_page) || 100;
            if (!_plantId || !_code) return writeJson(res, 400, { success: false, error: 'plant_id e code são obrigatórios' });

            const _token = await getShiftToken();
            const _upstreamUrl = `${_shiftBaseUrl}/api/schedules?plant_id=${_plantId}&code=${encodeURIComponent(_code)}&per_page=${_perPage}`;
            console.log(`[SHIFT][DEV][SCHEDULES] GET ${_upstreamUrl}`);
            console.log(`[SHIFT][DEV][SCHEDULES] plant_id=${_plantId}, code=${_code}, token=${getShiftTokenPreview(_token)}`);
            const _resp = await fetch(_upstreamUrl, {
              method: 'GET',
              headers: {
                Authorization: `Bearer ${_token}`,
                'Content-Type': 'application/json',
                Accept: 'application/json, text/plain, */*',
                'X-Requested-With': 'XMLHttpRequest',
                'x-requested_with': 'XLMHttpRequest'
              }
            });
            const _rawBody = await _resp.text();
            let _parsed: any;
            try { _parsed = JSON.parse(_rawBody); } catch { _parsed = _rawBody; }
            console.log(`[SHIFT][DEV][SCHEDULES] Response: status=${_resp.status}, body=${String(_rawBody).slice(0, 500)}`);
            return writeJson(res, 200, { success: _resp.ok, upstreamStatus: _resp.status, upstreamUrl: _upstreamUrl, tokenPreview: getShiftTokenPreview(_token), baseUrl: _shiftBaseUrl, plantId: _plantId, code: _code, data: _parsed, raw: typeof _parsed === 'string' ? _parsed : JSON.stringify(_parsed) });
          }

          if (_shiftAction === 'consolidation') {
            const _scheduleId = String(_shiftBody?.schedule_id || '').trim();
            if (!_scheduleId) return writeJson(res, 400, { success: false, error: 'schedule_id é obrigatório' });

            const _token = await getShiftToken();
            const _upstreamUrl = `${_shiftBaseUrl}/api/schedules/${encodeURIComponent(_scheduleId)}/consolidation`;
            console.log(`[SHIFT][DEV][CONSOLIDATION] GET ${_upstreamUrl}, token=${getShiftTokenPreview(_token)}`);
            const _resp = await fetch(_upstreamUrl, {
              method: 'GET',
              headers: {
                Authorization: `Bearer ${_token}`,
                'Content-Type': 'application/json',
                Accept: 'application/json, text/plain, */*',
                'X-Requested-With': 'XMLHttpRequest',
                'x-requested_with': 'XLMHttpRequest'
              }
            });
            const _raw = await _resp.text();
            let _parsed: any;
            try { _parsed = JSON.parse(_raw); } catch { _parsed = _raw; }

            const resultado: any[] = [];
            if (_parsed?.data?.shifts) {
              for (const shift of _parsed.data.shifts) {
                if (!shift?.drivers) continue;
                for (const driver of shift.drivers) {
                  if (!driver?.days) continue;
                  for (const [data, infoDia] of Object.entries(driver.days)) {
                    const dayInfo = infoDia as any;
                    if (dayInfo?.status === 'WORKING' && dayInfo?.routePlan) {
                      resultado.push({
                        motoristaId: driver.id, motorista: driver.name, data,
                        rotaId: dayInfo.routePlan.id, rota: dayInfo.routePlan.code,
                        inicioPrevisto: dayInfo.routePlan.expectedStart,
                        fimPrevisto: dayInfo.routePlan.expectedEnd,
                        operacao: _parsed.data.plantId
                      });
                    }
                  }
                }
              }
            }

            return writeJson(res, 200, { success: _resp.ok, upstreamStatus: _resp.status, upstreamUrl: _upstreamUrl, tokenPreview: getShiftTokenPreview(_token), resultado, data: _parsed, raw: _raw });
          }

          return writeJson(res, 400, { success: false, error: `Action desconhecida: ${_shiftAction}` });
        } catch (error: any) {
          console.error('[SHIFT][DEV] Erro:', error?.message || error);
          return writeJson(res, 500, { success: false, error: error?.message || 'Erro na Shift API' });
        }
      }

      // Maintenance events endpoint (dev proxy)
      if (req.method === 'POST' && pathname === '/api/maintenance-events') {
        try {
          const _maintBody = await readJsonBody(req);
          const events = await queryMaintenanceEvents(_maintBody?.items || []);
          return writeJson(res, 200, { success: true, events });
        } catch (error: any) {
          console.error('[MAINTENANCE_EVENTS][DEV] Erro:', error?.message || error);
          return writeJson(res, 500, { success: false, error: 'Erro ao consultar eventos de manutenção' });
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

