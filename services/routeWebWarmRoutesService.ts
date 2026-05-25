import { User } from '../types';
import { SharePointService } from './sharepointService';
import { getValidToken } from './tokenService';

export type RouteWebWarmQueryVariables = {
  perPage: number;
  strictDate: number;
  initialExpectedStartDate: string;
  finalExpectedStartDate: string;
};

export type RouteWebWarmRouteItem = Record<string, any> & {
  __plantId: number;
  __routeId: number | null;
  __filialName: string;
  __upstreamUrl?: string;
};

export type RouteWebWarmFailedPlant = {
  plantId: number;
  status?: number;
  error: string;
};

export type RouteWebWarmSnapshot = {
  key: string;
  userEmail: string;
  dateRef: string;
  status: 'idle' | 'loading' | 'ready' | 'error';
  query: RouteWebWarmQueryVariables;
  totalPlantIds: number;
  allowedPlantIds: number[];
  routes: RouteWebWarmRouteItem[];
  routeIds: number[];
  failedPlantIds: RouteWebWarmFailedPlant[];
  tokenPreview: string;
  bearerToken: string;
  error: string | null;
  startedAt: number;
  updatedAt: number;
  finishedAt: number | null;
};

export type RouteWebWarmEventEntry = {
  routeId: number;
  events: any[];
  upstreamUrl?: string;
  updatedAt: number;
};

const ROUTES_QUERY_DEFAULTS = {
  perPage: 60,
  strictDate: 1
};

const toNumericId = (value: unknown): number | null => {
  if (value == null) return null;
  if (typeof value === 'number' && Number.isFinite(value)) return Math.trunc(value);
  const raw = String(value).trim();
  if (!raw) return null;
  const match = raw.match(/-?\d+(?:[.,]\d+)?/);
  if (!match) return null;
  const parsed = Number(match[0].replace(',', '.'));
  return Number.isFinite(parsed) ? Math.trunc(parsed) : null;
};

const normalizeEmail = (email: string): string => String(email || '').toLowerCase().trim();

const getCurrentDayDate = (): string => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const buildRouteQueryVariables = (): RouteWebWarmQueryVariables => {
  const currentDay = getCurrentDayDate();
  return {
    perPage: ROUTES_QUERY_DEFAULTS.perPage,
    strictDate: ROUTES_QUERY_DEFAULTS.strictDate,
    initialExpectedStartDate: `${currentDay}T00:00:00Z`,
    finalExpectedStartDate: `${currentDay}T23:59:59Z`
  };
};

const buildCacheKey = (userEmail: string, dateRef: string): string => `${normalizeEmail(userEmail)}|${dateRef}`;

const pickRoutesArray = (payload: any): any[] => {
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload?.routes)) return payload.routes;
  if (Array.isArray(payload?.items)) return payload.items;
  if (Array.isArray(payload?.results)) return payload.results;
  if (Array.isArray(payload?.data?.routes)) return payload.data.routes;
  return [];
};

const fetchJsonWithTimeout = async (
  url: string,
  init: RequestInit,
  timeoutMs: number
): Promise<{ response: Response; data: any }> => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(url, {
      ...init,
      signal: controller.signal
    });
    const data = await response.json().catch(() => ({}));
    return { response, data };
  } finally {
    clearTimeout(timer);
  }
};

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

const createInitialSnapshot = (): RouteWebWarmSnapshot => {
  const dateRef = getCurrentDayDate();
  return {
    key: buildCacheKey('', dateRef),
    userEmail: '',
    dateRef,
    status: 'idle',
    query: buildRouteQueryVariables(),
    totalPlantIds: 0,
    allowedPlantIds: [],
    routes: [],
    routeIds: [],
    failedPlantIds: [],
    tokenPreview: '',
    bearerToken: '',
    error: null,
    startedAt: 0,
    updatedAt: Date.now(),
    finishedAt: null
  };
};

let warmSnapshot: RouteWebWarmSnapshot = createInitialSnapshot();
let warmPromise: Promise<RouteWebWarmSnapshot> | null = null;
let warmKey = '';
const warmEventsByKey = new Map<string, Map<number, RouteWebWarmEventEntry>>();
const warmEventsPromiseByKey = new Map<string, Promise<void>>();

const sortRouteIds = (routes: RouteWebWarmRouteItem[]): number[] => {
  return Array.from(
    new Set(
      routes
        .map((route) => toNumericId(route.__routeId ?? route.id))
        .filter((id): id is number => id != null)
    )
  ).sort((a, b) => a - b);
};

export const getRouteWebWarmRoutesSnapshot = (): RouteWebWarmSnapshot => warmSnapshot;

const getWarmEventsMapInternal = (key: string): Map<number, RouteWebWarmEventEntry> => {
  const existing = warmEventsByKey.get(key);
  if (existing) return existing;
  const created = new Map<number, RouteWebWarmEventEntry>();
  warmEventsByKey.set(key, created);
  return created;
};

export const getRouteWebWarmEventsMap = (currentUser: User): Map<number, RouteWebWarmEventEntry> => {
  const key = buildCacheKey(normalizeEmail(currentUser.email), getCurrentDayDate());
  return new Map(getWarmEventsMapInternal(key));
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

export const primeRouteWebWarmRoutes = async (currentUser: User): Promise<RouteWebWarmSnapshot> => {
  const userEmail = normalizeEmail(currentUser.email);
  const dateRef = getCurrentDayDate();
  const key = buildCacheKey(userEmail, dateRef);

  if (warmSnapshot.key === key && warmSnapshot.status === 'ready') {
    return warmSnapshot;
  }

  if (warmPromise && warmKey === key) {
    return warmPromise;
  }

  warmKey = key;
  warmPromise = (async () => {
    const queryVariables = buildRouteQueryVariables();
    const startedAt = Date.now();

    warmSnapshot = {
      key,
      userEmail,
      dateRef,
      status: 'loading',
      query: queryVariables,
      totalPlantIds: 0,
      allowedPlantIds: [],
      routes: [],
      routeIds: [],
      failedPlantIds: [],
      tokenPreview: '',
      bearerToken: '',
      error: null,
      startedAt,
      updatedAt: startedAt,
      finishedAt: null
    };

    try {
      const graphToken = (await getValidToken()) || currentUser.accessToken;
      if (!graphToken) {
        throw new Error('Sessão sem token válido para identificar Plant_id.');
      }

      const accessResult = await SharePointService.getRouteConfigsByAccess(graphToken, currentUser.email, false);
      const allowedPlantIds = Array.from(
        new Set(
          (accessResult.configs || [])
            .map((config) => toNumericId(config.plantId))
            .filter((id): id is number => id != null)
        )
      ).sort((a, b) => a - b);

      const plantDisplayMap = new Map<number, string>();
      (accessResult.configs || []).forEach((cfg) => {
        const plantId = toNumericId(cfg.plantId);
        if (plantId == null) return;
        const display = String(cfg.nomeExibicao || cfg.operacao || `Plant ${plantId}`).trim();
        if (display) {
          plantDisplayMap.set(plantId, display);
        }
      });

      warmSnapshot = {
        ...warmSnapshot,
        totalPlantIds: allowedPlantIds.length,
        allowedPlantIds,
        updatedAt: Date.now()
      };

      if (allowedPlantIds.length === 0) {
        warmSnapshot = {
          ...warmSnapshot,
          status: 'ready',
          routes: [],
          routeIds: [],
          failedPlantIds: [],
          tokenPreview: '',
          bearerToken: '',
          finishedAt: Date.now(),
          updatedAt: Date.now()
        };
        return warmSnapshot;
      }

      const { response: tokenResponse, data: tokenData } = await fetchJsonWithTimeout(
        '/api/route-web',
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ resource: 'token' })
        },
        15000
      );

      if (!tokenResponse.ok || !tokenData?.success || !tokenData?.token) {
        throw new Error(tokenData?.error || 'Falha ao obter token para consulta de rotas');
      }

      const bearerToken = String(tokenData.token).trim();
      const tokenPreview = String(tokenData.tokenPreview || '');
      const routes: RouteWebWarmRouteItem[] = [];
      const failedPlantIds: RouteWebWarmFailedPlant[] = [];

      const routeFetchConcurrency = 2;
      let cursor = 0;

      const workers = Array.from({ length: Math.min(routeFetchConcurrency, allowedPlantIds.length) }, async () => {
        while (true) {
          const nextIndex = cursor;
          cursor += 1;
          if (nextIndex >= allowedPlantIds.length) break;

          const plantId = allowedPlantIds[nextIndex];
          try {
            const { response, data: proxyData } = await fetchJsonWithTimeout(
              '/api/route-web',
              {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  resource: 'routes',
                  plantId,
                  perPage: queryVariables.perPage,
                  strictDate: queryVariables.strictDate,
                  initialExpectedStartDate: queryVariables.initialExpectedStartDate,
                  finalExpectedStartDate: queryVariables.finalExpectedStartDate,
                  bearerToken,
                  compact: true
                })
              },
              20000
            );

            if (!response.ok || !proxyData?.success) {
              failedPlantIds.push({
                plantId,
                status: proxyData?.upstreamStatus || response.status,
                error: proxyData?.error || `Falha ao consultar rotas para plant_id=${plantId}`
              });
              continue;
            }

            const routePayload = proxyData?.routes ?? proxyData?.response;
            const parsedRoutes = pickRoutesArray(routePayload).map((route: any) => {
              const routeId = toNumericId(route?.id);
              const filialName =
                String(
                  route?.plant?.display_name ||
                    route?.plant?.name ||
                    route?.plant_name ||
                    plantDisplayMap.get(plantId) ||
                    `Plant ${plantId}`
                ).trim() || `Plant ${plantId}`;

              return {
                ...route,
                __plantId: plantId,
                __routeId: routeId,
                __filialName: filialName,
                __upstreamUrl: proxyData?.upstreamUrl
              } as RouteWebWarmRouteItem;
            });

            routes.push(...parsedRoutes);
          } catch (error: any) {
            failedPlantIds.push({
              plantId,
              error:
                error?.name === 'AbortError'
                  ? `Timeout ao consultar rotas para plant_id=${plantId}`
                  : error?.message || `Erro ao consultar rotas para plant_id=${plantId}`
            });
          }
        }
      });

      await Promise.all(workers);

      const routeIds = sortRouteIds(routes);
      warmSnapshot = {
        ...warmSnapshot,
        status: 'ready',
        query: queryVariables,
        totalPlantIds: allowedPlantIds.length,
        allowedPlantIds,
        routes,
        routeIds,
        failedPlantIds,
        tokenPreview,
        bearerToken,
        error: null,
        updatedAt: Date.now(),
        finishedAt: Date.now()
      };

      return warmSnapshot;
    } catch (error: any) {
      warmSnapshot = {
        ...warmSnapshot,
        status: 'error',
        error: error?.message || 'Falha no pré-carregamento de rotas',
        updatedAt: Date.now(),
        finishedAt: Date.now()
      };
      return warmSnapshot;
    }
  })().finally(() => {
    warmPromise = null;
  });

  return warmPromise;
};

export const primeRouteWebWarmEvents = async (currentUser: User): Promise<void> => {
  const userEmail = normalizeEmail(currentUser.email);
  const dateRef = getCurrentDayDate();
  const key = buildCacheKey(userEmail, dateRef);
  const existingPromise = warmEventsPromiseByKey.get(key);
  if (existingPromise) return existingPromise;

  const warmEventPromise = (async () => {
    const routeSnapshot = await primeRouteWebWarmRoutes(currentUser);
    if (routeSnapshot.status !== 'ready') return;

    const bearerToken = String(routeSnapshot.bearerToken || '').trim();
    if (!bearerToken) return;

    const routeIds = routeSnapshot.routeIds || [];
    if (routeIds.length === 0) return;

    const maxRoutesToWarm = Math.min(routeIds.length, 40);
    const candidateIds = [...routeIds]
      .sort((a, b) => b - a)
      .slice(0, maxRoutesToWarm);
    const eventsMap = getWarmEventsMapInternal(key);
    let warmedCount = 0;

    for (const routeId of candidateIds) {
      if (eventsMap.has(routeId)) {
        continue;
      }

      try {
        const firstCall = await fetchJsonWithTimeout(
          '/api/route-web',
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              resource: 'route-events',
              routeId,
              withOccurrences: true,
              bearerToken,
              compact: true,
              nonCollectionOnly: true
            })
          },
          10000
        );

        let response = firstCall.response;
        let eventData = firstCall.data;

        if (!response.ok || !eventData?.success) {
          await sleep(150);
          const retryCall = await fetchJsonWithTimeout(
            '/api/route-web',
            {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                resource: 'route-events',
                routeId,
                bearerToken,
                compact: true,
                nonCollectionOnly: true
              })
            },
            16000
          );
          response = retryCall.response;
          eventData = retryCall.data;
        }

        if (!response.ok || !eventData?.success) {
          continue;
        }

        const events = pickEventsArray(eventData?.events || eventData?.response);
        eventsMap.set(routeId, {
          routeId,
          events,
          upstreamUrl: eventData?.upstreamUrl,
          updatedAt: Date.now()
        });
        warmedCount += 1;

        await sleep(80);
      } catch {
        // Warmup de eventos é opcional. Erros aqui não devem impactar o app.
      }
    }

  })().finally(() => {
    warmEventsPromiseByKey.delete(key);
  });

  warmEventsPromiseByKey.set(key, warmEventPromise);
  return warmEventPromise;
};
