import type { VercelRequest, VercelResponse } from '@vercel/node';
import {
  appendQueryToUrl,
  getRouteWebRoutesEndpointUrl,
  getRouteWebRoutesEnvDebug,
  getTokenPreview,
  requestRouteWebToken
} from '../utils/routeWebServer.js';

type RouteWebRoutesBody = {
  plantId?: number | string;
  perPage?: number | string;
  strictDate?: number | string;
  initialExpectedStartDate?: string;
  finalExpectedStartDate?: string;
  bearerToken?: string;
  compact?: boolean | number | string;
};

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

const ROUTES_SUCCESS_CACHE_TTL_MS = 2 * 60 * 1000;
const ROUTES_ERROR_CACHE_TTL_MS = 30 * 1000;
const routesCache = new Map<string, {
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
const routesInFlight = new Map<string, Promise<{
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

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const body = (req.body || {}) as RouteWebRoutesBody;

    const plantId = toOptionalInt(body.plantId);
    if (plantId == null) {
      return res.status(400).json({ success: false, error: 'plantId é obrigatório' });
    }

    const perPage = toOptionalInt(body.perPage) ?? 60;
    const strictDate = toOptionalInt(body.strictDate) ?? 1;
    const compact = toBoolean(body.compact, false);
    const initialExpectedStartDate = String(body.initialExpectedStartDate || '').trim();
    const finalExpectedStartDate = String(body.finalExpectedStartDate || '').trim();

    if (!initialExpectedStartDate || !finalExpectedStartDate) {
      return res.status(400).json({
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

    const manualToken = String(body.bearerToken || '').trim();
    const tokenResult = manualToken ? null : await requestRouteWebToken();
    const bearerToken = manualToken || String(tokenResult?.token || '').trim();
    const tokenSource = manualToken ? 'manual' : 'oauth';

    if (!bearerToken) {
      throw new Error('Token de acesso não disponível para consulta de rotas');
    }

    const routesEndpointUrl = getRouteWebRoutesEndpointUrl();
    const upstreamUrl = appendQueryToUrl(routesEndpointUrl, query);
    const requestQuery = Object.fromEntries(query.entries());
    const routesEnvDebug = getRouteWebRoutesEnvDebug();

    console.log('[ROUTE_WEB_ROUTES] Request:', {
      routesEndpointUrl,
      upstreamUrl,
      requestQuery,
      routesEnvDebug
    });

    const cacheKey = `${upstreamUrl}|${compact ? 1 : 0}`;
    const now = Date.now();
    let fromCache = false;
    const cached = routesCache.get(cacheKey);
    let cachedEntry:
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

    if (cached) {
      const maxAge = cached.ok ? ROUTES_SUCCESS_CACHE_TTL_MS : ROUTES_ERROR_CACHE_TTL_MS;
      if (now - cached.fetchedAt <= maxAge) {
        cachedEntry = cached;
        fromCache = true;
      }
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
          const entry = {
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
          routesCache.set(cacheKey, entry);
          return entry;
        })().finally(() => {
          routesInFlight.delete(cacheKey);
        });

        routesInFlight.set(cacheKey, promise);
        cachedEntry = await promise;
      }
    }

    const routes = cachedEntry.routes;
    const upstreamSnippet = cachedEntry.upstreamSnippet;

    console.log('[ROUTE_WEB_ROUTES] Response:', {
      upstreamUrl,
      status: cachedEntry.status,
      statusText: cachedEntry.statusText,
      contentType: cachedEntry.contentType,
      routesCount: routes.length,
      fromCache,
      snippet: upstreamSnippet
    });

    const errorMessage = cachedEntry.ok
      ? undefined
      : `Upstream ${cachedEntry.status} ${cachedEntry.statusText} em ${upstreamUrl}${upstreamSnippet ? ` | ${upstreamSnippet}` : ''}`;
    const payload: any = {
      success: cachedEntry.ok,
      error: errorMessage,
      upstreamStatus: cachedEntry.status,
      upstreamStatusText: cachedEntry.statusText,
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
      contentType: cachedEntry.contentType,
      count: routes.length,
      routes: compact ? routes.map(compactRoute) : routes
    };

    if (!compact) {
      payload.response = cachedEntry.data;
      payload.raw = cachedEntry.raw;
    }

    return res.status(200).json(payload);
  } catch (error: any) {
    console.error('[ROUTE_WEB_ROUTES] Erro ao consultar routes:', error?.message || error);
    return res.status(500).json({
      success: false,
      error: error?.message || 'Erro ao consultar routes do Route Web'
    });
  }
}
