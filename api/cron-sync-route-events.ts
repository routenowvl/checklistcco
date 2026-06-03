import type { VercelRequest, VercelResponse } from '@vercel/node';
import {
  appendQueryToUrl,
  getRouteWebRouteEventsEndpointUrl,
  getRouteWebRoutesEndpointUrl,
  requestRouteWebToken
} from '../utils/routeWebServer.js';
import { getPlantConfigsFromSharePoint, type PlantConfig } from '../utils/graphAppAuth.js';
import { upsertRouteWebEvents, closeRwePool, deleteDoneEvents, type RouteWebEventRow } from '../utils/rweDb.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const toOptionalInt = (value: unknown): number | null => {
  if (value == null) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  const parsed = Number(raw);
  if (!Number.isFinite(parsed)) return null;
  return Math.trunc(parsed);
};

const normalizeText = (value: unknown): string =>
  String(value ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

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
  if (NON_COLLECTION_EXCLUDED_REASON_PATTERNS.some((p) => descriptionNormalized.includes(normalizeText(p)))) {
    return false;
  }
  if (isScraperSmartQuestionOccurrence(occ)) return true;
  if (isTechnicalOccurrence(occ)) return false;
  const occurrenceTypeId = toOptionalInt(occ?.occurrence_type_id ?? occ?.occurrence_type?.id);
  return occurrenceTypeId != null || Boolean(description);
};

const getRouteCode = (route: any): string =>
  String(
    route?.roadmap_code ||
      route?.route_code ||
      route?.code ||
      route?.route ||
      route?.route_plan_id ||
      route?.id ||
      '-'
  );

const getDriverName = (route: any): string =>
  String(
    route?.driver_name ||
      route?.last_driver_name ||
      route?.driver?.name ||
      route?.last_driver?.name ||
      '-'
  );

const getPlate = (event: any, route: any): string => {
  const raw = String(
    route?.unloading_plate ||
      event?.trailer_plate ||
      event?.plate ||
      route?.trailer_plate ||
      route?.plate ||
      route?.vehicle_plate ||
      route?.last_vehicle?.plate ||
      '-'
  ).trim();
  return raw || '-';
};

const getScraperOccurrenceReason = (event: any): string => {
  if (!Array.isArray(event?.occurrences)) return '';
  for (const occ of event.occurrences) {
    if (normalizeText(occ?.inserted_by) !== 'scrapersmartquestion') continue;
    const reason = getOccurrenceDescription(occ);
    if (reason) return reason;
  }
  return '';
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

const pickEventsArray = (payload: any): any[] => {
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload?.events)) return payload.events;
  if (Array.isArray(payload?.items)) return payload.items;
  if (Array.isArray(payload?.results)) return payload.results;
  if (Array.isArray(payload?.data?.events)) return payload.data.events;
  return [];
};

const getCurrentDayDate = (): string => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const parseUpstreamResponse = async (response: Response): Promise<{ contentType: string; raw: string; data: any }> => {
  const contentType = String(response.headers.get('content-type') || '');
  const raw = await response.text();
  if (!raw) return { contentType, raw: '', data: null };
  try {
    return { contentType, raw, data: JSON.parse(raw) };
  } catch {
    return { contentType, raw, data: raw };
  }
};

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

const CRON_SYNC_SECRET = String(process.env.CRON_SYNC_SECRET || '').trim();

// ---------------------------------------------------------------------------
// Lógica de sincronização
// ---------------------------------------------------------------------------

const syncAll = async (): Promise<{
  success: boolean;
  totalRoutes: number;
  totalEvents: number;
  totalUpserted: number;
  errors: string[];
}> => {
  const errors: string[] = [];
  const startedAt = Date.now();
  const dateRef = getCurrentDayDate();

  console.log(`[CRON_SYNC] Iniciando sincronização para ${dateRef}`);

  // 1. Obter token Route Web
  let bearerToken: string;
  try {
    const tokenResult = await requestRouteWebToken();
    bearerToken = String(tokenResult.token || '').trim();
    if (!bearerToken) throw new Error('Token vazio');
  } catch (error: any) {
    const msg = `Token: ${error?.message || 'erro desconhecido'}`;
    console.error(`[CRON_SYNC] ${msg}`);
    return { success: false, totalRoutes: 0, totalEvents: 0, totalUpserted: 0, errors: [msg] };
  }

  // 2. Buscar plant configs do SharePoint
  let plantConfigs: PlantConfig[];
  try {
    plantConfigs = await getPlantConfigsFromSharePoint();
  } catch (error: any) {
    const msg = `SharePoint configs: ${error?.message || 'erro desconhecido'}`;
    console.error(`[CRON_SYNC] ${msg}`);
    return { success: false, totalRoutes: 0, totalEvents: 0, totalUpserted: 0, errors: [msg] };
  }

  if (plantConfigs.length === 0) {
    const msg = 'Nenhuma plant config encontrada no SharePoint';
    console.warn(`[CRON_SYNC] ${msg}`);
    return { success: false, totalRoutes: 0, totalEvents: 0, totalUpserted: 0, errors: [msg] };
  }

  console.log(`[CRON_SYNC] ${plantConfigs.length} plants obtidas do SharePoint`);

  const allRows: RouteWebEventRow[] = [];
  const doneKeys: { route_id: number; event_id: number }[] = [];
  let totalRoutes = 0;
  let totalEvents = 0;

  // 2. Buscar rotas por plant
  for (const config of plantConfigs) {
    try {
      const query = new URLSearchParams({
        plant_id: String(config.plantId),
        per_page: '60',
        strict_date: '1',
        initial_expected_start_date: `${dateRef}T00:00:00Z`,
        final_expected_start_date: `${dateRef}T23:59:59Z`
      });

      const routesEndpointUrl = getRouteWebRoutesEndpointUrl();
      const upstreamUrl = appendQueryToUrl(routesEndpointUrl, query);

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

      if (!response.ok) {
        const snippet = String(upstreamPayload.raw || '').slice(0, 300);
        errors.push(`Plant ${config.plantId}: upstream ${response.status} ${snippet}`);
        continue;
      }

      const routes = pickRoutesArray(upstreamPayload.data);
      totalRoutes += routes.length;

      console.log(`[CRON_SYNC] Plant ${config.plantId} (${config.filial}): ${routes.length} rotas`);

      // 3. Buscar eventos de cada rota (em lotes de 4)
      const batchSize = 4;
      for (let i = 0; i < routes.length; i += batchSize) {
        const batch = routes.slice(i, i + batchSize);

        await Promise.all(
          batch.map(async (route: any) => {
            const routeId = toOptionalInt(route?.id);
            if (routeId == null) return;

            try {
              const eventsEndpointUrl = getRouteWebRouteEventsEndpointUrl(routeId);
              const eventsQuery = new URLSearchParams({ with_occurrences: 'true' });
              const eventsUrl = appendQueryToUrl(eventsEndpointUrl, eventsQuery);

              const eventsResponse = await fetch(eventsUrl, {
                method: 'GET',
                headers: {
                  Authorization: `Bearer ${bearerToken}`,
                  'Content-Type': 'application/json',
                  'X-Requested-With': 'XMLHttpRequest',
                  'x-requested_with': 'XLMHttpRequest'
                }
              });

              const eventsPayload = await parseUpstreamResponse(eventsResponse);

              if (!eventsResponse.ok) return;

              const events = pickEventsArray(eventsPayload.data);
              const rotaCodigo = getRouteCode(route);
              const motorista = getDriverName(route);
              const filialName = String(
                route?.plant?.display_name ||
                  route?.plant?.name ||
                  route?.plant_name ||
                  config.filial
              ).trim();
              const operacao = String(config.operacao || '').trim();

              for (const event of events) {
                const typeNameNormalized = normalizeText(event?.type_name);
                if (typeNameNormalized !== 'coleta') continue;

                const occurrences = Array.isArray(event?.occurrences) ? event.occurrences : [];
                const nonCollectionOccs = occurrences.filter(isNonCollectionOccurrence);
                const placa = getPlate(event, route);
                const eventId = toOptionalInt(event?.id);
                const isScheduled = !event?.executed && normalizeText(event?.status) === 'scheduled';

                // Não coletas (ocorrências)
                if (nonCollectionOccs.length > 0) {
                  totalEvents += 1;
                  const scraperReason = getScraperOccurrenceReason(event);
                  const fallbackReason = getOccurrenceDescription(nonCollectionOccs[0]);
                  const motivo = scraperReason || fallbackReason || String(event?.motivo || event?.reason || '').trim();

                  for (const occ of nonCollectionOccs) {
                    allRows.push({
                      route_id: routeId,
                      event_id: eventId,
                      plant_id: config.plantId,
                      filial: filialName,
                      operacao,
                      rota_codigo: rotaCodigo,
                      motorista,
                      placa,
                      type_name: String(event?.type_name || ''),
                      reference: String(event?.reference || ''),
                      reference_code: String(event?.reference_code || ''),
                      status: String(event?.status || ''),
                      executed: event?.executed == null ? null : Boolean(event.executed),
                      expected_arrival: event?.expected_arrival || null,
                      actual_arrival: event?.actual_arrival || null,
                      expected_departure: event?.expected_departure || null,
                      actual_departure: event?.actual_departure || null,
                      motivo,
                      status_type: 'nao-coleta',
                      occurrence_id: toOptionalInt(occ?.id),
                      occurrence_type_id: toOptionalInt(occ?.occurrence_type_id ?? occ?.occurrence_type?.id),
                      occurrence_type_description: getOccurrenceDescription(occ),
                      occurrence_inserted_by: String(occ?.inserted_by || ''),
                      occurrence_inserted_at: occ?.inserted_at || null,
                      event_created_at: event?.created_at || null,
                      event_updated_at: event?.updated_at || null,
                      data_referencia: dateRef
                    });
                  }
                }

                // Coletas previstas (scheduled, not executed, sem ocorrências relevantes)
                if (isScheduled) {
                  totalEvents += 1;
                  allRows.push({
                    route_id: routeId,
                    event_id: eventId,
                    plant_id: config.plantId,
                    filial: filialName,
                    operacao,
                    rota_codigo: rotaCodigo,
                    motorista,
                    placa,
                    type_name: String(event?.type_name || ''),
                    reference: String(event?.reference || ''),
                    reference_code: String(event?.reference_code || ''),
                    status: String(event?.status || ''),
                    executed: false,
                    expected_arrival: event?.expected_arrival || null,
                    actual_arrival: null,
                    expected_departure: event?.expected_departure || null,
                    actual_departure: null,
                    motivo: 'Coleta prevista',
                    status_type: 'coleta-prevista',
                    occurrence_id: -1,
                    occurrence_type_id: null,
                    occurrence_type_description: '',
                    occurrence_inserted_by: '',
                    occurrence_inserted_at: null,
                    event_created_at: event?.created_at || null,
                    event_updated_at: event?.updated_at || null,
                    data_referencia: dateRef
                  });
                }

                // Eventos DONE/executed sem ocorrências de não-coleta: remover do banco
                if (event?.executed && normalizeText(event?.status) === 'done' && nonCollectionOccs.length === 0) {
                  doneKeys.push({ route_id: routeId, event_id: eventId });
                }
              }
            } catch (error: any) {
              errors.push(`Route ${routeId}: ${error?.message || 'erro'}`);
            }
          })
        );

        if (i + batchSize < routes.length) {
          await sleep(100);
        }
      }
    } catch (error: any) {
      errors.push(`Plant ${config.plantId}: ${error?.message || 'erro'}`);
    }
  }

  // 4. Persistir no banco
  let totalUpserted = 0;
  if (allRows.length > 0) {
    try {
      totalUpserted = await upsertRouteWebEvents(allRows);
    } catch (error: any) {
      errors.push(`DB: ${error?.message || 'erro ao persistir'}`);
    }
  }

  // 5. Remover eventos DONE do banco
  if (doneKeys.length > 0) {
    try {
      await deleteDoneEvents(doneKeys);
    } catch (error: any) {
      errors.push(`DB delete done: ${error?.message || 'erro ao deletar'}`);
    }
  }

  await closeRwePool().catch(() => {});

  const durationMs = Date.now() - startedAt;
  console.log(`[CRON_SYNC] Concluído em ${durationMs}ms: ${totalRoutes} rotas, ${totalEvents} eventos, ${totalUpserted} upserted, ${errors.length} erros`);

  return {
    success: errors.length === 0,
    totalRoutes,
    totalEvents,
    totalUpserted,
    errors
  };
};

// ---------------------------------------------------------------------------
// Handler
// ---------------------------------------------------------------------------

export default async function handler(req: VercelRequest, res: VercelResponse) {
  // Somente POST para chamadas manuais (desativado cron automático Vercel — sync via EC2)
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  // Se CRON_SYNC_SECRET está configurado, valida o header de autorização
  if (CRON_SYNC_SECRET) {
    const authHeader = String(req.headers['authorization'] || '').trim();
    const provided = authHeader.startsWith('Bearer ') ? authHeader.slice(7).trim() : '';
    if (provided !== CRON_SYNC_SECRET) {
      return res.status(401).json({ success: false, error: 'Unauthorized' });
    }
  }

  try {
    const result = await syncAll();
    return res.status(result.success ? 200 : 207).json({
      success: result.success,
      ...result
    });
  } catch (error: any) {
    console.error('[CRON_SYNC] Erro fatal:', error?.message || error);
    return res.status(500).json({
      success: false,
      error: error?.message || 'Erro fatal na sincronização'
    });
  }
}
