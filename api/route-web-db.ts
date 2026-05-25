import type { VercelRequest, VercelResponse } from '@vercel/node';
import {
  getRouteWebEventsByDateAndPlants,
  getRouteWebRoutesByDateAndPlants,
  closeRwePool,
  type RouteWebEventDbRow,
  type RouteWebRouteDbRow
} from './lib-rweDb.js';

/**
 * Endpoint consolidado Route Web DB.
 * Uso: GET/POST /api/route-web-db
 * Body/Query: { "entity": "events"|"routes", "dataReferencia": "YYYY-MM-DD", "plantIds": [...] }
 */

export type { RouteWebEventDbRow, RouteWebRouteDbRow };

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'GET' && req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const body = req.method === 'POST' ? req.body : req.query;
    const entity = String(body?.entity || '').trim();
    if (!entity) return res.status(400).json({ success: false, error: 'entity é obrigatório (events, routes)' });

    const dataReferencia = String(body?.dataReferencia || '').trim();
    if (!dataReferencia || !/^\d{4}-\d{2}-\d{2}$/.test(dataReferencia)) {
      return res.status(400).json({ success: false, error: 'dataReferencia inválida (YYYY-MM-DD)' });
    }

    let plantIds: number[] = [];
    const plantIdsRaw = body?.plantIds;
    if (Array.isArray(plantIdsRaw)) {
      plantIds = plantIdsRaw.map(Number).filter(Number.isFinite);
    } else if (plantIdsRaw != null) {
      const parsed = Number(plantIdsRaw);
      if (Number.isFinite(parsed)) plantIds = [parsed];
    }

    switch (entity) {
      case 'events': {
        const rows: RouteWebEventDbRow[] = await getRouteWebEventsByDateAndPlants(dataReferencia, plantIds);
        return res.status(200).json({ success: true, dataReferencia, plantIds, count: rows.length, events: rows });
      }
      case 'routes': {
        const rows: RouteWebRouteDbRow[] = await getRouteWebRoutesByDateAndPlants(dataReferencia, plantIds);
        return res.status(200).json({ success: true, dataReferencia, plantIds, count: rows.length, routes: rows });
      }
      default:
        return res.status(400).json({ success: false, error: `Entidade desconhecida: ${entity}` });
    }
  } catch (error: any) {
    console.error('[ROUTE_WEB_DB] Erro:', error?.message || error);
    return res.status(500).json({ success: false, error: error?.message || 'Erro ao consultar dados' });
  } finally {
    await closeRwePool().catch(() => {});
  }
}
