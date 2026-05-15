import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getRouteWebUpstreamUrl, getTokenPreview, requestRouteWebToken } from '../utils/routeWebServer.js';

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

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

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
  } catch (error: any) {
    console.error('[ROUTE_WEB_PLANTS] Erro ao consultar plants:', error?.message || error);
    return res.status(500).json({
      success: false,
      error: error?.message || 'Erro ao consultar plants do Route Web'
    });
  }
}
