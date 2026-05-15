import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getTokenPreview, getRouteWebTokenUrl, requestRouteWebToken } from '../utils/routeWebServer.js';

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
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
  } catch (error: any) {
    console.error('[ROUTE_WEB_TOKEN] Erro ao obter token:', error?.message || error);
    return res.status(500).json({
      success: false,
      error: error?.message || 'Erro ao obter token do Route Web'
    });
  }
}
