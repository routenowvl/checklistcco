import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getDepartures, upsertDeparture, deleteDeparture } from '../utils/checklistDb';

const validateToken = async (authHeader: string | undefined): Promise<boolean> => {
  if (!authHeader) return false;
  const token = authHeader.startsWith('Bearer ')
    ? authHeader.slice(7).trim()
    : authHeader.trim();
  if (!token) return false;
  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${token}` }
    });
    return res.ok;
  } catch {
    return false;
  }
};

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  const isAuthenticated = await validateToken(req.headers.authorization);
  if (!isAuthenticated) {
    return res.status(401).json({ success: false, error: 'Unauthorized' });
  }

  try {
    const { action } = req.body;

    switch (action) {
      case 'getAll': {
        const departures = await getDepartures();
        return res.status(200).json({ success: true, departures });
      }

      case 'upsert': {
        const { departure } = req.body;
        if (!departure || typeof departure !== 'object') {
          return res.status(400).json({ success: false, error: 'departure é obrigatório' });
        }
        const id = await upsertDeparture(departure);
        return res.status(200).json({ success: true, id });
      }

      case 'delete': {
        const { id } = req.body;
        if (!id) {
          return res.status(400).json({ success: false, error: 'id é obrigatório' });
        }
        await deleteDeparture(Number(id));
        return res.status(200).json({ success: true });
      }

      default:
        return res.status(400).json({ success: false, error: `Ação desconhecida: ${action}` });
    }
  } catch (error: any) {
    console.error('[CHECKLIST_DEPARTURES] Erro:', error?.message || error);
    return res.status(500).json({ success: false, error: 'Erro ao processar operação de departures' });
  }
}
