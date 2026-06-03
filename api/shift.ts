import type { VercelRequest, VercelResponse } from '@vercel/node';
import {
  getShiftApiBaseUrl,
  requestShiftToken,
  getTokenPreview
} from './lib-shiftApi.js';

/**
 * Endpoint Shift API — Escala de Motoristas.
 * Uso: POST /api/shift
 * Body: { "action": "schedules"|"consolidation", ... }
 *
 * Autentica-se diretamente na Shift API (SHIFT_API_URL/api/oauth/token)
 * usando credenciais da Route Web ou próprias (SHIFT_*).
 */

// ─── Helpers ──────────────────────────────────────────────────────────────

const parseResponse = async (response: Response): Promise<{ contentType: string; raw: string; data: any }> => {
  const contentType = String(response.headers.get('content-type') || '');
  const raw = await response.text();
  if (!raw) return { contentType, raw: '', data: null };
  try { return { contentType, raw, data: JSON.parse(raw) }; } catch { return { contentType, raw, data: raw }; }
};

const toOptionalInt = (value: unknown): number | null => {
  if (value == null) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? Math.trunc(parsed) : null;
};

// ─── Action: schedules ────────────────────────────────────────────────────

const handleSchedules = async (body: any, res: VercelResponse) => {
  const plantId = toOptionalInt(body.plant_id);
  if (plantId == null) return res.status(400).json({ success: false, error: 'plant_id é obrigatório' });

  const code = String(body.code || '').trim();
  if (!code) return res.status(400).json({ success: false, error: 'code é obrigatório (formato YYYYMM)' });

  const perPage = toOptionalInt(body.per_page) ?? 100;

  let token: string;
  try {
    token = await requestShiftToken();
    console.log(`[SHIFT][SCHEDULES] Token obtido com sucesso, preview: ${getTokenPreview(token)}`);
  } catch (tokenErr: any) {
    console.error(`[SHIFT][SCHEDULES] Falha ao obter token: ${tokenErr.message}`);
    return res.status(500).json({ success: false, error: `Falha ao obter token: ${tokenErr.message}` });
  }

  const baseUrl = getShiftApiBaseUrl();
  const upstreamUrl = `${baseUrl}/api/schedules?plant_id=${plantId}&code=${encodeURIComponent(code)}&per_page=${perPage}`;

  console.log(`[SHIFT][SCHEDULES] Request: GET ${upstreamUrl}`);
  console.log(`[SHIFT][SCHEDULES] plant_id=${plantId}, code=${code}, per_page=${perPage}`);
  console.log(`[SHIFT][SCHEDULES] Authorization: Bearer ${getTokenPreview(token)}`);

  const response = await fetch(upstreamUrl, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      Accept: 'application/json, text/plain, */*',
      'X-Requested-With': 'XMLHttpRequest',
      'x-requested_with': 'XLMHttpRequest'
    }
  });

  const parsed = await parseResponse(response);

  console.log(`[SHIFT][SCHEDULES] Response: status=${response.status}, content-type=${parsed.contentType}`);
  console.log(`[SHIFT][SCHEDULES] Response body (primeiros 500 chars): ${String(parsed.raw || '').slice(0, 500)}`);

  return res.status(200).json({
    success: response.ok,
    upstreamStatus: response.status,
    upstreamUrl,
    tokenPreview: getTokenPreview(token),
    baseUrl,
    plantId,
    code,
    data: parsed.data,
    raw: parsed.raw
  });
};

// ─── Action: consolidation ────────────────────────────────────────────────

const handleConsolidation = async (body: any, res: VercelResponse) => {
  const scheduleId = String(body.schedule_id || '').trim();
  if (!scheduleId) return res.status(400).json({ success: false, error: 'schedule_id é obrigatório' });

  const token = await requestShiftToken();
  const baseUrl = getShiftApiBaseUrl();
  const upstreamUrl = `${baseUrl}/api/schedules/${encodeURIComponent(scheduleId)}/consolidation`;

  const response = await fetch(upstreamUrl, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      Accept: 'application/json, text/plain, */*',
      'X-Requested-With': 'XMLHttpRequest',
      'x-requested_with': 'XLMHttpRequest'
    }
  });

  const parsed = await parseResponse(response);

  // Transforma a consolidação no formato simplificado para o frontend
  const resultado: any[] = [];
  const consolidation = parsed.data;

  if (consolidation?.data?.shifts) {
    for (const shift of consolidation.data.shifts) {
      if (!shift?.drivers) continue;
      for (const driver of shift.drivers) {
        if (!driver?.days) continue;
        for (const [data, infoDia] of Object.entries(driver.days)) {
          const dayInfo = infoDia as any;
          resultado.push({
            motoristaId: driver.id,
            motorista: driver.name,
            data,
            status: dayInfo?.status || 'UNKNOWN',
            rotaId: dayInfo?.routePlan?.id || null,
            rota: dayInfo?.routePlan?.code || null,
            inicioPrevisto: dayInfo?.routePlan?.expectedStart || null,
            fimPrevisto: dayInfo?.routePlan?.expectedEnd || null,
            operacao: consolidation.data.plantId
          });
        }
      }
    }
  }

  return res.status(200).json({
    success: response.ok,
    upstreamStatus: response.status,
    upstreamUrl,
    tokenPreview: getTokenPreview(token),
    resultado,
    data: parsed.data,
    raw: parsed.raw
  });
};

// ─── Handler ──────────────────────────────────────────────────────────────

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const { action } = req.body || {};
    if (!action) return res.status(400).json({ success: false, error: 'action é obrigatório (schedules, consolidation)' });

    switch (action) {
      case 'schedules': return await handleSchedules(req.body, res);
      case 'consolidation': return await handleConsolidation(req.body, res);
      default: return res.status(400).json({ success: false, error: `Action desconhecida: ${action}` });
    }
  } catch (error: any) {
    console.error('[SHIFT] Erro:', error?.message || error);
    return res.status(500).json({ success: false, error: error?.message || 'Erro na Shift API' });
  }
}
