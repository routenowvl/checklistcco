import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getGraphAppToken } from '../utils/graphAppAuth';
import { upsertDeparture } from '../utils/checklistDb';

/**
 * Endpoint de migração: lê a lista Dados_Saida_de_rotas do SharePoint
 * e copia todos os itens para a tabela departures no PostgreSQL.
 *
 * Uso: POST /api/migrate-departures-from-sharepoint
 * Body (opcional): { "mode": "all" | "range", "startDate": "YYYY-MM-DD", "endDate": "YYYY-MM-DD" }
 *
 * DELETAR este arquivo após a migração.
 */

const SITE_PATH = process.env.VITE_SHAREPOINT_SITE_PATH || '';
const LIST_NAME = 'Dados_Saida_de_rotas';

const graphFetch = async (endpoint: string, token: string): Promise<any> => {
  const url = endpoint.startsWith('https://')
    ? endpoint
    : `https://graph.microsoft.com/v1.0${endpoint}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API ${res.status}: ${text.slice(0, 400)}`);
  }
  return res.status === 204 ? null : res.json();
};

const normalizeString = (str: string): string =>
  str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '').trim();

const resolveFieldName = (mapping: Record<string, string>, target: string): string => {
  return mapping[normalizeString(target)] || target;
};

const findListByIdOrName = async (siteId: string, listIdOrName: string, token: string): Promise<any> => {
  try {
    return await graphFetch(`/sites/${siteId}/lists/${listIdOrName}`, token);
  } catch {
    const listsData = await graphFetch(`/sites/${siteId}/lists`, token);
    const found = (listsData.value || []).find(
      (l: any) =>
        l.name?.toLowerCase() === listIdOrName.toLowerCase() ||
        l.displayName?.toLowerCase() === listIdOrName.toLowerCase()
    );
    if (!found) throw new Error(`Lista ${listIdOrName} não encontrada`);
    return found;
  }
};

/** Converte data ISO do SharePoint para DD/MM/AAAA (formato que o upsertDeparture espera) */
const formatISOtoBR = (iso: any): string => {
  if (!iso) return '';
  const s = String(iso).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : s;
};

/** Converte time ISO "HH:MM:SS" ou datetime para "HH:MM:SS" */
const formatTime = (v: any): string => {
  if (!v) return '';
  const s = String(v).trim();
  const dtMatch = s.match(/T(\d{2}:\d{2})/);
  if (dtMatch) return dtMatch[1] + ':00';
  const tMatch = s.match(/^(\d{2}:\d{2})/);
  return tMatch ? tMatch[1] + ':00' : s;
};

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const appToken = await getGraphAppToken();

    // 1. Resolver site
    const siteData = await graphFetch(`/sites/${SITE_PATH}`, appToken);
    const siteId = siteData.id;

    // 2. Encontrar lista Dados_Saida_de_rotas
    const list = await findListByIdOrName(siteId, LIST_NAME, appToken);

    // 3. Obter mapeamento de colunas
    const columns = await graphFetch(`/sites/${siteId}/lists/${list.id}/columns`, appToken);
    const mapping: Record<string, string> = {};
    for (const col of columns.value || []) {
      mapping[normalizeString(col.name)] = col.name;
      mapping[normalizeString(col.displayName)] = col.name;
    }

    // 4. Montar query com filtro de data opcional
    const { mode = 'all', startDate, endDate } = req.body || {};
    let baseUrl = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=100`;

    if (mode === 'range' && startDate && endDate) {
      const colData = resolveFieldName(mapping, 'DataOperacao');
      const filter = `fields/${colData} ge '${startDate}T00:00:00Z' and fields/${colData} le '${endDate}T23:59:59Z'`;
      baseUrl += `&$filter=${filter}`;
    }

    // 5. Buscar todos os itens com paginação
    let allItems: any[] = [];
    let nextUrl: string | null = baseUrl;

    while (nextUrl) {
      const data = await graphFetch(nextUrl, appToken);
      allItems = allItems.concat(data.value || []);
      nextUrl = data['@odata.nextLink'] || null;
    }

    console.log(`[MIGRATE_DEPARTURES] ${allItems.length} itens encontrados na lista ${LIST_NAME}`);

    // 6. Inserir cada item no PG
    let inserted = 0;
    let errors: string[] = [];

    for (const item of allItems) {
      const f = item.fields || {};
      try {
        const colData = resolveFieldName(mapping, 'DataOperacao');
        const colOp = resolveFieldName(mapping, 'Operacao');

        const dataBR = formatISOtoBR(f[colData]);
        if (!dataBR) { errors.push(`Item ${item.id}: sem data`); continue; }

        const row: Record<string, unknown> = {
          operacao: String(f[colOp] || '').trim(),
          rota: String(f.Title || '').trim(),
          motorista: String(f[resolveFieldName(mapping, 'Motorista')] || '').trim(),
          placa: String(f[resolveFieldName(mapping, 'Placa')] || '').trim(),
          contato: String(f[resolveFieldName(mapping, 'Contato')] || '').replace(/\D/g, ''),
          inicio: formatTime(f[resolveFieldName(mapping, 'HorarioInicio')]),
          saida: formatTime(f[resolveFieldName(mapping, 'HorarioSaida')]),
          statusGeral: String(f[resolveFieldName(mapping, 'StatusGeral')] || '').trim(),
          motivo: String(f[resolveFieldName(mapping, 'MotivoAtraso')] || '').trim(),
          observacao: String(f[resolveFieldName(mapping, 'Observacao')] || '').trim(),
          data: dataBR,
          statusOp: String(f[resolveFieldName(mapping, 'StatusOp')] || 'Previsto').trim(),
          checklistMotorista: String(f[resolveFieldName(mapping, 'ChecklistMotorista')] || '').trim(),
          retornoMotorista: String(f[resolveFieldName(mapping, 'RetornoMotorista')] || '').trim(),
          causaRaiz: String(f[resolveFieldName(mapping, 'CausaRaiz')] || '').trim(),
          tempoResposta: String(f[resolveFieldName(mapping, 'TempoResposta')] || '').trim(),
          logTempoResposta: String(f[resolveFieldName(mapping, 'LogTempoResposta')] || '').trim(),
        };

        await upsertDeparture(row);
        inserted++;
      } catch (err: any) {
        errors.push(`Item ${item.id}: ${err.message}`);
      }
    }

    return res.status(200).json({
      success: true,
      total: allItems.length,
      inserted,
      errors: errors.length > 0 ? errors.slice(0, 20) : undefined
    });
  } catch (error: any) {
    console.error('[MIGRATE_DEPARTURES] Erro:', error?.message || error);
    return res.status(500).json({ success: false, error: error?.message || 'Erro na migração' });
  }
}
