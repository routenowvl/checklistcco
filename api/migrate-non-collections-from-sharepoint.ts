import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getGraphAppToken } from '../utils/graphAppAuth';
import { insertNonCollection } from '../utils/checklistDb';

/**
 * Endpoint de migração: lê a lista Dados_Nao_Coletas do SharePoint
 * (ID: 83e8cfb9-1982-47ae-b515-3fec112da457) e copia todos os itens para
 * a tabela non_collections no PostgreSQL.
 *
 * Uso: POST /api/migrate-non-collections-from-sharepoint
 * Body (opcional): { "mode": "all" | "range", "startDate": "YYYY-MM-DD", "endDate": "YYYY-MM-DD" }
 *
 * DELETAR este arquivo após a migração.
 */

const SITE_PATH = process.env.VITE_SHAREPOINT_SITE_PATH || '';
const LIST_ID = '83e8cfb9-1982-47ae-b515-3fec112da457';

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

/** Converte data ISO do SharePoint para DD/MM/AAAA */
const formatISOtoBR = (iso: any): string => {
  if (!iso) return '';
  const s = String(iso).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : s;
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

    // 2. Obter mapeamento de colunas da lista Dados_Nao_Coletas
    const columns = await graphFetch(`/sites/${siteId}/lists/${LIST_ID}/columns`, appToken);
    const mapping: Record<string, string> = {};
    for (const col of columns.value || []) {
      mapping[normalizeString(col.name)] = col.name;
      mapping[normalizeString(col.displayName)] = col.name;
    }

    // 3. Montar query com filtro de data opcional
    const { mode = 'all', startDate, endDate } = req.body || {};
    let baseUrl = `/sites/${siteId}/lists/${LIST_ID}/items?expand=fields&$top=100`;

    if (mode === 'range' && startDate && endDate) {
      const colData = resolveFieldName(mapping, 'Data');
      const filter = `fields/${colData} ge '${startDate}T00:00:00Z' and fields/${colData} le '${endDate}T23:59:59Z'`;
      baseUrl += `&$filter=${filter}`;
    }

    // 4. Buscar todos os itens com paginação
    let allItems: any[] = [];
    let nextUrl: string | null = baseUrl;

    while (nextUrl) {
      const data = await graphFetch(nextUrl, appToken);
      allItems = allItems.concat(data.value || []);
      nextUrl = data['@odata.nextLink'] || null;
    }

    console.log(`[MIGRATE_NC] ${allItems.length} itens encontrados na lista Dados_Nao_Coletas`);

    // 5. Inserir cada item no PG
    let inserted = 0;
    let errors: string[] = [];

    for (const item of allItems) {
      const f = item.fields || {};
      try {
        const colOp = resolveFieldName(mapping, 'Operacao');
        const colObs = resolveFieldName(mapping, 'Observacao');

        // Campo Data pode ser DateTime do SharePoint — converte para BR
        const dataRaw = f[resolveFieldName(mapping, 'Data')];
        const dataBR = formatISOtoBR(dataRaw);
        if (!dataBR) { errors.push(`Item ${item.id}: sem data`); continue; }

        const row: Record<string, unknown> = {
          operacao: String(f[colOp] || f['Opera_x00e7__x00e3_o'] || '').trim(),
          rota: String(f.Title || '').trim(),
          data: dataBR,
          semana: String(f[resolveFieldName(mapping, 'Semana')] || f.Title || '').trim(),
          codigo: String(f[resolveFieldName(mapping, 'Codigo')] || f['C_x00f3_digo'] || '').trim(),
          produtor: String(f[resolveFieldName(mapping, 'Produtor')] || '').trim(),
          motivo: String(f[resolveFieldName(mapping, 'Motivo')] || '').trim(),
          observacao: String(f[colObs] || f['Observa_x00e7__x00e3_o'] || '').trim(),
          acao: String(f[resolveFieldName(mapping, 'Acao')] || f['A_x00e7__x00e3_o'] || '').trim(),
          dataAcao: formatISOtoBR(f[resolveFieldName(mapping, 'DataAcao')] || f['DataA_x00e7__x00e3_o']),
          ultimaColeta: formatISOtoBR(f[resolveFieldName(mapping, 'UltimaColeta')] || f['_x00da_ltimaColeta']),
          Culpabilidade: String(f[resolveFieldName(mapping, 'Culpabilidade')] || '').trim(),
          causaRaiz: String(f[resolveFieldName(mapping, 'CausaRaiz')] || '').trim(),
        };

        await insertNonCollection(row);
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
    console.error('[MIGRATE_NC] Erro:', error?.message || error);
    return res.status(500).json({ success: false, error: error?.message || 'Erro na migração' });
  }
}
