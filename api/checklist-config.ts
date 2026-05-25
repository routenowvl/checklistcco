import type { VercelRequest, VercelResponse } from '@vercel/node';
import {
  getAllConfigs,
  getConfigByOperacao,
  updateConfigField,
  updateConfigFields,
  updateConteudoIfChanged,
  updateConteudoNcoletasIfChanged,
  getLockStatus,
  acquireLock,
  releaseLock
} from '../utils/checklistDb';

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
        const configs = await getAllConfigs();
        return res.status(200).json({ success: true, configs });
      }

      case 'getByOperacao': {
        const { operacao } = req.body;
        if (!operacao) {
          return res.status(400).json({ success: false, error: 'operacao é obrigatório' });
        }
        const config = await getConfigByOperacao(String(operacao));
        return res.status(200).json({ success: true, config });
      }

      case 'updateField': {
        const { operacao, field, value } = req.body;
        if (!operacao || !field) {
          return res.status(400).json({ success: false, error: 'operacao e field são obrigatórios' });
        }
        await updateConfigField(String(operacao), String(field), value);
        return res.status(200).json({ success: true });
      }

      case 'updateFields': {
        const { operacao, fields } = req.body;
        if (!operacao || !fields || typeof fields !== 'object') {
          return res.status(400).json({ success: false, error: 'operacao e fields são obrigatórios' });
        }
        await updateConfigFields(String(operacao), fields);
        return res.status(200).json({ success: true });
      }

      case 'updateConteudoIfChanged': {
        const { operacao, conteudo } = req.body;
        if (!operacao || conteudo === undefined) {
          return res.status(400).json({ success: false, error: 'operacao e conteudo são obrigatórios' });
        }
        const changed = await updateConteudoIfChanged(String(operacao), String(conteudo));
        return res.status(200).json({ success: true, changed });
      }

      case 'updateConteudoNcoletasIfChanged': {
        const { operacao, conteudoNcoletas } = req.body;
        if (!operacao || conteudoNcoletas === undefined) {
          return res.status(400).json({ success: false, error: 'operacao e conteudoNcoletas são obrigatórios' });
        }
        const changed = await updateConteudoNcoletasIfChanged(String(operacao), String(conteudoNcoletas));
        return res.status(200).json({ success: true, changed });
      }

      case 'getLockStatus': {
        const { operacao } = req.body;
        if (!operacao) {
          return res.status(400).json({ success: false, error: 'operacao é obrigatório' });
        }
        const lockStatus = await getLockStatus(String(operacao));
        return res.status(200).json({ success: true, lock: lockStatus });
      }

      case 'acquireLock': {
        const { operacao, userEmail, timestamp } = req.body;
        if (!operacao || !userEmail || !timestamp) {
          return res.status(400).json({ success: false, error: 'operacao, userEmail e timestamp são obrigatórios' });
        }
        await acquireLock(String(operacao), String(userEmail), String(timestamp));
        return res.status(200).json({ success: true });
      }

      case 'releaseLock': {
        const { operacao } = req.body;
        if (!operacao) {
          return res.status(400).json({ success: false, error: 'operacao é obrigatório' });
        }
        await releaseLock(String(operacao));
        return res.status(200).json({ success: true });
      }

      default:
        return res.status(400).json({ success: false, error: `Ação desconhecida: ${action}` });
    }
  } catch (error: any) {
    console.error('[CHECKLIST_CONFIG] Erro:', error?.message || error);
    return res.status(500).json({ success: false, error: 'Erro ao processar operação de config' });
  }
}
