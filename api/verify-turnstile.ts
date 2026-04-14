import type { VercelRequest, VercelResponse } from '@vercel/node';

/**
 * Serverless Function: /api/verify-turnstile
 *
 * Valida o token do Cloudflare Turnstile server-side.
 * A SECRET_KEY fica segura no servidor (Vercel Environment Variables).
 */
export default async function handler(req: VercelRequest, res: VercelResponse) {
  // Apenas POST
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  const { token } = req.body as { token?: string };

  if (!token) {
    return res.status(400).json({ success: false, error: 'Token não fornecido' });
  }

  const secretKey = process.env.SECRET_KEY;

  if (!secretKey) {
    console.error('[TURNSTILE] SECRET_KEY não configurada no ambiente Vercel');
    return res.status(500).json({ success: false, error: 'Configuração inválida' });
  }

  try {
    // Chama a API do Cloudflare para verificar o token
    const verifyUrl = 'https://challenges.cloudflare.com/turnstile/v0/siteverify';

    const result = await fetch(verifyUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        secret: secretKey,
        response: token,
        // Opcional: validar remoteip para mais segurança
        remoteip: req.headers['x-real-ip'] || req.headers['x-forwarded-for'] as string,
      }),
    });

    const data = await result.json();

    console.log('[TURNSTILE] Cloudflare response:', JSON.stringify(data));

    if (data.success) {
      return res.status(200).json({ success: true });
    } else {
      // Cloudflare retornou erros
      const errors = data['error-codes'] || ['unknown_error'];
      console.warn('[TURNSTILE] Falha na verificação. Erros:', errors);
      return res.status(403).json({
        success: false,
        errors,
      });
    }
  } catch (error: any) {
    console.error('[TURNSTILE] Erro ao verificar token:', error.message);
    return res.status(500).json({ success: false, error: 'Erro interno na verificação' });
  }
}
