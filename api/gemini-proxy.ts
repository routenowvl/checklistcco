import type { VercelRequest, VercelResponse } from '@vercel/node';
import { GoogleGenAI, Type } from '@google/genai';

const validateToken = async (authHeader: string | undefined): Promise<boolean> => {
  if (!authHeader) return false;

  const token = authHeader.startsWith('Bearer ')
    ? authHeader.slice(7).trim()
    : authHeader.trim();

  if (!token) return false;

  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${token}` },
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

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ success: false, error: 'API key not configured' });
  }

  const { prompt, responseSchema } = req.body as {
    prompt?: string;
    responseSchema?: Record<string, unknown>;
  };

  if (!prompt) {
    return res.status(400).json({ success: false, error: 'Missing prompt' });
  }

  if (prompt.length > 100000) {
    return res.status(400).json({ success: false, error: 'Prompt too long' });
  }

  // Validação rigorosa do responseSchema
  const ALLOWED_ROOT_TYPES = ['ARRAY', 'OBJECT'];
  const ALLOWED_PROP_TYPES = ['STRING', 'NUMBER', 'BOOLEAN', 'INTEGER', 'ARRAY'];
  const MAX_SCHEMA_DEPTH = 4;
  const MAX_PROPERTIES = 30;

  function validateSchema(schema: any, depth: number): string | null {
    if (depth > MAX_SCHEMA_DEPTH) return 'Schema muito profundo (max 4 niveis)';
    if (!schema || typeof schema !== 'object') return 'Schema invalido';

    const type = String(schema.type || '').toUpperCase();

    if (type === 'ARRAY') {
      if (schema.items) {
        const err = validateSchema(schema.items, depth + 1);
        if (err) return err;
      }
      return null;
    }

    if (type === 'OBJECT') {
      const props = schema.properties;
      if (!props || typeof props !== 'object') return 'OBJECT requer properties';
      const keys = Object.keys(props);
      if (keys.length > MAX_PROPERTIES) return `Maximo de ${MAX_PROPERTIES} propriedades`;

      for (const key of keys) {
        const prop = props[key];
        if (!prop || typeof prop !== 'object') return `Propriedade '${key}' invalida`;

        const propType = String(prop.type || '').toUpperCase();
        if (!ALLOWED_PROP_TYPES.includes(propType)) {
          return `Tipo '${propType}' nao permitido na propriedade '${key}'`;
        }

        // Valida enum se presente
        if (prop.enum && !Array.isArray(prop.enum)) return `enum de '${key}' deve ser array`;

        // Valida nested items (para ARRAY dentro de OBJECT)
        if (propType === 'ARRAY' && prop.items) {
          const err = validateSchema(prop.items, depth + 1);
          if (err) return `Em '${key}': ${err}`;
        }
      }

      return null;
    }

    if (!ALLOWED_ROOT_TYPES.includes(type)) return `Tipo raiz '${type}' nao permitido`;
    return null;
  }

  if (responseSchema) {
    const schemaError = validateSchema(responseSchema, 0);
    if (schemaError) {
      console.warn('[GEMINI_PROXY] Schema rejeitado:', schemaError);
      return res.status(400).json({ success: false, error: `Schema invalido: ${schemaError}` });
    }
  }

  try {
    const ai = new GoogleGenAI({ apiKey });

    const config: Record<string, unknown> = {
      responseMimeType: 'application/json',
    };

    if (responseSchema) {
      config.responseSchema = responseSchema;
    }

    const response = await ai.models.generateContent({
      model: 'gemini-3-flash-preview',
      contents: prompt,
      config,
    });

    const text = response.text || null;

    return res.status(200).json({ success: true, text });
  } catch (error: any) {
    console.error('[GEMINI_PROXY] Error:', error?.message || error);
    return res.status(500).json({ success: false, error: 'Gemini API error' });
  }
}
