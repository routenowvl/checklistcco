
import { Task, TaskStatus, RouteDeparture } from "../types";

const PROXY_URL = `${window.location.origin}/api/gemini-proxy`;

type SchemaProperty = {
  type: string;
  enum?: string[];
};

type SchemaObject = {
  type: string;
  properties: Record<string, SchemaProperty>;
  required?: string[];
};

const callGeminiProxy = async (
  prompt: string,
  token: string,
  responseSchema?: { type: string; items: SchemaObject }
): Promise<string | null> => {
  const response = await fetch(PROXY_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${token}`,
    },
    body: JSON.stringify({ prompt, responseSchema }),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error || `Gemini proxy error: ${response.status}`);
  }

  const data = await response.json();
  return data.text || null;
};

export const parseExcelContentToTasks = async (
  rawText: string,
  token: string
): Promise<Partial<Task>[]> => {
  const prompt = `Analise o texto bruto de uma planilha e extraia as tarefas em JSON. Texto: """${rawText}"""`;

  try {
    const text = await callGeminiProxy(prompt, token, {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          title: { type: "STRING" },
          description: { type: "STRING" },
          priority: { type: "STRING", enum: ["Baixa", "Média", "Alta"] },
          category: { type: "STRING" },
          dueDate: { type: "STRING" },
        },
        required: ["title", "priority", "category"],
      },
    });

    if (text) {
      const parsed = JSON.parse(text.trim());
      return parsed.map((item: any) => ({
        ...item,
        status: TaskStatus.TODO,
        createdAt: new Date().toISOString(),
      }));
    }
    return [];
  } catch (error) {
    console.error("Error parsing tasks:", error);
    throw error;
  }
};

/**
 * Robust Manual Parser for Excel data
 * Handles DD/MM/YYYY dates and missing columns
 */
export const parseRouteDeparturesManual = (rawText: string): Partial<RouteDeparture>[] => {
  const lines = rawText.split(/\r?\n/);
  const result: Partial<RouteDeparture>[] = [];

  const convertDate = (dateStr: string) => {
    if (!dateStr) return '';
    // Regex para DD/MM/YYYY ou DD-MM-YYYY
    const ddmmyyyy = dateStr.match(/(\d{2})[\/\-](\d{2})[\/\-](\d{4})/);
    if (ddmmyyyy) return `${ddmmyyyy[3]}-${ddmmyyyy[2]}-${ddmmyyyy[1]}`;

    // Regex para YYYY-MM-DD
    const yyyymmdd = dateStr.match(/(\d{4})[\/\-](\d{2})[\/\-](\d{2})/);
    if (yyyymmdd) return `${yyyymmdd[1]}-${yyyymmdd[2]}-${yyyymmdd[3]}`;

    return '';
  };

  const formatTime = (timeStr: string) => {
    if (!timeStr || timeStr.trim() === '' || timeStr === '00:00:00') return '00:00:00';
    const timeMatch = timeStr.match(/(\d{2}:\d{2}(?::\d{2})?)/);
    if (timeMatch) {
        let t = timeMatch[1];
        if (t.length === 5) t += ':00';
        return t;
    }
    return '00:00:00';
  };

  const dateRegex = /(\d{2}[\/\-]\d{2}[\/\-]\d{4})/;
  const timeRegex = /(\d{2}:\d{2}(?::\d{2})?)/g;

  for (let line of lines) {
    line = line.trim();
    if (!line) continue;

    // Se a linha tiver uma data, é uma nova rota
    const dateMatch = line.match(dateRegex);

    if (dateMatch) {
      const dateStr = dateMatch[1];
      const tabs = line.split('\t');

      if (tabs.length >= 3) {
          // Caso padrão: Colado direto do Excel (Tabulado)
          result.push({
            rota: tabs[0] || '',
            data: convertDate(tabs[1] || ''),
            inicio: formatTime(tabs[2] || '00:00:00'),
            motorista: tabs[3] || '',
            placa: (tabs[4] || '').replace(/[\s-]/g, '').toUpperCase(),
            saida: formatTime(tabs[5] || '00:00:00'),
            motivo: tabs[6] || '',
            observacao: tabs[7] || '',
            operacao: tabs[8] || ''
          });
      } else {
          // Caso fallback: Espaços irregulares (Heurística baseada em âncoras)
          const dateIdx = line.indexOf(dateStr);
          const rota = line.substring(0, dateIdx).trim();
          const afterDate = line.substring(dateIdx + dateStr.length).trim();

          const times = afterDate.match(timeRegex) || [];
          const inicio = times[0] || '00:00:00';
          const saida = times[1] || '00:00:00';

          // Tenta extrair motorista e placa entre os horários ou após a data
          const parts = afterDate.split(/\s+/).filter(p => !p.match(timeRegex));
          const placa = parts.length > 0 ? parts[parts.length - 1].replace(/[\s-]/g, '').toUpperCase() : '';
          const motorista = parts.length > 1 ? parts.slice(0, -1).join(' ') : '';

          result.push({
            rota: rota,
            data: convertDate(dateStr),
            inicio: formatTime(inicio),
            motorista: motorista,
            placa: placa,
            saida: formatTime(saida),
            motivo: '',
            observacao: '',
            operacao: ''
          });
      }
    } else if (result.length > 0) {
      // Linha sem data: Provavelmente continuação da Observação da linha anterior
      const last = result[result.length - 1];
      last.observacao = (last.observacao + ' ' + line).trim();
    }
  }

  // Filtra rotas que não conseguiram converter a data corretamente (evita Invalid Date no SharePoint)
  return result.filter(r => r.rota && r.data && r.data !== '');
};

export const parseRouteDepartures = async (
  rawText: string,
  token: string
): Promise<Partial<RouteDeparture>[]> => {
  const prompt = `Extraia dados logísticos de: """${rawText}""". Formato: ROTA, DATA(YYYY-MM-DD), INICIO, MOTORISTA, PLACA, SAIDA, MOTIVO, OBSERVAÇÃO, OPERAÇÃO.`;

  try {
    const text = await callGeminiProxy(prompt, token, {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          rota: { type: "STRING" },
          data: { type: "STRING" },
          inicio: { type: "STRING" },
          motorista: { type: "STRING" },
          placa: { type: "STRING" },
          saida: { type: "STRING" },
          motivo: { type: "STRING" },
          observacao: { type: "STRING" },
          operacao: { type: "STRING" },
        },
        required: ["rota", "data"],
      },
    });

    return text
      ? JSON.parse(text.trim()).map((item: any) => ({
          ...item,
          placa: (item.placa || '').replace(/[\s-]/g, '').toUpperCase(),
        }))
      : [];
  } catch (error) {
    console.error("Gemini Error:", error);
    throw error;
  }
};
