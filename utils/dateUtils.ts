/**
 * Utilitários de data com fuso horário de Brasília (America/Sao_Paulo)
 * 
 * O problema: new Date().toISOString() converte para UTC, causando inconsistências
 * quando o usuário está em UTC-3 e o servidor compara datas.
 * 
 * Solução: Sempre usar datas locais de Brasília para comparações e armazenamento.
 */

const BRAZIL_TIMEZONE = 'America/Sao_Paulo';

/**
 * Obtém a data atual no fuso de Brasília no formato YYYY-MM-DD
 * Ex: "2026-03-17"
 */
export function getBrazilDate(): string {
  const now = new Date();
  return now.toLocaleDateString('pt-BR', { timeZone: BRAZIL_TIMEZONE })
    .split('/')
    .reverse()
    .join('-');
}

/**
 * Obtém a data/hora atual no fuso de Brasília no formato YYYY-MM-DDTHH:mm:ss
 * Ex: "2026-03-17T10:00:00"
 */
export function getBrazilDateTime(): string {
  const now = new Date();
  const datePart = getBrazilDate();
  const timePart = now.toLocaleTimeString('pt-BR', { 
    timeZone: BRAZIL_TIMEZONE,
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false 
  });
  return `${datePart}T${timePart}`;
}

/**
 * Obtém a data/hora atual no fuso de Brasília no formato ISO completo
 * Ex: "2026-03-17T10:00:00.000-03:00"
 */
export function getBrazilISOString(): string {
  const now = new Date();
  const datePart = getBrazilDate();
  const timePart = now.toLocaleTimeString('pt-BR', { 
    timeZone: BRAZIL_TIMEZONE,
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    fractionalSecond: '3',
    hour12: false 
  });
  
  // Obtém o offset de Brasília em relação a UTC
  const utcDate = new Date(now.toLocaleString('en-US', { timeZone: 'UTC' }));
  const brazilDate = new Date(now.toLocaleString('en-US', { timeZone: BRAZIL_TIMEZONE }));
  const offsetMs = utcDate.getTime() - brazilDate.getTime();
  const offsetHours = Math.floor(offsetMs / 3600000);
  const offsetMins = Math.abs(offsetMs % 3600000) / 60000;
  const offsetSign = offsetHours <= 0 ? '+' : '-';
  const offsetStr = `${offsetSign}${String(Math.abs(offsetHours)).padStart(2, '0')}:${String(offsetMins).padStart(2, '0')}`;
  
  return `${datePart}T${timePart}${offsetStr}`;
}

/**
 * Obtém a hora atual no fuso de Brasília (0-23)
 * Usado para comparações de horário.
 */
export function getBrazilHours(): number {
  const now = new Date();
  const timeStr = now.toLocaleTimeString('pt-BR', {
    timeZone: BRAZIL_TIMEZONE,
    hour: '2-digit',
    minute: '2-digit',
    hour12: false
  });
  return parseInt(timeStr.split(':')[0], 10);
}

/**
 * Obtém os minutos atuais no fuso de Brasília (0-59)
 * Usado para comparações de horário mais precisas.
 */
export function getBrazilMinutes(): number {
  const now = new Date();
  const timeStr = now.toLocaleTimeString('pt-BR', {
    timeZone: BRAZIL_TIMEZONE,
    minute: '2-digit',
    hour12: false
  });
  return parseInt(timeStr.split(':')[1], 10);
}

/**
 * Verifica se já passou das 10:00h no fuso de Brasília
 * Retorna true se a hora atual for >= 10:00
 */
export function isAfter10amBrazil(): boolean {
  const now = new Date();
  const hours = getBrazilHours();
  const minutes = getBrazilMinutes();
  return hours >= 10;
}

/**
 * Obtém a hora e minuto atual no fuso de Brasília no formato HH:mm
 */
export function getBrazilTime(): string {
  const now = new Date();
  return now.toLocaleTimeString('pt-BR', { 
    timeZone: BRAZIL_TIMEZONE,
    hour: '2-digit',
    minute: '2-digit',
    hour12: false 
  });
}

/**
 * Converte uma string de data para o fuso de Brasília
 * @param dateString Data no formato YYYY-MM-DD ou ISO
 * @returns Data no formato YYYY-MM-DD (fuso de Brasília)
 */
export function toBrazilDate(dateString: string): string {
  if (!dateString) return getBrazilDate();
  
  const date = new Date(dateString);
  if (isNaN(date.getTime())) return getBrazilDate();
  
  return date.toLocaleDateString('pt-BR', { timeZone: BRAZIL_TIMEZONE })
    .split('/')
    .reverse()
    .join('-');
}

/**
 * Formata uma data ISO para exibição no formato brasileiro
 * @param isoString Data ISO
 * @returns Data formatada como DD/MM/YYYY HH:mm
 */
export function formatBrazilDateTime(isoString: string): string {
  if (!isoString) return '--/--/---- --:--';
  
  try {
    const date = new Date(isoString);
    if (isNaN(date.getTime())) return '--/--/---- --:--';
    
    return date.toLocaleDateString('pt-BR', { timeZone: BRAZIL_TIMEZONE }) + ' ' +
      date.toLocaleTimeString('pt-BR', { 
        timeZone: BRAZIL_TIMEZONE,
        hour: '2-digit',
        minute: '2-digit',
        hour12: false 
      });
  } catch {
    return '--/--/---- --:--';
  }
}

/**
 * Verifica se duas datas (YYYY-MM-DD) são iguais no fuso de Brasília
 */
export function isSameBrazilDate(date1: string, date2: string): boolean {
  return toBrazilDate(date1) === toBrazilDate(date2);
}

/**
 * Obtém o timestamp de uma data no fuso de Brasília
 * Usado para comparações e ordenações.
 */
export function getBrazilTimestamp(dateString: string): number {
  const date = new Date(dateString);
  const brazilDateStr = date.toLocaleDateString('pt-BR', { timeZone: BRAZIL_TIMEZONE });
  const [day, month, year] = brazilDateStr.split('/');
  return new Date(`${year}-${month}-${day}T00:00:00`).getTime();
}

/**
 * Obtém a data completa (YYYY-MM-DD) de um timestamp ISO no fuso de Brasília
 * @param isoString Data ISO (ex: "2026-03-17T10:00:00.000Z")
 * @returns Data no formato YYYY-MM-DD (fuso de Brasília)
 */
export function getBrazilDateFromISO(isoString: string): string {
  if (!isoString) return getBrazilDate();
  const date = new Date(isoString);
  return date.toLocaleDateString('pt-BR', { timeZone: BRAZIL_TIMEZONE })
    .split('/')
    .reverse()
    .join('-');
}

/**
 * Verifica se uma data ISO é do dia atual no fuso de Brasília
 * @param isoString Data ISO para verificar
 * @returns true se for hoje no fuso de Brasília
 */
export function isTodayBrazil(isoString: string): boolean {
  return getBrazilDateFromISO(isoString) === getBrazilDate();
}
