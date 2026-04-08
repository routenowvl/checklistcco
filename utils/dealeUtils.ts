import { RouteConfig } from '../types';

// Operações que compõem o grupo DEALE
const DEALE_OPERATIONS = ['ARATIBA', 'CATUIPE', 'ALMIRANTE'];
const DEALE_DISPLAY_NAME = 'DEALE';
// Operação "âncora" onde as infos de envio são salvas no SharePoint
const DEALE_ANCHOR_OPERATION = 'ALMIRANTE';

/**
 * Verifica se as configurações do usuário contêm TODAS as 3 operações DEALE.
 * Só retorna true se ARATIBA, CATUIPE e ALMIRANTE estiverem presentes.
 */
export function isDealeUser(configs: RouteConfig[]): boolean {
  const userOps = new Set(configs.map(c => c.operacao.toUpperCase()));
  return DEALE_OPERATIONS.every(op => userOps.has(op));
}

/**
 * Retorna as operações reais do grupo DEALE.
 */
export function getDealeRealOperations(): string[] {
  return [...DEALE_OPERATIONS];
}

/**
 * Retorna o nome de exibição para uma operação.
 * Se for um usuário DEALE e a operação fizer parte do grupo, retorna "DEALE".
 */
export function getDealeDisplayName(operacao: string, isDealeUser: boolean): string {
  if (isDealeUser && DEALE_OPERATIONS.includes(operacao.toUpperCase())) {
    return DEALE_DISPLAY_NAME;
  }
  return operacao;
}

/**
 * Retorna a operação âncora (ALMIRANTE) onde as configs de email e último envio
 * devem ser salvas no SharePoint para usuários DEALE.
 */
export function getDealeAnchorOperation(): string {
  return DEALE_ANCHOR_OPERATION;
}

/**
 * Para usuários DEALE, retorna a configuração da operação âncora (ALMIRANTE).
 * Para usuários normais, retorna a configuração da operação informada.
 */
export function getDealeEffectiveConfig(
  operacao: string,
  configs: RouteConfig[],
  dealeUser: boolean
): RouteConfig | undefined {
  if (dealeUser && DEALE_OPERATIONS.includes(operacao.toUpperCase())) {
    // Para usuários DEALE, sempre usa a config de ALMIRANTE
    return configs.find(c => c.operacao.toUpperCase() === DEALE_ANCHOR_OPERATION);
  }
  return configs.find(c => c.operacao === operacao);
}

/**
 * Para usuários DEALE, agrupa ARATIBA+CATUIPE+ALMIRANTE em uma única entrada "DEALE",
 * mas MANTÉM todas as outras operações do usuário intactas.
 * 
 * Ex: ['LAT-CWB', 'ARATIBA', 'CATUIPE', 'ALMIRANTE'] → ['LAT-CWB', 'DEALE']
 */
export function getDealeFilteredConfigs(configs: RouteConfig[]): RouteConfig[] {
  // Separa: operações NÃO-DEALE e operações DEALE
  const nonDealeConfigs = configs.filter(c =>
    !DEALE_OPERATIONS.includes(c.operacao.toUpperCase())
  );

  // Pega a config da operação âncora (ALMIRANTE) para criar a entrada DEALE
  const anchorConfig = configs.find(c =>
    c.operacao.toUpperCase() === DEALE_ANCHOR_OPERATION
  );

  const result: RouteConfig[] = [...nonDealeConfigs];

  // Se encontrou a âncora, adiciona DEALE no lugar das 3 operações
  if (anchorConfig) {
    result.push({
      ...anchorConfig,
      operacao: DEALE_DISPLAY_NAME,
      nomeExibicao: DEALE_DISPLAY_NAME
    });
  }

  return result;
}

/**
 * Para usuários DEALE, filtra as rotas para agrupar as 3 operações sob DEALE.
 * As rotas mantêm sua operacao original internamente, mas são exibidas como DEALE.
 */
export function getDealeGroupedRoutes(
  departures: any[],
  dealeUser: boolean
): any[] {
  if (!dealeUser) return departures;

  // Para usuários DEALE, retorna todas as rotas das 3 operações
  // mas com operacao renomeada para "DEALE" no visual
  return departures
    .filter(d => DEALE_OPERATIONS.includes(d.operacao?.toUpperCase()))
    .map(d => ({
      ...d,
      _originalOperacao: d.operacao, // Preserva a operação original internamente
      operacao: DEALE_DISPLAY_NAME   // Nome de exibição
    }));
}

/**
 * Para envio: retorna as operações reais que devem ser enviadas quando
 * o usuário seleciona "DEALE".
 */
export function getDealeOperationsToSend(): string[] {
  return [...DEALE_OPERATIONS];
}

/**
 * Para usuários DEALE, retorna o último envio combinado das 3 operações
 * (pega o mais recente entre ARATIBA, CATUIPE e ALMIRANTE).
 */
export function getDealeCombinedLastEnvio(configs: RouteConfig[]): string {
  const dealeConfigs = configs.filter(c =>
    DEALE_OPERATIONS.includes(c.operacao.toUpperCase())
  );

  let latestEnvio = '';
  let latestTimestamp = 0;

  dealeConfigs.forEach(config => {
    if (config.ultimoEnvioSaida) {
      let parsedDate: Date | null = null;

      if (config.ultimoEnvioSaida.includes('T')) {
        parsedDate = new Date(config.ultimoEnvioSaida);
      } else if (config.ultimoEnvioSaida.includes('/')) {
        const [data, hora] = config.ultimoEnvioSaida.split(' ');
        const [dia, mes, ano] = data.split('/');
        const [h, m, s] = hora ? hora.split(':') : ['00', '00', '00'];
        parsedDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(h), Number(m), Number(s));
      }

      if (parsedDate && !isNaN(parsedDate.getTime()) && parsedDate.getTime() > latestTimestamp) {
        latestTimestamp = parsedDate.getTime();
        latestEnvio = config.ultimoEnvioSaida;
      }
    }
  });

  return latestEnvio;
}
