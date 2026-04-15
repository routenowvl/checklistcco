
import React, { useState, useEffect, useMemo } from 'react';
import { SharePointService } from '../services/sharepointService';
import { getValidToken } from '../services/tokenService';
import { getBrazilDate, getBrazilHours, isAfter10amBrazil } from '../utils/dateUtils';
import { isDealeUser, getDealeFilteredConfigs, getDealeAnchorOperation, getDealeOperationsToSend, getDealeRealOperations, getDealeCombinedLastEnvio, getDealeEffectiveConfig } from '../utils/dealeUtils';
import { RouteDeparture, Task, User, RouteConfig, ColetaPrevista } from '../types';
import {
  TowerControl, Send, RefreshCw, Loader2,
  CheckCircle2, AlertCircle, Clock, Filter,
  ChevronDown, X
} from 'lucide-react';

interface SummaryItem {
  id: string;
  operacao: string;
  timestamp: string;
  relativeTime: string;
  status: string;
  statusColor: string;
  webhookStatus?: string; // Status retornado pelo webhook: "OK" ou "Atualizar"
  ultimoEnvioFormatado?: string;
}

const SendReportView: React.FC<{ currentUser: User }> = ({ currentUser }) => {
  const [isLoading, setIsLoading] = useState(true);
  const [departures, setDepartures] = useState<RouteDeparture[]>([]);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [lastSync, setLastSync] = useState(new Date());

  // Estado para usuários DEALE (ARATIBA + CATUIPE + ALMIRANTE)
  const [isDeale, setIsDeale] = useState(false);

  // Estados para seleção e envio
  const [selectedOperacao, setSelectedOperacao] = useState<string>('');
  const [selectedOperacaoNC, setSelectedOperacaoNC] = useState<string>('');
  const [isSending, setIsSending] = useState(false);

  // Feedback separado para cada tipo de envio
  const [sendError, setSendError] = useState<string | null>(null);       // Saídas (lado esquerdo)
  const [sendSuccess, setSendSuccess] = useState(false);                  // Saídas (lado esquerdo)
  const [ncSendError, setNcSendError] = useState<string | null>(null);   // Não Coletas (lado direito)
  const [ncSendSuccess, setNcSendSuccess] = useState(false);              // Não Coletas (lado direito)
  const [isAtualizacao, setIsAtualizacao] = useState(false);
  const [isSendingSummary, setIsSendingSummary] = useState(false);
  const [isSendingNCSummary, setIsSendingNCSummary] = useState(false);
  const [coletasPrevistas, setColetasPrevistas] = useState<ColetaPrevista[]>([]);

  const WEBHOOK_URL = import.meta.env.VITE_WEBHOOK_SAIDAS_URL || "https://n8n.datastack.viagroup.com.br/webhook/8cb1f3e1-833d-42a7-a3f0-2f959ea390d6";
  const WEBHOOK_URL_NAO_COLETAS = "https://n8n.datastack.viagroup.com.br/webhook-test/d712d06e-b81f-40f4-9ca8-5b2403a90fdd";
  const WEBHOOK_URL_RESUMO = import.meta.env.VITE_WEBHOOK_RESUMO_URL || "https://n8n.datastack.viagroup.com.br/webhook/8cb1f3e1-833d-42a7-a3f0-2f959ea390d6";
  const WEBHOOK_URL_NC_RESUMO = "https://n8n.datastack.viagroup.com.br/webhook-test/20541afc-08c7-4799-b3e9-26dd3afdbb5a";

  // Estado para não coletas reais do SharePoint (usado na coluna da direita)
  const [realNonCollections, setRealNonCollections] = useState<any[]>([]);

  const fetchAllData = async (forceRefresh: boolean = false) => {
    setIsLoading(true);
    const token = await getValidToken();
    if (!token) return;

    try {
      console.log('[FETCH_ALL] Buscando dados completos...', forceRefresh ? '(force refresh)' : '');
      console.log('[FETCH_ALL] Usuário logado:', currentUser.email);

      const [depData, configs, spNonCollections] = await Promise.all([
        SharePointService.getDepartures(token, forceRefresh),
        SharePointService.getRouteConfigs(token, currentUser.email, forceRefresh),
        SharePointService.getNonCollections(token, currentUser.email)
      ]);

      console.log('[FETCH_ALL] Total de rotas brutas do SharePoint:', depData?.length || 0);
      console.log('[FETCH_ALL] Configurações carregadas:', configs?.length || 0);
      console.log('[FETCH_ALL] Operações do usuário:', configs?.map(c => c.operacao));

      // Detecta se é usuário DEALE (tem ARATIBA + CATUIPE + ALMIRANTE)
      const deale = isDealeUser(configs || []);
      setIsDeale(deale);
      console.log('[FETCH_ALL] É usuário DEALE?', deale);

      // FILTRA rotas APENAS das operações do usuário logado
      const myOps = new Set((configs || []).map(c => c.operacao));
      const filteredRoutes = (depData || []).filter(route => {
        // Se não houver operações configuradas, retorna array vazio (segurança)
        if (myOps.size === 0) {
          console.warn('[FETCH_ALL] ⚠️ Nenhuma operação configurada para este usuário - retornando array vazio');
          return false;
        }
        return myOps.has(route.operacao);
      });

      console.log('[FETCH_ALL] Rotas filtradas por usuário:', filteredRoutes.length);
      console.log('[FETCH_ALL] Operações nas rotas filtradas:', Array.from(new Set(filteredRoutes.map(r => r.operacao))));

      setDepartures(filteredRoutes);

      // Manter TODAS as configs originais do usuário
      setUserConfigs(configs);

      // Salvar não coletas reais do SharePoint
      setRealNonCollections(spNonCollections || []);
      console.log('[FETCH_ALL] Não coletas reais carregadas:', (spNonCollections || []).length);

      setLastSync(new Date());
      console.log('[FETCH_ALL] Dados atualizados com sucesso');
    } catch (e) {
      console.error("Erro ao carregar resumo:", e);
    } finally {
      setIsLoading(false);
    }
  };

  // Polling para atualizar apenas as configs (últimos envios) a cada 5 segundos
  useEffect(() => {
    const fetchConfigsOnly = async (force: boolean = false) => {
      const token = await getValidToken();
      if (!token) return;

      try {
        console.log('[POLLING_CONFIGS] Buscando configs atualizadas...', force ? '(force refresh)' : '');
        const configs = await SharePointService.getRouteConfigs(token, currentUser.email, force);
        console.log('[DEBUG_CONFIGS] Configs carregadas:', configs);
        configs.forEach(c => {
          console.log(`[DEBUG_CONFIG] ${c.operacao}: ultimoEnvioSaida = "${c.ultimoEnvioSaida}" | Status = "${c.Status}" | Envio = "${c.Envio}" | Copia = "${c.Copia}" | UltimoEnvioResumoSaida = "${c.UltimoEnvioResumoSaida}" | StatusResumoSaida = "${c.StatusResumoSaida}"`);
        });

        // Detecta DEALE e mantém estado atualizado
        const deale = isDealeUser(configs || []);
        setIsDeale(deale);

        // Manter TODAS as configs originais
        setUserConfigs(configs);

        setLastSync(new Date());
      } catch (e) {
        console.error("Erro ao atualizar configs:", e);
      }
    };

    // Polling para atualizar não coletas reais a cada 10 segundos
    const fetchNonCollectionsOnly = async () => {
      const token = await getValidToken();
      if (!token) return;

      try {
        const spNC = await SharePointService.getNonCollections(token, currentUser.email);
        setRealNonCollections(spNC || []);
      } catch (e) {
        console.error("Erro ao atualizar não coletas:", e);
      }
    };

    // Carrega dados iniciais com force refresh
    fetchAllData(true);

    // Polling das configs a cada 15 segundos (otimizado para reduzir chamadas à API)
    const configsInterval = setInterval(() => fetchConfigsOnly(false), 15000);

    // Polling de não coletas a cada 30 segundos (otimizado para reduzir chamadas à API)
    const ncInterval = setInterval(() => fetchNonCollectionsOnly(), 30000);

    // Atualização completa a cada 60 segundos (otimizado para reduzir chamadas à API)
    const fullInterval = setInterval(() => {
      console.log('[POLLING] Atualização completa dos dados');
      fetchAllData(false); // Sem forceRefresh — usa cache do serviço
    }, 60000);

    return () => {
      clearInterval(configsInterval);
      clearInterval(ncInterval);
      clearInterval(fullInterval);
    };
  }, []);

  // Define o estado inicial do botão de Atualização baseado no horário atual (Brasília)
  useEffect(() => {
    // Entre 12:00h e 23:59h (Brasília) = ativado, entre 00:00h e 11:59h = desativado
    setIsAtualizacao(isAfter10amBrazil());
  }, []);

  // Função para enviar saídas da operação selecionada
  const handleSendDepartures = async () => {
    if (!selectedOperacao) {
      setSendError("Selecione uma operação para enviar.");
      setTimeout(() => setSendError(null), 3000);
      return;
    }

    // Para usuários DEALE, a operação selecionada "DEALE" corresponde às 3 operações reais
    // Para operações NÃO-DEALE, trata normalmente
    const isDealeSelection = isDeale && selectedOperacao === 'DEALE';
    const realOperations = isDealeSelection
      ? getDealeOperationsToSend()
      : [selectedOperacao];

    // Para DEALE, a operação âncora (ALMIRANTE) é usada para trava e configs
    // Para operações normais, usa a própria operação selecionada
    const anchorOperation = isDealeSelection ? getDealeAnchorOperation() : selectedOperacao;

    // VALIDAÇÃO DE SEGURANÇA: Verifica se a operação selecionada pertence ao usuário
    // Para DEALE, verifica se o usuário tem TODAS as 3 operações do grupo
    const myOps = new Set(userConfigs.map(c => c.operacao));
    if (isDealeSelection) {
      // Para DEALE, verifica se o usuário tem todas as operações do grupo
      const hasAllDealeOps = getDealeOperationsToSend().every(op => myOps.has(op));
      if (!hasAllDealeOps) {
        console.error(`[SEND_DEPARTURES_BLOCKED] Usuário tentou enviar DEALE sem ter todas as operações necessárias`);
        setSendError(`Erro: Você não tem permissão para enviar esta operação.`);
        setTimeout(() => setSendError(null), 5000);
        return;
      }
    } else if (!myOps.has(selectedOperacao)) {
      console.error(`[SEND_DEPARTURES_BLOCKED] Usuário tentou enviar operação não pertencente: ${selectedOperacao}`);
      setSendError(`Erro: Você não tem permissão para enviar esta operação.`);
      setTimeout(() => setSendError(null), 5000);
      return;
    }

    // VERIFICA TRAVA PARA EVITAR ENVIO DUPLICADO (usa operação âncora para DEALE)
    const token = await getValidToken() || currentUser.accessToken;
    if (!token) {
      setSendError("Erro de autenticação. Tente novamente.");
      return;
    }

    // Verifica se já há envio em andamento para esta operação
    const lockResult = await SharePointService.checkSendLock(token, anchorOperation);
    if (lockResult && lockResult.locked && !lockResult.expired) {
      const errorMsg = `⚠️ Já existe envio em andamento para ${selectedOperacao} por ${lockResult.user}. Aguarde alguns segundos e tente novamente.`;
      setSendError(errorMsg);
      console.warn('[SEND_DEPARTURES] Envio bloqueado por trava:', lockResult.user);
      setTimeout(() => setSendError(null), 8000);
      return;
    }

    // Adquire trava (na operação âncora para DEALE)
    const acquireResult = await SharePointService.acquireSendLock(token, anchorOperation, currentUser.email);
    if (!acquireResult.success) {
      setSendError(`⚠️ Não foi possível adquirir trava para ${selectedOperacao}: ${acquireResult.message}. Tente novamente.`);
      setTimeout(() => setSendError(null), 8000);
      return;
    }

    console.log(`[SEND_DEPARTURES] ✅ Trava adquirida para ${anchorOperation} (selecionado: ${selectedOperacao})`);

    setIsSending(true);
    setSendError(null);

    // Para DEALE, pega todas as rotas das 3 operações
    const selectedDepartures = departures.filter(d => realOperations.includes(d.operacao));
    // Configs: para DEALE (selecionando DEALE), usa config de ALMIRANTE; para ops normais, usa a config da própria operação
    const config = isDealeSelection
      ? userConfigs.find(c => c.operacao.toUpperCase() === getDealeAnchorOperation())
      : userConfigs.find(c => c.operacao === selectedOperacao);

    // Debug DEALE: log completo de todas as configs do usuário
    if (isDealeSelection) {
      console.log('[SEND_DEPARTURES][DEALE_DEBUG] Todas as userConfigs:', userConfigs.map(c => ({
        operacao: c.operacao,
        Envio: c.Envio,
        Copia: c.Copia,
        nomeExibicao: c.nomeExibicao
      })));
      console.log('[SEND_DEPARTURES][DEALE_DEBUG] Config encontrada para ALMIRANTE:', config);
      if (!config) {
        console.error('[SEND_DEPARTURES][DEALE_DEBUG] ❌ Config de ALMIRANTE NÃO encontrada!');
      }
    }

    console.log('[SEND_DEPARTURES] === ENVIANDO SAÍDAS ===');
    console.log('[SEND_DEPARTURES] Operação selecionada:', selectedOperacao);
    console.log('[SEND_DEPARTURES] Operações reais sendo enviadas:', realOperations);
    console.log('[SEND_DEPARTURES] Rotas encontradas:', selectedDepartures.length);
    console.log('[SEND_DEPARTURES] Config (efetiva):', config);
    console.log('[SEND_DEPARTURES] Config - Envio:', config?.Envio, '| Copia:', config?.Copia);

    if (selectedDepartures.length === 0) {
      setSendError("Nenhuma saída encontrada para esta operação.");
      setTimeout(() => setSendError(null), 3000);
      setIsSending(false);
      return;
    }

    // VERIFICA SE HÁ ROTAS PENDENTES DE SAIR NO DIA (saida vazia) em TODAS as operações reais
    // Se houver, o status deve ser "Atualizar" em vez de "OK"
    const today = getBrazilDate();
    const hasPendingRoute = selectedDepartures.some(d => {
      const routeDate = d.data || '';
      if (routeDate !== today) return false;

      // Verifica se a coluna saida está vazia (nula, undefined, string vazia, ou apenas espaços)
      // IMPORTANTE: "00:00:00" é um horário válido (meia-noite) e NÃO é considerado vazio
      // Se tiver "-" na coluna saida, considera como rota que já saiu (não é pendente)
      const saidaVazia = !d.saida || d.saida.trim() === '';

      return saidaVazia;
    });

    // Determina o status baseado na verificação de rotas pendentes
    const statusDeterminado = hasPendingRoute ? 'Atualizar' : 'OK';
    console.log(`[SEND_DEPARTURES] Status determinado para ${selectedOperacao}: ${statusDeterminado} (hasPendingRoute: ${hasPendingRoute})`);

    const payload = {
      tipo: "SAIDA",
      operacao: isDealeSelection ? realOperations.join(',') : selectedOperacao,
      nomeExibicao: config?.nomeExibicao || selectedOperacao,
      tolerancia: config?.tolerancia || "00:00:00",
      atualizacao: isAtualizacao ? "sim" : "não",
      usuario: currentUser.email,
      dataEnvio: new Date().toISOString(),
      envio: config?.Envio || "",
      copia: config?.Copia || "",
      saidas: selectedDepartures.map(d => ({
        rota: d.rota,
        data: d.data,
        inicio: d.inicio,
        motorista: d.motorista,
        placa: d.placa,
        saida: d.saida,
        motivo: d.motivo,
        observacao: d.observacao,
        status: d.statusOp,
        operacaoReal: d.operacao
      }))
    };

    try {
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (response.ok) {
        // Processa a resposta do webhook para pegar data/hora de envio
        let responseData;
        try {
          responseData = await response.json();
        } catch (jsonError) {
          console.warn('[WEBHOOK] Resposta não é JSON válido, usando dados do payload:', jsonError);
          // Se o webhook não retorna JSON, usa os dados do payload
          responseData = { sucesso: true, data: new Date().toLocaleDateString('pt-BR'), horario: new Date().toLocaleTimeString('pt-BR') };
        }
        
        console.log('[WEBHOOK_RESPONSE]', responseData);

        // Tenta pegar a data/hora de envio de diferentes campos possíveis
        // Agora o webhook retorna data e hora separados: dataEnvioEmail e horarioEnvioEmail
        let dataHoraEnvio = '';
        
        const dataEnvio = 
          responseData[0]?.dataEnvioEmail ||
          responseData[0]?.data ||
          responseData.dataEnvioEmail ||
          responseData.data;
        
        const horarioEnvio = 
          responseData[0]?.horarioEnvioEmail ||
          responseData[0]?.horario ||
          responseData.horarioEnvioEmail ||
          responseData.horario;
        
        // Junta data e hora no formato DD/MM/YYYY HH:MM:SS
        if (dataEnvio && horarioEnvio) {
          dataHoraEnvio = `${dataEnvio} ${horarioEnvio}`;
        } else if (dataEnvio) {
          // Se só tem data, adiciona horário zerado
          dataHoraEnvio = `${dataEnvio} 00:00:00`;
        } else if (horarioEnvio) {
          // Se só tem hora, usa data atual (fuso de Brasília)
          const hoje = getBrazilDate();
          dataHoraEnvio = `${hoje} ${horarioEnvio}`;
        }

        console.log('[DEBUG_DATA] dataEnvio:', dataEnvio, 'horarioEnvio:', horarioEnvio, 'dataHoraEnvio final:', dataHoraEnvio);

        // Se o webhook retornou data/hora de envio, atualiza no SharePoint
        // Para DEALE, salva na operação âncora (ALMIRANTE)
        if (dataHoraEnvio) {
          const token = await getValidToken() || currentUser.accessToken;
          if (token) {
            try {
              const opParaSalvar = isDealeSelection ? anchorOperation : selectedOperacao;
              console.log(`[ULTIMO_ENVIO] Enviando para atualização: ${dataHoraEnvio} (operação: ${opParaSalvar})`);
              await SharePointService.updateUltimoEnvioSaida(
                token,
                opParaSalvar,
                dataHoraEnvio
              );
              console.log(`[ULTIMO_ENVIO] ✅ Atualizado com sucesso na operação ${opParaSalvar}: ${dataHoraEnvio}`);

              // Para DEALE, também atualiza as outras 2 operações
              if (isDealeSelection) {
                const realOps = getDealeRealOperations();
                for (const op of realOps) {
                  if (op !== anchorOperation) {
                    try {
                      await SharePointService.updateUltimoEnvioSaida(token, op, dataHoraEnvio);
                      console.log(`[ULTIMO_ENVIO_DEALE] ✅ Atualizado também ${op}: ${dataHoraEnvio}`);
                    } catch (err) {
                      console.warn(`[ULTIMO_ENVIO_DEALE] Falha ao atualizar ${op}:`, err);
                    }
                  }
                }
              }
            } catch (err: any) {
              console.error('Erro ao atualizar UltimoEnvioSaida:', err.message);
            }
          }
        } else {
          console.warn('[WEBHOOK] Campo de data/hora de envio não encontrado na resposta');
        }

        // Processa e salva o status retornado pelo webhook OU o status determinado localmente
        const webhookStatus = responseData[0]?.status || responseData.status;

        // Usa o status do webhook se disponível, senão usa o status determinado localmente
        let statusFinal = '';

        if (webhookStatus) {
          // Webhook retornou status - usa o retorno
          statusFinal = webhookStatus.toLowerCase() === 'atualizar' ? 'Atualizar' :
                        webhookStatus.toLowerCase() === 'ok' ? 'OK' : webhookStatus;
          console.log('[STATUS_WEBHOOK] Status retornado pelo webhook:', statusFinal);
        } else {
          // Webhook não retornou status - usa o status determinado localmente
          statusFinal = statusDeterminado;
          console.log('[STATUS_WEBHOOK] Webhook não retornou status - usando status determinado localmente:', statusFinal);
        }

        const token = await getValidToken() || currentUser.accessToken;
        if (token) {
          try {
            const opParaSalvar = isDealeSelection ? anchorOperation : selectedOperacao;
            console.log(`[STATUS] Atualizando Status no SharePoint para ${opParaSalvar}:`, statusFinal);
            await SharePointService.updateStatusOperacao(
              token,
              opParaSalvar,
              statusFinal
            );
            console.log(`[STATUS] ✅ Status atualizado no SharePoint para ${opParaSalvar}:`, statusFinal);

            // Para DEALE, também atualiza as outras 2 operações
            if (isDealeSelection) {
              const realOps = getDealeRealOperations();
              for (const op of realOps) {
                if (op !== anchorOperation) {
                  try {
                    await SharePointService.updateStatusOperacao(token, op, statusFinal);
                    console.log(`[STATUS_DEALE] ✅ Status atualizado também para ${op}:`, statusFinal);
                  } catch (err) {
                    console.warn(`[STATUS_DEALE] Falha ao atualizar ${op}:`, err);
                  }
                }
              }
            }
          } catch (err: any) {
            console.error('Erro ao atualizar Status:', err.message);
          }
        }

        setSendSuccess(true);
        setTimeout(() => setSendSuccess(false), 3000);
      } else {
        throw new Error(`Erro na resposta do webhook: ${response.status}`);
      }
    } catch (error: any) {
      console.error("Erro ao enviar webhook:", error);
      setSendError(error.message || "Falha ao enviar dados.");
      setTimeout(() => setSendError(null), 5000);
    } finally {
      setIsSending(false);

      // Libera trava após o envio (sucesso ou erro) - usa operação âncora para DEALE
      const token = await getValidToken() || currentUser.accessToken;
      if (token) {
        const opParaLiberar = isDealeSelection ? anchorOperation : selectedOperacao;
        await SharePointService.releaseSendLock(token, opParaLiberar);
        console.log(`[SEND_DEPARTURES] 🔓 Trava liberada para ${opParaLiberar}`);
      }
    }
  };

  // Função para enviar não coletas
  const handleSendNonCollections = async () => {
    if (!selectedOperacaoNC) {
      setNcSendError("Selecione uma operação para enviar.");
      setTimeout(() => setNcSendError(null), 3000);
      return;
    }

    // Para usuários DEALE, a operação selecionada "DEALE" corresponde às 3 operações reais
    // Para operações NÃO-DEALE, trata normalmente
    const isDealeSelectionNC = isDeale && selectedOperacaoNC === 'DEALE';
    const realOperations = isDealeSelectionNC
      ? getDealeOperationsToSend()
      : [selectedOperacaoNC];

    // Para DEALE, a operação âncora (ALMIRANTE) é usada para trava e configs
    // Para operações normais, usa a própria operação selecionada
    const anchorOperation = isDealeSelectionNC ? getDealeAnchorOperation() : selectedOperacaoNC;

    // VALIDAÇÃO DE SEGURANÇA: Verifica se a operação selecionada pertence ao usuário
    // Para DEALE, verifica se o usuário tem TODAS as 3 operações do grupo
    const myOps = new Set(userConfigs.map(c => c.operacao));
    if (isDealeSelectionNC) {
      // Para DEALE, verifica se o usuário tem todas as operações do grupo
      const hasAllDealeOps = getDealeOperationsToSend().every(op => myOps.has(op));
      if (!hasAllDealeOps) {
        console.error(`[SEND_NAO_COLETA_BLOCKED] Usuário tentou enviar DEALE sem ter todas as operações necessárias`);
        setNcSendError(`Erro: Você não tem permissão para enviar esta operação.`);
        setTimeout(() => setNcSendError(null), 5000);
        return;
      }
    } else if (!myOps.has(selectedOperacaoNC)) {
      console.error(`[SEND_NAO_COLETA_BLOCKED] Usuário tentou enviar operação não pertencente: ${selectedOperacaoNC}`);
      setNcSendError(`Erro: Você não tem permissão para enviar esta operação.`);
      setTimeout(() => setNcSendError(null), 5000);
      return;
    }

    // VERIFICA TRAVA PARA EVITAR ENVIO DUPLICADO (usa operação âncora para DEALE)
    const token = await getValidToken() || currentUser.accessToken;
    if (!token) {
      setNcSendError("Erro de autenticação. Tente novamente.");
      return;
    }

    // Verifica se já há envio em andamento para esta operação
    const lockResult = await SharePointService.checkSendLock(token, anchorOperation);
    if (lockResult && lockResult.locked && !lockResult.expired) {
      const errorMsg = `⚠️ Já existe envio em andamento para ${selectedOperacaoNC} por ${lockResult.user}. Aguarde alguns segundos e tente novamente.`;
      setNcSendError(errorMsg);
      console.warn('[SEND_NAO_COLETA] Envio bloqueado por trava:', lockResult.user);
      setTimeout(() => setNcSendError(null), 8000);
      return;
    }

    // Adquire trava (na operação âncora para DEALE)
    const acquireResult = await SharePointService.acquireSendLock(token, anchorOperation, currentUser.email);
    if (!acquireResult.success) {
      setNcSendError(`⚠️ Não foi possível adquirir trava para ${selectedOperacaoNC}: ${acquireResult.message}. Tente novamente.`);
      setTimeout(() => setNcSendError(null), 8000);
      return;
    }

    console.log(`[SEND_NAO_COLETA] ✅ Trava adquirida para ${anchorOperation} (selecionado: ${selectedOperacaoNC})`);

    setIsSending(true);
    setNcSendError(null);

    // Para DEALE, pega todas as rotas das 3 operações
    const selectedDepartures = departures.filter(d => realOperations.includes(d.operacao));
    // Configs: para DEALE (selecionando DEALE), usa config de ALMIRANTE; para ops normais, usa a config da própria operação
    const config = isDealeSelectionNC
      ? userConfigs.find(c => c.operacao.toUpperCase() === getDealeAnchorOperation())
      : userConfigs.find(c => c.operacao === selectedOperacaoNC);

    console.log('[SEND_NAO_COLETA] === ENVIANDO NÃO COLETAS ===');
    console.log('[SEND_NAO_COLETA] Operação selecionada:', selectedOperacaoNC);
    console.log('[SEND_NAO_COLETA] Operações reais sendo enviadas:', realOperations);
    console.log('[SEND_NAO_COLETA] Total de rotas:', selectedDepartures.length);
    console.log('[SEND_NAO_COLETA] Config (efetiva):', config);
    console.log('[SEND_NAO_COLETA] Config - Envio:', config?.Envio, '| Copia:', config?.Copia);

    // BUSCA NÃO COLETAS REAIS NO SHAREPOINT PARA ESTA OPERAÇÃO
    const spNonCollections = await SharePointService.getNonCollections(token, currentUser.email);
    const ncFiltradas = spNonCollections.filter(nc => realOperations.includes(nc.operacao));

    console.log('[SEND_NAO_COLETA] Não coletas encontradas no SharePoint:', ncFiltradas.length);

    // BUSCA COLETAS PREVISTAS DA DATA DAS NÃO COLETAS PARA INCLUIR NO PAYLOAD
    // Pega a data da primeira não coleta encontrada (todas são da mesma data)
    const ncDate = ncFiltradas[0]?.data || '';
    let coletasPrev: ColetaPrevista[] = coletasPrevistas;
    try {
      // Converte DD/MM/YYYY → YYYY-MM-DD para a busca no SharePoint
      let dataISO = '';
      if (ncDate.includes('/')) {
        const [dia, mes, ano] = ncDate.split('/');
        dataISO = `${ano}-${mes}-${dia}`;
      } else if (ncDate.includes('-')) {
        dataISO = ncDate;
      }

      if (dataISO) {
        coletasPrev = await SharePointService.getColetasPrevistas(token, dataISO, currentUser.email);
        setColetasPrevistas(coletasPrev);
        console.log('[SEND_NAO_COLETA] Coletas previstas para data:', dataISO, '— encontradas:', coletasPrev.length);
      }
    } catch (err) {
      console.warn('[SEND_NAO_COLETA] Erro ao buscar coletas previstas, usando cache:', err);
    }

    // Soma coletas previstas por operação
    const coletasPorOperacao: Record<string, number> = {};
    realOperations.forEach(op => {
      const previstas = coletasPrev.filter(c => c.Title === op);
      coletasPorOperacao[op] = previstas.reduce((sum, c) => sum + (c.QntColeta || 0), 0);
    });

    const totalColetasPrevistas = Object.values(coletasPorOperacao).reduce((sum, n) => sum + n, 0);

    console.log('[SEND_NAO_COLETA] Coletas previstas por operação:', coletasPorOperacao);
    console.log('[SEND_NAO_COLETA] Total coletas previstas:', totalColetasPrevistas);

    if (ncFiltradas.length === 0) {
      setNcSendError(`⚠️ Nenhuma não coleta lançada para ${selectedOperacaoNC} na tabela. Não há dados para enviar.`);
      setTimeout(() => setNcSendError(null), 6000);
      setIsSending(false);
      // Libera a trava já que não há dados para enviar
      await SharePointService.releaseSendLock(token, anchorOperation);
      return;
    }

    // Não coletas da tabela de saídas (rotas NOK) — usado para contexto adicional
    const nonCollections = selectedDepartures.filter(d => d.statusGeral === 'NOK');

    // VERIFICA SE HÁ ROTAS PENDENTES DE SAIR NO DIA (saida vazia)
    // Se houver, o status deve ser "Atualizar" em vez de "OK"
    const today = getBrazilDate();
    const hasPendingRoute = nonCollections.some(d => {
      const routeDate = d.data || '';
      if (routeDate !== today) return false;
      
      // Verifica se a coluna saida está vazia (nula, undefined, string vazia, ou apenas espaços)
      // IMPORTANTE: "00:00:00" é um horário válido (meia-noite) e NÃO é considerado vazio
      // Se tiver "-" na coluna saida, considera como rota que já saiu (não é pendente)
      const saidaVazia = !d.saida || d.saida.trim() === '';
      
      return saidaVazia;
    });

    // Determina o status baseado na verificação de rotas pendentes
    const statusDeterminado = hasPendingRoute ? 'Atualizar' : 'OK';
    console.log(`[SEND_NAO_COLETA] Status determinado para ${selectedOperacaoNC}: ${statusDeterminado} (hasPendingRoute: ${hasPendingRoute})`);

    const payload = {
      tipo: "NAO_COLETA",
      operacao: isDealeSelectionNC ? realOperations.join(',') : selectedOperacaoNC,
      nomeExibicao: config?.nomeExibicao || selectedOperacaoNC,
      tolerancia: config?.tolerancia || "00:00:00",
      atualizacao: isAtualizacao ? "sim" : "não",
      usuario: currentUser.email,
      dataEnvio: new Date().toISOString(),
      envio: config?.Envio || "",
      copia: config?.Copia || "",
      totalColetasPrevistas: totalColetasPrevistas,
      coletasPorOperacao: coletasPorOperacao,
      totalNaoColetas: ncFiltradas.length,
      naoColetas: ncFiltradas.map(nc => ({
        semana: nc.semana,
        rota: nc.rota,
        data: nc.data,
        codigo: nc.codigo,
        produtor: nc.produtor,
        motivo: nc.motivo,
        observacao: nc.observacao,
        acao: nc.acao,
        dataAcao: nc.dataAcao,
        ultimaColeta: nc.ultimaColeta,
        culpabilidade: nc.Culpabilidade,
        operacao: nc.operacao
      }))
    };

    try {
      const response = await fetch(WEBHOOK_URL_NAO_COLETAS, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (response.ok) {
        // Processa a resposta do webhook para pegar data/hora de envio
        let responseData;
        try {
          responseData = await response.json();
        } catch (jsonError) {
          console.warn('[WEBHOOK_NAO_COLETAS] Resposta não é JSON válido, usando dados do payload:', jsonError);
          // Se o webhook não retorna JSON, usa os dados do payload
          responseData = { sucesso: true, data: new Date().toLocaleDateString('pt-BR'), horario: new Date().toLocaleTimeString('pt-BR') };
        }
        
        console.log('[WEBHOOK_RESPONSE_NAO_COLETAS]', responseData);

        // Tenta pegar a data/hora de envio de diferentes campos possíveis
        // Agora o webhook retorna data e hora separados: dataEnvioEmail e horarioEnvioEmail
        let dataHoraEnvio = '';
        
        const dataEnvio = 
          responseData[0]?.dataEnvioEmail ||
          responseData[0]?.data ||
          responseData.dataEnvioEmail ||
          responseData.data;
        
        const horarioEnvio = 
          responseData[0]?.horarioEnvioEmail ||
          responseData[0]?.horario ||
          responseData.horarioEnvioEmail ||
          responseData.horario;
        
        // Junta data e hora no formato DD/MM/YYYY HH:MM:SS
        if (dataEnvio && horarioEnvio) {
          dataHoraEnvio = `${dataEnvio} ${horarioEnvio}`;
        } else if (dataEnvio) {
          // Se só tem data, adiciona horário zerado
          dataHoraEnvio = `${dataEnvio} 00:00:00`;
        } else if (horarioEnvio) {
          // Se só tem hora, usa data atual (fuso de Brasília)
          const hoje = getBrazilDate();
          dataHoraEnvio = `${hoje} ${horarioEnvio}`;
        }

        console.log('[DEBUG_DATA_NAO_COLETAS] dataEnvio:', dataEnvio, 'horarioEnvio:', horarioEnvio, 'dataHoraEnvio final:', dataHoraEnvio);

        // Se o webhook retornou data/hora de envio, atualiza no SharePoint
        // Para DEALE, salva na operação âncora (ALMIRANTE)
        if (dataHoraEnvio) {
          const token = await getValidToken() || currentUser.accessToken;
          if (token) {
            try {
              const opParaSalvar = isDealeSelectionNC ? anchorOperation : selectedOperacaoNC;
              console.log(`[ULTIMO_ENVIO_NAO_COLETAS] Enviando para atualização: ${dataHoraEnvio} (operação: ${opParaSalvar})`);
              await SharePointService.updateUltimoEnvioNaoColetas(
                token,
                opParaSalvar,
                dataHoraEnvio
              );
              console.log(`[ULTIMO_ENVIO_NAO_COLETAS] ✅ Atualizado com sucesso na operação ${opParaSalvar}: ${dataHoraEnvio}`);

              // Para DEALE, também atualiza as outras 2 operações
              if (isDealeSelectionNC) {
                const realOps = getDealeRealOperations();
                for (const op of realOps) {
                  if (op !== anchorOperation) {
                    try {
                      await SharePointService.updateUltimoEnvioNaoColetas(token, op, dataHoraEnvio);
                      console.log(`[ULTIMO_ENVIO_DEALE_NC] ✅ Atualizado também ${op}: ${dataHoraEnvio}`);
                    } catch (err) {
                      console.warn(`[ULTIMO_ENVIO_DEALE_NC] Falha ao atualizar ${op}:`, err);
                    }
                  }
                }
              }
            } catch (err: any) {
              console.error('Erro ao atualizar UltimoEnvioNaoColetas:', err.message);
            }
          }
        } else {
          console.warn('[WEBHOOK] Campo de data/hora de envio não encontrado na resposta');
        }

        // Processa e salva o status retornado pelo webhook OU o status determinado localmente
        const webhookStatus = responseData[0]?.status || responseData.status;

        // Usa o status do webhook se disponível, senão usa o status determinado localmente
        let statusFinal = '';

        if (webhookStatus) {
          // Webhook retornou status - usa o retorno
          statusFinal = webhookStatus.toLowerCase() === 'atualizar' ? 'Atualizar' :
                        webhookStatus.toLowerCase() === 'ok' ? 'OK' : webhookStatus;
          console.log('[STATUS_WEBHOOK_NAO_COLETAS] Status retornado pelo webhook:', statusFinal);
        } else {
          // Webhook não retornou status - usa o status determinado localmente
          statusFinal = statusDeterminado;
          console.log('[STATUS_WEBHOOK_NAO_COLETAS] Webhook não retornou status - usando status determinado localmente:', statusFinal);
        }

        const token = await getValidToken() || currentUser.accessToken;
        if (token) {
          try {
            const opParaSalvar = isDealeSelectionNC ? anchorOperation : selectedOperacaoNC;
            console.log(`[STATUS_NAO_COLETAS] Atualizando Status no SharePoint para ${opParaSalvar}:`, statusFinal);
            await SharePointService.updateStatusOperacao(
              token,
              opParaSalvar,
              statusFinal
            );
            console.log(`[STATUS_NAO_COLETAS] ✅ Status atualizado no SharePoint para ${opParaSalvar}:`, statusFinal);

            // Para DEALE, também atualiza as outras 2 operações
            if (isDealeSelectionNC) {
              const realOps = getDealeRealOperations();
              for (const op of realOps) {
                if (op !== anchorOperation) {
                  try {
                    await SharePointService.updateStatusOperacao(token, op, statusFinal);
                    console.log(`[STATUS_DEALE_NC] ✅ Status atualizado também para ${op}:`, statusFinal);
                  } catch (err) {
                    console.warn(`[STATUS_DEALE_NC] Falha ao atualizar ${op}:`, err);
                  }
                }
              }
            }
          } catch (err: any) {
            console.error('Erro ao atualizar Status:', err.message);
          }
        }

        setNcSendSuccess(true);
        setTimeout(() => setNcSendSuccess(false), 3000);
      } else {
        throw new Error(`Erro na resposta do webhook: ${response.status}`);
      }
    } catch (error: any) {
      console.error("[SEND_NAO_COLETA] Erro ao enviar webhook:", error);
      setNcSendError(error.message || "Falha ao enviar dados.");
      setTimeout(() => setNcSendError(null), 5000);
    } finally {
      setIsSending(false);

      // Libera trava após o envio (sucesso ou erro) - usa operação âncora para DEALE
      const token = await getValidToken() || currentUser.accessToken;
      if (token) {
        const opParaLiberar = isDealeSelectionNC ? anchorOperation : selectedOperacaoNC;
        await SharePointService.releaseSendLock(token, opParaLiberar);
        console.log(`[SEND_NAO_COLETA] 🔓 Trava liberada para ${opParaLiberar}`);
      }
    }
  };

  // Função para enviar resumo GERAL de todas as operações
  const handleSendSummary = async () => {
    // FILTRA rotas APENAS das operações do usuário logado (validação de segurança)
    const myOps = new Set(userConfigs.map(c => c.operacao));

    console.log('[RESUMO_GERAL] === INICIANDO ENVIO DE RESUMO ===');
    console.log('[RESUMO_GERAL] Operações configuradas para este usuário:', userConfigs.map(c => c.operacao));
    console.log('[RESUMO_GERAL] Total de rotas no estado departures:', departures.length);

    const userRoutes = departures.filter(r => {
      const pertence = !r.operacao || myOps.has(r.operacao);
      if (!pertence) {
        console.log(`[RESUMO_GERAL] 🚫 Rota ${r.rota} da operação ${r.operacao} NÃO pertence a este usuário - será ignorada`);
      }
      return pertence;
    });

    console.log('[RESUMO_GERAL] ✅ Rotas filtradas para envio (apenas do usuário):', userRoutes.length);
    console.log('[RESUMO_GERAL] Operações nas rotas filtradas:', Array.from(new Set(userRoutes.map(r => r.operacao))));

    if (userRoutes.length === 0) {
      setSendError("Não há rotas das suas operações para enviar.");
      setTimeout(() => setSendError(null), 3000);
      return;
    }

    setIsSendingSummary(true);
    setSendError(null);

    // Agrupa rotas por operação (apenas do usuário logado)
    const routesByOperation: Record<string, RouteDeparture[]> = {};
    userRoutes.forEach(r => {
      if (!routesByOperation[r.operacao]) {
        routesByOperation[r.operacao] = [];
      }
      routesByOperation[r.operacao].push(r);
    });

    console.log('[RESUMO_GERAL] 📦 Operações sendo enviadas:', Object.keys(routesByOperation));
    console.log('[RESUMO_GERAL] 📊 Total de rotas por operação:', 
      Object.entries(routesByOperation).map(([op, rotas]) => `${op}: ${rotas.length}`)
    );

    // Prepara payload com todas as operações DO USUÁRIO
    const payload = {
      tipo: "RESUMO_GERAL",  // Tipo para resumo geral de todas as operações
      usuario: currentUser.email,
      dataEnvio: new Date().toISOString(),
      totalRotas: userRoutes.length,
      operacoes: Object.keys(routesByOperation).length,
      rotasPorOperacao: Object.entries(routesByOperation)
        .filter(([operacao, _]) => myOps.has(operacao)) // FILTRO EXTRA DE SEGURANÇA
        .map(([operacao, rotas]) => {
        const config = userConfigs.find(c => c.operacao === operacao);
        return {
          operacao: operacao,
          nomeExibicao: config?.nomeExibicao || operacao,
          tolerancia: config?.tolerancia || "00:00:00",
          envio: config?.Envio || "",
          copia: config?.Copia || "",
          totalRotas: rotas.length,
          rotasOK: rotas.filter(r => r.statusOp === 'OK').length,
          rotasPrevistas: rotas.filter(r => r.statusOp === 'Previsto').length,
          rotasAtrasadas: rotas.filter(r => r.statusOp === 'Atrasada').length,
          saidas: rotas.map(r => ({
            rota: r.rota,
            data: r.data,
            inicio: r.inicio,
            motorista: r.motorista,
            placa: r.placa,
            saida: r.saida,
            motivo: r.motivo,
            observacao: r.observacao,
            status: r.statusOp
          }))
        };
      })
    };

    // LOG FINAL DE CONFIRMAÇÃO ANTES DO ENVIO
    console.log('[RESUMO_GERAL] ========================================');
    console.log('[RESUMO_GERAL] 📋 RESUMO FINAL DO PAYLOAD:');
    console.log('[RESUMO_GERAL]    Usuário:', currentUser.email);
    console.log('[RESUMO_GERAL]    Total de rotas:', payload.totalRotas);
    console.log('[RESUMO_GERAL]    Total de operações:', payload.operacoes);
    console.log('[RESUMO_GERAL]    Operações sendo enviadas:');
    payload.rotasPorOperacao.forEach((op: any) => {
      console.log(`[RESUMO_GERAL]      - ${op.operacao}: ${op.totalRotas} rotas`);
    });
    console.log('[RESUMO_GERAL] ========================================');

    console.log('[RESUMO_GERAL] Enviando payload:', payload);

    try {
      const response = await fetch(WEBHOOK_URL_RESUMO, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (response.ok) {
        let responseData;
        try {
          responseData = await response.json();
        } catch {
          console.warn("[RESUMO_GERAL] Resposta não é JSON válido");
          responseData = { sucesso: true };
        }

        console.log('[RESUMO_GERAL] Resposta recebida:', responseData);

        // Processa a data/hora de envio retornada pelo webhook
        const dataEnvio = responseData[0]?.dataEnvioEmail || responseData.dataEnvioEmail || responseData.data;
        const horarioEnvio = responseData[0]?.horarioEnvioEmail || responseData.horarioEnvioEmail || responseData.horario;
        const statusRetorno = responseData[0]?.status || responseData.status; // "atualizar" ou "ok"
        
        let dataHoraEnvio = '';
        if (dataEnvio && horarioEnvio) {
          dataHoraEnvio = `${dataEnvio} ${horarioEnvio}`;
        } else if (dataEnvio) {
          dataHoraEnvio = `${dataEnvio} 00:00:00`;
        } else if (horarioEnvio) {
          dataHoraEnvio = `${new Date().toLocaleDateString('pt-BR')} ${horarioEnvio}`;
        }

        // Normaliza o status retornado pelo webhook
        let statusResumo = '';
        if (statusRetorno) {
          const statusLower = statusRetorno.toLowerCase().trim();
          if (statusLower === 'atualizar') {
            statusResumo = 'Atualizar'; // Azul
          } else if (statusLower === 'ok') {
            statusResumo = 'OK'; // Verde
          }
        }

        console.log('[RESUMO_GERAL] Status retornado:', statusRetorno, '→ Status resumo:', statusResumo);

        // Atualiza UltimoEnvioResumoSaida e StatusResumoSaida para todas as operações enviadas
        if (dataHoraEnvio) {
          const token = await getValidToken() || currentUser.accessToken;
          if (token) {
            console.log('[ULTIMO_ENVIO_RESUMO] Atualizando para operações:', Object.keys(routesByOperation));
            
            // Atualiza para cada operação
            for (const operacao of Object.keys(routesByOperation)) {
              try {
                // Atualiza data/hora
                await SharePointService.updateUltimoEnvioResumoSaida(token, operacao, dataHoraEnvio);
                console.log(`[ULTIMO_ENVIO_RESUMO] ✅ UltimoEnvioResumoSaida atualizado para ${operacao}: ${dataHoraEnvio}`);
                
                // Atualiza status se houver
                if (statusResumo) {
                  await SharePointService.updateStatusResumoSaida(token, operacao, statusResumo);
                  console.log(`[ULTIMO_ENVIO_RESUMO] ✅ StatusResumoSaida atualizado para ${operacao}: ${statusResumo}`);
                }
              } catch (err: any) {
                console.error(`[ULTIMO_ENVIO_RESUMO] Erro ao atualizar ${operacao}:`, err.message);
              }
            }
          }
        }
      } else {
        throw new Error(`Erro na resposta do webhook: ${response.status}`);
      }
    } catch (e: any) {
      console.error('[RESUMO_GERAL] Erro ao enviar:', e.message);
      setSendError(e.message || "Falha ao enviar resumo");
      setTimeout(() => setSendError(null), 5000);
    } finally {
      setIsSendingSummary(false);
    }
  };

  // Função para enviar resumo GERAL de todas as não coletas do usuário
  const handleSendNCSummary = async () => {
    const myOps = new Set(userConfigs.map(c => c.operacao));

    console.log('[RESUMO_NC] === INICIANDO ENVIO DE RESUMO NÃO COLETAS ===');
    console.log('[RESUMO_NC] Operações configuradas para este usuário:', userConfigs.map(c => c.operacao));
    console.log('[RESUMO_NC] Total de não coletas no estado realNonCollections:', realNonCollections.length);

    if (myOps.size === 0) {
      setNcSendError("Nenhuma operação configurada para este usuário.");
      setTimeout(() => setNcSendError(null), 3000);
      return;
    }

    setIsSendingNCSummary(true);
    setNcSendError(null);

    // Filtra não coletas apenas das operações do usuário logado
    const userNCs = realNonCollections.filter(nc => {
      const pertence = !nc.operacao || myOps.has(nc.operacao);
      if (!pertence) {
        console.log(`[RESUMO_NC] Não coleta da operação ${nc.operacao} NÃO pertence a este usuário - será ignorada`);
      }
      return pertence;
    });

    // Busca coletas previstas da DATA das não coletas (usa a data da primeira NC encontrada)
    const ncDate = userNCs.length > 0 ? userNCs[0].data : '';
    let coletasPrev: ColetaPrevista[] = coletasPrevistas;
    try {
      const token = await getValidToken();
      if (token) {
        // Converte DD/MM/YYYY → YYYY-MM-DD para a busca no SharePoint
        let dataISO = '';
        if (ncDate.includes('/')) {
          const [dia, mes, ano] = ncDate.split('/');
          dataISO = `${ano}-${mes}-${dia}`;
        } else if (ncDate.includes('-')) {
          dataISO = ncDate;
        } else {
          // Fallback: usa data de hoje
          dataISO = getBrazilDate();
        }

        coletasPrev = await SharePointService.getColetasPrevistas(token, dataISO, currentUser.email);
        setColetasPrevistas(coletasPrev);
        console.log('[RESUMO_NC] Coletas previstas para data:', dataISO, '— encontradas:', coletasPrev.length);
      }
    } catch (err) {
      console.warn('[RESUMO_NC] Erro ao buscar coletas previstas, usando cache:', err);
    }

    // Agrupa não coletas por operação
    const ncsByOperation: Record<string, any[]> = {};
    userNCs.forEach(nc => {
      const op = nc.operacao || 'SEM_OPERACAO';
      if (!ncsByOperation[op]) {
        ncsByOperation[op] = [];
      }
      ncsByOperation[op].push(nc);
    });

    // Lista com os nomes de todas as operações do login
    const listaOperacoes = userConfigs.map(c => c.operacao);

    // Envia TODAS as operações do login — inclusive as que não têm não coletas
    const payload = {
      tipo: "RESUMO_NAO_COLETAS",
      usuario: currentUser.email,
      dataEnvio: new Date().toISOString(),
      totalNaoColetas: userNCs.length,
      operacoes: listaOperacoes.length,
      operacoesLogin: listaOperacoes,
      naoColetasPorOperacao: Array.from(myOps).map(operacao => {
        const config = userConfigs.find(c => c.operacao === operacao);
        const ncs = ncsByOperation[operacao] || [];
        const previstas = coletasPrev.filter(c => c.Title === operacao);
        const totalPrevistas = previstas.reduce((sum, c) => sum + (c.QntColeta || 0), 0);
        return {
          operacao: operacao,
          nomeExibicao: config?.nomeExibicao || operacao,
          totalColetasPrevistas: totalPrevistas,
          totalNaoColetas: ncs.length,
          naoColetas: ncs.map(nc => ({
            semana: nc.semana,
            rota: nc.rota,
            data: nc.data,
            codigo: nc.codigo,
            produtor: nc.produtor,
            motivo: nc.motivo,
            observacao: nc.observacao,
            acao: nc.acao,
            dataAcao: nc.dataAcao,
            ultimaColeta: nc.ultimaColeta,
            culpabilidade: nc.Culpabilidade,
            operacao: nc.operacao
          }))
        };
      })
    };

    console.log('[RESUMO_NC] Operações sendo enviadas:', payload.naoColetasPorOperacao.map((op: any) => `${op.operacao}: ${op.totalNaoColetas} NCs`));
    console.log('[RESUMO_NC] Enviando payload:', payload);

    try {
      const response = await fetch(WEBHOOK_URL_NC_RESUMO, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (response.ok) {
        console.log('[RESUMO_NC] ✅ Resumo de não coletas enviado com sucesso');
        setNcSendSuccess(true);
        setTimeout(() => setNcSendSuccess(false), 3000);
      } else {
        throw new Error(`Erro na resposta do webhook: ${response.status}`);
      }
    } catch (e: any) {
      console.error('[RESUMO_NC] Erro ao enviar:', e.message);
      setNcSendError(e.message || "Falha ao enviar resumo de não coletas");
      setTimeout(() => setNcSendError(null), 5000);
    } finally {
      setIsSendingNCSummary(false);
    }
  };

  const getRelativeTime = (dateStr: string) => {
    if (!dateStr) return "há -- horas";
    const date = new Date(dateStr);
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    return `há ${diffHours} horas`;
  };

  // Função auxiliar para formatar data/hora para exibição
  const formatarDataHora = (dataISO: string) => {
    try {
      let date: Date;
      if (dataISO.includes('T')) {
        date = new Date(dataISO);
      } else if (dataISO.includes('/')) {
        const [data, hora] = dataISO.split(' ');
        const [dia, mes, ano] = data.split('/');
        const [h, m, s] = hora ? hora.split(':') : ['00', '00', '00'];
        date = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(h), Number(m), Number(s));
      } else {
        date = new Date(dataISO);
      }

      if (isNaN(date.getTime())) return "Data inválida";

      const dataFormatada = date.toLocaleDateString('pt-BR');
      const horaFormatada = date.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
      return `${dataFormatada} ${horaFormatada}`;
    } catch {
      return "Data inválida";
    }
  };

  // Lógica para processar a lista de SAÍDAS - agrupa DEALE se aplicável
  const departuresSummary = useMemo(() => {
    console.log('[DEBUG_SUMMARY] userConfigs:', userConfigs);

    // Separa configs DEALE e não-DEALE
    const dealeOps = new Set(['ARATIBA', 'CATUIPE', 'ALMIRANTE']);
    const dealeConfigs = userConfigs.filter(c => dealeOps.has(c.operacao.toUpperCase()));
    const nonDealeConfigs = userConfigs.filter(c => !dealeOps.has(c.operacao.toUpperCase()));

    const result: SummaryItem[] = [];

    // Adiciona operações não-DEALE normalmente
    nonDealeConfigs.forEach(config => {
      const ultimoEnvio = config.ultimoEnvioSaida || "";
      const webhookStatus = config.Status || "";

      let timestamp: string;
      let relativeTime: string;

      if (ultimoEnvio) {
        let parsedDate: Date | null = null;
        if (ultimoEnvio.includes('T')) {
          parsedDate = new Date(ultimoEnvio);
        } else if (ultimoEnvio.includes('/')) {
          const [data, hora] = ultimoEnvio.split(' ');
          const [dia, mes, ano] = data.split('/');
          const [h, m, s] = hora ? hora.split(':') : ['00', '00', '00'];
          parsedDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(h), Number(m), Number(s));
        } else {
          parsedDate = new Date(ultimoEnvio);
        }

        if (parsedDate && !isNaN(parsedDate.getTime())) {
          timestamp = parsedDate.toISOString();
          const now = new Date();
          const diffMs = now.getTime() - parsedDate.getTime();
          const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
          relativeTime = `há ${diffHours} horas`;
        } else {
          timestamp = new Date().toISOString();
          relativeTime = "há -- horas";
        }
      } else {
        timestamp = new Date().toISOString();
        relativeTime = "Nunca enviado";
      }

      let status = "PREVISTO";
      let color = "bg-slate-300 text-slate-600";

      const todayBrazil = getBrazilDate();
      const [todayY, todayM, todayD] = todayBrazil.split('-').map(Number);
      const today = new Date(todayY, todayM - 1, todayD);
      today.setHours(0, 0, 0, 0);

      let envioDateObj: Date | null = null;
      if (ultimoEnvio) {
        if (ultimoEnvio.includes('T')) {
          envioDateObj = new Date(ultimoEnvio);
        } else if (ultimoEnvio.includes('/')) {
          const [data, hora] = ultimoEnvio.split(' ');
          const [dia, mes, ano] = data.split('/');
          const [h, m, s] = hora ? hora.split(':') : ['00', '00', '00'];
          envioDateObj = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(h), Number(m), Number(s));
        }
      }

      if (envioDateObj && !isNaN(envioDateObj.getTime())) {
        envioDateObj.setHours(0, 0, 0, 0);
        const isToday = envioDateObj.getTime() === today.getTime();

        if (isToday && webhookStatus) {
          if (webhookStatus.toUpperCase() === 'OK') {
            status = "OK";
            color = "bg-emerald-500 text-white";
          } else if (webhookStatus.toUpperCase() === 'ATUALIZAR') {
            status = "ATUALIZAR";
            color = "bg-blue-500 text-white";
          } else {
            status = webhookStatus.toUpperCase();
            color = "bg-slate-500 text-white";
          }
        }
      }

      result.push({
        id: config.operacao,
        operacao: config.operacao,
        timestamp,
        relativeTime,
        status,
        statusColor: color,
        webhookStatus,
        ultimoEnvioFormatado: ultimoEnvio ? formatarDataHora(ultimoEnvio) : "Nunca"
      });
    });

    // Se é usuário DEALE e tem configs DEALE, agrupa em uma entrada DEALE
    if (dealeConfigs.length > 0) {
      // Pega o último envio mais recente entre as 3 operações
      const ultimoEnvio = getDealeCombinedLastEnvio(dealeConfigs);

      // Pega o webhookStatus da operação âncora (ALMIRANTE)
      const anchorConfig = dealeConfigs.find(c => c.operacao.toUpperCase() === 'ALMIRANTE');
      const webhookStatus = anchorConfig?.Status || "";

      let timestamp: string;
      let relativeTime: string;

      if (ultimoEnvio) {
        let parsedDate: Date | null = null;
        if (ultimoEnvio.includes('T')) {
          parsedDate = new Date(ultimoEnvio);
        } else if (ultimoEnvio.includes('/')) {
          const [data, hora] = ultimoEnvio.split(' ');
          const [dia, mes, ano] = data.split('/');
          const [h, m, s] = hora ? hora.split(':') : ['00', '00', '00'];
          parsedDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(h), Number(m), Number(s));
        } else {
          parsedDate = new Date(ultimoEnvio);
        }

        if (parsedDate && !isNaN(parsedDate.getTime())) {
          timestamp = parsedDate.toISOString();
          const now = new Date();
          const diffMs = now.getTime() - parsedDate.getTime();
          const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
          relativeTime = `há ${diffHours} horas`;
        } else {
          timestamp = new Date().toISOString();
          relativeTime = "há -- horas";
        }
      } else {
        timestamp = new Date().toISOString();
        relativeTime = "Nunca enviado";
      }

      let status = "PREVISTO";
      let color = "bg-slate-300 text-slate-600";

      const todayBrazil = getBrazilDate();
      const [todayY, todayM, todayD] = todayBrazil.split('-').map(Number);
      const today = new Date(todayY, todayM - 1, todayD);
      today.setHours(0, 0, 0, 0);

      let envioDateObj: Date | null = null;
      if (ultimoEnvio) {
        if (ultimoEnvio.includes('T')) {
          envioDateObj = new Date(ultimoEnvio);
        } else if (ultimoEnvio.includes('/')) {
          const [data, hora] = ultimoEnvio.split(' ');
          const [dia, mes, ano] = data.split('/');
          const [h, m, s] = hora ? hora.split(':') : ['00', '00', '00'];
          envioDateObj = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(h), Number(m), Number(s));
        }
      }

      if (envioDateObj && !isNaN(envioDateObj.getTime())) {
        envioDateObj.setHours(0, 0, 0, 0);
        const isToday = envioDateObj.getTime() === today.getTime();

        if (isToday && webhookStatus) {
          if (webhookStatus.toUpperCase() === 'OK') {
            status = "OK";
            color = "bg-emerald-500 text-white";
          } else if (webhookStatus.toUpperCase() === 'ATUALIZAR') {
            status = "ATUALIZAR";
            color = "bg-blue-500 text-white";
          } else {
            status = webhookStatus.toUpperCase();
            color = "bg-slate-500 text-white";
          }
        }
      }

      result.push({
        id: 'DEALE',
        operacao: 'DEALE',
        timestamp,
        relativeTime,
        status,
        statusColor: color,
        webhookStatus,
        ultimoEnvioFormatado: ultimoEnvio ? formatarDataHora(ultimoEnvio) : "Nunca"
      });
    }

    return result;
  }, [departures, userConfigs]);

  // Lógica para processar a lista de NÃO COLETAS (agrupa DEALE se aplicável)
  // Usa NÃO COLETAS REAIS do SharePoint, não rotas com status NOK
  const nonCollectionsSummary = useMemo(() => {
    const dealeOps = new Set(['ARATIBA', 'CATUIPE', 'ALMIRANTE']);
    const allOps = Array.from(new Set(userConfigs.map(c => c.operacao)));
    const nonDealeOps = allOps.filter(op => !dealeOps.has(op.toUpperCase()));
    const hasDealeOps = allOps.some(op => dealeOps.has(op.toUpperCase()));

    const result = nonDealeOps.map(op => {
      // Conta não coletas REAIS desta operação
      const ncCount = realNonCollections.filter(nc => nc.operacao === op).length;

      return {
        id: op,
        operacao: op,
        timestamp: new Date().toISOString(),
        relativeTime: getRelativeTime(new Date().toISOString()),
        status: ncCount > 0 ? `${ncCount} NÃO COLETAS` : "TODOS COLETADOS",
        statusColor: ncCount > 0 ? "bg-red-500 text-white" : "bg-emerald-500 text-white"
      };
    });

    // Adiciona DEALE agrupado se aplicável
    if (hasDealeOps) {
      const dealeNcCount = realNonCollections.filter(nc => dealeOps.has(nc.operacao?.toUpperCase())).length;

      result.push({
        id: 'DEALE',
        operacao: 'DEALE',
        timestamp: new Date().toISOString(),
        relativeTime: getRelativeTime(new Date().toISOString()),
        status: dealeNcCount > 0 ? `${dealeNcCount} NÃO COLETAS` : "TODOS COLETADOS",
        statusColor: dealeNcCount > 0 ? "bg-red-500 text-white" : "bg-emerald-500 text-white"
      });
    }

    return result;
  }, [realNonCollections, userConfigs]);

  // Status do Resumo (pega da operação com UltimoEnvioResumoSaida mais recente)
  const resumoStatus = useMemo(() => {
    if (userConfigs.length === 0) return { status: 'NÃO ENVIADO', color: 'bg-slate-500 text-white', data: '' };
    
    // Filtra configs que tem UltimoEnvioResumoSaida e StatusResumoSaida
    const comResumo = userConfigs.filter(c => c.UltimoEnvioResumoSaida && c.StatusResumoSaida);
    
    if (comResumo.length === 0) {
      return { status: 'NÃO ENVIADO', color: 'bg-slate-500 text-white', data: '' };
    }
    
    // Ordena por UltimoEnvioResumoSaida (mais recente primeiro)
    comResumo.sort((a, b) => {
      const dateA = new Date(a.UltimoEnvioResumoSaida!).getTime();
      const dateB = new Date(b.UltimoEnvioResumoSaida!).getTime();
      return dateB - dateA;
    });
    
    // Pega o mais recente
    const maisRecente = comResumo[0];
    const statusResumo = maisRecente.StatusResumoSaida || '';
    const dataEnvio = maisRecente.UltimoEnvioResumoSaida || '';
    
    if (!statusResumo) {
      return { status: 'NÃO ENVIADO', color: 'bg-slate-500 text-white', data: dataEnvio };
    } else if (statusResumo.toUpperCase() === 'OK') {
      return { status: 'OK', color: 'bg-emerald-500 text-white', data: dataEnvio };
    } else if (statusResumo.toUpperCase() === 'ATUALIZAR') {
      return { status: 'ATUALIZAR', color: 'bg-blue-500 text-white', data: dataEnvio };
    }
    
    return { status: 'NÃO ENVIADO', color: 'bg-slate-500 text-white', data: dataEnvio };
  }, [userConfigs]);

  // Lista de operações para os selects de envio (agrupa DEALE se aplicável)
  const sendOptions = useMemo(() => {
    const dealeOps = new Set(['ARATIBA', 'CATUIPE', 'ALMIRANTE']);
    const allOps = Array.from(new Set(userConfigs.map(c => c.operacao)));
    const nonDealeOps = allOps.filter(op => !dealeOps.has(op.toUpperCase()));
    const hasDealeOps = allOps.some(op => dealeOps.has(op.toUpperCase()));

    const options = nonDealeOps.map(op => ({ value: op, label: op }));
    if (hasDealeOps) {
      options.push({ value: 'DEALE', label: 'DEALE' });
    }
    return options;
  }, [userConfigs]);

  if (isLoading && departures.length === 0) {
    return (
      <div className="h-full flex flex-col items-center justify-center text-primary-500 gap-4">
        <Loader2 size={48} className="animate-spin" />
        <p className="font-black text-xs uppercase tracking-widest">Sincronizando Resumo...</p>
      </div>
    );
  }

  return (
    <div className="h-full flex flex-col bg-[#F8FAFC] dark:bg-slate-950 p-4 font-sans overflow-hidden">
      {/* Header Estilo Print */}
      <div className="flex justify-between items-center mb-6 px-2">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-[#075985] text-white rounded-xl shadow-lg">
            <TowerControl size={24} />
          </div>
          <div>
            <h1 className="text-xl font-black text-[#075985] dark:text-sky-400 uppercase tracking-tight">
              Envio de Saídas e Não Coletas
            </h1>
            <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest flex items-center gap-2">
              {isLoading ? (
                <><Loader2 size={12} className="animate-spin text-blue-500" /> Atualizando...</>
              ) : (
                <><CheckCircle2 size={12} className="text-green-500" /> Atualizado às {lastSync.toLocaleTimeString()}</>
              )}
            </p>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => {
              console.log('[USER_ACTION] Refresh manual acionado');
              fetchAllData(true);
            }}
            className="p-2 text-slate-400 hover:text-primary-600 transition-colors relative"
            title="Atualizar dados agora"
          >
            <RefreshCw size={20} className={isLoading ? 'animate-spin' : ''} />
          </button>
        </div>
      </div>

      {/* Grid Principal */}
      <div className="flex-1 grid grid-cols-1 md:grid-cols-2 gap-6 min-h-0">
        
        {/* Coluna Saídas */}
        <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm flex flex-col overflow-hidden">
          <div className="p-4 border-b dark:border-slate-800 flex justify-between items-center bg-slate-50/50 dark:bg-slate-800/50">
            <div className="flex items-center gap-2">
              <Filter size={16} className="text-slate-400" />
              <h2 className="font-black text-[#075985] dark:text-sky-400 uppercase tracking-widest text-sm">Saídas</h2>
            </div>
            <button
              onClick={handleSendSummary}
              disabled={isSendingSummary || departures.length === 0}
              className="flex items-center gap-2 px-4 py-2 bg-emerald-700 text-white rounded-lg hover:bg-emerald-600 font-bold border border-emerald-600 uppercase text-[10px] tracking-wide transition-all shadow-sm disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {isSendingSummary ? (
                <><Loader2 size={16} className="animate-spin" /> Enviando...</>
              ) : (
                <><Send size={16} /> Enviar Resumo</>
              )}
            </button>
          </div>
          
          <div className="flex-1 overflow-y-auto p-4 space-y-3 scrollbar-thin">
            {departuresSummary.map(item => (
              <div key={item.id} className="flex justify-between items-center p-3 rounded-xl border border-slate-100 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-all">
                <div>
                  <h3 className="font-black text-slate-700 dark:text-slate-200 text-sm">{item.operacao}</h3>
                  <p className="text-[9px] text-slate-400 font-medium">Último envio: {item.ultimoEnvioFormatado}</p>
                  {item.webhookStatus && (
                    <p className="text-[8px] font-bold text-slate-500 mt-1">
                      Status: <span className={item.webhookStatus === 'OK' ? 'text-emerald-600' : item.webhookStatus === 'Atualizar' ? 'text-blue-600' : 'text-slate-600'}>{item.webhookStatus}</span>
                    </p>
                  )}
                </div>
                <div className="text-right">
                  <p className="text-[9px] font-bold text-slate-500 mb-1">{item.relativeTime}</p>
                  <span className={`px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-tighter ${item.statusColor}`}>
                    {item.status}
                  </span>
                </div>
              </div>
            ))}
          </div>

          <div className="p-4 bg-slate-50 dark:bg-slate-800/80 border-t dark:border-slate-800 space-y-3">
             <div className="flex items-center gap-2">
                <select
                  value={selectedOperacao}
                  onChange={(e) => setSelectedOperacao(e.target.value)}
                  className="flex-1 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg p-2 text-xs font-bold outline-none appearance-none cursor-pointer"
                >
                  <option value="">SELECIONAR OPERAÇÃO...</option>
                  {sendOptions.map(opt => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
                </select>
                <button 
                  onClick={handleSendDepartures}
                  disabled={isSending}
                  className="bg-[#075985] text-white px-6 py-2 rounded-lg font-black uppercase text-[10px] flex items-center gap-2 hover:bg-sky-800 transition-all disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {isSending ? <Loader2 size={14} className="animate-spin" /> : <><Send size={14} /> Enviar</>}
                </button>
             </div>
             {sendSuccess && (
                <div className="flex items-center gap-2 text-green-600 dark:text-green-400 bg-green-50 dark:bg-green-900/20 px-3 py-2 rounded-lg border border-green-200 dark:border-green-800">
                  <CheckCircle2 size={14} />
                  <span className="text-[10px] font-bold">Enviado com sucesso!</span>
                </div>
             )}
             {sendError && (
                <div className="flex items-center gap-2 text-red-600 dark:text-red-400 bg-red-50 dark:bg-red-900/20 px-3 py-2 rounded-lg border border-red-200 dark:border-red-800">
                  <AlertCircle size={14} />
                  <span className="text-[10px] font-bold">{sendError}</span>
                </div>
             )}
             <div className="flex items-center justify-between">
                <label className="flex items-center gap-2 cursor-pointer">
                  <span className="text-[10px] font-black text-slate-500 uppercase">Atualização?</span>
                  <div 
                    onClick={() => setIsAtualizacao(!isAtualizacao)}
                    className={`relative w-10 h-5 rounded-full transition-colors ${isAtualizacao ? 'bg-green-500' : 'bg-slate-300 dark:bg-slate-600'}`}
                  >
                    <div className={`absolute top-0.5 w-4 h-4 bg-white rounded-full shadow-sm transition-transform ${isAtualizacao ? 'left-5' : 'left-0.5'}`}></div>
                  </div>
                </label>
                <span className={`text-[9px] font-black uppercase ${isAtualizacao ? 'text-green-600' : 'text-slate-400'}`}>
                  {isAtualizacao ? 'SIM' : 'NÃO'}
                </span>
             </div>
          </div>
        </div>

        {/* Coluna Não Coletas */}
        <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm flex flex-col overflow-hidden">
          <div className="p-4 border-b dark:border-slate-800 flex justify-between items-center bg-slate-50/50 dark:bg-slate-800/50">
            <h2 className="font-black text-[#075985] dark:text-sky-400 uppercase tracking-widest text-sm">Não Coletas</h2>
            <button
              onClick={handleSendNCSummary}
              disabled={isSendingNCSummary || userConfigs.length === 0}
              className="flex items-center gap-2 px-4 py-2 bg-emerald-700 text-white rounded-lg hover:bg-emerald-600 font-bold border border-emerald-600 uppercase text-[10px] tracking-wide transition-all shadow-sm disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {isSendingNCSummary ? (
                <><Loader2 size={16} className="animate-spin" /> Enviando...</>
              ) : (
                <><Send size={16} /> Enviar Resumo</>
              )}
            </button>
          </div>

          <div className="flex-1 overflow-y-auto p-4 space-y-3 scrollbar-thin">
            {nonCollectionsSummary.map(item => (
              <div key={item.id} className="flex justify-between items-center p-3 rounded-xl border border-slate-100 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-all">
                <div>
                  <h3 className="font-black text-slate-700 dark:text-slate-200 text-sm">{item.operacao}</h3>
                  <p className="text-[10px] text-slate-400 font-medium">{new Date(item.timestamp).toLocaleString()}</p>
                </div>
                <div className="text-right">
                  <p className="text-[10px] font-bold text-slate-500 mb-1">{item.relativeTime}</p>
                  <span className={`px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-tighter ${item.statusColor}`}>
                    {item.status}
                  </span>
                </div>
              </div>
            ))}
          </div>

          <div className="p-4 bg-slate-50 dark:bg-slate-800/80 border-t dark:border-slate-800 flex flex-col gap-2">
            <div className="flex items-center gap-2">
              <select
                value={selectedOperacaoNC}
                onChange={(e) => setSelectedOperacaoNC(e.target.value)}
                className="flex-1 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg p-2 text-xs font-bold outline-none appearance-none cursor-pointer"
              >
                <option value="">SELECIONAR OPERAÇÃO...</option>
                {sendOptions.map(opt => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
              </select>
              <button
                onClick={handleSendNonCollections}
                disabled={isSending}
                className="bg-[#075985] text-white px-6 py-2 rounded-lg font-black uppercase text-[10px] flex items-center gap-2 hover:bg-sky-800 transition-all disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {isSending ? <Loader2 size={14} className="animate-spin" /> : <><Send size={14} /> Enviar</>}
              </button>
            </div>
            {/* Feedback de envio de NÃO COLETAS — aparece apenas no lado direito */}
            {ncSendSuccess && (
              <div className="flex items-center gap-2 text-green-600 dark:text-green-400 bg-green-50 dark:bg-green-900/20 px-3 py-2 rounded-lg border border-green-200 dark:border-green-800">
                <CheckCircle2 size={14} />
                <span className="text-[10px] font-bold">Não coletas enviadas com sucesso!</span>
              </div>
            )}
            {ncSendError && (
              <div className="flex items-center gap-2 text-red-600 dark:text-red-400 bg-red-50 dark:bg-red-900/20 px-3 py-2 rounded-lg border border-red-200 dark:border-red-800">
                <AlertCircle size={14} />
                <span className="text-[10px] font-bold">{ncSendError}</span>
              </div>
            )}
          </div>
        </div>

      </div>

      {/* Footer Status Bar */}
      <div className="mt-6 bg-white dark:bg-slate-900 p-2 rounded-xl border border-slate-200 dark:border-slate-800 flex items-center justify-center gap-4">
        <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Status Resumo</span>
        <span className={`px-4 py-1 rounded-full text-[10px] font-black uppercase flex items-center gap-2 ${resumoStatus.color}`}>
          {resumoStatus.status === 'OK' && <CheckCircle2 size={12} />}
          {resumoStatus.status === 'ATUALIZAR' && <RefreshCw size={12} className="animate-spin" />}
          {resumoStatus.status}
        </span>
        {resumoStatus.data && (
          <span className="text-[9px] text-slate-500 font-medium">
            {new Date(resumoStatus.data).toLocaleString('pt-BR')}
          </span>
        )}
      </div>
    </div>
  );
};

export default SendReportView;
