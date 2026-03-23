
import React, { useState, useEffect, useMemo } from 'react';
import { SharePointService } from '../services/sharepointService';
import { getValidToken } from '../services/tokenService';
import { getBrazilDate, getBrazilHours, isAfter10amBrazil } from '../utils/dateUtils';
import { RouteDeparture, Task, User, RouteConfig } from '../types';
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

  // Estados para seleção e envio
  const [selectedOperacao, setSelectedOperacao] = useState<string>('');
  const [selectedOperacaoNC, setSelectedOperacaoNC] = useState<string>('');
  const [isSending, setIsSending] = useState(false);
  const [sendError, setSendError] = useState<string | null>(null);
  const [sendSuccess, setSendSuccess] = useState(false);
  const [isAtualizacao, setIsAtualizacao] = useState(false);
  const [isSendingSummary, setIsSendingSummary] = useState(false);

  const WEBHOOK_URL = import.meta.env.VITE_WEBHOOK_SAIDAS_URL || "https://n8n.datastack.viagroup.com.br/webhook/8cb1f3e1-833d-42a7-a3f0-2f959ea390d6";
  const WEBHOOK_URL_RESUMO = import.meta.env.VITE_WEBHOOK_RESUMO_URL || "https://n8n.datastack.viagroup.com.br/webhook/8cb1f3e1-833d-42a7-a3f0-2f959ea390d6";

  const fetchAllData = async (forceRefresh: boolean = false) => {
    setIsLoading(true);
    const token = await getValidToken();
    if (!token) return;

    try {
      console.log('[FETCH_ALL] Buscando dados completos...', forceRefresh ? '(force refresh)' : '');
      console.log('[FETCH_ALL] Usuário logado:', currentUser.email);
      
      const [depData, configs] = await Promise.all([
        SharePointService.getDepartures(token, forceRefresh),
        SharePointService.getRouteConfigs(token, currentUser.email, forceRefresh)
      ]);

      console.log('[FETCH_ALL] Total de rotas brutas do SharePoint:', depData?.length || 0);
      console.log('[FETCH_ALL] Configurações carregadas:', configs?.length || 0);
      console.log('[FETCH_ALL] Operações do usuário:', configs?.map(c => c.operacao));

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
      setUserConfigs(configs);
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
          console.log(`[DEBUG_CONFIG] ${c.operacao}: ultimoEnvioSaida = "${c.ultimoEnvioSaida}" | Status = "${c.Status}" | UltimoEnvioResumoSaida = "${c.UltimoEnvioResumoSaida}" | StatusResumoSaida = "${c.StatusResumoSaida}"`);
        });
        setUserConfigs(configs);
        setLastSync(new Date());
      } catch (e) {
        console.error("Erro ao atualizar configs:", e);
      }
    };

    // Carrega dados iniciais com force refresh
    fetchAllData();

    // Polling das configs a cada 5 segundos (mais frequente para atualizar em tempo real)
    const configsInterval = setInterval(() => fetchConfigsOnly(true), 5000);

    // Atualização completa a cada 30 segundos
    const fullInterval = setInterval(() => {
      console.log('[POLLING] Atualização completa dos dados');
      fetchAllData(true);
    }, 30000);

    return () => {
      clearInterval(configsInterval);
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

    // VALIDAÇÃO DE SEGURANÇA: Verifica se a operação selecionada pertence ao usuário
    const myOps = new Set(userConfigs.map(c => c.operacao));
    if (!myOps.has(selectedOperacao)) {
      console.error(`[SEND_DEPARTURES_BLOCKED] Usuário tentou enviar operação não pertencente: ${selectedOperacao}`);
      setSendError(`Erro: Você não tem permissão para enviar esta operação.`);
      setTimeout(() => setSendError(null), 5000);
      return;
    }

    setIsSending(true);
    setSendError(null);

    const selectedDepartures = departures.filter(d => d.operacao === selectedOperacao);
    const config = userConfigs.find(c => c.operacao === selectedOperacao);

    console.log('[SEND_DEPARTURES] === ENVIANDO SAÍDAS ===');
    console.log('[SEND_DEPARTURES] Operação:', selectedOperacao);
    console.log('[SEND_DEPARTURES] Rotas encontradas:', selectedDepartures.length);
    console.log('[SEND_DEPARTURES] Config:', config);

    if (selectedDepartures.length === 0) {
      setSendError("Nenhuma saída encontrada para esta operação.");
      setTimeout(() => setSendError(null), 3000);
      setIsSending(false);
      return;
    }

    const payload = {
      tipo: "SAIDA",  // Tipo para envio de uma única filial/operação
      operacao: selectedOperacao,
      nomeExibicao: config?.nomeExibicao || selectedOperacao,
      tolerancia: config?.tolerancia || "00:00:00",
      atualizacao: isAtualizacao ? "sim" : "não",
      usuario: currentUser.email,
      dataEnvio: new Date().toISOString(),
      envio: config?.Envio || "", // Emails para envio principal
      copia: config?.Copia || "", // Emails para cópia
      saidas: selectedDepartures.map(d => ({
        rota: d.rota,
        data: d.data,
        inicio: d.inicio,
        motorista: d.motorista,
        placa: d.placa,
        saida: d.saida,
        motivo: d.motivo,
        observacao: d.observacao,
        status: d.statusOp
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
        if (dataHoraEnvio) {
          const token = await getValidToken() || currentUser.accessToken;
          if (token) {
            try {
              console.log('[ULTIMO_ENVIO] Enviando para atualização:', dataHoraEnvio);
              await SharePointService.updateUltimoEnvioSaida(
                token,
                selectedOperacao,
                dataHoraEnvio
              );
              console.log('[ULTIMO_ENVIO] ✅ Atualizado com sucesso:', dataHoraEnvio);
            } catch (err: any) {
              console.error('Erro ao atualizar UltimoEnvioSaida:', err.message);
            }
          }
        } else {
          console.warn('[WEBHOOK] Campo de data/hora de envio não encontrado na resposta');
        }

        // Processa e salva o status retornado pelo webhook
        const webhookStatus = responseData[0]?.status || responseData.status;
        if (webhookStatus) {
          const token = await getValidToken() || currentUser.accessToken;
          if (token) {
            try {
              console.log('[STATUS_WEBHOOK] Status retornado:', webhookStatus);
              // Normaliza o status: "atualizar" → "Atualizar", "ok" → "OK"
              const normalizedStatus = webhookStatus.toLowerCase() === 'atualizar' ? 'Atualizar' : 
                                       webhookStatus.toLowerCase() === 'ok' ? 'OK' : webhookStatus;
              
              await SharePointService.updateStatusOperacao(
                token,
                selectedOperacao,
                normalizedStatus
              );
              console.log('[STATUS_WEBHOOK] ✅ Status atualizado no SharePoint:', normalizedStatus);
            } catch (err: any) {
              console.error('Erro ao atualizar Status:', err.message);
            }
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
    }
  };

  // Função para enviar não coletas
  const handleSendNonCollections = async () => {
    if (!selectedOperacaoNC) {
      setSendError("Selecione uma operação para enviar.");
      setTimeout(() => setSendError(null), 3000);
      return;
    }

    // VALIDAÇÃO DE SEGURANÇA: Verifica se a operação selecionada pertence ao usuário
    const myOps = new Set(userConfigs.map(c => c.operacao));
    if (!myOps.has(selectedOperacaoNC)) {
      console.error(`[SEND_NAO_COLETA_BLOCKED] Usuário tentou enviar operação não pertencente: ${selectedOperacaoNC}`);
      setSendError(`Erro: Você não tem permissão para enviar esta operação.`);
      setTimeout(() => setSendError(null), 5000);
      return;
    }

    setIsSending(true);
    setSendError(null);

    const selectedDepartures = departures.filter(d => d.operacao === selectedOperacaoNC);
    const config = userConfigs.find(c => c.operacao === selectedOperacaoNC);

    console.log('[SEND_NAO_COLETA] === ENVIANDO NÃO COLETAS ===');
    console.log('[SEND_NAO_COLETA] Operação:', selectedOperacaoNC);
    console.log('[SEND_NAO_COLETA] Total de rotas:', selectedDepartures.length);

    const nonCollections = selectedDepartures.filter(d => d.statusGeral === 'NOK');

    console.log('[SEND_NAO_COLETA] Não coletas encontradas:', nonCollections.length);

    if (nonCollections.length === 0) {
      setSendError("Nenhuma não coleta encontrada para esta operação.");
      setTimeout(() => setSendError(null), 3000);
      setIsSending(false);
      return;
    }

    const payload = {
      tipo: "NAO_COLETA",  // Tipo para envio de não coletas de uma única filial/operação
      operacao: selectedOperacaoNC,
      nomeExibicao: config?.nomeExibicao || selectedOperacaoNC,
      tolerancia: config?.tolerancia || "00:00:00",
      atualizacao: isAtualizacao ? "sim" : "não",
      usuario: currentUser.email,
      dataEnvio: new Date().toISOString(),
      envio: config?.Envio || "", // Emails para envio principal
      copia: config?.Copia || "", // Emails para cópia
      naoColetas: nonCollections.map(d => ({
        rota: d.rota,
        data: d.data,
        inicio: d.inicio,
        motorista: d.motorista,
        placa: d.placa,
        saida: d.saida,
        motivo: d.motivo,
        observacao: d.observacao,
        status: d.statusOp
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
        if (dataHoraEnvio) {
          const token = await getValidToken() || currentUser.accessToken;
          if (token) {
            try {
              console.log('[ULTIMO_ENVIO_NAO_COLETAS] Enviando para atualização:', dataHoraEnvio);
              await SharePointService.updateUltimoEnvioSaida(
                token,
                selectedOperacaoNC,
                dataHoraEnvio
              );
              console.log('[ULTIMO_ENVIO_NAO_COLETAS] ✅ Atualizado com sucesso:', dataHoraEnvio);
            } catch (err: any) {
              console.error('Erro ao atualizar UltimoEnvioSaida:', err.message);
            }
          }
        } else {
          console.warn('[WEBHOOK] Campo de data/hora de envio não encontrado na resposta');
        }

        // Processa e salva o status retornado pelo webhook
        const webhookStatus = responseData[0]?.status || responseData.status;
        if (webhookStatus) {
          const token = await getValidToken() || currentUser.accessToken;
          if (token) {
            try {
              console.log('[STATUS_WEBHOOK_NAO_COLETAS] Status retornado:', webhookStatus);
              const normalizedStatus = webhookStatus.toLowerCase() === 'atualizar' ? 'Atualizar' : 
                                       webhookStatus.toLowerCase() === 'ok' ? 'OK' : webhookStatus;
              
              await SharePointService.updateStatusOperacao(
                token,
                selectedOperacaoNC,
                normalizedStatus
              );
              console.log('[STATUS_WEBHOOK_NAO_COLETAS] ✅ Status atualizado no SharePoint:', normalizedStatus);
            } catch (err: any) {
              console.error('Erro ao atualizar Status:', err.message);
            }
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

  // Lógica para processar a lista de SAÍDAS - usa o campo ultimoEnvioSaida das configs
  const departuresSummary = useMemo(() => {
    console.log('[DEBUG_SUMMARY] userConfigs:', userConfigs);
    return userConfigs.map(config => {
      const lastRoute = departures.filter(d => d.operacao === config.operacao).pop();

      // Usa o campo ultimoEnvioSaida da config (vem do SharePoint no formato ISO ou DD/MM/YYYY HH:MM:SS)
      const ultimoEnvio = config.ultimoEnvioSaida || "";

      console.log(`[DEBUG_SUMMARY] ${config.operacao}: ultimoEnvio = "${ultimoEnvio}" | Status = "${config.Status}"`);

      // Converte para Date para exibição
      let timestamp: string;
      let relativeTime: string;

      if (ultimoEnvio) {
        // Tenta converter de diferentes formatos
        let parsedDate: Date | null = null;

        // Formato ISO: 2026-03-12T16:08:26.000Z
        if (ultimoEnvio.includes('T')) {
          parsedDate = new Date(ultimoEnvio);
        }
        // Formato brasileiro: 12/03/2026 16:08:26
        else if (ultimoEnvio.includes('/')) {
          const [data, hora] = ultimoEnvio.split(' ');
          const [dia, mes, ano] = data.split('/');
          const [h, m, s] = hora ? hora.split(':') : ['00', '00', '00'];
          parsedDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(h), Number(m), Number(s));
        }
        // Formato americano: 03/12/2026 (MM/DD/YYYY) - fallback
        else {
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

      // Status do webhook (padrão: "Previsto" se não houver envio no dia)
      const webhookStatus = config.Status || "";
      let status = "PREVISTO";
      let color = "bg-slate-300 text-slate-600";

      // Verifica se houve envio HOJE comparando datas no fuso de Brasília
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

      // Se houve envio hoje (mesma data), usa o status do webhook
      if (envioDateObj && !isNaN(envioDateObj.getTime())) {
        envioDateObj.setHours(0, 0, 0, 0);
        const isToday = envioDateObj.getTime() === today.getTime();
        
        console.log(`[DEBUG_STATUS] ${config.operacao}: envioDate = ${envioDateObj.toISOString()}, isToday = ${isToday}, webhookStatus = "${webhookStatus}"`);
        
        if (isToday && webhookStatus) {
          // Usa o status retornado pelo webhook
          if (webhookStatus.toUpperCase() === 'OK') {
            status = "OK";
            color = "bg-emerald-500 text-white";
          } else if (webhookStatus.toUpperCase() === 'ATUALIZAR') {
            status = "ATUALIZAR";
            color = "bg-blue-500 text-white";
          } else {
            // Status desconhecido, mas houve envio hoje
            status = webhookStatus.toUpperCase();
            color = "bg-slate-500 text-white";
          }
        }
      }

      return {
        id: config.operacao,
        operacao: config.operacao,
        timestamp: timestamp,
        relativeTime: relativeTime,
        status: status,
        statusColor: color,
        webhookStatus: webhookStatus,
        ultimoEnvioFormatado: ultimoEnvio ? formatarDataHora(ultimoEnvio) : "Nunca"
      };
    });
  }, [departures, userConfigs]);

  // Lógica para processar a lista de NÃO COLETAS (Simulada com base no status das plantas)
  const nonCollectionsSummary = useMemo(() => {
    const ops = Array.from(new Set(userConfigs.map(c => c.operacao)));
    return ops.map(op => {
      // Simulamos a contagem de não coletas baseado em rotas NOK ou dados pendentes
      const opRoutes = departures.filter(d => d.operacao === op);
      const nokCount = opRoutes.filter(r => r.statusGeral === 'NOK').length;

      return {
        id: op,
        operacao: op,
        timestamp: new Date().toISOString(),
        relativeTime: getRelativeTime(new Date().toISOString()),
        status: nokCount > 0 ? `${nokCount} NÃO COLETAS` : "TODOS COLETADOS",
        statusColor: nokCount > 0 ? "bg-red-500 text-white" : "bg-emerald-500 text-white"
      };
    });
  }, [departures, userConfigs]);

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
            onClick={handleSendSummary}
            disabled={isSendingSummary || departures.length === 0}
            className="flex items-center gap-2 px-4 py-2 bg-emerald-700 text-white rounded-lg hover:bg-emerald-600 font-bold border border-emerald-600 uppercase text-[10px] tracking-wide transition-all shadow-sm disabled:opacity-50 disabled:cursor-not-allowed"
          >
            <Send size={16} /> Enviar Resumo
          </button>
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
                  {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
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
          <div className="p-4 border-b dark:border-slate-800 flex justify-center items-center bg-slate-50/50 dark:bg-slate-800/50">
            <h2 className="font-black text-[#075985] dark:text-sky-400 uppercase tracking-widest text-sm">Não Coletas</h2>
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

          <div className="p-4 bg-slate-50 dark:bg-slate-800/80 border-t dark:border-slate-800 flex items-center gap-2">
            <select 
              value={selectedOperacaoNC}
              onChange={(e) => setSelectedOperacaoNC(e.target.value)}
              className="flex-1 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg p-2 text-xs font-bold outline-none appearance-none cursor-pointer"
            >
              <option value="">SELECIONAR OPERAÇÃO...</option>
              {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
            </select>
            <button 
              onClick={handleSendNonCollections}
              disabled={isSending}
              className="bg-[#075985] text-white px-6 py-2 rounded-lg font-black uppercase text-[10px] flex items-center gap-2 hover:bg-sky-800 transition-all disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {isSending ? <Loader2 size={14} className="animate-spin" /> : <><Send size={14} /> Enviar</>}
            </button>
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
