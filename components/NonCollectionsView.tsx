import React, { useState, useEffect, useRef, useMemo } from 'react';
import { RouteConfig, User } from '../types';
import { SharePointService } from '../services/sharepointService';
import { getValidToken } from '../services/tokenService';
import * as XLSX from 'xlsx';
import { getBrazilDate, getWeekString } from '../utils/dateUtils';
import {
  Clock, X, Loader2, RefreshCw, ShieldCheck,
  CheckCircle2, ChevronDown,
  Filter, Search, CheckSquare, Square,
  ChevronRight, Maximize2, Minimize2,
  Archive, Database, Save, LinkIcon,
  Layers, Trash2, Settings2, Check, Table, SortAsc,
  Sun, Moon, AlertTriangle, Plus, Milk
} from 'lucide-react';

interface NonCollection {
  id: string;
  semana: string;
  rota: string;
  data: string;
  codigo: string;
  produtor: string;
  motivo: string;
  observacao: string;
  acao: string;
  dataAcao: string;
  ultimaColeta: string;
  culpabilidade: string;
  operacao: string;
}

// Lista de MOTIVOS no padrÃ£o "MOTIVO - CULPABILIDADE"
const MOTIVOS_CULPABILIDADES: Record<string, string> = {
  'Rota Atrasada/FÃ¡brica': 'Cliente',
  'Rota Atrasada/LogÃ­stica': 'VIA',
  'Rota Atrasada/ManutenÃ§Ã£o': 'VIA',
  'Rota Atrasada/MÃ£o De Obra': 'VIA',
  'Eqp. Cheio/Coleta Extra': 'VIA',
  'Eqp. Cheio/Eqp Menor': 'VIA',
  'Eqp. Cheio/Rota Atrasada': 'VIA',
  'Alizarol Positivo': 'Outros',
  'Leite Com AntibiÃ³tico': 'Outros',
  'Leite Congelado': 'Outros',
  'Leite Descartado': 'Outros',
  'Leite Quente': 'Outros',
  'NÃ£o Coletado - SaÃºde Do Motorista': 'VIA',
  'NÃ£o Coletado - Greve': 'Outros',
  'NÃ£o Coletado - Solicitado Pelo Sdl': 'Cliente',
  'Objeto Ou Sujeira No Leite': 'Outros',
  'Parou De Fornecer': 'Outros',
  'Produtor Suspenso': 'Cliente',
  'Passou Para 48Hrs': 'VIA',
  'Problemas MecÃ¢nicos Equipamento': 'VIA',
  'Produtor Solicitou A NÃ£o Coleta': 'Outros',
  'Resfriador Vazio': 'Outros',
  'Volume Insuficiente Para Medida': 'Outros',
  'Coletado Por Outra Transportadora': 'VIA',
  'Descumprimento de roteirizaÃ§Ã£o': 'VIA',
  'A rota nÃ£o foi realizada': 'Outros',
  'Falta De Acesso': 'Outros',
  'Jornada Excedida': 'Outros',
  'Eqp. Cheio/Aumento de Vol.': 'Outros',
  'CorreÃ§Ã£o de RoteirzaÃ§Ã£o': 'VIA',
  'Crioscopia': 'Outros',
  'Rota Atrasada/Infraestrutura': 'Outros',
  'Eqp. Cheio/ManutenÃ§Ã£o': 'VIA'
};

// Motivos que mostram popup de causa raiz
const MOTIVOS_COM_CAUSA_RAIZ = [
  'Eqp. Cheio/Coleta Extra',
  'Eqp. Cheio/Eqp Menor',
  'Eqp. Cheio/Rota Atrasada',
  'Coletado Por Outra Transportadora',
  'Descumprimento de roteirizaÃ§Ã£o'
];

const CULPABILIDADE_OPCOES = ['VIA', 'Cliente', 'Outros'];

const NonCollectionsView: React.FC<{
  currentUser: User;
}> = ({ currentUser }) => {
  const [nonCollections, setNonCollections] = useState<NonCollection[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isDarkMode, setIsDarkMode] = useState(() => {
    const saved = localStorage.getItem('non_collections_dark_mode');
    return saved !== 'false';
  });

  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>({});
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [isSortByDataEnabled, setIsSortByDataEnabled] = useState(true);
  const [colWidths, setColWidths] = useState<Record<string, number>>({
    semana: 120,
    rota: 100,
    data: 120,
    codigo: 90,
    produtor: 220,
    motivo: 180,
    observacao: 300,
    acao: 200,
    dataAcao: 120,
    ultimaColeta: 120,
    culpabilidade: 130,
    operacao: 120
  });

  const [hiddenColumns, setHiddenColumns] = useState<Set<string>>(() => {
    const saved = localStorage.getItem('non_collections_hidden_cols');
    if (saved) {
      return new Set(JSON.parse(saved));
    }
    return new Set();
  });

  const [contextMenu, setContextMenu] = useState<{ visible: boolean; x: number; y: number; col: string | null }>({ visible: false, x: 0, y: 0, col: null });
  const filterDropdownRef = useRef<HTMLDivElement>(null);
  const contextMenuRef = useRef<HTMLDivElement>(null);
  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);

  // Estados para modal de adicionar nÃ£o coleta
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [newNonCollectionData, setNewNonCollectionData] = useState<{
    rota: string;
    data: string;
    codigo: string;
    produtor: string;
    operacao: string;
  }>({
    rota: '',
    data: getBrazilDate().split('-').reverse().join('/'),
    codigo: '',
    produtor: '',
    operacao: ''
  });

  // Ghost Row para adiÃ§Ã£o rÃ¡pida via paste
  const [ghostRow, setGhostRow] = useState<Partial<NonCollection>>({
    id: 'ghost',
    semana: '',
    rota: '',
    data: getBrazilDate().split('-').reverse().join('/'),
    codigo: '',
    produtor: '',
    motivo: '',
    observacao: '',
    acao: '',
    dataAcao: '',
    ultimaColeta: '',
    culpabilidade: '',
    operacao: ''
  });

  const [showCausaRaiz, setShowCausaRaiz] = useState(false);

  // Estados para modal de seleÃ§Ã£o de operaÃ§Ã£o (bulk paste)
  const [isOperationModalOpen, setIsOperationModalOpen] = useState(false);
  const [pendingBulkRoutes, setPendingBulkRoutes] = useState<string[]>([]);

  const updateGhostCell = (field: keyof NonCollection, value: string) => {
    const updatedGhost = { ...ghostRow, [field]: value };

    if (field === 'motivo') {
      // Preenche culpabilidade automaticamente baseado no motivo
      const culpabilidadeAuto = MOTIVOS_CULPABILIDADES[value];
      if (culpabilidadeAuto) {
        updatedGhost.culpabilidade = culpabilidadeAuto;
      }
      // Verifica se Ã© motivo com causa raiz
      setShowCausaRaiz(MOTIVOS_COM_CAUSA_RAIZ.includes(value));
    }

    setGhostRow(updatedGhost);
  };

  const handleAddFromGhost = async () => {
    if (!ghostRow.rota || !ghostRow.data || !ghostRow.produtor || !ghostRow.operacao) {
      alert('Preencha todos os campos obrigatÃ³rios na linha de criaÃ§Ã£o!');
      return;
    }

    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) {
        alert('Erro: Token nÃ£o encontrado');
        return;
      }

      const dataParaSemana = ghostRow.data!.split('/').reverse().join('-');
      const semana = getWeekString(dataParaSemana);

      // Calcula dataAcao automÃ¡tica: data + 2 dias se motivo gerar "SerÃ¡ coletado na rota..."
      const calcularDataAcao = () => {
        const motivo = (ghostRow.motivo || '').trim();
        const dataNaoColeta = ghostRow.data!;

        if (motivo && dataNaoColeta) {
          // Motivos que geram "Leite Descartado" ou "Aguardando autorizaÃ§Ã£o" ficam com hÃ­fen
          const motivoLower = motivo.toLowerCase();
          if (motivoLower === 'parou de fornecer' || motivoLower === 'produtor suspenso' || motivoLower === 'alizarol positivo') {
            return '-';
          }

          // Se tem rota e data, calcula data + 2 dias
          if (ghostRow.rota) {
            const match = dataNaoColeta.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
            if (match) {
              const [, day, month, year] = match;
              const dateObj = new Date(Number(year), Number(month) - 1, Number(day));
              dateObj.setDate(dateObj.getDate() + 2);
              const d = String(dateObj.getDate()).padStart(2, '0');
              const m = String(dateObj.getMonth() + 1).padStart(2, '0');
              const y = dateObj.getFullYear();
              return `${d}/${m}/${y}`;
            }
          }
        }

        return ''; // Sem motivo definido, fica vazio
      };

      const newRecord: NonCollection = {
        id: Date.now().toString(),
        semana,
        rota: ghostRow.rota!,
        data: ghostRow.data!,
        codigo: ghostRow.codigo || `P${String(nonCollections.length + 1).padStart(3, '0')}`,
        produtor: ghostRow.produtor!,
        motivo: ghostRow.motivo || '',
        observacao: ghostRow.observacao || '',
        acao: ghostRow.acao || '',
        dataAcao: calcularDataAcao(),
        ultimaColeta: '',
        culpabilidade: ghostRow.culpabilidade || 'NÃ£o se aplica',
        operacao: ghostRow.operacao!
      };

      // Salva no SharePoint e obtÃ©m o ID real
      const spId = await SharePointService.saveNonCollection(token, newRecord);

      // Adiciona localmente com o ID real do SharePoint
      setNonCollections(prev => [...prev, { ...newRecord, id: spId }]);

      // Limpa ghost row
      setGhostRow({
        id: 'ghost',
        semana: '',
        rota: '',
        data: getBrazilDate().split('-').reverse().join('/'),
        codigo: '',
        produtor: '',
        motivo: '',
        observacao: '',
        acao: '',
        dataAcao: '',
        ultimaColeta: '',
        culpabilidade: '',
        operacao: ''
      });
      setShowCausaRaiz(false);
    } catch (e: any) {
      console.error('[ADD_NON_COLLECTION] Error:', e);
      alert(`Erro ao adicionar nÃ£o coleta: ${e.message}`);
    }
  };

  /**
   * Cria novas linhas de nÃ£o coleta apÃ³s o usuÃ¡rio selecionar a operaÃ§Ã£o no modal.
   * SALVA CADA REGISTRO NO SHAREPOINT IMEDIATAMENTE.
   */
  const createBulkRecordsWithOperation = async (operacao: string) => {
    const dataFormatada = getBrazilDate().split('-').reverse().join('/');
    const dataParaSemana = dataFormatada.split('/').reverse().join('-');
    const semana = getWeekString(dataParaSemana);

    // Tenta salvar no SharePoint imediatamente
    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) {
        console.error('[BULK_PASTE] Token nÃ£o encontrado, criando apenas localmente');
        // Fallback: cria localmente (nÃ£o serÃ¡ salvo no SharePoint)
        const newRecords: NonCollection[] = pendingBulkRoutes.map((rota, i) => ({
          id: Date.now().toString() + i,
          semana,
          rota,
          data: dataFormatada,
          codigo: '',
          produtor: '',
          motivo: '',
          observacao: '',
          acao: '',
          dataAcao: '',
          ultimaColeta: '',
          culpabilidade: 'NÃ£o se aplica',
          operacao
        }));
        setNonCollections(prev => [...prev, ...newRecords]);
        setIsOperationModalOpen(false);
        setPendingBulkRoutes([]);
        return;
      }

      console.log('[BULK_PASTE] Criando', pendingBulkRoutes.length, 'registros no SharePoint...');
      const savedRecords: NonCollection[] = [];

      // Salva cada registro no SharePoint sequencialmente
      for (let i = 0; i < pendingBulkRoutes.length; i++) {
        const rota = pendingBulkRoutes[i];
        const tempRecord: NonCollection = {
          id: 'temp', // ID temporÃ¡rio, serÃ¡ substituÃ­do pelo ID real do SharePoint
          semana,
          rota,
          data: dataFormatada,
          codigo: '',
          produtor: '',
          motivo: '',
          observacao: '',
          acao: '',
          dataAcao: '',
          ultimaColeta: '',
          culpabilidade: 'NÃ£o se aplica',
          operacao
        };

        try {
          const spId = await SharePointService.saveNonCollection(token, tempRecord);
          console.log('[BULK_PASTE] âœ… Criado no SharePoint:', rota, 'ID:', spId);
          savedRecords.push({ ...tempRecord, id: spId }); // Usa o ID real do SharePoint
        } catch (e: any) {
          console.error('[BULK_PASTE] Erro ao criar registro no SharePoint:', rota, e.message);
          // Adiciona localmente com ID temporÃ¡rio mesmo com erro
          savedRecords.push({ ...tempRecord, id: (Date.now() + i).toString() });
        }
      }

      setNonCollections(prev => [...prev, ...savedRecords]);
      console.log('[BULK_PASTE] âœ…', savedRecords.length, 'linhas criadas e salvas com operaÃ§Ã£o:', operacao);
    } catch (e: any) {
      console.error('[BULK_PASTE] Erro crÃ­tico ao salvar no SharePoint:', e.message);
      // Fallback: cria localmente
      const newRecords: NonCollection[] = pendingBulkRoutes.map((rota, i) => ({
        id: Date.now().toString() + i,
        semana,
        rota,
        data: dataFormatada,
        codigo: '',
        produtor: '',
        motivo: '',
        observacao: '',
        acao: '',
        dataAcao: '',
        ultimaColeta: '',
        culpabilidade: 'NÃ£o se aplica',
        operacao
      }));
      setNonCollections(prev => [...prev, ...newRecords]);
    }

    // Limpa estados do modal
    setIsOperationModalOpen(false);
    setPendingBulkRoutes([]);
  };

  /**
   * Cancela o modal de seleÃ§Ã£o de operaÃ§Ã£o sem criar linhas.
   */
  const cancelBulkPaste = () => {
    setIsOperationModalOpen(false);
    setPendingBulkRoutes([]);
  };

  // Fecha dropdown ao clicar fora
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (filterDropdownRef.current && !filterDropdownRef.current.contains(e.target as Node)) {
        setActiveFilterCol(null);
      }
      if (contextMenuRef.current && !contextMenuRef.current.contains(e.target as Node)) {
        setContextMenu(prev => ({ ...prev, visible: false }));
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Redimensionamento de colunas
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!resizingRef.current) return;
      const { col, startX, startWidth } = resizingRef.current;
      const diff = e.clientX - startX;
      const newWidth = Math.max(50, startWidth + diff);
      setColWidths(prev => ({ ...prev, [col]: newWidth }));
    };

    const handleMouseUp = () => {
      resizingRef.current = null;
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, []);

  // Salvar preferÃªncias de colunas
  useEffect(() => {
    localStorage.setItem('non_collections_hidden_cols', JSON.stringify(Array.from(hiddenColumns)));
  }, [hiddenColumns]);

  // Carrega dados ao montar e quando usuÃ¡rio muda
  useEffect(() => {
    loadData();
  }, [currentUser]);

  const loadData = async () => {
    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) {
        console.error('[NonCollections] Token nÃ£o encontrado');
        return;
      }

      console.log('[NonCollections] Carregando dados...', currentUser.email);

      // Carrega configuraÃ§Ãµes do usuÃ¡rio
      const configs = await SharePointService.getRouteConfigs(token, currentUser.email, true);
      setUserConfigs(configs || []);
      console.log('[NonCollections] OperaÃ§Ãµes do usuÃ¡rio:', configs?.map(c => c.operacao));

      // Carrega nÃ£o coletas do SharePoint
      const spNonCollections = await SharePointService.getNonCollections(token, currentUser.email);
      console.log('[NonCollections] Total bruto do SharePoint:', spNonCollections.length);

      // Filtra APENAS nÃ£o coletas das operaÃ§Ãµes do usuÃ¡rio logado
      const myOps = new Set((configs || []).map(c => c.operacao));
      const filtered = (spNonCollections || []).filter(nc => {
        if (myOps.size === 0) return true; // Fallback se config nÃ£o carregou
        return myOps.has(nc.operacao);
      });

      console.log('[NonCollections] NÃ£o coletas filtradas por usuÃ¡rio:', filtered.length);
      console.log('[NonCollections] OperaÃ§Ãµes nos dados filtrados:', Array.from(new Set(filtered.map(r => r.operacao))));

      setNonCollections(filtered);
      console.log('[NonCollections] âœ… Dados carregados com sucesso');
    } catch (e: any) {
      console.error('[NonCollections] Erro ao carregar dados:', e.message);
    } finally {
      setIsLoading(false);
    }
  };

  // Dark mode effect
  useEffect(() => {
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
    localStorage.setItem('non_collections_dark_mode', String(isDarkMode));
  }, [isDarkMode]);

  const getUniqueValues = (field: keyof NonCollection) => {
    return Array.from(new Set(nonCollections.map(r => String(r[field] || '')))).filter(v => v).sort();
  };

  const toggleFilterValue = (field: keyof NonCollection, value: string) => {
    const current = selectedFilters[field] || [];
    const updated = current.includes(value)
      ? current.filter(v => v !== value)
      : [...current, value];
    setSelectedFilters({ ...selectedFilters, [field]: updated });
  };

  const clearFilters = () => {
    setColFilters({});
    setSelectedFilters({});
  };

  const hasActiveFilters = Object.keys(selectedFilters).some(col => (selectedFilters[col] || []).length > 0);

  const filteredData = useMemo(() => {
    let result = [...nonCollections];

    // Aplica filtros de texto
    result = result.filter(r => {
      return (Object.entries(colFilters) as [string, string][]).every(([col, val]) => {
        if (!val) return true;
        const field = col as keyof NonCollection;
        return String(r[field] || '').toLowerCase().includes(val.toLowerCase());
      });
    });

    // Aplica filtros selecionados
    result = result.filter(r => {
      return (Object.entries(selectedFilters) as [string, string[]][]).every(([col, vals]) => {
        if (!vals || vals.length === 0) return true;
        const field = col as keyof NonCollection;
        return vals.includes(String(r[field] || ''));
      });
    });

    // OrdenaÃ§Ã£o por data
    if (isSortByDataEnabled) {
      result.sort((a, b) => {
        const parseDate = (dateStr: string) => {
          if (!dateStr) return 0;
          const match = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
          if (match) {
            const [, day, month, year] = match;
            return new Date(Number(year), Number(month) - 1, Number(day)).getTime();
          }
          return 0;
        };
        return parseDate(a.data) - parseDate(b.data);
      });
    }

    return result;
  }, [nonCollections, colFilters, selectedFilters, isSortByDataEnabled]);

  const handleAddNonCollection = async () => {
    if (!newNonCollectionData.rota || !newNonCollectionData.data || !newNonCollectionData.produtor || !newNonCollectionData.operacao) {
      alert('Preencha todos os campos obrigatÃ³rios!');
      return;
    }

    try {
      const dataParaSemana = newNonCollectionData.data.split('/').reverse().join('-');
      const semana = getWeekString(dataParaSemana);

      const newRecord: NonCollection = {
        id: Date.now().toString(),
        semana,
        rota: newNonCollectionData.rota,
        data: newNonCollectionData.data,
        codigo: newNonCollectionData.codigo || `P${String(nonCollections.length + 1).padStart(3, '0')}`,
        produtor: newNonCollectionData.produtor,
        motivo: '',
        observacao: '',
        acao: '',
        dataAcao: '',
        ultimaColeta: '',
        culpabilidade: 'NÃ£o se aplica',
        operacao: newNonCollectionData.operacao
      };

      setNonCollections(prev => [...prev, newRecord]);
      setIsAddModalOpen(false);
      setNewNonCollectionData({
        rota: '',
        data: getBrazilDate().split('-').reverse().join('/'),
        codigo: '',
        produtor: '',
        operacao: ''
      });
    } catch (e) {
      console.error('[ADD_NON_COLLECTION] Error:', e);
      alert('Erro ao adicionar nÃ£o coleta');
    }
  };

  /**
   * Tenta separar cÃ³digo e produtor de uma string colada.
   * PadrÃ£o esperado: cÃ³digo alfanumÃ©rico no inÃ­cio seguido de texto (ex: "202520769MARCELO DINIZ COUTO" ou "P0274001RODRIGO ALVES PEREIRA")
   * Retorna null se nÃ£o conseguir identificar o padrÃ£o.
   */
  const parseCodigoProdutor = (line: string): { codigo: string; produtor: string } | null => {
    const trimmed = line.trim();
    
    // Regex: captura caracteres alfanumÃ©ricos no inÃ­cio (cÃ³digo) seguidos de texto (produtor)
    // O cÃ³digo pode ter letras e nÃºmeros (ex: "P0274001", "202520769")
    // O produtor Ã© todo o restante da string
    // Ex: "P0274001RODRIGO ALVES PEREIRA (IN )" â†’ codigo: "P0274001", produtor: "RODRIGO ALVES PEREIRA (IN )"
    const match = trimmed.match(/^([A-Za-z0-9]+)(.+)$/);
    
    if (match) {
      return {
        codigo: match[1].toUpperCase(),
        produtor: match[2].trim()
      };
    }
    
    return null;
  };

  const handleBulkPaste = async (field: keyof NonCollection, value: string) => {
    const lines = value.split(/[\n\r]/).map(l => l.trim()).filter(Boolean);
    if (lines.length === 0) return;

    console.log('[BULK_PASTE] Campo:', field, 'Valores:', lines);
    console.log('[BULK_PASTE] nonCollections.length:', nonCollections.length);

    // Colunas que SEMPRE criam novas linhas (mesmo se jÃ¡ houver dados)
    const colunasQueCriamLinhas: (keyof NonCollection)[] = ['rota'];
    const criaNovasLinhas = colunasQueCriamLinhas.includes(field);

    console.log('[BULK_PASTE] criaNovasLinhas:', criaNovasLinhas);

    if (criaNovasLinhas) {
      // COMPORTAMENTO 1: Criar novas linhas (apenas para ROTA)
      // Abre modal para selecionar operaÃ§Ã£o antes de criar as linhas
      setPendingBulkRoutes(lines);
      setIsOperationModalOpen(true);
    } else {
      // COMPORTAMENTO 2: Atualizar linhas existentes (para CÃ“DIGO, PRODUTOR, MOTIVO, OBSERVAÃ‡ÃƒO, etc.)
      console.log('[BULK_PASTE] Atualizando linhas existentes...');
      const updatedRecords: NonCollection[] = [];

      // Colunas que devem ser tratadas individualmente (sem tentar separar cÃ³digo/produtor)
      const colunasIndividuais: (keyof NonCollection)[] = ['codigo', 'produtor'];
      const isColunaIndividual = colunasIndividuais.includes(field);

      for (let i = 0; i < Math.min(lines.length, nonCollections.length); i++) {
        const record = nonCollections[i];
        let finalValue = lines[i];
        let codigo = '';
        let produtor = '';

        // Tenta separar cÃ³digo e produtor SOMENTE se NÃƒO for coluna individual
        if (!isColunaIndividual) {
          const parsed = parseCodigoProdutor(lines[i]);
          if (parsed) {
            codigo = parsed.codigo;
            produtor = parsed.produtor;
            console.log('[BULK_PASTE] âœ… CÃ³digo e produtor separados:', parsed);
          }
        }

        // Formata se for data
        if (field === 'data' || field === 'dataAcao' || field === 'ultimaColeta') {
          if (finalValue.includes('-')) {
            const [year, month, day] = finalValue.split('-');
            finalValue = `${day}/${month}/${year}`;
          }
        }

        // CÃ“DIGO: converte para maiÃºsculo
        if (field === 'codigo') {
          finalValue = finalValue.toUpperCase();
        }

        // Se Ã© coluna individual, aplica direto. Se separou cÃ³digo/produtor, aplica ambos.
        if (isColunaIndividual) {
          updatedRecords.push({
            ...record,
            [field]: finalValue
          });
        } else {
          updatedRecords.push({
            ...record,
            ...(codigo && produtor ? { codigo, produtor } : { [field]: finalValue })
          });
        }
        console.log('[BULK_PASTE] Atualizando linha', i, 'campo', field, 'valor', isColunaIndividual ? finalValue : (codigo && produtor ? `${codigo} + ${produtor}` : finalValue));
      }

      // MantÃ©m as linhas que nÃ£o foram atualizadas
      const remainingRecords = nonCollections.slice(updatedRecords.length);

      // Atualiza estado local
      setNonCollections([...updatedRecords, ...remainingRecords]);
      console.log('[BULK_PASTE] âœ…', updatedRecords.length, 'linhas atualizadas', remainingRecords.length, 'mantidas');

      // Salva no SharePoint APENAS registros que jÃ¡ existem no SharePoint
      // Registros criados localmente (com IDs baseados em timestamp) sÃ£o ignorados
      // atÃ© serem salvos via "Adicionar NÃ£o Coleta" ou similar
      try {
        const token = await getValidToken() || currentUser.accessToken;
        if (!token) {
          console.error('[BULK_PASTE] Token nÃ£o encontrado, pulando salvamento no SharePoint');
          return;
        }

        // Filtra apenas registros que jÃ¡ existem no SharePoint (IDs numÃ©ricos vÃ¡lidos)
        // IDs locais gerados por Date.now() sÃ£o strings grandes (13+ dÃ­gitos)
        const recordsToSave = updatedRecords.filter(r => {
          const id = parseInt(r.id);
          // SharePoint IDs sÃ£o inteiros pequenos (< 1 milhÃ£o), locais sÃ£o timestamps grandes
          return !isNaN(id) && id < 1000000;
        });

        if (recordsToSave.length === 0) {
          console.log('[BULK_PASTE] Nenhum registro para salvar no SharePoint (todos sÃ£o locais)');
          return;
        }

        console.log('[BULK_PASTE] Salvando', recordsToSave.length, 'registros no SharePoint...');
        const savePromises = recordsToSave.map(async (record) => {
          try {
            await SharePointService.updateNonCollection(token, record);
            console.log('[BULK_PASTE] âœ… Salvo:', record.rota, '-', record.codigo);
          } catch (e: any) {
            console.error('[BULK_PASTE] Erro ao salvar registro:', record.rota, e.message);
          }
        });
        await Promise.all(savePromises);
        console.log('[BULK_PASTE] âœ… Todos os registros salvos no SharePoint');
      } catch (e: any) {
        console.error('[BULK_PASTE] Erro ao salvar no SharePoint:', e.message);
      }
    }
  };

  const handleColumnResize = (col: string, startX: number, startWidth: number) => {
    resizingRef.current = { col, startX, startWidth };
  };

  const handleContextMenu = (e: React.MouseEvent, col: string) => {
    e.preventDefault();
    setContextMenu({ visible: true, x: e.clientX, y: e.clientY, col });
  };

  const toggleColumn = (col: string) => {
    setHiddenColumns(prev => {
      const next = new Set(prev);
      if (next.has(col)) {
        next.delete(col);
      } else {
        next.add(col);
      }
      return next;
    });
  };

  const formatDateToBR = (dateString: string) => {
    if (!dateString) return '--/--/----';
    if (dateString.includes('/')) return dateString;
    try {
      const [year, month, day] = dateString.split('-');
      return `${day}/${month}/${year}`;
    } catch {
      return dateString;
    }
  };

  const columns: { key: keyof NonCollection; label: string }[] = [
    { key: 'semana', label: 'SEMANA' },
    { key: 'rota', label: 'ROTA' },
    { key: 'data', label: 'DATA' },
    { key: 'codigo', label: 'CÃ“DIGO' },
    { key: 'produtor', label: 'PRODUTOR' },
    { key: 'motivo', label: 'MOTIVO' },
    { key: 'observacao', label: 'OBSERVAÃ‡ÃƒO' },
    { key: 'acao', label: 'AÃ‡ÃƒO' },
    { key: 'dataAcao', label: 'DATA AÃ‡ÃƒO' },
    { key: 'ultimaColeta', label: 'ÃšLTIMA COLETA' },
    { key: 'culpabilidade', label: 'CULPABILIDADE' },
    { key: 'operacao', label: 'OPERAÃ‡ÃƒO' }
  ];

  if (isLoading) {
    return (
      <div className="h-full flex items-center justify-center">
        <div className="flex flex-col items-center gap-4">
          <Loader2 size={40} className="animate-spin text-blue-600" />
          <p className="font-bold uppercase text-sm tracking-widest text-slate-500">Carregando...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="h-full flex flex-col bg-gradient-to-br from-white via-slate-50 to-slate-100 dark:from-slate-950 dark:via-slate-900 dark:to-slate-950 overflow-hidden">
      {/* Header */}
      <div className="flex items-center justify-between p-6 pb-4">
        <div className="flex items-center gap-6">
          <div>
            <h1 className="text-2xl font-black uppercase text-slate-800 dark:text-white tracking-tight">
              NÃ£o Coletas
            </h1>
            <p className="text-xs font-bold text-slate-500 dark:text-slate-400 mt-1 uppercase tracking-widest">
              Acompanhamento de ocorrÃªncias
            </p>
          </div>

          {/* Cards de Indicadores */}
          <div className="flex items-center gap-3 ml-8">
            <div className={`flex items-center gap-3 px-6 py-3 rounded-2xl min-w-[140px] ${
              isDarkMode ? 'bg-blue-900/30 border border-blue-700/50' : 'bg-blue-100 border border-blue-300'
            }`}>
              <div className="text-center flex-1">
                <p className={`text-[9px] font-black uppercase tracking-wider mb-1 ${isDarkMode ? 'text-blue-400' : 'text-blue-700'}`}>Total</p>
                <p className={`text-2xl font-black leading-none ${isDarkMode ? 'text-blue-400' : 'text-blue-700'}`}>{nonCollections.length}</p>
              </div>
              <div className="w-2 h-2 bg-blue-500 rounded-full animate-pulse shrink-0"></div>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={() => setIsAddModalOpen(true)}
            className="flex items-center gap-2 px-4 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-black uppercase text-[10px] tracking-widest transition-all shadow-lg shadow-blue-500/20"
          >
            <Plus size={16} />
            Adicionar NÃ£o Coleta
          </button>

          <button
            onClick={loadData}
            className="p-3 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-sm hover:shadow-md transition-all"
            title="Recarregar"
          >
            <RefreshCw size={20} className="text-slate-600 dark:text-slate-400" />
          </button>

          <button
            onClick={() => setIsDarkMode(!isDarkMode)}
            className="p-3 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-sm hover:shadow-md transition-all"
          >
            {isDarkMode ? (
              <Sun size={20} className="text-amber-400" />
            ) : (
              <Moon size={20} className="text-slate-600" />
            )}
          </button>

          <button
            onClick={() => setIsSortByDataEnabled(!isSortByDataEnabled)}
            className={`p-3 rounded-xl border shadow-sm hover:shadow-md transition-all ${
              isSortByDataEnabled
                ? 'bg-blue-600 border-blue-700 text-white'
                : 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-400'
            }`}
            title="Ordenar por Data"
          >
            <SortAsc size={20} />
          </button>

          {hasActiveFilters && (
            <button
              onClick={clearFilters}
              className="px-4 py-3 rounded-xl bg-red-50 dark:bg-red-900/30 border border-red-200 dark:border-red-800 text-red-600 dark:text-red-400 font-black uppercase text-[10px] tracking-widest hover:bg-red-100 dark:hover:bg-red-900/50 transition-all"
            >
              Limpar Filtros
            </button>
          )}
        </div>
      </div>

      {/* Table Container */}
      <div className="flex-1 mx-6 mt-4 mb-6 bg-white/95 dark:bg-slate-900/95 backdrop-blur-sm rounded-[2rem] shadow-2xl border border-white/50 dark:border-slate-800 overflow-hidden flex flex-col">
        <div className="overflow-x-auto flex-1 scrollbar-thin" id="table-container">
          <table className="w-full border-collapse">
            <thead className="sticky top-0 z-20">
              <tr className="bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900">
                {columns.map(({ key, label }) => (
                  <th
                    key={key}
                    className={`relative px-4 py-4 text-left border-r border-slate-700/50 last:border-r-0 select-none ${
                      hiddenColumns.has(key) ? 'hidden' : ''
                    }`}
                    style={{ width: colWidths[key] }}
                    onContextMenu={(e) => handleContextMenu(e, key)}
                  >
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] font-black text-slate-300 uppercase tracking-widest whitespace-nowrap">
                        {label}
                      </span>

                      <button
                        onClick={() => setActiveFilterCol(activeFilterCol === key ? null : key)}
                        className={`p-1 rounded transition-colors ${
                          (selectedFilters[key] || []).length > 0
                            ? 'bg-blue-600 text-white'
                            : 'hover:bg-slate-700 text-slate-400'
                        }`}
                      >
                        <Filter size={12} />
                      </button>

                      <div
                        className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-500/50 transition-colors"
                        onMouseDown={(e) => {
                          e.preventDefault();
                          handleColumnResize(key, e.clientX, colWidths[key]);
                        }}
                      />
                    </div>

                    {activeFilterCol === key && (
                      <div
                        ref={filterDropdownRef}
                        className="absolute top-full left-0 mt-2 z-50 w-56 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-2xl p-3 animate-in fade-in zoom-in-95 duration-150"
                      >
                        <div className="flex items-center gap-2 mb-3 p-2 bg-slate-50 dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700">
                          <Search size={14} className="text-slate-400" />
                          <input
                            type="text"
                            placeholder="Filtrar..."
                            autoFocus
                            value={colFilters[key] || ''}
                            onChange={(e) => setColFilters({ ...colFilters, [key]: e.target.value })}
                            className="w-full bg-transparent outline-none text-[10px] font-bold text-slate-800 dark:text-white"
                          />
                        </div>

                        <div className="max-h-48 overflow-y-auto space-y-1 scrollbar-thin border-t border-slate-100 dark:border-slate-700 py-2">
                          {getUniqueValues(key).map((value) => (
                            <div
                              key={value}
                              onClick={() => toggleFilterValue(key, value)}
                              className="flex items-center gap-2 p-2 hover:bg-slate-50 dark:hover:bg-slate-700 rounded-lg cursor-pointer transition-all"
                            >
                              {(selectedFilters[key] || []).includes(value) ? (
                                <CheckSquare size={14} className="text-blue-600" />
                              ) : (
                                <Square size={14} className="text-slate-300" />
                              )}
                              <span className="text-[10px] font-bold uppercase truncate text-slate-700 dark:text-slate-300">
                                {value}
                              </span>
                            </div>
                          ))}
                        </div>

                        <button
                          onClick={() => {
                            setColFilters({ ...colFilters, [key]: '' });
                            setSelectedFilters({ ...selectedFilters, [key]: [] });
                            setActiveFilterCol(null);
                          }}
                          className="w-full mt-2 py-2 text-[10px] font-black uppercase text-red-600 bg-red-50 dark:bg-red-900/30 hover:bg-red-100 dark:hover:bg-red-900/50 rounded-lg border border-red-100 dark:border-red-900/50 transition-colors"
                        >
                          Limpar Filtro
                        </button>
                      </div>
                    )}
                  </th>
                ))}
                <th className="relative p-0 border border-slate-700/50 text-[10px] font-black uppercase text-center bg-slate-900/50" style={{ width: 60 }}>
                  <Settings2 size={14} className="mx-auto opacity-40" />
                </th>
              </tr>
            </thead>

            <tbody>
              {filteredData.map((row, index) => (
                <tr
                  key={row.id}
                  className={`border-b border-slate-200/50 dark:border-slate-800/50 transition-colors ${
                    index % 2 === 0 ? '' : 'bg-black/[0.02] dark:bg-white/[0.02]'
                  }`}
                >
                  {columns.map(({ key }) => {
                    const inputClass = `w-full bg-transparent outline-none border-none px-3 py-2 text-[11px] transition-all ${
                      isDarkMode ? 'text-slate-200' : 'text-slate-900'
                    }`;

                    // ROTA - Input editÃ¡vel
                    if (key === 'rota') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={row.rota}
                            onChange={(e) => {
                              const updated = { ...row, rota: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('rota', val);
                              }
                            }}
                            className={`${inputClass} font-black text-center`}
                          />
                        </td>
                      );
                    }

                    // DATA - Input editÃ¡vel com mÃ¡scara
                    if (key === 'data') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={row.data}
                            onChange={(e) => {
                              let val = e.target.value.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              const updated = { ...row, data: val };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('data', val);
                              }
                            }}
                            className={`${inputClass} text-center font-mono`}
                          />
                        </td>
                      );
                    }

                    // PRODUTOR - Input editÃ¡vel
                    if (key === 'produtor') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={row.produtor}
                            onChange={(e) => {
                              const updated = { ...row, produtor: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              console.log('[PASTE PRODUTOR]', val);
                              
                              // MÃºltiplas linhas: bulk paste
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('produtor', val);
                                return;
                              }
                              
                              // Linha Ãºnica com cÃ³digo + produtor: separa automaticamente
                              const parsed = parseCodigoProdutor(val);
                              if (parsed) {
                                e.preventDefault();
                                const updated = { ...row, codigo: parsed.codigo, produtor: parsed.produtor };
                                setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                                console.log('[PASTE PRODUTOR] âœ… CÃ³digo e produtor separados:', parsed);
                                return;
                              }
                              
                              // Se nÃ£o tiver padrÃ£o cÃ³digo+produtor, deixa colar normalmente
                            }}
                            className={`${inputClass} font-bold`}
                          />
                        </td>
                      );
                    }

                    // CÃ“DIGO - Input editÃ¡vel
                    if (key === 'codigo') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={row.codigo}
                            onChange={(e) => {
                              const updated = { ...row, codigo: e.target.value.toUpperCase() };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              console.log('[PASTE CODIGO ROW]', val);
                              
                              // MÃºltiplas linhas: bulk paste
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('codigo', val);
                                return;
                              }
                              
                              // Linha Ãºnica com cÃ³digo + produtor: separa automaticamente
                              const parsed = parseCodigoProdutor(val);
                              if (parsed) {
                                e.preventDefault();
                                const updated = { ...row, codigo: parsed.codigo, produtor: parsed.produtor };
                                setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                                console.log('[PASTE CODIGO] âœ… CÃ³digo e produtor separados:', parsed);
                                return;
                              }
                              
                              // Se nÃ£o tiver padrÃ£o cÃ³digo+produtor, deixa colar normalmente (jÃ¡ converte para upper)
                            }}
                            className={`${inputClass} font-bold text-center uppercase`}
                          />
                        </td>
                      );
                    }

                    // MOTIVO - Select editÃ¡vel
                    if (key === 'motivo') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <select
                            value={row.motivo}
                            onChange={(e) => {
                              const selectedMotivo = e.target.value;
                              const culpabilidadeAuto = MOTIVOS_CULPABILIDADES[selectedMotivo];
                              const updated = {
                                ...row,
                                motivo: selectedMotivo,
                                culpabilidade: culpabilidadeAuto || row.culpabilidade || ''
                              };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            className={`w-full px-3 py-2 text-[11px] text-left truncate transition-all cursor-pointer outline-none border-none bg-slate-100 dark:bg-slate-800 text-slate-900 dark:text-slate-200 hover:bg-slate-200 dark:hover:bg-slate-700 focus:ring-2 focus:ring-blue-500 rounded ${
                              isDarkMode ? 'dark-mode-select' : ''
                            }`}
                          >
                            <option value="" className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">Selecione...</option>
                            {Object.keys(MOTIVOS_CULPABILIDADES).map(label => (
                              <option
                                key={label}
                                value={label}
                                className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white"
                              >
                                {label}
                              </option>
                            ))}
                          </select>
                        </td>
                      );
                    }

                    // OBSERVAÃ‡ÃƒO - Textarea editÃ¡vel
                    if (key === 'observacao') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <textarea
                            value={row.observacao}
                            onChange={(e) => {
                              const updated = { ...row, observacao: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('observacao', val);
                              }
                            }}
                            rows={1}
                            className={`${inputClass} resize-none whitespace-pre-wrap break-words min-h-[48px]`}
                          />
                        </td>
                      );
                    }

                    // AÃ‡ÃƒO - Campo automÃ¡tico baseado no MOTIVO (editÃ¡vel)
                    if (key === 'acao') {
                      // Calcula o valor automÃ¡tico da aÃ§Ã£o baseado no motivo
                      const getAcaoAutomatica = () => {
                        const motivo = (row.motivo || '').trim();
                        const rota = (row.rota || '').trim();

                        // Se nÃ£o tem motivo definido, retorna vazio
                        if (!motivo) return '';

                        // Regras especÃ­ficas
                        if (motivo.toLowerCase() === 'parou de fornecer') {
                          return 'Retirado da roteirizaÃ§Ã£o';
                        }
                        if (motivo.toLowerCase() === 'produtor suspenso') {
                          return 'Aguardando autorizaÃ§Ã£o';
                        }
                        if (motivo.toLowerCase() === 'alizarol positivo') {
                          return 'Leite Descartado';
                        }

                        // Se tem rota, gera "SerÃ¡ coletado na rota X"
                        if (rota) {
                          return `SerÃ¡ coletado na rota ${rota}`;
                        }

                        return '';
                      };

                      const acaoValue = row.acao || getAcaoAutomatica();

                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <textarea
                            value={acaoValue}
                            onChange={(e) => {
                              const updated = { ...row, acao: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('acao', val);
                              }
                            }}
                            rows={1}
                            className={`${inputClass} resize-none whitespace-pre-wrap break-words min-h-[48px]`}
                          />
                        </td>
                      );
                    }

                    // DATA AÃ‡ÃƒO - Input editÃ¡vel com mÃ¡scara (preenchida automaticamente apenas para "SerÃ¡ coletado na rota...")
                    if (key === 'dataAcao') {
                      // Calcula data aÃ§Ã£o automÃ¡tica: data + 2 dias APENAS para motivos que geram "SerÃ¡ coletado na rota..."
                      const getDataAcaoAutomatica = () => {
                        const motivo = (row.motivo || '').trim();
                        const dataNaoColeta = (row.data || '').trim();

                        // Se o usuÃ¡rio jÃ¡ preencheu dataAcao manualmente, respeita
                        if (row.dataAcao) return row.dataAcao;

                        // Para "parou de fornecer", "produtor suspenso" e "alizarol positivo", coloca hÃ­fen
                        const motivoLower = motivo.toLowerCase();
                        if (motivoLower === 'parou de fornecer' || motivoLower === 'produtor suspenso' || motivoLower === 'alizarol positivo') {
                          return '-';
                        }

                        // Se tem motivo e data, calcula data + 2 dias
                        if (motivo && dataNaoColeta && row.rota) {
                          const match = dataNaoColeta.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                          if (match) {
                            const [, day, month, year] = match;
                            const dateObj = new Date(Number(year), Number(month) - 1, Number(day));
                            dateObj.setDate(dateObj.getDate() + 2);
                            const d = String(dateObj.getDate()).padStart(2, '0');
                            const m = String(dateObj.getMonth() + 1).padStart(2, '0');
                            const y = dateObj.getFullYear();
                            return `${d}/${m}/${y}`;
                          }
                        }

                        return '';
                      };

                      const dataAcaoValue = getDataAcaoAutomatica();

                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={dataAcaoValue}
                            onChange={(e) => {
                              let val = e.target.value.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              const updated = { ...row, dataAcao: val };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('dataAcao', val);
                              }
                            }}
                            className={`${inputClass} text-center font-mono`}
                          />
                        </td>
                      );
                    }

                    // ÃšLTIMA COLETA - Input editÃ¡vel com mÃ¡scara
                    if (key === 'ultimaColeta') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={row.ultimaColeta}
                            onChange={(e) => {
                              let val = e.target.value.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              const updated = { ...row, ultimaColeta: val };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('ultimaColeta', val);
                              }
                            }}
                            className={`${inputClass} text-center font-mono`}
                          />
                        </td>
                      );
                    }

                    // OPERAÃ‡ÃƒO - Select editÃ¡vel
                    if (key === 'operacao') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <select
                            value={row.operacao}
                            onChange={(e) => {
                              const updated = { ...row, operacao: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            className={`${inputClass} text-center font-bold cursor-pointer bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-200`}
                          >
                            <option value="" className="bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-200">Selecione...</option>
                            {userConfigs.map(op => (
                              <option key={op.operacao} value={op.operacao} className="bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-200">{op.operacao}</option>
                            ))}
                          </select>
                        </td>
                      );
                    }

                    // CULPABILIDADE - Dropdown editÃ¡vel com opÃ§Ãµes fixas
                    if (key === 'culpabilidade') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <select
                            value={row.culpabilidade || ''}
                            onChange={(e) => {
                              const updated = { ...row, culpabilidade: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            className={`${inputClass} text-center cursor-pointer`}
                          >
                            <option value="" className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">Selecione...</option>
                            {CULPABILIDADE_OPCOES.map(culp => (
                              <option key={culp} value={culp} className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">{culp}</option>
                            ))}
                          </select>
                        </td>
                      );
                    }

                    // SEMANA - Somente leitura
                    if (key === 'semana') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <span className={`${inputClass} text-center block py-2`}>
                            {row.semana || '-'}
                          </span>
                        </td>
                      );
                    }

                    // Default - Texto
                    return (
                      <td
                        key={key}
                        className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                          hiddenColumns.has(key) ? 'hidden' : ''
                        }`}
                        style={{ minWidth: colWidths[key] }}
                      >
                        <span className={`${inputClass} px-3 py-2 block`}>
                          {row[key as keyof NonCollection]}
                        </span>
                      </td>
                    );
                  })}
                  {/* Coluna de aÃ§Ã£o */}
                  <td className="p-0 border border-slate-200/30 dark:border-slate-800/30 text-center">
                    <button
                      onClick={() => {
                        setNonCollections(prev => prev.filter(r => r.id !== row.id));
                      }}
                      className="p-2 text-red-600 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition-colors"
                      title="Excluir linha"
                    >
                      <Trash2 size={16} />
                    </button>
                  </td>
                </tr>
              ))}

              {/* Ghost Row - AdiÃ§Ã£o rÃ¡pida */}
              <tr
                key="ghost"
                className="border-b-2 border-blue-200 dark:border-blue-800 transition-colors bg-slate-50 dark:bg-slate-800 italic text-slate-400 border-l-4 border-dashed border-slate-300 dark:border-slate-600"
              >
                {columns.map(({ key }) => {
                  const inputClass = `w-full bg-transparent outline-none border-none px-3 py-2 text-[11px] font-semibold uppercase transition-all ${
                    isDarkMode ? 'text-slate-200 placeholder-slate-500' : 'text-slate-900 placeholder-slate-400'
                  }`;

                  if (key === 'rota') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow.rota || ''}
                          onChange={(e) => updateGhostCell('rota', e.target.value)}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('rota', val);
                            }
                          }}
                          placeholder="Cole rotas..."
                          className={`${inputClass} font-black text-center`}
                        />
                      </td>
                    );
                  }

                  if (key === 'data') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow.data || ''}
                          onChange={(e) => {
                            let val = e.target.value.replace(/\D/g, '');
                            if (val.length > 8) val = val.slice(0, 8);
                            if (val.length >= 8) {
                              val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                            }
                            updateGhostCell('data', val);
                          }}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('data', val);
                            }
                          }}
                          placeholder="DD/MM/AAAA"
                          maxLength={10}
                          className={`${inputClass} text-center font-mono`}
                        />
                      </td>
                    );
                  }

                  if (key === 'operacao') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <select
                          value={ghostRow.operacao || ''}
                          onChange={(e) => updateGhostCell('operacao', e.target.value)}
                          className={`${inputClass} text-center font-bold cursor-pointer`}
                        >
                          <option value="">Selecione...</option>
                          {userConfigs.map(op => (
                            <option key={op.operacao} value={op.operacao}>{op.operacao}</option>
                          ))}
                        </select>
                      </td>
                    );
                  }

                  if (key === 'produtor') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow.produtor || ''}
                          onChange={(e) => updateGhostCell('produtor', e.target.value)}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('produtor', val);
                            }
                          }}
                          placeholder="Cole produtores..."
                          className={`${inputClass} font-bold`}
                        />
                      </td>
                    );
                  }

                  if (key === 'codigo') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow.codigo || ''}
                          onChange={(e) => updateGhostCell('codigo', e.target.value.toUpperCase())}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('codigo', val);
                            }
                          }}
                          placeholder="Cole cÃ³digos..."
                          className={`${inputClass} font-bold text-center uppercase`}
                        />
                      </td>
                    );
                  }

                  if (key === 'motivo') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <select
                          value={ghostRow.motivo || ''}
                          onChange={(e) => updateGhostCell('motivo', e.target.value)}
                          className={`${inputClass} text-left cursor-pointer`}
                        >
                          <option value="" className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">Selecione...</option>
                          {Object.keys(MOTIVOS_CULPABILIDADES).map(label => (
                            <option key={label} value={label} className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">{label}</option>
                          ))}
                        </select>
                      </td>
                    );
                  }

                  if (key === 'culpabilidade') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <select
                          value={ghostRow.culpabilidade || ''}
                          onChange={(e) => updateGhostCell('culpabilidade', e.target.value)}
                          className={`${inputClass} text-center cursor-pointer`}
                        >
                          <option value="" className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">Selecione...</option>
                          {CULPABILIDADE_OPCOES.map(culp => (
                            <option key={culp} value={culp} className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">{culp}</option>
                          ))}
                        </select>
                      </td>
                    );
                  }

                  if (key === 'observacao') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <textarea
                          value={ghostRow.observacao || ''}
                          onChange={(e) => updateGhostCell('observacao', e.target.value)}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('observacao', val);
                            }
                          }}
                          rows={1}
                          className={`${inputClass} resize-none whitespace-pre-wrap break-words min-h-[48px]`}
                        />
                      </td>
                    );
                  }

                  if (key === 'acao') {
                    // Calcula o valor automÃ¡tico da aÃ§Ã£o baseado no motivo (ghost row)
                    const getAcaoAutomatica = () => {
                      const motivo = (ghostRow.motivo || '').trim();
                      const rota = (ghostRow.rota || '').trim();

                      // Se nÃ£o tem motivo definido, retorna vazio
                      if (!motivo) return '';

                      // Regras especÃ­ficas
                      if (motivo.toLowerCase() === 'parou de fornecer') {
                        return 'Retirado da roteirizaÃ§Ã£o';
                      }
                      if (motivo.toLowerCase() === 'produtor suspenso') {
                        return 'Aguardando autorizaÃ§Ã£o';
                      }
                      if (motivo.toLowerCase() === 'alizarol positivo') {
                        return 'Leite Descartado';
                      }

                      // Se tem rota, gera "SerÃ¡ coletado na rota X"
                      if (rota) {
                        return `SerÃ¡ coletado na rota ${rota}`;
                      }

                      return '';
                    };

                    const acaoValue = ghostRow.acao || getAcaoAutomatica();

                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <textarea
                          value={acaoValue}
                          onChange={(e) => updateGhostCell('acao', e.target.value)}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('acao', val);
                            }
                          }}
                          rows={1}
                          className={`${inputClass} resize-none whitespace-pre-wrap break-words min-h-[48px]`}
                        />
                      </td>
                    );
                  }

                  if (key === 'semana') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <span className={`${inputClass} text-center block py-2`}>
                          {ghostRow.data ? getWeekString(ghostRow.data.split('/').reverse().join('-')) : '-'}
                        </span>
                      </td>
                    );
                  }

                  if (key === 'dataAcao' || key === 'ultimaColeta') {
                    // Para dataAcao, calcula automaticamente data + 2 dias APENAS para "SerÃ¡ coletado na rota..."
                    if (key === 'dataAcao') {
                      const getDataAcaoAutomatica = () => {
                        // Se o usuÃ¡rio jÃ¡ preencheu manualmente, respeita
                        if (ghostRow.dataAcao) return ghostRow.dataAcao;

                        const motivo = (ghostRow.motivo || '').trim();
                        const dataNaoColeta = (ghostRow.data || '').trim();

                        // Motivos que geram "Leite Descartado" ou "Aguardando autorizaÃ§Ã£o" ficam com hÃ­fen
                        const motivoLower = motivo.toLowerCase();
                        if (motivoLower === 'parou de fornecer' || motivoLower === 'produtor suspenso' || motivoLower === 'alizarol positivo') {
                          return '-';
                        }

                        // Se tem motivo, data e rota, calcula data + 2 dias
                        if (motivo && dataNaoColeta && ghostRow.rota) {
                          const match = dataNaoColeta.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                          if (match) {
                            const [, day, month, year] = match;
                            const dateObj = new Date(Number(year), Number(month) - 1, Number(day));
                            dateObj.setDate(dateObj.getDate() + 2);
                            const d = String(dateObj.getDate()).padStart(2, '0');
                            const m = String(dateObj.getMonth() + 1).padStart(2, '0');
                            const y = dateObj.getFullYear();
                            return `${d}/${m}/${y}`;
                          }
                        }

                        return '';
                      };

                      const dataAcaoValue = getDataAcaoAutomatica();

                      return (
                        <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                          <input
                            type="text"
                            value={dataAcaoValue}
                            onChange={(e) => {
                              let val = e.target.value.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              updateGhostCell(key, val);
                            }}
                            maxLength={10}
                            className={`${inputClass} text-center font-mono`}
                          />
                        </td>
                      );
                    }

                    // ultimaColeta mantÃ©m comportamento normal (vazio para usuÃ¡rio preencher)
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow[key] || ''}
                          onChange={(e) => {
                            let val = e.target.value.replace(/\D/g, '');
                            if (val.length > 8) val = val.slice(0, 8);
                            if (val.length >= 8) {
                              val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                            }
                            updateGhostCell(key, val);
                          }}
                          placeholder="DD/MM/AAAA"
                          maxLength={10}
                          className={`${inputClass} text-center font-mono`}
                        />
                      </td>
                    );
                  }

                  return (
                    <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                      <input
                        type="text"
                        value={ghostRow[key] || ''}
                        onChange={(e) => updateGhostCell(key, e.target.value)}
                        placeholder=""
                        className={`${inputClass}`}
                      />
                    </td>
                  );
                })}
                {/* Coluna de aÃ§Ã£o */}
                <td className="p-0 border border-slate-200/30 dark:border-slate-800/30 text-center">
                  <button
                    onClick={handleAddFromGhost}
                    className="p-2 text-blue-600 hover:bg-blue-50 dark:hover:bg-blue-900/20 rounded-lg transition-colors"
                    title="Adicionar"
                  >
                    <Plus size={18} />
                  </button>
                </td>
              </tr>

              {filteredData.length === 0 && nonCollections.length === 0 && (
                <tr>
                  <td colSpan={columns.length + 1} className="px-4 py-16 text-center">
                    <div className="flex flex-col items-center gap-3">
                      <div className="w-16 h-16 rounded-full bg-slate-100 dark:bg-slate-800 flex items-center justify-center">
                        <Milk size={32} className="text-slate-400" />
                      </div>
                      <p className="text-sm font-bold text-slate-500 dark:text-slate-400 uppercase tracking-widest">
                        Nenhuma nÃ£o coleta registrada
                      </p>
                      <p className="text-xs font-medium text-slate-400 dark:text-slate-500">
                        Clique em "Adicionar NÃ£o Coleta" ou cole dados do Excel
                      </p>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Context Menu */}
      {contextMenu.visible && (
        <div
          ref={contextMenuRef}
          className="fixed z-[1000] bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-2xl py-2 min-w-[200px] animate-in fade-in zoom-in-95 duration-150"
          style={{ top: contextMenu.y, left: contextMenu.x }}
        >
          <div className="px-3 py-2 border-b border-slate-200 dark:border-slate-700">
            <p className="text-[10px] font-black uppercase text-slate-400">Colunas</p>
          </div>
          {columns.map(({ key, label }) => (
            <button
              key={key}
              onClick={() => toggleColumn(key)}
              className="w-full px-3 py-2 text-left hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors flex items-center gap-2"
            >
              {hiddenColumns.has(key) ? (
                <Square size={14} className="text-slate-400" />
              ) : (
                <CheckSquare size={14} className="text-blue-600" />
              )}
              <span className="text-[10px] font-bold uppercase text-slate-700 dark:text-slate-300">{label}</span>
            </button>
          ))}
        </div>
      )}

      {/* Modal de Adicionar NÃ£o Coleta */}
      {isAddModalOpen && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[200] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-lg">
            <div className="bg-[#1e293b] p-6 flex justify-between items-center text-white">
              <div className="flex items-center gap-3">
                <Milk size={24} />
                <h3 className="font-black uppercase tracking-widest text-base">Adicionar NÃ£o Coleta</h3>
              </div>
              <button
                onClick={() => setIsAddModalOpen(false)}
                className="p-2 hover:bg-slate-700 rounded-lg transition-colors"
              >
                <X size={24} />
              </button>
            </div>
            <div className="p-6 space-y-4">
              {/* OperaÃ§Ã£o */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  OperaÃ§Ã£o *
                </label>
                <select
                  value={newNonCollectionData.operacao}
                  onChange={e => setNewNonCollectionData({ ...newNonCollectionData, operacao: e.target.value })}
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors"
                >
                  <option value="">Selecione a operaÃ§Ã£o</option>
                  {userConfigs.map(config => (
                    <option key={config.operacao} value={config.operacao}>
                      {config.nomeExibicao || config.operacao}
                    </option>
                  ))}
                </select>
              </div>

              {/* Rota */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Rota *
                </label>
                <input
                  type="text"
                  value={newNonCollectionData.rota}
                  onChange={e => setNewNonCollectionData({ ...newNonCollectionData, rota: e.target.value })}
                  placeholder="Ex: ROTA 01"
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors"
                />
              </div>

              {/* Data */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Data *
                </label>
                <input
                  type="text"
                  value={newNonCollectionData.data}
                  onChange={e => {
                    let val = e.target.value.replace(/\D/g, '');
                    if (val.length > 8) val = val.slice(0, 8);
                    if (val.length >= 8) {
                      val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                    }
                    setNewNonCollectionData({ ...newNonCollectionData, data: val });
                  }}
                  placeholder="DD/MM/AAAA"
                  maxLength={10}
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors font-mono text-center"
                />
              </div>

              {/* CÃ³digo */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  CÃ³digo
                </label>
                <input
                  type="text"
                  value={newNonCollectionData.codigo}
                  onChange={e => setNewNonCollectionData({ ...newNonCollectionData, codigo: e.target.value.toUpperCase() })}
                  placeholder="Ex: 123456 ou P001"
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors uppercase"
                />
              </div>

              {/* Produtor */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Produtor *
                </label>
                <input
                  type="text"
                  value={newNonCollectionData.produtor}
                  onChange={e => setNewNonCollectionData({ ...newNonCollectionData, produtor: e.target.value })}
                  placeholder="Nome do produtor"
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors"
                />
              </div>
            </div>
            <div className="p-6 pt-0 flex gap-3">
              <button
                onClick={() => setIsAddModalOpen(false)}
                className="flex-1 py-3 bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 rounded-xl font-black uppercase text-[10px] tracking-widest hover:bg-slate-200 dark:hover:bg-slate-700 transition-all"
              >
                Cancelar
              </button>
              <button
                onClick={handleAddNonCollection}
                className="flex-1 py-3 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-black uppercase text-[10px] tracking-widest transition-all shadow-lg shadow-blue-500/20"
              >
                Adicionar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Footer Info */}
      <div className="px-6 py-3 text-center">
        <p className="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest">
          VisualizaÃ§Ã£o apenas â€¢ Dados nÃ£o persistidos
        </p>
      </div>

      {/* Modal de SeleÃ§Ã£o de OperaÃ§Ã£o (Bulk Paste) */}
      {isOperationModalOpen && (
        <div className="fixed inset-0 bg-slate-950/90 backdrop-blur-md z-[300] flex items-center justify-center p-4 animate-in fade-in zoom-in-95 duration-200">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-8 w-full max-w-md border border-blue-500/50 shadow-2xl animate-in zoom-in duration-300">
            <div className="flex items-center gap-3 text-blue-600 dark:text-blue-400 mb-6 font-black uppercase text-xs">
              <Milk size={24} />
              Selecionar OperaÃ§Ã£o
            </div>
            <p className="text-sm text-slate-600 dark:text-slate-400 mb-2 font-medium">
              {pendingBulkRoutes.length} rota(s) colada(s). Selecione a operaÃ§Ã£o:
            </p>
            <div className="mb-6 max-h-32 overflow-y-auto scrollbar-thin bg-slate-50 dark:bg-slate-800 rounded-xl p-3 border border-slate-200 dark:border-slate-700">
              {pendingBulkRoutes.slice(0, 10).map((rota, i) => (
                <div key={i} className="text-[10px] font-bold text-slate-600 dark:text-slate-400 py-1 truncate">
                  â€¢ {rota}
                </div>
              ))}
              {pendingBulkRoutes.length > 10 && (
                <div className="text-[10px] font-bold text-slate-400 italic mt-1">
                  ...e mais {pendingBulkRoutes.length - 10} rota(s)
                </div>
              )}
            </div>

            {userConfigs.length === 0 ? (
              <div className="text-center py-6 text-slate-400 dark:text-slate-500">
                <AlertTriangle size={32} className="mx-auto mb-3 opacity-50" />
                <p className="text-sm font-bold">Nenhuma operaÃ§Ã£o disponÃ­vel</p>
                <p className="text-[10px] mt-1">Contate o administrador para configurar suas operaÃ§Ãµes.</p>
              </div>
            ) : (
              <div className="grid grid-cols-2 gap-3">
                {userConfigs.map(op => (
                  <button
                    key={op.operacao}
                    onClick={() => createBulkRecordsWithOperation(op.operacao)}
                    className="p-4 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-2xl hover:bg-blue-600 hover:text-white hover:border-blue-700 dark:hover:bg-blue-600 dark:hover:border-blue-500 transition-all font-black text-xs uppercase text-slate-700 dark:text-slate-300"
                  >
                    {op.operacao}
                  </button>
                ))}
              </div>
            )}

            <button
              onClick={cancelBulkPaste}
              className="w-full mt-6 py-4 text-[10px] font-black uppercase text-slate-400 hover:text-slate-600 dark:hover:text-slate-300 transition-colors"
            >
              Cancelar
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default NonCollectionsView;
