import React, { useState, useEffect, useRef } from 'react';
import { User } from '../types';
import {
  Search, Filter, CheckSquare, Square, ChevronDown,
  Sun, Moon, Table, SortAsc, Plus, X, Loader2, CheckCircle2, Milk, RefreshCw
} from 'lucide-react';
import { getBrazilDate, getWeekString } from '../utils/dateUtils';
import { SharePointService } from '../services/sharepointService';
import { getValidToken } from '../services/tokenService';
import { RouteConfig } from '../types';

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

// Lista de MOTIVOS no padrão "MOTIVO - CULPABILIDADE"
const MOTIVOS_CULPABILIDADES = [
  { label: 'Rota Atrasada/Fábrica', culpabilidades: ['Cliente'] },
  { label: 'Rota Atrasada/Logística', culpabilidades: ['VIA'] },
  { label: 'Rota Atrasada/Manutenção', culpabilidades: ['VIA'] },
  { label: 'Rota Atrasada/Mão De Obra', culpabilidades: ['VIA'] },
  { label: 'Eqp. Cheio/Coleta Extra', culpabilidades: ['VIA', 'Cliente', 'Outros'], causaRaiz: true },
  { label: 'Eqp. Cheio/Eqp Menor', culpabilidades: ['VIA', 'Cliente'], causaRaiz: true },
  { label: 'Eqp. Cheio/Rota Atrasada', culpabilidades: ['VIA', 'Cliente', 'Outros'], causaRaiz: true },
  { label: 'Alizarol Positivo', culpabilidades: ['Outros'] },
  { label: 'Leite Com Antibiótico', culpabilidades: ['Outros'] },
  { label: 'Leite Congelado', culpabilidades: ['Outros'] },
  { label: 'Leite Descartado', culpabilidades: ['Outros'] },
  { label: 'Leite Quente', culpabilidades: ['Outros'] },
  { label: 'Não Coletado - Saúde Do Motorista', culpabilidades: ['VIA'] },
  { label: 'Não Coletado - Greve', culpabilidades: ['Outros'] },
  { label: 'Não Coletado - Solicitado Pelo Sdl', culpabilidades: ['Cliente'] },
  { label: 'Objeto Ou Sujeira No Leite', culpabilidades: ['Outros'] },
  { label: 'Parou De Fornecer', culpabilidades: ['Outros'] },
  { label: 'Produtor suspenso', culpabilidades: ['Cliente'] },
  { label: 'Passou Para 48Hrs', culpabilidades: ['Outros'] },
  { label: 'Problemas Mecânicos Equipamento', culpabilidades: ['VIA'] },
  { label: 'Produtor Solicitou A Não Coleta', culpabilidades: ['Outros'] },
  { label: 'Resfriador Vazio', culpabilidades: ['Outros'] },
  { label: 'Volume Insuficiente Para Medida', culpabilidades: ['Outros'] },
  { label: 'Coletado Por Outra Transportadora', culpabilidades: ['VIA', 'Cliente', 'Outros'], causaRaiz: true },
  { label: 'Descumprimento de roteirização', culpabilidades: ['VIA'] },
  { label: 'A rota não foi realizada', culpabilidades: ['VIA', 'Cliente'], causaRaiz: true },
  { label: 'Falta De Acesso', culpabilidades: ['Outros'] },
  { label: 'Jornada Excedida', culpabilidades: ['VIA'] },
  { label: 'Eqp. Cheio/Aumento de Vol.', culpabilidades: ['Outros'] },
  { label: 'Correção de Roteirzação', culpabilidades: ['VIA'] },
  { label: 'Crioscopia', culpabilidades: ['Outros'] },
  { label: 'Rota Atrasada/Infraestrutura', culpabilidades: ['Outros'] }
];

const CULPABILIDADE_OPCOES = ['VIA', 'Cliente', 'Outros'];

const NonCollectionsView: React.FC<{
  currentUser: User;
}> = ({ currentUser }) => {
  // Dados reais (inicialmente vazio)
  const [nonCollections, setNonCollections] = useState<NonCollection[]>([]);
  const [userOps, setUserOps] = useState<RouteConfig[]>([]);
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

  // Ghost Row para adição rápida
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
    dataAcao: getBrazilDate().split('-').reverse().join('/'),
    ultimaColeta: getBrazilDate().split('-').reverse().join('/'),
    culpabilidade: '',
    operacao: ''
  });

  // Estado para mostrar campo de causa raiz na ghost row
  const [showCausaRaiz, setShowCausaRaiz] = useState(false);

  const filterDropdownRef = useRef<HTMLDivElement>(null);

  // Fecha dropdown ao clicar fora
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (filterDropdownRef.current && !filterDropdownRef.current.contains(e.target as Node)) {
        setActiveFilterCol(null);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Carrega operações do usuário ao montar
  useEffect(() => {
    loadUserOperations();
  }, [currentUser]);

  const loadUserOperations = async () => {
    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) {
        console.error('[NonCollections] Token não encontrado');
        return;
      }

      const configs = await SharePointService.getRouteConfigs(token, currentUser.email, true);
      setUserOps(configs || []);
      console.log('[NonCollections] Operações carregadas:', configs?.map(c => c.operacao));
    } catch (e) {
      console.error('[NonCollections] Erro ao carregar operações:', e);
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

  const filteredData = React.useMemo(() => {
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

    // Ordenação por data
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

  const handleColumnResize = (col: string, startX: number, startWidth: number) => {
    const handleMouseMove = (e: MouseEvent) => {
      const newWidth = startWidth + (e.clientX - startX);
      setColWidths(prev => ({ ...prev, [col]: Math.max(50, newWidth) }));
    };

    const handleMouseUp = () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  const getRowStyle = (row: NonCollection | Partial<NonCollection>) => {
    // Ghost row style
    if (row.id === 'ghost') {
      return isDarkMode
        ? 'bg-slate-800 italic text-slate-400 border-l-4 border-dashed border-slate-600'
        : 'bg-slate-50 italic text-slate-500 border-l-4 border-dashed border-slate-300';
    }

    const culpabilidade = row.culpabilidade?.toLowerCase();
    
    if (culpabilidade === 'produtor') {
      return isDarkMode
        ? 'bg-orange-900/30 text-orange-100 border-l-[12px] border-orange-600'
        : 'bg-orange-200 text-orange-900 border-l-[12px] border-orange-600';
    }
    
    if (culpabilidade === 'logística' || culpabilidade === 'logistica') {
      return isDarkMode
        ? 'bg-yellow-900/30 text-yellow-100 border-l-[12px] border-yellow-600'
        : 'bg-yellow-200 text-yellow-900 border-l-[12px] border-yellow-600';
    }
    
    if (culpabilidade === 'clima') {
      return isDarkMode
        ? 'bg-blue-900/30 text-blue-100 border-l-[12px] border-blue-600'
        : 'bg-blue-200 text-blue-900 border-l-[12px] border-blue-600';
    }

    return isDarkMode
      ? 'bg-slate-800 border-l-4 border-slate-600 text-slate-300'
      : 'bg-white border-l-4 border-slate-300 text-slate-800';
  };

  // Processa colagem em massa e adiciona múltiplas não coletas
  const handleBulkPaste = async (pastedText: string) => {
    const lines = pastedText.trim().split('\n');
    const newRecords: NonCollection[] = [];

    for (const line of lines) {
      const cols = line.split('\t').map(c => c.trim());
      
      // Espera pelo menos: Rota, Data, Produtor, Motivo, Operação
      if (cols.length < 5) continue;

      const [rota, data, produtor, motivo, operacao, codigo, observacao, acao, culpabilidade] = cols;

      // Validação mínima
      if (!rota || !produtor || !motivo || !operacao) continue;

      // Formata data se necessário
      let dataFormatada = data || getBrazilDate().split('-').reverse().join('/');
      if (dataFormatada.includes('-')) {
        const [year, month, day] = dataFormatada.split('-');
        dataFormatada = `${day}/${month}/${year}`;
      }

      // Calcula semana
      const dataParaSemana = dataFormatada.split('/').reverse().join('-');
      const semana = getWeekString(dataParaSemana);

      // Busca culpabilidade do motivo
      const motivoData = MOTIVOS_CULPABILIDADES.find(m => m.label === motivo);
      const culpabilidadeFinal = culpabilidade || (motivoData?.culpabilidades.length === 1 ? motivoData.culpabilidades[0] : 'Outros');

      newRecords.push({
        id: Date.now().toString() + Math.random().toString(36).substr(2, 5),
        semana,
        rota,
        data: dataFormatada,
        codigo: codigo || `P${String(nonCollections.length + newRecords.length + 1).padStart(3, '0')}`,
        produtor,
        motivo,
        observacao: observacao || '',
        acao: acao || '',
        dataAcao: dataFormatada,
        ultimaColeta: dataFormatada,
        culpabilidade: culpabilidadeFinal,
        operacao
      });
    }

    if (newRecords.length > 0) {
      setNonCollections(prev => [...prev, ...newRecords]);
      alert(`✅ ${newRecords.length} não coleta(s) adicionada(s)!`);
      
      // Reseta ghost row
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
        dataAcao: getBrazilDate().split('-').reverse().join('/'),
        ultimaColeta: getBrazilDate().split('-').reverse().join('/'),
        culpabilidade: '',
        operacao: ''
      });
      setShowCausaRaiz(false);
    } else {
      alert('⚠️ Nenhum dado válido encontrado na colagem. Verifique o formato.');
    }
  };

  // Adiciona nova não coleta a partir da ghost row
  const handleAddFromGhost = async () => {
    // Validação básica
    if (!ghostRow.rota || !ghostRow.data || !ghostRow.produtor || !ghostRow.motivo || !ghostRow.operacao) {
      alert('Preencha todos os campos obrigatórios na linha de criação!');
      return;
    }

    // Validação de causa raiz se necessário
    const motivoSelecionado = MOTIVOS_CULPABILIDADES.find(m => m.label === ghostRow.motivo);
    if (motivoSelecionado?.causaRaiz && !ghostRow.observacao?.includes('Causa raiz:')) {
      alert('⚠️ Este motivo requer que você informe a causa raiz na observação!');
      return;
    }

    try {
      // Calcula a semana automaticamente baseado na data
      const dataParaSemana = ghostRow.data!.split('/').reverse().join('-');
      const semana = getWeekString(dataParaSemana);

      const newRecord: NonCollection = {
        id: Date.now().toString(),
        semana,
        rota: ghostRow.rota!,
        data: ghostRow.data!,
        codigo: ghostRow.codigo || `P${String(nonCollections.length + 1).padStart(3, '0')}`,
        produtor: ghostRow.produtor!,
        motivo: ghostRow.motivo!,
        observacao: ghostRow.observacao || '',
        acao: ghostRow.acao || '',
        dataAcao: ghostRow.dataAcao || ghostRow.data!,
        ultimaColeta: ghostRow.ultimaColeta || ghostRow.data!,
        culpabilidade: ghostRow.culpabilidade || 'Não se aplica',
        operacao: ghostRow.operacao!
      };

      setNonCollections(prev => [...prev, newRecord]);

      // Reseta ghost row
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
        dataAcao: getBrazilDate().split('-').reverse().join('/'),
        ultimaColeta: getBrazilDate().split('-').reverse().join('/'),
        culpabilidade: '',
        operacao: ''
      });
      setShowCausaRaiz(false);
    } catch (e) {
      console.error('[ADD_NON_COLLECTION] Error:', e);
      alert('Erro ao adicionar não coleta');
    }
  };

  // Atualiza célula da ghost row
  const updateGhostCell = (field: keyof NonCollection, value: string) => {
    const updatedGhost = { ...ghostRow, [field]: value };

    // Se mudou o motivo, atualiza culpabilidade automaticamente
    if (field === 'motivo') {
      const motivoData = MOTIVOS_CULPABILIDADES.find(m => m.label === value);
      if (motivoData) {
        // Atualiza culpabilidade se houver apenas uma opção
        if (motivoData.culpabilidades.length === 1) {
          updatedGhost.culpabilidade = motivoData.culpabilidades[0];
        }
        // Mostra campo de causa raiz se necessário
        setShowCausaRaiz(!!motivoData.causaRaiz);
      }
    }

    setGhostRow(updatedGhost);
  };

  const columns: { key: keyof NonCollection; label: string }[] = [
    { key: 'semana', label: 'SEMANA' },
    { key: 'rota', label: 'ROTA' },
    { key: 'data', label: 'DATA' },
    { key: 'codigo', label: 'CÓDIGO' },
    { key: 'produtor', label: 'PRODUTOR' },
    { key: 'motivo', label: 'MOTIVO' },
    { key: 'observacao', label: 'OBSERVAÇÃO' },
    { key: 'acao', label: 'AÇÃO' },
    { key: 'dataAcao', label: 'DATA AÇÃO' },
    { key: 'ultimaColeta', label: 'ÚLTIMA COLETA' },
    { key: 'culpabilidade', label: 'CULPABILIDADE' },
    { key: 'operacao', label: 'OPERAÇÃO' }
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
        <div>
          <h1 className="text-2xl font-black uppercase text-slate-800 dark:text-white tracking-tight">
            Não Coletas
          </h1>
          <p className="text-xs font-bold text-slate-500 dark:text-slate-400 mt-1 uppercase tracking-widest">
            Acompanhamento de ocorrências
          </p>
        </div>

        <div className="flex items-center gap-3">
          {/* Refresh Button */}
          <button
            onClick={loadUserOperations}
            className="p-3 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-sm hover:shadow-md transition-all"
            title="Recarregar operações"
          >
            <RefreshCw size={20} className="text-slate-600 dark:text-slate-400" />
          </button>

          {/* Toggle Dark/Light Mode */}
          <button
            onClick={() => setIsDarkMode(!isDarkMode)}
            className="p-3 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-sm hover:shadow-md transition-all"
            title={isDarkMode ? 'Modo Claro' : 'Modo Escuro'}
          >
            {isDarkMode ? (
              <Sun size={20} className="text-amber-400" />
            ) : (
              <Moon size={20} className="text-slate-600" />
            )}
          </button>

          {/* Toggle Sort */}
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

          {/* Clear Filters */}
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

      {/* Info Bar - Operações do usuário */}
      <div className="flex items-center gap-3 px-6 py-2">
        <div className="flex items-center gap-2 px-3 py-1.5 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800 rounded-lg">
          <CheckSquare size={14} className="text-blue-600 dark:text-blue-400" />
          <span className="text-[10px] font-bold text-blue-700 dark:text-blue-300 uppercase">
            {userOps.length} operação(ões) disponível(eis)
          </span>
        </div>
        {userOps.length > 0 && (
          <div className="flex gap-1 flex-wrap">
            {userOps.map(op => (
              <span key={op.operacao} className="px-2 py-0.5 bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded text-[9px] font-bold text-slate-600 dark:text-slate-400 uppercase">
                {op.operacao}
              </span>
            ))}
          </div>
        )}
      </div>

      {/* Filters Bar */}
      <div className="flex items-center gap-3 px-6 py-3">
        <div className="flex items-center gap-2 px-4 py-2.5 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-sm flex-1 max-w-md">
          <Search size={16} className="text-slate-400" />
          <input
            type="text"
            placeholder="Buscar..."
            className="flex-1 bg-transparent outline-none text-sm font-medium text-slate-700 dark:text-slate-300 placeholder-slate-400"
          />
        </div>

        <div className="flex items-center gap-2 px-4 py-2.5 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-sm">
          <Table size={16} className="text-slate-400" />
          <span className="text-xs font-bold text-slate-600 dark:text-slate-400 uppercase">
            {filteredData.length} registro(s)
          </span>
        </div>
      </div>

      {/* Table Container */}
      <div className="flex-1 mx-6 mt-4 mb-6 bg-white/95 dark:bg-slate-900/95 backdrop-blur-sm rounded-[2rem] shadow-2xl border border-white/50 dark:border-slate-800 overflow-hidden flex flex-col">
        {/* Table Header */}
        <div className="overflow-x-auto flex-1 scrollbar-thin">
          <table className="w-full border-collapse">
            <thead className="sticky top-0 z-20">
              <tr className="bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900">
                {columns.map(({ key, label }) => (
                  <th
                    key={key}
                    className="relative px-4 py-4 text-left border-r border-slate-700/50 last:border-r-0 select-none"
                    style={{ width: colWidths[key] }}
                  >
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] font-black text-slate-300 uppercase tracking-widest whitespace-nowrap">
                        {label}
                      </span>
                      
                      {/* Filter Button */}
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

                      {/* Resize Handle */}
                      <div
                        className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-500/50 transition-colors"
                        onMouseDown={(e) => {
                          e.preventDefault();
                          handleColumnResize(key, e.clientX, colWidths[key]);
                        }}
                      />
                    </div>

                    {/* Filter Dropdown */}
                    {activeFilterCol === key && (
                      <div
                        ref={filterDropdownRef}
                        className="absolute top-full left-0 mt-2 z-50 w-56 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-2xl p-3 animate-in fade-in zoom-in-95 duration-150"
                      >
                        {/* Search */}
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

                        {/* Filter Options */}
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

                        {/* Clear Filter Button */}
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
              </tr>
            </thead>

            {/* Table Body */}
            <tbody>
              {/* Dados reais */}
              {filteredData.map((row, index) => (
                <tr
                  key={row.id}
                  className={`border-b border-slate-200/50 dark:border-slate-800/50 transition-colors ${getRowStyle(row)} ${
                    index % 2 === 0 ? '' : 'bg-black/[0.02] dark:bg-white/[0.02]'
                  }`}
                >
                  {columns.map(({ key }) => (
                    <td
                      key={key}
                      className="px-4 py-3 border-r border-slate-200/30 dark:border-slate-800/30 last:border-r-0"
                      style={{ minWidth: colWidths[key] }}
                    >
                      {key === 'culpabilidade' ? (
                        <span className={`inline-flex items-center px-2.5 py-1 rounded-full text-[9px] font-black uppercase tracking-wider ${
                          row.culpabilidade?.toLowerCase() === 'produtor'
                            ? 'bg-orange-100 dark:bg-orange-900/40 text-orange-700 dark:text-orange-300'
                            : row.culpabilidade?.toLowerCase() === 'logística' || row.culpabilidade?.toLowerCase() === 'logistica'
                            ? 'bg-yellow-100 dark:bg-yellow-900/40 text-yellow-700 dark:text-yellow-300'
                            : row.culpabilidade?.toLowerCase() === 'clima'
                            ? 'bg-blue-100 dark:bg-blue-900/40 text-blue-700 dark:text-blue-300'
                            : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400'
                        }`}>
                          {row[key]}
                        </span>
                      ) : key === 'motivo' ? (
                        <span className="text-[10px] font-bold uppercase tracking-wide">
                          {row[key]}
                        </span>
                      ) : (
                        <span className="text-[10px] font-medium">
                          {row[key]}
                        </span>
                      )}
                    </td>
                  ))}
                </tr>
              ))}

              {/* Ghost Row (sempre visível no final) */}
              <tr
                key="ghost"
                className={`border-b-2 border-blue-200 dark:border-blue-800 transition-colors ${getRowStyle(ghostRow)}`}
                onPaste={(e) => {
                  e.preventDefault();
                  const text = e.clipboardData.getData('text');
                  if (text.includes('\n') || text.split('\t').length >= 5) {
                    handleBulkPaste(text);
                  }
                }}
              >
                {columns.map(({ key }) => {
                  const isGhost = true;
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
                          placeholder="Cole ou digite..."
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
                            } else if (val.length >= 4) {
                              val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4)}`;
                            } else if (val.length >= 2) {
                              val = `${val.slice(0, 2)}/${val.slice(2)}`;
                            }
                            updateGhostCell('data', val);
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
                          {userOps.map(op => (
                            <option key={op.operacao} value={op.operacao}>{op.operacao}</option>
                          ))}
                        </select>
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
                          <option value="">Selecione...</option>
                          {MOTIVOS_CULPABILIDADES.map(m => (
                            <option key={m.label} value={m.label}>{m.label}</option>
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
                          disabled={!ghostRow.motivo}
                        >
                          <option value="">Selecione...</option>
                          {(() => {
                            const motivoData = MOTIVOS_CULPABILIDADES.find(m => m.label === ghostRow.motivo);
                            const opcoes = motivoData?.culpabilidades || CULPABILIDADE_OPCOES;
                            return opcoes.map(culp => (
                              <option key={culp} value={culp}>{culp}</option>
                            ));
                          })()}
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
                          placeholder="Nome do produtor"
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
                          placeholder="P001"
                          className={`${inputClass} font-bold text-center uppercase`}
                        />
                      </td>
                    );
                  }

                  if (key === 'observacao' || key === 'acao') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <textarea
                          value={ghostRow[key] || ''}
                          onChange={(e) => updateGhostCell(key, e.target.value)}
                          placeholder={key === 'observacao' ? 'Detalhes + causa raiz...' : 'Ação...'}
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
                          {ghostRow.data ? (() => {
                            const dataParaSemana = ghostRow.data.split('/').reverse().join('-');
                            return getWeekString(dataParaSemana);
                          })() : '-'}
                        </span>
                      </td>
                    );
                  }

                  if (key === 'dataAcao' || key === 'ultimaColeta') {
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

                  // Default input
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
              </tr>

              {/* Botão de adicionar na linha ghost */}
              <tr className={`${isDarkMode ? 'bg-slate-900' : 'bg-slate-100'}`}>
                <td colSpan={columns.length} className="p-3 text-center">
                  <button
                    onClick={handleAddFromGhost}
                    className="px-6 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-black uppercase text-[10px] tracking-widest transition-all flex items-center gap-2 mx-auto shadow-lg shadow-blue-500/30"
                  >
                    <Plus size={14} /> Adicionar Não Coleta
                  </button>
                  <p className="text-[9px] font-bold text-slate-400 dark:text-slate-500 mt-2 uppercase">
                    Ou cole múltiplas linhas (Ctrl+V) com colunas separadas por tab
                  </p>
                </td>
              </tr>

              {filteredData.length === 0 && nonCollections.length === 0 && (
                <tr>
                  <td colSpan={columns.length} className="px-4 py-16 text-center">
                    <div className="flex flex-col items-center gap-3">
                      <div className="w-16 h-16 rounded-full bg-slate-100 dark:bg-slate-800 flex items-center justify-center">
                        <Milk size={32} className="text-slate-400" />
                      </div>
                      <p className="text-sm font-bold text-slate-500 dark:text-slate-400 uppercase tracking-widest">
                        Nenhuma não coleta registrada
                      </p>
                      <p className="text-xs font-medium text-slate-400 dark:text-slate-500">
                        Preencha a linha acima ou cole dados do Excel
                      </p>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Footer Info */}
      <div className="px-6 py-3 text-center">
        <p className="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest">
          Visualização apenas • Dados não persistidos
        </p>
      </div>
    </div>
  );
};

export default NonCollectionsView;
