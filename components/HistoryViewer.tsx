
import { HistoryRecord, Task, User as AppUser, RouteConfig, OperationStatus } from '../types';
import { SharePointService } from '../services/sharepointService';
import { getBrazilDate, toBrazilDate } from '../utils/dateUtils';
import { History, Calendar, Clock, ChevronRight, ChevronDown, CheckCircle2, User, Loader2, MapPin, Eye, FileSearch, Filter, X, Search } from 'lucide-react';
import React, { useState, useEffect, useMemo, useRef } from 'react';

const STATUS_CONFIG: Record<string, { label: string, color: string }> = {
  'PR': { label: 'PR', color: 'bg-slate-200 text-slate-600 border-slate-300 dark:bg-slate-700 dark:text-slate-300 dark:border-slate-600' },
  'OK': { label: 'OK', color: 'bg-green-200 text-green-800 border-green-300 dark:bg-green-900/60 dark:text-green-300 dark:border-green-800' },
  'EA': { label: 'EA', color: 'bg-yellow-200 text-yellow-800 border-yellow-300 dark:bg-yellow-900/60 dark:text-yellow-300 dark:border-yellow-800' },
  'AR': { label: 'AR', color: 'bg-orange-200 text-orange-800 border-orange-300 dark:bg-orange-900/60 dark:text-orange-300 dark:border-orange-800' },
  'ATT': { label: 'ATT', color: 'bg-blue-200 text-blue-800 border-blue-300 dark:bg-blue-900/60 dark:text-blue-300 dark:border-blue-800' },
  'AT': { label: 'AT', color: 'bg-red-500 text-white border-red-600 dark:bg-red-800 dark:text-white dark:border-red-700' },
};

interface HistoryViewerProps {
    currentUser: AppUser;
}

type EnhancedHistoryRecord = HistoryRecord & {
    hasPartial?: boolean;
    partialRecord?: HistoryRecord;
};

const HistoryViewer: React.FC<HistoryViewerProps> = ({ currentUser }) => {
  const [history, setHistory] = useState<HistoryRecord[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [selectedRecord, setSelectedRecord] = useState<EnhancedHistoryRecord | null>(null);
  const [viewingPartial, setViewingPartial] = useState(false);
  const [collapsedCategories, setCollapsedCategories] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(true);

  // Estados para filtros das colunas (igual à tabela de saídas)
  const [colFilters, setColFilters] = useState<Record<string, string[]>>({});
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>({});
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const filterDropdownRef = useRef<HTMLDivElement>(null);

  // Filtros de busca manual
  const [startDate, setStartDate] = useState(() => {
    // Usa data brasileira diretamente
    const today = getBrazilDate();
    const [year, month, day] = today.split('-').map(Number);
    const startDateObj = new Date(year, month - 1, day - 7);
    return startDateObj.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' }).split('/').reverse().join('-');
  });
  const [endDate, setEndDate] = useState(getBrazilDate());

  const fetchHistory = async () => {
    const token = currentUser.accessToken || (window as any).__access_token; 
    if (!token) return;
    
    setIsLoading(true);
    try {
        // Carrega configurações de permissão do usuário
        const configs = await SharePointService.getRouteConfigs(token, currentUser.email);
        setUserConfigs(configs);

        // Busca registros históricos
        const data = await SharePointService.getHistory(token, currentUser.email);
        setHistory(data);
    } catch (e) {
        console.error("Erro ao carregar histórico:", e);
    } finally {
        setIsLoading(false);
    }
  };

  useEffect(() => { fetchHistory(); }, [currentUser]);

  const displayTimestamp = (timestamp: string) => {
    try {
        if (!timestamp) return "--/--/---- --:--";
        if (timestamp.includes('/') && !timestamp.includes('T')) return timestamp;
        const date = new Date(timestamp);
        // Usa fuso de Brasília para exibição consistente
        return date.toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    } catch(e) { return timestamp; }
  };

  // Formata data ISO ou YYYY-MM-DD para DD/MM/AAAA
  const formatDateBR = (dateString: string) => {
    if (!dateString) return "--/--/----";
    try {
        // Se já estiver no formato DD/MM/AAAA, retorna como está
        if (dateString.includes('/') && /^\d{2}\/\d{2}\/\d{4}$/.test(dateString)) {
            return dateString;
        }
        // Se for ISO completo (com T e hora), usa new Date normal
        if (dateString.includes('T')) {
            const date = new Date(dateString);
            if (isNaN(date.getTime())) return dateString;
            return date.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });
        }
        // Se for apenas data (YYYY-MM-DD), parse manualmente para evitar problema de fuso
        const match = dateString.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (match) {
            const [, year, month, day] = match;
            // Cria data com timezone local explícito (meio-dia para evitar DST issues)
            return `${day}/${month}/${year}`;
        }
        // Fallback
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString;
        return date.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    } catch {
        return dateString;
    }
  };

  // Funções de filtro das colunas (igual à tabela de saídas)
  const getColUniqueValues = (col: string) => {
    if (!selectedRecord) return [];
    const tasks = viewingPartial && selectedRecord.partialRecord 
      ? selectedRecord.partialRecord.tasks 
      : selectedRecord.tasks || [];
    
    if (col === 'operacao') {
      const allOps = Array.from(new Set(tasks.flatMap(t => Object.keys(t.operations))));
      return allOps.filter(op => userConfigs.some(c => c.operacao === op)).sort();
    }
    
    if (col === 'categoria') {
      return Array.from(new Set(tasks.map(t => t.category || 'Geral'))).sort();
    }
    
    return [];
  };

  const toggleColFilter = (col: string, value: string) => {
    setSelectedFilters(prev => {
      const current = prev[col] || [];
      const updated = current.includes(value)
        ? current.filter(v => v !== value)
        : [...current, value];
      return { ...prev, [col]: updated };
    });
  };

  const clearColFilters = () => {
    setSelectedFilters({});
    setColFilters({});
  };

  const hasActiveColFilters = Object.keys(selectedFilters).some(col => (selectedFilters[col] || []).length > 0);

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

  // Aplica filtros às tarefas exibidas
  const filteredTasksToDisplay = useMemo(() => {
    let tasks = viewingPartial && selectedRecord?.partialRecord 
      ? selectedRecord.partialRecord.tasks 
      : selectedRecord?.tasks || [];

    if (!hasActiveColFilters) return tasks;

    return tasks.filter(task => {
      // Filtro por categoria
      if (selectedFilters['categoria']?.length > 0) {
        if (!selectedFilters['categoria'].includes(task.category || 'Geral')) {
          return false;
        }
      }

      // Filtro por operação (status em uma operação específica)
      if (selectedFilters['operacao']?.length > 0) {
        const taskOps = Object.entries(task.operations).filter(([op, status]) => 
          selectedFilters['operacao']?.includes(op)
        );
        if (taskOps.length === 0) return false;
      }

      return true;
    });
  }, [selectedRecord, viewingPartial, selectedFilters, userConfigs]);

  const processedHistory = useMemo(() => {
    const groupedByDay: Record<string, HistoryRecord[]> = {};
    
    history.forEach(rec => {
        const dateStr = displayTimestamp(rec.timestamp).split(',')[0].trim();
        if (!groupedByDay[dateStr]) groupedByDay[dateStr] = [];
        groupedByDay[dateStr].push(rec);
    });

    const finalHistory: EnhancedHistoryRecord[] = [];
    Object.keys(groupedByDay).forEach(date => {
        const dayRecs = groupedByDay[date];
        const partials = dayRecs.filter(r => r.resetBy === 'Salvamento automático (10:00h)');
        const mains = dayRecs.filter(r => r.resetBy !== 'Salvamento automático (10:00h)');

        if (mains.length > 0) {
            mains.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
            mains.forEach((main, idx) => {
                const item: EnhancedHistoryRecord = { ...main };
                if (idx === 0 && partials.length > 0) {
                    item.hasPartial = true;
                    item.partialRecord = partials[0];
                }
                finalHistory.push(item);
            });
        } else { finalHistory.push(...partials); }
    });

    const sorted = finalHistory.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
    if (sorted.length > 0 && !selectedRecord) setSelectedRecord(sorted[0]);
    return sorted;
  }, [history]);

  // Filtro Automático por Permissão de Célula
  const filteredAllowedLocations = useMemo(() => {
    if (!selectedRecord) return [];
    const recordTasks = viewingPartial && selectedRecord?.partialRecord ? selectedRecord.partialRecord.tasks : selectedRecord?.tasks || [];
    const allLocsInRecord = Array.from(new Set(recordTasks.flatMap(t => Object.keys(t.operations))));
    const myAllowedOps = new Set(userConfigs.map(c => c.operacao));
    return allLocsInRecord.filter(loc => myAllowedOps.has(loc));
  }, [selectedRecord, viewingPartial, userConfigs]);

  const currentTasksToDisplay = filteredTasksToDisplay;

  const getGroupedTasks = (tasks: Task[]) => {
    return tasks.reduce((acc, task) => {
      const cat = task.category || 'Geral';
      if (!acc[cat]) acc[cat] = [];
      acc[cat].push(task);
      return acc;
    }, {} as Record<string, Task[]>);
  };

  return (
    <div id="history-container" className="flex flex-col md:flex-row h-full gap-4 animate-fade-in bg-slate-50 dark:bg-slate-950 p-2">
      {/* Sidebar - Filtros de Data e Lista */}
      <div className="w-full md:w-80 bg-white dark:bg-slate-900 rounded-2xl shadow-xl border border-gray-200 dark:border-slate-800 flex flex-col h-[400px] md:h-full">
        <div className="p-5 border-b dark:border-slate-800 bg-slate-50 dark:bg-slate-800/50">
            <h2 className="font-black text-gray-800 dark:text-white flex items-center gap-2 uppercase text-xs tracking-widest">
                <History size={18} className="text-primary-500"/> Histórico Digital
            </h2>
            <div className="mt-4 space-y-3">
                <div className="grid grid-cols-2 gap-2">
                    <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} className="p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-900 text-[10px] font-bold dark:text-white outline-none" />
                    <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} className="p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-900 text-[10px] font-bold dark:text-white outline-none" />
                </div>
                <button onClick={fetchHistory} className="w-full py-2 bg-primary-600 text-white font-black uppercase text-[10px] rounded-lg hover:bg-primary-700 transition-all flex items-center justify-center gap-2">
                    {isLoading ? <Loader2 size={14} className="animate-spin" /> : <FileSearch size={14} />} BUSCAR AGORA
                </button>
            </div>
        </div>
        
        <div className="flex-1 overflow-y-auto p-3 space-y-2 scrollbar-thin">
          {isLoading ? (
              <div className="py-20 flex flex-col items-center gap-3 text-primary-500">
                  <Loader2 className="animate-spin" size={32}/>
                  <span className="text-[10px] font-black uppercase tracking-tighter">Sincronizando...</span>
              </div>
          ) : processedHistory.map(record => (
            <button
              key={record.id}
              onClick={() => { setSelectedRecord(record); setViewingPartial(false); }}
              className={`w-full text-left p-4 rounded-2xl transition-all border flex flex-col gap-2 relative group
                ${selectedRecord?.id === record.id 
                  ? 'bg-primary-50 dark:bg-primary-900/20 border-primary-200 dark:border-primary-800 shadow-lg scale-[1.02]' 
                  : 'bg-white dark:bg-slate-900 border-transparent hover:bg-gray-50 dark:hover:bg-slate-800 hover:border-slate-200 dark:hover:border-slate-700'
                }
              `}
            >
              <div className="flex items-center gap-3">
                  <div className={`p-2 rounded-xl ${selectedRecord?.id === record.id ? 'bg-primary-600 text-white' : 'bg-gray-100 dark:bg-slate-800 text-slate-400'}`}>
                    <Calendar size={16} />
                  </div>
                  <div className="flex-1">
                    <div className="font-black text-xs dark:text-white">{formatDateBR(record.timestamp)}</div>
                    <div className="text-[9px] text-slate-500 font-bold uppercase">{displayTimestamp(record.timestamp).split(', ')[1] || ''}</div>
                  </div>
                  {record.hasPartial && <div className="w-2 h-2 rounded-full bg-amber-500 animate-pulse"></div>}
              </div>
              <div className="text-[9px] font-black uppercase bg-slate-100 dark:bg-slate-800 px-2 py-1 rounded-lg dark:text-slate-400 truncate">
                  {record.resetBy}
              </div>
            </button>
          ))}
        </div>
      </div>

      {/* Main Content - Tabela em Dark Mode */}
      <div className="flex-1 bg-white dark:bg-slate-900 rounded-2xl shadow-2xl border border-gray-200 dark:border-slate-800 overflow-hidden flex flex-col">
        {!selectedRecord ? (
           <div className="flex-1 flex flex-col items-center justify-center text-slate-400">
              <History size={64} className="mb-4 opacity-5"/>
              <p className="font-black uppercase text-[10px] tracking-widest">Selecione para recuperar dados</p>
           </div>
        ) : (
          <>
            <div className="p-5 bg-slate-50 dark:bg-slate-950 border-b dark:border-slate-800 flex justify-between items-center">
               <div className="flex items-center gap-4">
                  <div className={`w-12 h-12 ${viewingPartial ? 'bg-amber-500' : 'bg-primary-600'} rounded-2xl flex items-center justify-center text-white shadow-xl`}>
                      {viewingPartial ? <FileSearch size={24} /> : <User size={24} />}
                  </div>
                  <div>
                    <h3 className="font-black text-gray-800 dark:text-white text-sm uppercase">
                        {viewingPartial ? "Salvamento Parcial (10:00h)" : `Resp: ${selectedRecord.resetBy}`}
                    </h3>
                    <div className="flex gap-4 mt-1">
                        <span className="text-[9px] font-black uppercase text-slate-500 flex items-center gap-1.5">
                            <Calendar size={10} /> {formatDateBR(viewingPartial && selectedRecord.partialRecord ? selectedRecord.partialRecord.timestamp : selectedRecord.timestamp)}
                        </span>
                    </div>
                  </div>
               </div>

               <div className="flex items-center gap-2">
                  {hasActiveColFilters && (
                    <button onClick={clearColFilters} className="px-3 py-2 text-[9px] font-bold text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-xl uppercase flex items-center gap-1">
                      <X size={12} /> Limpar Filtros
                    </button>
                  )}

                  {selectedRecord.hasPartial && (
                    <button
                      onClick={() => setViewingPartial(!viewingPartial)}
                      className={`px-4 py-2 rounded-xl text-[10px] font-black uppercase border transition-all flex items-center gap-2 ${viewingPartial
                          ? 'bg-primary-600 text-white' : 'bg-amber-50 dark:bg-amber-900/20 text-amber-600 border-amber-200 dark:border-amber-800'}`}
                    >
                        {viewingPartial ? "Snapshot Principal" : "Snap Parcial (10h)"}
                    </button>
                  )}
               </div>
            </div>

            <div className="flex-1 overflow-auto bg-slate-100 dark:bg-slate-950 scrollbar-thin">
                {/* Indicador de filtros ativos */}
                {hasActiveColFilters && (
                  <div className="bg-primary-50 dark:bg-primary-900/20 border-b border-primary-200 dark:border-primary-800 px-4 py-2 flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <Filter size={12} className="text-primary-600 dark:text-primary-400" />
                      <span className="text-[9px] font-bold text-primary-600 dark:text-primary-400 uppercase">
                        Filtros ativos:
                      </span>
                      {Object.entries(selectedFilters).filter(([_, values]) => values.length > 0).map(([col, values]) => (
                        <span key={col} className="text-[9px] font-bold text-slate-600 dark:text-slate-400 bg-white dark:bg-slate-800 px-2 py-0.5 rounded">
                          {col}: {values.join(', ')}
                        </span>
                      ))}
                    </div>
                    <button onClick={clearColFilters} className="text-[9px] font-bold text-red-500 hover:text-red-600 uppercase flex items-center gap-1">
                      <X size={12} /> Limpar
                    </button>
                  </div>
                )}

                {/* Filtros das Colunas */}
                <div className="bg-slate-50 dark:bg-slate-900 border-b dark:border-slate-800 px-4 py-3 flex items-center gap-4">
                  <div className="flex items-center gap-2 text-[10px] font-bold text-slate-500 uppercase">
                    <Filter size={12} />
                    <span>Filtros:</span>
                  </div>

                  {/* Filtro de Categoria */}
                  <div className="relative" ref={filterDropdownRef}>
                    <button
                      onClick={() => setActiveFilterCol(activeFilterCol === 'categoria' ? null : 'categoria')}
                      className={`px-3 py-1.5 rounded-lg text-[10px] font-bold uppercase flex items-center gap-1.5 transition-all ${
                        activeFilterCol === 'categoria' 
                          ? 'bg-primary-600 text-white' 
                          : 'bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-400 border border-slate-200 dark:border-slate-700 hover:border-primary-500'
                      }`}
                    >
                      Categoria
                      {selectedFilters['categoria']?.length > 0 && (
                        <span className="w-2 h-2 bg-primary-500 rounded-full"></span>
                      )}
                      <ChevronDown size={12} />
                    </button>

                    {activeFilterCol === 'categoria' && (
                      <div className="absolute top-full left-0 mt-1 w-48 bg-white dark:bg-slate-800 rounded-xl shadow-2xl border border-slate-200 dark:border-slate-700 overflow-hidden z-[100] animate-in fade-in slide-in-from-top-2">
                        <div className="p-2 border-b dark:border-slate-700 flex justify-between items-center">
                          <span className="text-[9px] font-black uppercase text-slate-600 dark:text-slate-400 px-2">Categoria</span>
                          {selectedFilters['categoria']?.length > 0 && (
                            <button 
                              onClick={() => setSelectedFilters(prev => ({ ...prev, categoria: [] }))}
                              className="text-[8px] font-bold text-primary-600 hover:text-primary-700 uppercase px-2"
                            >
                              Limpar
                            </button>
                          )}
                        </div>
                        <div className="max-h-48 overflow-y-auto p-2 space-y-1">
                          {getColUniqueValues('categoria').map(cat => (
                            <button
                              key={cat}
                              onClick={() => toggleColFilter('categoria', cat)}
                              className={`w-full text-left px-3 py-2 rounded-lg text-[10px] font-bold transition-all flex items-center justify-between ${
                                selectedFilters['categoria']?.includes(cat)
                                  ? 'bg-primary-100 dark:bg-primary-900/30 text-primary-600 dark:text-primary-400'
                                  : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-700'
                              }`}
                            >
                              {cat}
                              {selectedFilters['categoria']?.includes(cat) && <CheckCircle2 size={12} />}
                            </button>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Filtro de Operação */}
                  <div className="relative" ref={filterDropdownRef}>
                    <button
                      onClick={() => setActiveFilterCol(activeFilterCol === 'operacao' ? null : 'operacao')}
                      className={`px-3 py-1.5 rounded-lg text-[10px] font-bold uppercase flex items-center gap-1.5 transition-all ${
                        activeFilterCol === 'operacao' 
                          ? 'bg-primary-600 text-white' 
                          : 'bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-400 border border-slate-200 dark:border-slate-700 hover:border-primary-500'
                      }`}
                    >
                      Operação
                      {selectedFilters['operacao']?.length > 0 && (
                        <span className="w-2 h-2 bg-primary-500 rounded-full"></span>
                      )}
                      <ChevronDown size={12} />
                    </button>

                    {activeFilterCol === 'operacao' && (
                      <div className="absolute top-full left-0 mt-1 w-40 bg-white dark:bg-slate-800 rounded-xl shadow-2xl border border-slate-200 dark:border-slate-700 overflow-hidden z-[100] animate-in fade-in slide-in-from-top-2">
                        <div className="p-2 border-b dark:border-slate-700 flex justify-between items-center">
                          <span className="text-[9px] font-black uppercase text-slate-600 dark:text-slate-400 px-2">Operação</span>
                          {selectedFilters['operacao']?.length > 0 && (
                            <button 
                              onClick={() => setSelectedFilters(prev => ({ ...prev, operacao: [] }))}
                              className="text-[8px] font-bold text-primary-600 hover:text-primary-700 uppercase px-2"
                            >
                              Limpar
                            </button>
                          )}
                        </div>
                        <div className="max-h-48 overflow-y-auto p-2 space-y-1">
                          {getColUniqueValues('operacao').map(op => (
                            <button
                              key={op}
                              onClick={() => toggleColFilter('operacao', op)}
                              className={`w-full text-left px-3 py-2 rounded-lg text-[10px] font-bold transition-all flex items-center justify-between ${
                                selectedFilters['operacao']?.includes(op)
                                  ? 'bg-primary-100 dark:bg-primary-900/30 text-primary-600 dark:text-primary-400'
                                  : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-700'
                              }`}
                            >
                              {op.replace('LAT-', '').replace('ITA-', '')}
                              {selectedFilters['operacao']?.includes(op) && <CheckCircle2 size={12} />}
                            </button>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                </div>

                <table className="w-full border-collapse bg-white dark:bg-slate-900 text-[10px]">
                  <thead className="sticky top-0 z-20 bg-slate-800 dark:bg-slate-950 text-white shadow-lg">
                    <tr>
                      <th className="p-4 text-left w-[30%] min-w-[300px] border-r border-slate-700 dark:border-slate-800 sticky left-0 bg-slate-800 dark:bg-slate-950 font-black uppercase text-[9px]">
                        Ação / Tarefa
                      </th>
                      {filteredAllowedLocations.map(loc => (
                        <th key={loc} className="p-1 text-center min-w-[60px] border-r border-slate-700 dark:border-slate-800 font-black uppercase text-[9px]">
                            {loc.replace('LAT-', '').replace('ITA-', '')}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                    {(() => {
                        const grouped = getGroupedTasks(currentTasksToDisplay);
                        return Object.keys(grouped).map(category => {
                            const isCollapsed = collapsedCategories.includes(category);
                            return (
                                <React.Fragment key={category}>
                                    <tr className="bg-slate-600 dark:bg-slate-800 text-white cursor-pointer hover:bg-slate-700" onClick={() => setCollapsedCategories(prev => prev.includes(category) ? prev.filter(c => c !== category) : [...prev, category])}>
                                        <td colSpan={1 + filteredAllowedLocations.length} className="px-4 py-2 font-black uppercase text-[10px] sticky left-0">
                                            <div className="flex items-center gap-2">{isCollapsed ? <ChevronRight size={14}/> : <ChevronDown size={14}/>} {category}</div>
                                        </td>
                                    </tr>
                                    {!isCollapsed && grouped[category].map(task => (
                                        <tr key={task.id} className="hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors">
                                            <td className="p-4 border-r dark:border-slate-800 sticky left-0 bg-white dark:bg-slate-900 z-10 shadow-sm">
                                                <div className="font-bold dark:text-white text-[11px]">{task.title}</div>
                                                <div className="text-slate-400 text-[9px] mt-1 italic">{task.description}</div>
                                            </td>
                                            {filteredAllowedLocations.map(loc => {
                                                const statusKey = task.operations[loc] || 'PR';
                                                const config = STATUS_CONFIG[statusKey] || STATUS_CONFIG['PR'];
                                                return (
                                                    <td key={`${task.id}-${loc}`} className="p-0 border-r dark:border-slate-800 relative h-12">
                                                        <div className={`absolute inset-[3px] rounded-lg flex items-center justify-center text-[9px] font-black border transition-all ${config.color} uppercase`}>
                                                            {config.label}
                                                        </div>
                                                    </td>
                                                );
                                            })}
                                        </tr>
                                    ))}
                                </React.Fragment>
                            );
                        });
                    })()}
                  </tbody>
                </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

export default HistoryViewer;
