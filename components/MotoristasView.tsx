import React, { useEffect, useMemo, useRef, useState } from 'react';
import { ChevronLeft, ChevronRight, Loader2, Search, ArrowUpDown } from 'lucide-react';
import { Motorista, User } from '../types';
import { SharePointService } from '../services/sharepointService';
import { getValidToken } from '../services/tokenService';

const PAGE_SIZE = 100;
const TWO_HOURS_MS = 2 * 60 * 60 * 1000;
const BACKGROUND_SYNC_DELAY_MS = 5000;

const formatContactMask = (value: string): string => {
  const digits = String(value || '').replace(/\D/g, '').slice(0, 11);
  if (!digits) return '';
  if (digits.length <= 2) return `(${digits}`;
  if (digits.length <= 6) return `(${digits.slice(0, 2)}) ${digits.slice(2)}`;
  if (digits.length <= 10) return `(${digits.slice(0, 2)}) ${digits.slice(2, 6)}-${digits.slice(6)}`;
  return `(${digits.slice(0, 2)}) ${digits.slice(2, 7)}-${digits.slice(7)}`;
};

const sanitizeContact = (value: string): string => String(value || '').replace(/\D/g, '').slice(0, 11);
const normalizeText = (value: string): string =>
  String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();

type CadastroFilterMode = 'complete' | 'incomplete';

const hasOperationData = (value: string): boolean => String(value || '').trim().length > 0;
const hasContactData = (value: string): boolean => {
  const digits = sanitizeContact(value);
  return digits.length >= 10;
};

const MotoristasView: React.FC<{ currentUser: User }> = ({ currentUser }) => {
  const [motoristas, setMotoristas] = useState<Motorista[]>([]);
  const [userOperations, setUserOperations] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [page, setPage] = useState(1);
  const [error, setError] = useState('');
  const [editingOperationId, setEditingOperationId] = useState<string | null>(null);
  const [editingContactId, setEditingContactId] = useState<string | null>(null);
  const [contactDraft, setContactDraft] = useState('');
  const [savingRowId, setSavingRowId] = useState<string | null>(null);
  const [lastUpdatedAt, setLastUpdatedAt] = useState<Date | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [sortMode, setSortMode] = useState<'default' | 'az' | 'za'>('default');
  const [cadastroFilter, setCadastroFilter] = useState<CadastroFilterMode>('incomplete');
  const backgroundSyncTimeoutRef = useRef<number | null>(null);

  const filteredAndSortedRows = useMemo(() => {
    const query = normalizeText(searchTerm);
    let rows = motoristas;

    if (query) {
      rows = rows.filter((row) => {
        const contactMasked = formatContactMask(row.contato);
        return [
          row.codigo,
          row.motorista,
          row.operacao,
          row.contato,
          contactMasked
        ].some((field) => normalizeText(field).includes(query));
      });
    }

    rows = rows.filter((row) => {
      const operationOk = hasOperationData(row.operacao);
      const contactOk = hasContactData(row.contato);
      if (cadastroFilter === 'complete') return operationOk && contactOk;
      return !operationOk || !contactOk;
    });

    if (sortMode === 'az') {
      return [...rows].sort((a, b) =>
        String(a.motorista || '').localeCompare(String(b.motorista || ''), 'pt-BR', { sensitivity: 'base' })
      );
    }

    if (sortMode === 'za') {
      return [...rows].sort((a, b) =>
        String(b.motorista || '').localeCompare(String(a.motorista || ''), 'pt-BR', { sensitivity: 'base' })
      );
    }

    return rows;
  }, [motoristas, searchTerm, sortMode, cadastroFilter]);

  const totalPages = Math.max(1, Math.ceil(filteredAndSortedRows.length / PAGE_SIZE));

  const paginatedRows = useMemo(() => {
    const start = (page - 1) * PAGE_SIZE;
    return filteredAndSortedRows.slice(start, start + PAGE_SIZE);
  }, [filteredAndSortedRows, page]);

  const loadData = async (forceRefresh = false, showMainLoader = false): Promise<void> => {
    const token = (await getValidToken()) || currentUser.accessToken;
    if (!token) {
      setError('Token de acesso não encontrado.');
      return;
    }

    if (showMainLoader) setIsLoading(true);
    if (!showMainLoader) setIsRefreshing(true);
    setError('');

    try {
      const [rows, ops] = await Promise.all([
        SharePointService.getMotoristas(token, forceRefresh),
        SharePointService.getOperations(token, currentUser.email)
      ]);

      setMotoristas(rows);
      setUserOperations(ops.map(op => op.Title).filter(Boolean));
      setLastUpdatedAt(new Date());
      setPage(prev => Math.min(prev, Math.max(1, Math.ceil(rows.length / PAGE_SIZE))));
    } catch (err: any) {
      setError(err?.message || 'Falha ao carregar motoristas.');
    } finally {
      setIsLoading(false);
      setIsRefreshing(false);
    }
  };

  useEffect(() => {
    loadData(false, true);
  }, []);

  useEffect(() => {
    const interval = setInterval(() => {
      loadData(true, false);
    }, TWO_HOURS_MS);
    return () => clearInterval(interval);
  }, [currentUser.email]);

  useEffect(() => {
    setPage(1);
  }, [searchTerm, sortMode, cadastroFilter]);

  const scheduleBackgroundSync = (): void => {
    if (backgroundSyncTimeoutRef.current !== null) {
      window.clearTimeout(backgroundSyncTimeoutRef.current);
      backgroundSyncTimeoutRef.current = null;
    }

    backgroundSyncTimeoutRef.current = window.setTimeout(() => {
      backgroundSyncTimeoutRef.current = null;
      void loadData(true, false);
    }, BACKGROUND_SYNC_DELAY_MS);
  };

  useEffect(() => {
    return () => {
      if (backgroundSyncTimeoutRef.current !== null) {
        window.clearTimeout(backgroundSyncTimeoutRef.current);
      }
    };
  }, []);

  const saveOperation = async (row: Motorista, nextOperation: string): Promise<void> => {
    const token = (await getValidToken()) || currentUser.accessToken;
    if (!token) return;

    const normalizedOperation = String(nextOperation || '').trim();
    const previousOperation = row.operacao;

    if (normalizedOperation === previousOperation) {
      setEditingOperationId(null);
      return;
    }

    setError('');
    setEditingOperationId(null);
    setSavingRowId(row.id);
    setMotoristas((prev) =>
      prev.map((item) =>
        item.id === row.id ? { ...item, operacao: normalizedOperation } : item
      )
    );

    try {
      await SharePointService.updateMotorista(token, row.id, { operacao: normalizedOperation });
      setLastUpdatedAt(new Date());
      scheduleBackgroundSync();
    } catch (err: any) {
      setMotoristas((prev) =>
        prev.map((item) =>
          item.id === row.id ? { ...item, operacao: previousOperation } : item
        )
      );
      setError(err?.message || 'Falha ao atualizar operação do motorista.');
    } finally {
      setSavingRowId(null);
    }
  };

  const saveContact = async (row: Motorista): Promise<void> => {
    const token = (await getValidToken()) || currentUser.accessToken;
    if (!token) return;

    const cleaned = sanitizeContact(contactDraft);
    const previousContact = row.contato;

    setEditingContactId(null);
    setContactDraft('');

    if (cleaned === previousContact) {
      return;
    }

    setError('');
    setSavingRowId(row.id);
    setMotoristas((prev) =>
      prev.map((item) =>
        item.id === row.id ? { ...item, contato: cleaned } : item
      )
    );

    try {
      await SharePointService.updateMotorista(token, row.id, { contato: cleaned });
      setLastUpdatedAt(new Date());
      scheduleBackgroundSync();
    } catch (err: any) {
      setMotoristas((prev) =>
        prev.map((item) =>
          item.id === row.id ? { ...item, contato: previousContact } : item
        )
      );
      setError(err?.message || 'Falha ao atualizar contato do motorista.');
    } finally {
      setSavingRowId(null);
    }
  };

  const toggleSortMode = (): void => {
    setSortMode((prev) => {
      if (prev === 'default') return 'az';
      if (prev === 'az') return 'za';
      return 'default';
    });
  };

  if (isLoading) {
    return (
      <div className="h-full w-full flex items-center justify-center bg-slate-50 dark:bg-slate-950">
        <Loader2 className="animate-spin text-blue-600" size={30} />
      </div>
    );
  }

  return (
    <div className="h-full overflow-auto bg-slate-50 dark:bg-slate-950 p-6">
      <div className="max-w-[1700px] mx-auto space-y-4">
        <div className="bg-white/95 dark:bg-slate-900/95 backdrop-blur-sm border border-white/50 dark:border-slate-800 rounded-[2rem] p-5 shadow-2xl">
          <div className="flex items-center justify-between gap-3 flex-wrap">
            <div>
              <h1 className="text-xl font-black uppercase text-slate-800 dark:text-white tracking-tight">Motoristas</h1>
              <p className="text-[11px] font-bold uppercase text-slate-500 dark:text-slate-400">
                Total: {filteredAndSortedRows.length} registros
                {isRefreshing
                  ? ' • Atualizando dados...'
                  : (lastUpdatedAt ? ` • Atualizado às ${lastUpdatedAt.toLocaleTimeString('pt-BR')}` : '')
                }
              </p>
            </div>
          </div>

          {error && (
            <div className="mt-3 p-3 rounded-xl bg-red-50 border border-red-200 text-red-700 text-[11px] font-bold">
              {error}
            </div>
          )}

          <div className="mt-4 grid grid-cols-1 md:grid-cols-2 gap-3">
            <div>
              <label className="text-[10px] font-black uppercase text-slate-400">Pesquisa</label>
              <div className="mt-1 flex items-center gap-2 px-3 py-2 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 shadow-sm">
                <Search size={14} className="text-slate-400" />
                <input
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  placeholder="Buscar por código, motorista, operação ou contato"
                  className="w-full bg-transparent outline-none text-[11px] font-bold text-slate-800 dark:text-white placeholder:text-slate-400"
                />
              </div>
            </div>
            <div>
              <label className="text-[10px] font-black uppercase text-slate-400">Filtro de cadastro</label>
              <select
                value={cadastroFilter}
                onChange={(e) => setCadastroFilter(e.target.value as CadastroFilterMode)}
                className="mt-1 w-full px-3 py-2 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 text-[11px] font-bold text-slate-800 dark:text-white outline-none"
              >
                <option value="complete">Cadastro completo</option>
                <option value="incomplete">Cadastro incompleto (operação ou contato)</option>
              </select>
            </div>
          </div>
        </div>

        <div className="bg-white/95 dark:bg-slate-900/95 backdrop-blur-sm border border-white/50 dark:border-slate-800 rounded-[2rem] overflow-hidden shadow-2xl flex flex-col">
          <div className="overflow-x-auto flex-1 scrollbar-thin">
            <table className="w-full min-w-[900px] border-collapse">
              <thead className="sticky top-0 z-20">
                <tr className="bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900">
                  <th className="px-4 py-4 text-left border-r border-slate-700/50 text-[10px] font-black text-slate-300 uppercase tracking-widest">Código</th>
                  <th className="px-4 py-4 text-left border-r border-slate-700/50 text-[10px] font-black text-slate-300 uppercase tracking-widest">
                    <button
                      type="button"
                      onClick={toggleSortMode}
                      className="inline-flex items-center gap-2 text-slate-300 hover:text-white transition-colors"
                    >
                      <span>Motorista</span>
                      <ArrowUpDown size={12} />
                      <span className="text-[9px] font-black tracking-normal">
                        {sortMode === 'default' && 'A/Z'}
                        {sortMode === 'az' && 'A->Z'}
                        {sortMode === 'za' && 'Z->A'}
                      </span>
                    </button>
                  </th>
                  <th className="px-4 py-4 text-left border-r border-slate-700/50 text-[10px] font-black text-slate-300 uppercase tracking-widest">Operação</th>
                  <th className="px-4 py-4 text-left text-[10px] font-black text-slate-300 uppercase tracking-widest">Contato</th>
                </tr>
              </thead>
              <tbody>
                {paginatedRows.length === 0 ? (
                  <tr>
                    <td colSpan={4} className="p-0 border border-slate-200/30 dark:border-slate-800/30">
                      <div className="p-6 text-center text-[11px] font-bold text-slate-500 dark:text-slate-400">
                        Nenhum registro encontrado para o filtro aplicado.
                      </div>
                    </td>
                  </tr>
                ) : (
                  paginatedRows.map((row, index) => (
                    <tr
                      key={row.id}
                      className={`border-b border-slate-200/50 dark:border-slate-800/50 transition-colors ${
                        index % 2 === 0 ? '' : 'bg-black/[0.02] dark:bg-white/[0.02]'
                      }`}
                    >
                      <td className="p-0 border border-slate-200/30 dark:border-slate-800/30">
                        <div className="px-3 py-2 text-[11px] font-bold text-slate-700 dark:text-slate-300">{row.codigo}</div>
                      </td>
                      <td className="p-0 border border-slate-200/30 dark:border-slate-800/30">
                        <div className="px-3 py-2 text-[11px] font-bold text-slate-800 dark:text-white">{row.motorista || '-'}</div>
                      </td>
                      <td className="p-0 border border-slate-200/30 dark:border-slate-800/30 relative">
                        {editingOperationId === row.id ? (
                          <select
                            autoFocus
                            value={row.operacao || ''}
                            onChange={(e) => saveOperation(row, e.target.value)}
                            onBlur={() => setEditingOperationId(null)}
                            disabled={savingRowId === row.id}
                            className="w-full bg-slate-900 text-slate-100 border-none px-3 py-2 text-[11px] font-bold outline-none"
                          >
                            <option value="">Selecione</option>
                            {userOperations.map((op) => (
                              <option key={op} value={op}>{op}</option>
                            ))}
                          </select>
                        ) : (
                          <button
                            onClick={() => setEditingOperationId(row.id)}
                            className="w-full text-left px-3 py-2 text-[11px] font-bold text-slate-700 dark:text-slate-300 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition-colors"
                          >
                            {savingRowId === row.id ? 'Salvando...' : (row.operacao || 'Clique para definir')}
                          </button>
                        )}
                      </td>
                      <td className="p-0 border border-slate-200/30 dark:border-slate-800/30">
                        {editingContactId === row.id ? (
                          <input
                            autoFocus
                            value={contactDraft}
                            onChange={(e) => setContactDraft(formatContactMask(e.target.value))}
                            onBlur={() => saveContact(row)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') saveContact(row);
                              if (e.key === 'Escape') {
                                setEditingContactId(null);
                                setContactDraft('');
                              }
                            }}
                            disabled={savingRowId === row.id}
                            className="w-full bg-slate-900 text-slate-100 border-none px-3 py-2 text-[11px] font-bold outline-none"
                            placeholder="(00) 00000-0000"
                          />
                        ) : (
                          <button
                            onClick={() => {
                              setEditingContactId(row.id);
                              setContactDraft(formatContactMask(row.contato));
                            }}
                            className="w-full text-left px-3 py-2 text-[11px] font-bold text-slate-700 dark:text-slate-300 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition-colors"
                          >
                            {savingRowId === row.id ? 'Salvando...' : (formatContactMask(row.contato) || 'Clique para informar')}
                          </button>
                        )}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>

          <div className="p-4 bg-slate-50 dark:bg-slate-900/70 border-t border-slate-200 dark:border-slate-800 flex items-center justify-between gap-3 flex-wrap">
            <p className="text-[10px] font-black uppercase text-slate-500 dark:text-slate-400">
              Página {page} de {totalPages}
            </p>
            <div className="flex items-center gap-2">
              <button
                onClick={() => setPage(p => Math.max(1, p - 1))}
                disabled={page === 1}
                className="px-3 py-2 rounded-lg border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 disabled:opacity-50"
              >
                <ChevronLeft size={14} />
              </button>
              <button
                onClick={() => setPage(p => Math.min(totalPages, p + 1))}
                disabled={page === totalPages}
                className="px-3 py-2 rounded-lg border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 disabled:opacity-50"
              >
                <ChevronRight size={14} />
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default MotoristasView;
