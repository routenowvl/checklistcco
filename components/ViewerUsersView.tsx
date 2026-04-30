import React, { useEffect, useMemo, useState } from 'react';
import { Check, Loader2, Save, Trash2, UserPlus } from 'lucide-react';
import { RouteConfig, User, ViewerAccessEntry } from '../types';
import { SharePointService } from '../services/sharepointService';
import { getValidToken } from '../services/tokenService';

type ViewerAccessRow = {
  email: string;
  operacoes: string[];
};

const normalizeEmail = (value: string): string => String(value || '').trim().toLowerCase();
const isValidEmail = (value: string): boolean => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizeEmail(value));

const ViewerUsersView: React.FC<{ currentUser: User }> = ({ currentUser }) => {
  const [configs, setConfigs] = useState<RouteConfig[]>([]);
  const [viewerEntries, setViewerEntries] = useState<ViewerAccessEntry[]>([]);
  const [email, setEmail] = useState('');
  const [selectedOperations, setSelectedOperations] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');

  const availableOperations = useMemo(() => {
    return Array.from(
      new Set(
        configs
          .map((cfg) => String(cfg.operacao || '').trim())
          .filter(Boolean)
      )
    ).sort((a, b) => a.localeCompare(b, 'pt-BR'));
  }, [configs]);

  const viewers = useMemo<ViewerAccessRow[]>(() => {
    const availableSet = new Set(availableOperations.map((op) => op.toUpperCase()));
    const map = new Map<string, Set<string>>();

    viewerEntries.forEach((entry) => {
      const viewerEmail = normalizeEmail(entry.email);
      const operacao = String(entry.operacao || '').trim();
      if (!viewerEmail || !operacao) return;
      if (!availableSet.has(operacao.toUpperCase())) return;
      if (!map.has(viewerEmail)) {
        map.set(viewerEmail, new Set<string>());
      }
      map.get(viewerEmail)?.add(operacao);
    });

    return Array.from(map.entries())
      .map(([viewerEmail, ops]) => ({
        email: viewerEmail,
        operacoes: Array.from(ops).sort((a, b) => a.localeCompare(b, 'pt-BR'))
      }))
      .sort((a, b) => a.email.localeCompare(b.email, 'pt-BR'));
  }, [viewerEntries, availableOperations]);

  const loadData = async (forceRefresh = false): Promise<void> => {
    const token = (await getValidToken()) || currentUser.accessToken;
    if (!token) {
      setError('Token de acesso não encontrado.');
      setIsLoading(false);
      return;
    }

    if (forceRefresh) {
      setIsSaving(true);
    } else {
      setIsLoading(true);
    }
    setError('');

    try {
      const [editableConfigs, viewerAccess] = await Promise.all([
        SharePointService.getRouteConfigs(token, currentUser.email, forceRefresh),
        SharePointService.getViewerAccessEntries(token, forceRefresh)
      ]);
      setConfigs(editableConfigs || []);
      setViewerEntries(viewerAccess || []);
    } catch (err: any) {
      setError(err?.message || 'Falha ao carregar operações disponíveis.');
    } finally {
      setIsLoading(false);
      setIsSaving(false);
    }
  };

  useEffect(() => {
    void loadData(false);
  }, [currentUser.email]);

  const toggleOperation = (operation: string): void => {
    setSelectedOperations((prev) =>
      prev.includes(operation)
        ? prev.filter((item) => item !== operation)
        : [...prev, operation]
    );
  };

  const fillFormFromViewer = (row: ViewerAccessRow): void => {
    setEmail(row.email);
    setSelectedOperations(row.operacoes);
    setSuccess('');
    setError('');
  };

  const resetForm = (): void => {
    setEmail('');
    setSelectedOperations([]);
  };

  const saveViewerAccess = async (): Promise<void> => {
    const normalizedEmail = normalizeEmail(email);
    if (!isValidEmail(normalizedEmail)) {
      setError('Informe um e-mail válido para cadastro.');
      setSuccess('');
      return;
    }

    if (selectedOperations.length === 0) {
      setError('Selecione ao menos uma filial para visualização.');
      setSuccess('');
      return;
    }

    const allowedOps = new Set(availableOperations);
    const validSelected = selectedOperations
      .map((op) => String(op || '').trim())
      .filter((op) => allowedOps.has(op));

    if (validSelected.length === 0) {
      setError('Selecione filiais válidas da operação do seu login.');
      setSuccess('');
      return;
    }

    const token = (await getValidToken()) || currentUser.accessToken;
    if (!token) {
      setError('Token de acesso não encontrado.');
      setSuccess('');
      return;
    }

    setIsSaving(true);
    setError('');
    setSuccess('');

    try {
      await SharePointService.replaceViewerAccessForEmail(
        token,
        normalizedEmail,
        validSelected,
        availableOperations
      );
      await loadData(true);
      setSuccess('Acessos atualizados com sucesso.');
    } catch (err: any) {
      setError(err?.message || 'Falha ao salvar acessos de visualização.');
    } finally {
      setIsSaving(false);
    }
  };

  const removeViewerFromAllOperations = async (viewerEmail: string): Promise<void> => {
    const normalizedEmail = normalizeEmail(viewerEmail);
    if (!normalizedEmail) return;

    const token = (await getValidToken()) || currentUser.accessToken;
    if (!token) {
      setError('Token de acesso não encontrado.');
      return;
    }

    setIsSaving(true);
    setError('');
    setSuccess('');

    try {
      await SharePointService.replaceViewerAccessForEmail(
        token,
        normalizedEmail,
        [],
        availableOperations
      );
      await loadData(true);
      if (normalizeEmail(email) === normalizedEmail) {
        resetForm();
      }
      setSuccess('Usuário removido das filiais selecionadas.');
    } catch (err: any) {
      setError(err?.message || 'Falha ao remover usuário.');
    } finally {
      setIsSaving(false);
    }
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
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-blue-100 dark:bg-blue-900/30 text-blue-700 dark:text-blue-300 flex items-center justify-center">
              <UserPlus size={18} />
            </div>
            <div>
              <h1 className="text-xl font-black uppercase text-slate-800 dark:text-white tracking-tight">Cadastro de Usuários</h1>
              <p className="text-[11px] font-bold uppercase text-slate-500 dark:text-slate-400">
                Visualização somente leitura em Saídas e Não Coletas
              </p>
            </div>
          </div>

          {error && (
            <div className="mt-4 p-3 rounded-xl bg-red-50 border border-red-200 text-red-700 text-[11px] font-bold">
              {error}
            </div>
          )}

          {success && (
            <div className="mt-4 p-3 rounded-xl bg-emerald-50 border border-emerald-200 text-emerald-700 text-[11px] font-bold">
              {success}
            </div>
          )}

          <div className="mt-4 grid grid-cols-1 lg:grid-cols-3 gap-4">
            <div className="lg:col-span-1 space-y-3">
              <label className="text-[10px] font-black uppercase text-slate-400">E-mail do usuário</label>
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="usuario@empresa.com.br"
                className="w-full px-3 py-2 rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 text-[11px] font-bold text-slate-800 dark:text-white outline-none"
              />

              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={saveViewerAccess}
                  disabled={isSaving}
                  className="flex-1 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-[11px] font-black uppercase tracking-wide disabled:opacity-60 flex items-center justify-center gap-2"
                >
                  {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
                  Salvar
                </button>
                <button
                  type="button"
                  onClick={resetForm}
                  disabled={isSaving}
                  className="px-3 py-2 rounded-xl border border-slate-200 dark:border-slate-700 text-[11px] font-black uppercase text-slate-600 dark:text-slate-300 disabled:opacity-60"
                >
                  Limpar
                </button>
              </div>
            </div>

            <div className="lg:col-span-2">
              <label className="text-[10px] font-black uppercase text-slate-400">Filiais permitidas</label>
              <div className="mt-2 grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-2">
                {availableOperations.map((op) => {
                  const isActive = selectedOperations.includes(op);
                  return (
                    <button
                      key={op}
                      type="button"
                      onClick={() => toggleOperation(op)}
                      className={`px-3 py-2 rounded-xl border text-left text-[11px] font-black uppercase tracking-wide transition-colors ${
                        isActive
                          ? 'bg-blue-600 border-blue-600 text-white'
                          : 'bg-white dark:bg-slate-900 border-slate-200 dark:border-slate-700 text-slate-700 dark:text-slate-300 hover:border-blue-400'
                      }`}
                    >
                      <span className="inline-flex items-center gap-2">
                        {isActive ? <Check size={13} /> : <span className="w-[13px]" />}
                        {op}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
        </div>

        <div className="bg-white/95 dark:bg-slate-900/95 backdrop-blur-sm border border-white/50 dark:border-slate-800 rounded-[2rem] overflow-hidden shadow-2xl">
          <div className="px-5 py-4 border-b border-slate-200/70 dark:border-slate-800/70">
            <h2 className="text-sm font-black uppercase tracking-wide text-slate-700 dark:text-slate-200">
              Usuários cadastrados para visualização ({viewers.length})
            </h2>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full min-w-[840px] border-collapse">
              <thead>
                <tr className="bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900">
                  <th className="px-4 py-3 text-left text-[10px] font-black uppercase tracking-widest text-slate-300 border-r border-slate-700/50">E-mail</th>
                  <th className="px-4 py-3 text-left text-[10px] font-black uppercase tracking-widest text-slate-300 border-r border-slate-700/50">Filiais</th>
                  <th className="px-4 py-3 text-right text-[10px] font-black uppercase tracking-widest text-slate-300">Ações</th>
                </tr>
              </thead>
              <tbody>
                {viewers.length === 0 ? (
                  <tr>
                    <td colSpan={3} className="px-4 py-8 text-center text-[11px] font-bold text-slate-500 dark:text-slate-400">
                      Nenhum usuário de visualização cadastrado.
                    </td>
                  </tr>
                ) : (
                  viewers.map((row, idx) => (
                    <tr
                      key={row.email}
                      className={`border-t border-slate-200/50 dark:border-slate-800/50 ${idx % 2 === 0 ? '' : 'bg-black/[0.02] dark:bg-white/[0.02]'}`}
                    >
                      <td className="px-4 py-3 text-[11px] font-black text-slate-800 dark:text-slate-100">{row.email}</td>
                      <td className="px-4 py-3">
                        <div className="flex flex-wrap gap-2">
                          {row.operacoes.map((op) => (
                            <span
                              key={`${row.email}-${op}`}
                              className="inline-flex px-2 py-1 rounded-full border border-blue-200 dark:border-blue-800 bg-blue-100 dark:bg-blue-900/30 text-[10px] font-black uppercase text-blue-700 dark:text-blue-300"
                            >
                              {op}
                            </span>
                          ))}
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        <div className="flex justify-end gap-2">
                          <button
                            type="button"
                            onClick={() => fillFormFromViewer(row)}
                            disabled={isSaving}
                            className="px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 text-[10px] font-black uppercase text-slate-600 dark:text-slate-300 hover:border-blue-400 disabled:opacity-60"
                          >
                            Editar
                          </button>
                          <button
                            type="button"
                            onClick={() => removeViewerFromAllOperations(row.email)}
                            disabled={isSaving}
                            className="px-3 py-1.5 rounded-lg border border-red-200 dark:border-red-800 text-[10px] font-black uppercase text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20 disabled:opacity-60 inline-flex items-center gap-1.5"
                          >
                            <Trash2 size={12} />
                            Remover
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ViewerUsersView;
