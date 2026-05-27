import { useRef, useCallback } from 'react';
import { SharePointService } from '../services/sharepointService';

type SyncStatus = 'idle' | 'syncing' | 'error';

type PendingEntry = {
  route: any;
  timer: ReturnType<typeof setTimeout> | null;
  status: SyncStatus;
  retryCount: number;
};

const DEBOUNCE_MS = 500;
const MAX_RETRIES = 2;

export const useWriteBehind = (getToken: () => Promise<string> | string) => {
  const pendingRef = useRef<Map<number, PendingEntry>>(new Map());

  const flushRoute = useCallback(async (routeId: number) => {
    const entry = pendingRef.current.get(routeId);
    if (!entry) return;
    if (entry.timer) {
      clearTimeout(entry.timer);
      entry.timer = null;
    }

    const route = entry.route;
    entry.status = 'syncing';

    try {
      const token = await getToken();
      await SharePointService.updateDeparture(token, route);
      entry.status = 'idle';
      pendingRef.current.delete(routeId);
      console.log(`[WriteBehind] sync OK for route ${routeId}`, { saida: route.saida, id: route.id });
    } catch (err: any) {
      console.error(`[WriteBehind] sync FAILED for route ${routeId}:`, err?.message || err, err);
      if (entry.retryCount < MAX_RETRIES) {
        entry.retryCount++;
        entry.status = 'idle';
        // Retry after a short delay
        entry.timer = setTimeout(() => flushRoute(routeId), 2000);
      } else {
        entry.status = 'error';
      }
    }
  }, [getToken]);

  const enqueue = useCallback((routeId: number, route: any) => {
    const existing = pendingRef.current.get(routeId);
    if (existing) {
      // Update with latest snapshot, reset debounce timer
      existing.route = route;
      if (existing.timer) clearTimeout(existing.timer);
      existing.timer = setTimeout(() => flushRoute(routeId), DEBOUNCE_MS);
    } else {
      const entry: PendingEntry = {
        route,
        timer: setTimeout(() => flushRoute(routeId), DEBOUNCE_MS),
        status: 'idle',
        retryCount: 0,
      };
      pendingRef.current.set(routeId, entry);
    }
  }, [flushRoute]);

  const flushAll = useCallback(async () => {
    const promises: Promise<void>[] = [];
    for (const [routeId] of pendingRef.current) {
      promises.push(flushRoute(routeId));
    }
    await Promise.allSettled(promises);
  }, [flushRoute]);

  const getStatus = useCallback((routeId: number): SyncStatus => {
    return pendingRef.current.get(routeId)?.status || 'idle';
  }, []);

  const cancelAll = useCallback(() => {
    for (const [, entry] of pendingRef.current) {
      if (entry.timer) clearTimeout(entry.timer);
    }
    pendingRef.current.clear();
  }, []);

  return { enqueue, flushRoute, flushAll, getStatus, cancelAll };
};
