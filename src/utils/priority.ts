import { PriorityInfo, PriorityLevel } from '../types';

/**
 * Calculates priority based on target Entry Date and "Today" (Fixed at 2026-06-26 for demo consistency).
 * Formula:
 * - Days remaining <= 1: High Priority (🔴 Tinggi)
 * - Days remaining 2 to 4: Medium Priority (🟠 Sedang)
 * - Days remaining > 4: Low Priority (🟢 Rendah)
 */
/**
 * Calculates priority based on target Entry Date and a Reference Date (defaults to system current date).
 * Formula:
 * - Days remaining <= 1: High Priority (🔴 Tinggi)
 * - Days remaining 2 to 4: Medium Priority (🟠 Sedang)
 * - Days remaining > 4: Low Priority (🟢 Rendah)
 */
export function getPriorityInfo(entryMadinahDateStr: string, referenceDateStr?: string): PriorityInfo {
  if (!entryMadinahDateStr) {
    return {
      level: 'Belum Ada',
      daysRemaining: NaN,
      badgeColor: 'bg-slate-100 text-slate-500 border-slate-200 dark:bg-zinc-800 dark:text-zinc-400 dark:border-zinc-750',
      dotColor: 'bg-slate-400',
    };
  }

  // Use referenceDateStr if provided, otherwise default to current local system time
  const today = referenceDateStr ? new Date(referenceDateStr) : new Date();
  today.setHours(0, 0, 0, 0);

  const entryDate = new Date(entryMadinahDateStr);
  entryDate.setHours(0, 0, 0, 0);

  const diffTime = entryDate.getTime() - today.getTime();
  const daysRemaining = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  if (isNaN(daysRemaining)) {
    return {
      level: 'Belum Ada',
      daysRemaining: NaN,
      badgeColor: 'bg-slate-100 text-slate-500 border-slate-200 dark:bg-zinc-800 dark:text-zinc-400 dark:border-zinc-750',
      dotColor: 'bg-slate-400',
    };
  }

  if (daysRemaining <= 1) {
    return {
      level: 'Tinggi',
      daysRemaining,
      badgeColor: 'bg-red-50 text-red-700 border-red-200/60 dark:bg-red-500/10 dark:text-red-400 dark:border-red-500/20',
      dotColor: 'bg-red-500',
    };
  } else if (daysRemaining >= 2 && daysRemaining <= 4) {
    return {
      level: 'Sedang',
      daysRemaining,
      badgeColor: 'bg-amber-50 text-amber-700 border-amber-200/60 dark:bg-amber-500/10 dark:text-amber-400 dark:border-amber-500/20',
      dotColor: 'bg-amber-500',
    };
  } else {
    return {
      level: 'Rendah',
      daysRemaining,
      badgeColor: 'bg-blue-50 text-blue-700 border-blue-200/60 dark:bg-blue-500/10 dark:text-blue-400 dark:border-blue-500/20',
      dotColor: 'bg-blue-500',
    };
  }
}

/**
 * Custom sorter: sorts 'Tinggi' first, then 'Sedang', then 'Rendah', and 'Belum Ada' at the end.
 * In case of tie, sorts by closest entryMadinah date.
 */
export function sortJamaahByPriorityAndDate(
  a: { entryMadinah: string },
  b: { entryMadinah: string },
  referenceDateStr?: string
): number {
  const pA = getPriorityInfo(a.entryMadinah, referenceDateStr);
  const pB = getPriorityInfo(b.entryMadinah, referenceDateStr);
  
  const priorityWeight = { 'Tinggi': 4, 'Sedang': 3, 'Rendah': 2, 'Belum Ada': 1 };
  const diff = priorityWeight[pB.level] - priorityWeight[pA.level];
  if (diff !== 0) return diff;
  
  // Tie-breaker: earlier entry date comes first (if both are invalid, they are equal)
  const timeA = new Date(a.entryMadinah).getTime();
  const timeB = new Date(b.entryMadinah).getTime();
  const validA = !isNaN(timeA);
  const validB = !isNaN(timeB);
  if (!validA && !validB) return 0;
  if (!validA) return 1; // puts empty dates at the end
  if (!validB) return -1;
  return timeA - timeB;
}

