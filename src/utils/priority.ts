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
  // Use referenceDateStr if provided, otherwise default to current local system time
  const today = referenceDateStr ? new Date(referenceDateStr) : new Date();
  today.setHours(0, 0, 0, 0);

  const entryDate = new Date(entryMadinahDateStr);
  entryDate.setHours(0, 0, 0, 0);

  const diffTime = entryDate.getTime() - today.getTime();
  const daysRemaining = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

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
 * Custom sorter: sorts 'Tinggi' first, then 'Sedang', then 'Rendah'.
 * In case of tie, sorts by closest entryMadinah date.
 */
export function sortJamaahByPriorityAndDate(
  a: { entryMadinah: string },
  b: { entryMadinah: string },
  referenceDateStr?: string
): number {
  const pA = getPriorityInfo(a.entryMadinah, referenceDateStr);
  const pB = getPriorityInfo(b.entryMadinah, referenceDateStr);
  
  const priorityWeight = { 'Tinggi': 3, 'Sedang': 2, 'Rendah': 1 };
  const diff = priorityWeight[pB.level] - priorityWeight[pA.level];
  if (diff !== 0) return diff;
  
  // Tie-breaker: earlier entry date comes first
  return new Date(a.entryMadinah).getTime() - new Date(b.entryMadinah).getTime();
}

