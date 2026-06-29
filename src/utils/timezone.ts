/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 *
 * Timezone helpers for Raudhah scheduling.
 *
 * Key idea: the "Jadwal Slot Raudhah (Nusuk)" value is wall-clock time in MADINAH (GMT+3, no DST).
 * The team operates in WITA (GMT+8). A bare datetime-local string like "2026-06-27T08:00" has no
 * timezone, so `new Date(...)` would (wrongly) read it as the browser's local time. We instead pin
 * it to +03:00 so every countdown/reminder is computed from the correct absolute instant, regardless
 * of where the team's computer is.
 */

export const TZ_MADINAH = 'Asia/Riyadh'; // GMT+3
export const TZ_WITA = 'Asia/Makassar'; // GMT+8 (Waktu Indonesia Tengah)

const MADINAH_OFFSET = '+03:00';

/**
 * Parse a datetime-local string ("YYYY-MM-DDTHH:MM") that represents MADINAH wall-clock time
 * into an absolute Date instant. Returns null for empty/invalid input.
 */
export function parseMadinahSlot(slot?: string | null): Date | null {
  if (!slot) return null;
  // datetime-local may be "YYYY-MM-DDTHH:MM" (16 chars) or already include seconds.
  const normalized = slot.length === 16 ? `${slot}:00` : slot;
  const d = new Date(`${normalized}${MADINAH_OFFSET}`);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Absolute instant the team must act = Raudhah slot minus the lead time (default 2 hours).
 * This is when they should log into Nusuk and download the QR.
 */
export function getDistributionInstant(slot?: string | null, leadHours = 2): Date | null {
  const slotInstant = parseMadinahSlot(slot);
  if (!slotInstant) return null;
  return new Date(slotInstant.getTime() - leadHours * 3600_000);
}

/** Format an absolute instant as "HH:MM" (optionally with "DD Mon") in a given IANA timezone. */
export function formatInZone(date: Date, timeZone: string, withDate = false): string {
  return date.toLocaleString('id-ID', {
    timeZone,
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
    ...(withDate ? { day: '2-digit', month: 'short' } : {}),
  });
}

/** Convenience: format an instant in both Madinah and WITA, e.g. "06:00 Madinah / 11:00 WITA". */
export function formatDualZone(date: Date, withDate = false): string {
  return `${formatInZone(date, TZ_MADINAH, withDate)} Madinah / ${formatInZone(date, TZ_WITA, withDate)} WITA`;
}

/** Time as "HH:MM" with a colon (24-hour) in a given timezone, e.g. "22:19". */
export function formatTimeColon(date: Date, timeZone: string): string {
  return date.toLocaleTimeString('en-GB', { timeZone, hour: '2-digit', minute: '2-digit' });
}

/** Full date + time in a given timezone, e.g. "2 Juli 2026, 17:19". */
export function formatFullInZone(date: Date, timeZone: string): string {
  const d = date.toLocaleDateString('id-ID', { timeZone, day: 'numeric', month: 'long', year: 'numeric' });
  return `${d}, ${formatTimeColon(date, timeZone)}`;
}

/** Format an instant's calendar date with weekday in a given timezone, e.g. "Sen, 29 Jun". */
export function formatDateInZone(date: Date, timeZone: string): string {
  return date.toLocaleDateString('id-ID', {
    timeZone,
    weekday: 'short',
    day: '2-digit',
    month: 'short',
  });
}
