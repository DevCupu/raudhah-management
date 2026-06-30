/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import {
  LayoutDashboard,
  Users,
  FileSpreadsheet,
  UserCheck,
  Settings,
  Search,
  Plus,
  X,
  Filter,
  Upload,
  CheckCircle2,
  AlertTriangle,
  Clock,
  Phone,
  Calendar,
  ChevronRight,
  ChevronDown,
  Sparkles,
  Check,
  Copy,
  FileText,
  Eye,
  EyeOff,
  Loader2,
  Trash2,
  Edit,
  ArrowRight,
  Grid,
  BookOpen,
  Menu,
  Building2,
  Laptop,
  Key,
  Moon,
  Sun,
  Download,
  Bell,
  Activity,
  Gauge,
  Volume2
} from 'lucide-react';
import { Jamaah, Operator, JamaahStatus, Gender, CustomField } from './types';
import { getPriorityInfo, sortJamaahByPriorityAndDate } from './utils/priority';
import { getDistributionInstant, formatInZone, formatFullInZone, formatTimeColon, TZ_MADINAH, TZ_WITA } from './utils/timezone';
import {
  getSupabase,
  initSupabaseClient,
  getSupabaseConfig,
  mapJamaahToDb,
  mapJamaahFromDb,
  mapOperatorToDb,
  mapOperatorFromDb,
  mapCustomFieldToDb,
  mapCustomFieldFromDb
} from './utils/supabaseClient';
import { INITIAL_JAMAAH, INITIAL_OPERATORS } from './data/mockData';
import {
  buildVisaTextPrompt,
  extractTextFromPdf,
  requestVisaExtraction,
} from './utils/visaScan';
import * as XLSX from 'xlsx';
// @ts-ignore
import * as pdfjsLib from 'pdfjs-dist';
// @ts-ignore
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

// Configure pdfjs worker locally via Vite asset bundler
if (typeof window !== 'undefined') {
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;
}

// Uniform default password applied when generating default jamaah accounts
const DEFAULT_JAMAAH_PASSWORD = 'Visa2424@';

// --- IndexedDB Configuration for Jamaah Storage ---
const DB_NAME = 'raudhah_db';
const STORE_NAME = 'jamaahs';

const initDB = (): Promise<IDBDatabase> => {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = (e: any) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
    request.onsuccess = (e: any) => {
      resolve(e.target.result);
    };
    request.onerror = (e: any) => {
      reject(e.target.error);
    };
  });
};

const saveJamaahsToDB = async (data: Jamaah[]) => {
  try {
    const db = await initDB();
    const transaction = db.transaction(STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORE_NAME);
    store.put(data, 'raudhah_jamaahs');
  } catch (err) {
    console.error('IndexedDB save error:', err);
  }
};

const loadJamaahsFromDB = (): Promise<Jamaah[] | null> => {
  return new Promise(async (resolve) => {
    try {
      const db = await initDB();
      const transaction = db.transaction(STORE_NAME, 'readonly');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get('raudhah_jamaahs');
      request.onsuccess = () => {
        resolve(request.result || null);
      };
      request.onerror = () => {
        resolve(null);
      };
    } catch (err) {
      console.error('IndexedDB load error:', err);
      resolve(null);
    }
  });
};

// Helper to compress base64 images to stay within storage budget
const compressImage = (dataUrl: string, maxWidth = 800, quality = 0.7): Promise<string> => {
  return new Promise((resolve) => {
    if (!dataUrl || !dataUrl.startsWith('data:image')) {
      resolve(dataUrl);
      return;
    }
    const img = new Image();
    img.src = dataUrl;
    img.onload = () => {
      const canvas = document.createElement('canvas');
      let width = img.width;
      let height = img.height;

      if (width > maxWidth) {
        height = Math.round((height * maxWidth) / width);
        width = maxWidth;
      }

      canvas.width = width;
      canvas.height = height;

      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL('image/jpeg', quality));
      } else {
        resolve(dataUrl);
      }
    };
    img.onerror = () => resolve(dataUrl);
  });
};

// Normalize any stored entry value into the "YYYY-MM-DD" a <input type="date"> needs (drops any time part).
const toDateValue = (s: string): string => {
  if (!s) return '';
  return s.trim().replace(' ', 'T').split('T')[0];
};

type QrReminderTone = 'done' | 'past' | 'now' | 'soon' | 'warn' | 'normal';

// Human-friendly duration, e.g. "2 jam 10 menit lagi", "1 hari 3 jam lagi", "25 menit lagi".
const formatDurationLabel = (mins: number): string => {
  if (mins < 60) return `${Math.max(1, Math.ceil(mins))} menit lagi`;
  const totalMin = Math.round(mins);
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  const d = Math.floor(h / 24);
  if (d > 0) return `${d} hari ${h % 24} jam lagi`;
  return `${h} jam ${m} menit lagi`;
};

// Compute the QR-distribution reminder state for a jamaah, relative to `nowMs`.
// distInstant = slot − leadHours; everything is derived from the Madinah-pinned slot instant.
const getQrReminder = (
  raudhahSlot: string | null | undefined,
  status: string,
  nowMs: number,
  leadHours: number,
): { dist: Date; slot: Date; tone: QrReminderTone; countdownLabel: string } | null => {
  const dist = getDistributionInstant(raudhahSlot, leadHours);
  const slot = getDistributionInstant(raudhahSlot, 0);
  if (!dist || !slot) return null;
  const minsToDist = (dist.getTime() - nowMs) / 60000;
  const minsToSlot = (slot.getTime() - nowMs) / 60000;
  let tone: QrReminderTone;
  let countdownLabel: string;
  if (status === 'QR Berhasil') {
    tone = 'done'; countdownLabel = 'QR sudah siap';
  } else if (minsToSlot <= 0) {
    tone = 'past'; countdownLabel = 'Slot Raudhah terlewat';
  } else if (minsToDist <= 0) {
    tone = 'now'; countdownLabel = 'Download QR SEKARANG!';
  } else {
    countdownLabel = formatDurationLabel(minsToDist);
    tone = minsToDist <= 60 ? 'soon' : minsToDist <= 180 ? 'warn' : 'normal';
  }
  return { dist, slot, tone, countdownLabel };
};

// Tailwind classes per reminder tone for the countdown chip (text + bg + border).
const QR_REMINDER_BADGE: Record<QrReminderTone, string> = {
  done: 'text-emerald-700 dark:text-emerald-400 bg-emerald-50 dark:bg-emerald-500/10 border-emerald-100 dark:border-emerald-500/20',
  past: 'text-slate-500 dark:text-zinc-400 bg-slate-50 dark:bg-zinc-700/40 border-slate-200 dark:border-zinc-600/50',
  now: 'text-rose-700 dark:text-rose-300 bg-rose-100 dark:bg-rose-500/15 border-rose-300 dark:border-rose-500/40 animate-pulse',
  soon: 'text-rose-700 dark:text-rose-400 bg-rose-50 dark:bg-rose-500/10 border-rose-200 dark:border-rose-500/30',
  warn: 'text-amber-700 dark:text-amber-300 bg-amber-50 dark:bg-amber-500/10 border-amber-200 dark:border-amber-500/30',
  normal: 'text-sky-800 dark:text-sky-300 bg-sky-50 dark:bg-sky-900/15 border-sky-200 dark:border-sky-900/40',
};

// True if the Raudhah slot date falls outside the Madinah stay range [entry, exit].
// Returns false when any value is missing (nothing to validate against).
const isSlotOutsideStay = (slot?: string | null, entry?: string | null, exit?: string | null): boolean => {
  const s = toDateValue(slot || '');
  const e = toDateValue(entry || '');
  const x = toDateValue(exit || '');
  if (!s || !e || !x) return false;
  return s < e || s > x; // lexicographic compare is valid for YYYY-MM-DD
};

// "27 Jun 2026" from a date-only string (noon avoids any timezone day-shift).
const formatDateLabel = (s: string): string => {
  const d = toDateValue(s);
  if (!d) return '';
  const dt = new Date(`${d}T12:00:00`);
  if (isNaN(dt.getTime())) return d;
  return dt.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
};

const formatIndonesianDate = (dateStr: string): string => {
  try {
    const months = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    const parts = dateStr.split('-');
    if (parts.length === 3) {
      const day = parseInt(parts[2], 10);
      const month = months[parseInt(parts[1], 10) - 1];
      const year = parts[0];
      return `${day} ${month} ${year}`;
    }
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    const day = d.getDate();
    const month = months[d.getMonth()];
    const year = d.getFullYear();
    return `${day} ${month} ${year}`;
  } catch {
    return dateStr;
  }
};

export default function App() {
  // --- Persistent States ---
  const [jamaahs, setJamaahs] = useState<Jamaah[]>([]);
  const [isDbLoaded, setIsDbLoaded] = useState(false);
  const [isMobileSidebarOpen, setIsMobileSidebarOpen] = useState(false);

  const [operators, setOperators] = useState<Operator[]>(() => {
    const saved = localStorage.getItem('raudhah_operators');
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed.some((o: any) => o.id === 'op-1' || o.id === 'op-2')) {
        localStorage.removeItem('raudhah_operators');
        return [];
      }
      return parsed;
    }
    return INITIAL_OPERATORS;
  });

  const [currentTab, setCurrentTab] = useState<string>('dashboard');
  const [activeSettingSubTab, setActiveSettingSubTab] = useState<string>('umum');

  // --- Search & Filters ---
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('All');
  const [operatorFilter, setOperatorFilter] = useState<string>('All');
  const [genderFilter, setGenderFilter] = useState<string>('All');
  const [entryFilter, setEntryFilter] = useState('');
  const [travelFilter, setTravelFilter] = useState<string>('All'); // New filter for travel agent / rombongan
  const [priorityFilter, setPriorityFilter] = useState<string>('All'); // Filter by priority level (Tinggi/Sedang/Rendah)
  const [qrFilter, setQrFilter] = useState<string>('All'); // Filter by QR/file status (uploaded/pending)
  const [expandedTravels, setExpandedTravels] = useState<{ [key: string]: boolean }>({});

  // --- Data jamaah pagination ---
  const [jamaahPage, setJamaahPage] = useState(1);
  const [jamaahPageSize, setJamaahPageSize] = useState(25);

  // --- Detail / Modal States ---
  const [selectedJamaah, setSelectedJamaah] = useState<Jamaah | null>(null);
  const [showAddModal, setShowAddModal] = useState(false);
  const [editingJamaah, setEditingJamaah] = useState<Jamaah | null>(null);
  const [showEditModal, setShowEditModal] = useState(false);
  const [selectedJamaahIds, setSelectedJamaahIds] = useState<{ [id: string]: boolean }>({});

  // --- Batch Scan States ---
  const [showBatchScanModal, setShowBatchScanModal] = useState(false);
  const [showDuplicateModal, setShowDuplicateModal] = useState(false);
  const [copiedField, setCopiedField] = useState<string | null>(null);
  const [batchScanFiles, setBatchScanFiles] = useState<File[]>([]);
  const [batchScanResults, setBatchScanResults] = useState<any[]>([]);
  const [batchScanErrors, setBatchScanErrors] = useState<{ fileName: string; error: string }[]>([]);
  const [batchScanSuccessFilesCount, setBatchScanSuccessFilesCount] = useState<number>(0);
  const [batchScanFailedFilesCount, setBatchScanFailedFilesCount] = useState<number>(0);
  const [batchScanTotalFilesCount, setBatchScanTotalFilesCount] = useState<number>(0);
  const [isBatchScanning, setIsBatchScanning] = useState(false);
  const [batchScanProgress, setBatchScanProgress] = useState({ current: 0, total: 0, status: '' });
  const [isBatchScanPaused, setIsBatchScanPaused] = useState(false);
  const [batchScanTimeElapsed, setBatchScanTimeElapsed] = useState(0);
  const [batchScanFileMeta, setBatchScanFileMeta] = useState<{
    name: string;
    size: string;
    pages: number;
    charCount: number;
    type: 'PDF (Teks Digital)' | 'PDF (Gambar/Scan)' | 'Gambar (OCR Visual)' | 'Format Lain';
  } | null>(null);
  const [batchActiveFileProgress, setBatchActiveFileProgress] = useState(0);
  
  const cancelBatchRef = useRef<boolean>(false);
  const pauseBatchRef = useRef<boolean>(false);

  // --- Manual Jamaah Form State ---
  const [newJamaah, setNewJamaah] = useState<Partial<Jamaah>>({
    name: '',
    passport: '',
    visa: '',
    gender: 'Laki-laki' as Gender,
    phone: '',
    entryMadinah: '', // Tanggal masuk Madinah (kosong, diisi user)
    exitMadinah: '', // Tanggal keluar Madinah (kosong, diisi user)
    operatorId: '',
    notes: '',
    travel: '', // Nama travel/rombongan (kosong, diisi user)
    email: '',
    password: localStorage.getItem('raudhah_default_password') || 'Visa2424@', // Default password (auto-filled, editable)
    raudhahSlot: localStorage.getItem('raudhah_default_raudhah_slot') || '', // Booked Raudhah slot datetime string (YYYY-MM-DDTHH:MM)
    customValues: {},
  });

  // --- Operator Form State ---
  const [newOperatorName, setNewOperatorName] = useState('');
  const [newOperatorPhone, setNewOperatorPhone] = useState('');
  const [newOperatorPassword, setNewOperatorPassword] = useState('123456');

  // --- Auth & Login States ---
  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(() => {
    return sessionStorage.getItem('raudhah_is_logged_in') === 'true';
  });
  const [activeOperatorId, setActiveOperatorId] = useState<string | null>(() => {
    const saved = sessionStorage.getItem('raudhah_active_operator_id');
    return saved === 'null' || saved === null ? null : saved;
  });
  const [loginTargetId, setLoginTargetId] = useState('admin');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  const [showLoginPassword, setShowLoginPassword] = useState(false);
  const [adminPassword, setAdminPassword] = useState(() => {
    return localStorage.getItem('raudhah_admin_password') || 'admin123';
  });

  // --- Operator Form State Toggles ---
  const [showNewOperatorPassword, setShowNewOperatorPassword] = useState(false);

  // --- Import States ---
  const [importFile, setImportFile] = useState<File | null>(null);
  const [previewRows, setPreviewRows] = useState<any[]>([]);
  const [importStats, setImportStats] = useState({
    total: 0,
    valid: 0,
    duplicate: 0,
    incomplete: 0,
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  // --- Settings states ---
  const [settingsTravelName, setSettingsTravelName] = useState(() => {
    const saved = localStorage.getItem('raudhah_travel_name');
    return saved !== null ? saved : 'Raudhah Al-Haramain Travel';
  });
  const [settingsNusukLimit, setSettingsNusukLimit] = useState(() => {
    const saved = localStorage.getItem('raudhah_nusuk_limit');
    return saved ? parseInt(saved, 10) : 15;
  });
  // Lead time (hours) before the Raudhah slot when the team must download & distribute the QR. Default 2h.
  const [settingsQrLeadHours, setSettingsQrLeadHours] = useState(() => {
    const saved = localStorage.getItem('raudhah_qr_lead_hours');
    return saved ? parseFloat(saved) : 2;
  });
  const [settingsEnableSound, setSettingsEnableSound] = useState(() => {
    return localStorage.getItem('raudhah_enable_sound') !== 'false';
  });

  // Synthesizes a pleasant double-chime tone (E5 -> A5) using Web Audio API (no asset dependencies)
  const playNotificationSound = () => {
    try {
      const audioCtx = new (window.AudioContext || (window as any).webkitAudioContext)();
      const playTone = (freq: number, start: number, duration: number) => {
        const osc = audioCtx.createOscillator();
        const gain = audioCtx.createGain();
        osc.type = 'sine';
        osc.frequency.value = freq;
        gain.gain.setValueAtTime(0, start);
        gain.gain.linearRampToValueAtTime(0.15, start + 0.05);
        gain.gain.exponentialRampToValueAtTime(0.0001, start + duration);
        osc.connect(gain);
        gain.connect(audioCtx.destination);
        osc.start(start);
        osc.stop(start + duration);
      };
      const t = audioCtx.currentTime;
      playTone(659.25, t, 0.4); // E5
      playTone(880.00, t + 0.12, 0.6); // A5
    } catch (err) {
      console.warn('Failed to play notification sound:', err);
    }
  };

  // Live clock: one ticking Date drives both Madinah (GMT+3) and Indonesia/WITA (GMT+8) displays.
  // Both are the same instant — only the timezone label differs, so they can never drift apart (selisih tetap 5 jam).
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const timer = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);
  // Tracks which QR-distribution reminders have already fired (key: jamaahId|slot) so we notify once.
  const notifiedDistRef = useRef<Set<string>>(new Set());
  const isInitializedRef = useRef(false);

  // Ask for browser notification permission once (best-effort; ignored if unsupported/denied).
  useEffect(() => {
    if (typeof Notification !== 'undefined' && Notification.permission === 'default') {
      Notification.requestPermission().catch(() => {});
    }
  }, []);

  // Fire a browser notification & play sound when a jamaah's QR-distribution time arrives (slot − lead hours),
  // and the Raudhah slot hasn't passed yet. Driven by the 1s clock tick.
  useEffect(() => {
    if (!isDbLoaded) return;
    const t = now.getTime();

    // Silently register currently urgent slots at startup to prevent mass alerting on load
    if (!isInitializedRef.current) {
      jamaahs.forEach(j => {
        if (!j.raudhahSlot || j.status === 'QR Berhasil') return;
        const dist = getDistributionInstant(j.raudhahSlot, settingsQrLeadHours);
        const slot = getDistributionInstant(j.raudhahSlot, 0);
        if (!dist || !slot) return;
        const minsToDist = (dist.getTime() - t) / 60000;
        const minsToSlot = (slot.getTime() - t) / 60000;
        if (minsToDist <= 0 && minsToSlot > 0) {
          const key = `${j.id}|${j.raudhahSlot}`;
          notifiedDistRef.current.add(key);
        }
      });
      isInitializedRef.current = true;
      return;
    }

    let shouldPlaySound = false;
    jamaahs.forEach(j => {
      if (!j.raudhahSlot || j.status === 'QR Berhasil') return;
      const dist = getDistributionInstant(j.raudhahSlot, settingsQrLeadHours);
      const slot = getDistributionInstant(j.raudhahSlot, 0);
      if (!dist || !slot) return;
      const minsToDist = (dist.getTime() - t) / 60000;
      const minsToSlot = (slot.getTime() - t) / 60000;
      const key = `${j.id}|${j.raudhahSlot}`;
      
      // Distribution window is open (reached, slot not yet passed) and we haven't notified for it.
      if (minsToDist <= 0 && minsToSlot > 0 && !notifiedDistRef.current.has(key)) {
        notifiedDistRef.current.add(key);
        shouldPlaySound = true;
        
        if (typeof Notification !== 'undefined' && Notification.permission === 'granted') {
          try {
            new Notification('Waktunya Distribusi QR Raudhah', {
              body: `${j.name} — masuk Nusuk, download & berikan QR sekarang. Slot Raudhah ${formatInZone(slot, TZ_WITA)} WITA (${formatInZone(slot, TZ_MADINAH)} Madinah).`,
              tag: key,
            });
          } catch { /* notification may be blocked; ignore */ }
        }
      }
    });

    if (shouldPlaySound && settingsEnableSound) {
      playNotificationSound();
    }
  }, [now, jamaahs, settingsQrLeadHours, isDbLoaded, settingsEnableSound]);
  const [settingsReferenceDate, setSettingsReferenceDate] = useState(() => {
    const saved = localStorage.getItem('raudhah_reference_date');
    if (saved) return saved;
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  });
  // Global default password for jamaah accounts (editable, persisted)
  const [settingsDefaultPassword, setSettingsDefaultPassword] = useState(() => {
    return localStorage.getItem('raudhah_default_password') || DEFAULT_JAMAAH_PASSWORD;
  });
  const [showDefaultPassword, setShowDefaultPassword] = useState(false);
  // Default Raudhah entrance slot time (editable, persisted)
  const [settingsDefaultRaudhahSlot, setSettingsDefaultRaudhahSlot] = useState(() => {
    return localStorage.getItem('raudhah_default_raudhah_slot') || '';
  });
  const [settingsGeminiApiKey, setSettingsGeminiApiKey] = useState(() => {
    return localStorage.getItem('raudhah_gemini_api_key') || '';
  });
  const [settingsGeminiModel, setSettingsGeminiModel] = useState(() => {
    return localStorage.getItem('raudhah_gemini_model') || 'gemini-2.0-flash';
  });

  const [settingsSupabaseUrl, setSettingsSupabaseUrl] = useState(() => {
    return (import.meta as any).env?.VITE_SUPABASE_URL || localStorage.getItem('raudhah_supabase_url') || '';
  });
  const [settingsSupabaseAnonKey, setSettingsSupabaseAnonKey] = useState(() => {
    return (import.meta as any).env?.VITE_SUPABASE_ANON_KEY || localStorage.getItem('raudhah_supabase_anon_key') || '';
  });
  const [isSupabaseConnected, setIsSupabaseConnected] = useState(false);
  const [isSupabaseLoading, setIsSupabaseLoading] = useState(false);
  const [supabaseError, setSupabaseError] = useState<string | null>(null);
  const [showGeminiApiKey, setShowGeminiApiKey] = useState(false);


  const [settingsExportColumns, setSettingsExportColumns] = useState(() => {
    const saved = localStorage.getItem('raudhah_export_columns');
    return saved ? JSON.parse(saved) : {
      name: true,
      passport: true,
      visa: true,
      gender: true,
      phone: true,
      email: true,
      entryMadinah: true,
      exitMadinah: true,
      travel: true,
      password: true,
      status: true,
      operator: true,
      notes: true,
    };
  });

  const [customFields, setCustomFields] = useState<CustomField[]>(() => {
    const saved = localStorage.getItem('raudhah_custom_fields');
    return saved ? JSON.parse(saved) : [];
  });

  const [newCustomFieldLabel, setNewCustomFieldLabel] = useState('');
  const [editingCustomFieldId, setEditingCustomFieldId] = useState<string | null>(null);
  const [editingCustomFieldLabel, setEditingCustomFieldLabel] = useState('');

  const [isDarkMode, setIsDarkMode] = useState(() => {
    return localStorage.getItem('raudhah_dark_mode') === 'true';
  });
  // --- Database cleanup modal states ---
  const [showResetModal, setShowResetModal] = useState(false);
  const [resetScopes, setResetScopes] = useState({
    jamaah: true,
    operators: true,
    settings: true,
    credentials: true,
  });
  const [resetConfirmed, setResetConfirmed] = useState(false);

  const [isScanningVisa, setIsScanningVisa] = useState(false);
  const [scanVisaStatus, setScanVisaStatus] = useState('');
  const [ocrError, setOcrError] = useState<string | null>(null);
  const [isTestingApi, setIsTestingApi] = useState(false);
  const [apiTestResult, setApiTestResult] = useState<any | null>(null);
  
  // --- Google Gemini Quota Tracking States (Free Tier Limits) ---
  const [settingsDailyUsageCount, setSettingsDailyUsageCount] = useState<number>(() => {
    const saved = localStorage.getItem('raudhah_daily_usage_count');
    const resetDate = localStorage.getItem('raudhah_daily_usage_reset_date');
    const today = new Date().toDateString();

    if (resetDate !== today) {
      localStorage.setItem('raudhah_daily_usage_reset_date', today);
      localStorage.setItem('raudhah_daily_usage_count', '0');
      return 0;
    }
    return saved ? parseInt(saved, 10) : 0;
  });

  const [sessionMinuteRequests, setSessionMinuteRequests] = useState<number[]>([]);

  const recordGeminiApiCall = () => {
    const now = Date.now();
    setSettingsDailyUsageCount(prev => {
      const next = prev + 1;
      localStorage.setItem('raudhah_daily_usage_count', String(next));
      return next;
    });
    setSessionMinuteRequests(prev => {
      const oneMinuteAgo = now - 60000;
      const filtered = prev.filter(t => t > oneMinuteAgo);
      return [...filtered, now];
    });
  };

  const getRequestsPerMinute = () => {
    const now = Date.now();
    const oneMinuteAgo = now - 60000;
    return sessionMinuteRequests.filter(t => t > oneMinuteAgo).length;
  };

  // --- Deteksi rate-limit NYATA (saat Google benar-benar menolak dgn 429) ---
  // Berbeda dari hitungan RPM/RPD di atas yang hanya estimasi. Ini dipicu oleh
  // respons 429 sungguhan, menyimpan kapan cooldown berakhir (pakai saran RetryInfo Google).
  const [geminiRateLimitedUntil, setGeminiRateLimitedUntil] = useState<number>(() => {
    const v = localStorage.getItem('raudhah_gemini_ratelimit_until');
    return v ? Number(v) : 0;
  });
  const markGeminiRateLimited = (retrySec?: number) => {
    const until = Date.now() + Math.max(retrySec || 0, 15) * 1000;
    setGeminiRateLimitedUntil(until);
    localStorage.setItem('raudhah_gemini_ratelimit_until', String(until));
  };
  const clearGeminiRateLimited = () => {
    setGeminiRateLimitedUntil(0);
    localStorage.removeItem('raudhah_gemini_ratelimit_until');
  };
  const isRateLimitedNow = geminiRateLimitedUntil > Date.now();
  const rateLimitSecondsLeft = isRateLimitedNow ? Math.ceil((geminiRateLimitedUntil - Date.now()) / 1000) : 0;

  const [tick, setTick] = useState(0);
  useEffect(() => {
    const timer = setInterval(() => {
      setTick(prev => prev + 1);
    }, 15000);
    return () => clearInterval(timer);
  }, []);

  const isProModel = settingsGeminiModel.includes('pro');
  const limitRPM = isProModel ? 2 : 15;
  const limitRPD = isProModel ? 50 : 1500;
  const limitTPM = isProModel ? '32.000 TPM' : '1.000.000 TPM';
  const currentRPM = getRequestsPerMinute();
  const currentRPD = settingsDailyUsageCount;
  const rpmPercent = Math.min(100, (currentRPM / limitRPM) * 100);
  const rpdPercent = Math.min(100, (currentRPD / limitRPD) * 100);

  // --- Load and migrate data asynchronously from IndexedDB ---
  useEffect(() => {
    const loadData = async () => {
      const dbData = await loadJamaahsFromDB();
      if (dbData) {
        setJamaahs(dbData);
      } else {
        const legacy = localStorage.getItem('raudhah_jamaahs');
        if (legacy) {
          try {
            const parsed = JSON.parse(legacy);
            setJamaahs(parsed);
            await saveJamaahsToDB(parsed);
            localStorage.removeItem('raudhah_jamaahs');
          } catch (e) {
            setJamaahs(INITIAL_JAMAAH);
            await saveJamaahsToDB(INITIAL_JAMAAH);
          }
        } else {
          setJamaahs(INITIAL_JAMAAH);
          await saveJamaahsToDB(INITIAL_JAMAAH);
        }
      }
      setIsDbLoaded(true);
    };
    loadData();
  }, []);

  // --- Effects ---
  useEffect(() => {
    if (isDbLoaded && !isSupabaseConnected) {
      saveJamaahsToDB(jamaahs);
    }
  }, [jamaahs, isDbLoaded, isSupabaseConnected]);

  // Helper to update setting in Supabase
  const updateSettingInSupabase = async (key: string, value: string) => {
    const supabase = getSupabase();
    if (!supabase || !isSupabaseConnected) return;
    try {
      await supabase.from('settings').upsert({ key, value });
    } catch (e) {
      console.error('Failed to update setting in Supabase:', e);
    }
  };

  // Supabase Real-time Sync and Initial Fetch Effect
  useEffect(() => {
    let jamaahChannel: any = null;
    let operatorChannel: any = null;
    let customFieldChannel: any = null;
    let settingsChannel: any = null;

    const setupSupabase = async () => {
      const client = initSupabaseClient(settingsSupabaseUrl, settingsSupabaseAnonKey);
      if (!client) {
        setIsSupabaseConnected(false);
        return;
      }

      setIsSupabaseLoading(true);
      setSupabaseError(null);

      try {
        // Test connection
        const { error: testError } = await client
          .from('settings')
          .select('key')
          .limit(1);

        if (testError && testError.code !== 'PGRST116' && testError.code !== '42P01') {
          throw testError;
        }

        setIsSupabaseConnected(true);

        // Fetch operators
        const { data: dbOperators, error: opError } = await client
          .from('operators')
          .select('*');
        if (opError) throw opError;
        if (dbOperators) {
          setOperators(dbOperators.map(mapOperatorFromDb));
        }

        // Fetch custom fields
        const { data: dbCustomFields, error: cfError } = await client
          .from('custom_fields')
          .select('*');
        if (cfError) throw cfError;
        if (dbCustomFields) {
          setCustomFields(dbCustomFields.map(mapCustomFieldFromDb));
        }

        // Fetch settings
        const { data: dbSettings, error: setErr } = await client
          .from('settings')
          .select('*');
        if (!setErr && dbSettings) {
          dbSettings.forEach((s: any) => {
            if (s.key === 'travel_name') setSettingsTravelName(prev => prev !== s.value ? s.value : prev);
            if (s.key === 'nusuk_limit') setSettingsNusukLimit(prev => {
              const val = parseInt(s.value) || 15;
              return prev !== val ? val : prev;
            });
            if (s.key === 'qr_lead_hours') setSettingsQrLeadHours(prev => {
              const val = parseFloat(s.value) || 2;
              return prev !== val ? val : prev;
            });
            if (s.key === 'reference_date') setSettingsReferenceDate(prev => prev !== s.value ? s.value : prev);
            if (s.key === 'default_password') setSettingsDefaultPassword(prev => prev !== s.value ? s.value : prev);
            if (s.key === 'default_raudhah_slot') setSettingsDefaultRaudhahSlot(prev => prev !== s.value ? s.value : prev);
            if (s.key === 'admin_password') setAdminPassword(prev => prev !== s.value ? s.value : prev);
            if (s.key === 'enable_sound') setSettingsEnableSound(prev => {
              const val = s.value !== 'false';
              return prev !== val ? val : prev;
            });
          });
        }

        // Fetch jamaahs
        const { data: dbJamaahs, error: jamError } = await client
          .from('jamaahs')
          .select('*');
        if (jamError) throw jamError;
        if (dbJamaahs) {
          setJamaahs(dbJamaahs.map(mapJamaahFromDb));
        }

        // Subscribe to real-time changes
        jamaahChannel = client
          .channel('jamaah-changes')
          .on(
            'postgres_changes',
            { event: '*', schema: 'public', table: 'jamaahs' },
            (payload) => {
              const { eventType, new: newRec, old: oldRec } = payload;
              if (eventType === 'INSERT' || eventType === 'UPDATE') {
                const mapped = mapJamaahFromDb(newRec);
                setJamaahs((prev) => {
                  const exists = prev.some((j) => j.id === mapped.id);
                  if (exists) {
                    return prev.map((j) => (j.id === mapped.id && JSON.stringify(j) !== JSON.stringify(mapped) ? mapped : j));
                  } else {
                    return [mapped, ...prev];
                  }
                });
                setSelectedJamaah((prev) => {
                  if (prev && prev.id === mapped.id && JSON.stringify(prev) !== JSON.stringify(mapped)) {
                    return mapped;
                  }
                  return prev;
                });
              } else if (eventType === 'DELETE') {
                setJamaahs((prev) => prev.filter((j) => j.id !== oldRec.id));
                setSelectedJamaah((prev) => (prev && prev.id === oldRec.id ? null : prev));
              }
            }
          )
          .subscribe();

        operatorChannel = client
          .channel('operator-changes')
          .on(
            'postgres_changes',
            { event: '*', schema: 'public', table: 'operators' },
            (payload) => {
              const { eventType, new: newRec, old: oldRec } = payload;
              if (eventType === 'INSERT' || eventType === 'UPDATE') {
                const mapped = mapOperatorFromDb(newRec);
                setOperators((prev) => {
                  const exists = prev.some((o) => o.id === mapped.id);
                  if (exists) {
                    return prev.map((o) => (o.id === mapped.id && JSON.stringify(o) !== JSON.stringify(mapped) ? mapped : o));
                  } else {
                    return [...prev, mapped];
                  }
                });
              } else if (eventType === 'DELETE') {
                setOperators((prev) => prev.filter((o) => o.id !== oldRec.id));
              }
            }
          )
          .subscribe();

        customFieldChannel = client
          .channel('cf-changes')
          .on(
            'postgres_changes',
            { event: '*', schema: 'public', table: 'custom_fields' },
            (payload) => {
              const { eventType, new: newRec, old: oldRec } = payload;
              if (eventType === 'INSERT' || eventType === 'UPDATE') {
                const mapped = mapCustomFieldFromDb(newRec);
                setCustomFields((prev) => {
                  const exists = prev.some((cf) => cf.id === mapped.id);
                  if (exists) {
                    return prev.map((cf) => (cf.id === mapped.id && JSON.stringify(cf) !== JSON.stringify(mapped) ? mapped : cf));
                  } else {
                    return [...prev, mapped];
                  }
                });
              } else if (eventType === 'DELETE') {
                setCustomFields((prev) => prev.filter((cf) => cf.id !== oldRec.id));
              }
            }
          )
          .subscribe();

        settingsChannel = client
          .channel('settings-changes')
          .on(
            'postgres_changes',
            { event: '*', schema: 'public', table: 'settings' },
            (payload) => {
              const { eventType, new: newRec } = payload;
              if (eventType === 'INSERT' || eventType === 'UPDATE') {
                const { key, value } = newRec;
                if (key === 'travel_name') setSettingsTravelName(prev => prev !== value ? value : prev);
                if (key === 'nusuk_limit') setSettingsNusukLimit(prev => {
                  const val = parseInt(value) || 15;
                  return prev !== val ? val : prev;
                });
                if (key === 'qr_lead_hours') setSettingsQrLeadHours(prev => {
                  const val = parseFloat(value) || 2;
                  return prev !== val ? val : prev;
                });
                if (key === 'reference_date') setSettingsReferenceDate(prev => prev !== value ? value : prev);
                if (key === 'default_password') setSettingsDefaultPassword(prev => prev !== value ? value : prev);
                if (key === 'default_raudhah_slot') setSettingsDefaultRaudhahSlot(prev => prev !== value ? value : prev);
                if (key === 'admin_password') setAdminPassword(prev => prev !== value ? value : prev);
                if (key === 'enable_sound') setSettingsEnableSound(prev => {
                  const val = value !== 'false';
                  return prev !== val ? val : prev;
                });
              }
            }
          )
          .subscribe();

      } catch (err: any) {
        console.error('Supabase connection error:', err);
        setSupabaseError(err.message || String(err));
        setIsSupabaseConnected(false);
      } finally {
        setIsSupabaseLoading(false);
      }
    };

    setupSupabase();

    return () => {
      const client = getSupabase();
      if (client) {
        if (jamaahChannel) client.removeChannel(jamaahChannel);
        if (operatorChannel) client.removeChannel(operatorChannel);
        if (customFieldChannel) client.removeChannel(customFieldChannel);
        if (settingsChannel) client.removeChannel(settingsChannel);
      }
    };
  }, [settingsSupabaseUrl, settingsSupabaseAnonKey]);

  useEffect(() => {
    localStorage.setItem('raudhah_operators', JSON.stringify(operators));
  }, [operators]);

  useEffect(() => {
    localStorage.setItem('raudhah_travel_name', settingsTravelName);
    updateSettingInSupabase('travel_name', settingsTravelName);
  }, [settingsTravelName]);

  useEffect(() => {
    localStorage.setItem('raudhah_nusuk_limit', String(settingsNusukLimit));
    updateSettingInSupabase('nusuk_limit', String(settingsNusukLimit));
  }, [settingsNusukLimit]);

  useEffect(() => {
    localStorage.setItem('raudhah_qr_lead_hours', String(settingsQrLeadHours));
    updateSettingInSupabase('qr_lead_hours', String(settingsQrLeadHours));
  }, [settingsQrLeadHours]);

  useEffect(() => {
    localStorage.setItem('raudhah_enable_sound', String(settingsEnableSound));
    updateSettingInSupabase('enable_sound', String(settingsEnableSound));
  }, [settingsEnableSound]);

  useEffect(() => {
    localStorage.setItem('raudhah_admin_password', adminPassword);
    updateSettingInSupabase('admin_password', adminPassword);
  }, [adminPassword]);

  useEffect(() => {
    localStorage.setItem('raudhah_reference_date', settingsReferenceDate);
    updateSettingInSupabase('reference_date', settingsReferenceDate);
  }, [settingsReferenceDate]);


  useEffect(() => {
    localStorage.setItem('raudhah_gemini_api_key', settingsGeminiApiKey);
  }, [settingsGeminiApiKey]);

  useEffect(() => {
    localStorage.setItem('raudhah_gemini_model', settingsGeminiModel);
  }, [settingsGeminiModel]);



  useEffect(() => {
    localStorage.setItem('raudhah_supabase_url', settingsSupabaseUrl);
    initSupabaseClient(settingsSupabaseUrl, settingsSupabaseAnonKey);
  }, [settingsSupabaseUrl, settingsSupabaseAnonKey]);

  useEffect(() => {
    localStorage.setItem('raudhah_supabase_anon_key', settingsSupabaseAnonKey);
    initSupabaseClient(settingsSupabaseUrl, settingsSupabaseAnonKey);
  }, [settingsSupabaseUrl, settingsSupabaseAnonKey]);


  // Reset pagination to first page whenever filters/search change
  useEffect(() => {
    setJamaahPage(1);
  }, [searchQuery, statusFilter, operatorFilter, genderFilter, entryFilter, travelFilter, priorityFilter, qrFilter, jamaahPageSize]);

  useEffect(() => {
    localStorage.setItem('raudhah_default_password', settingsDefaultPassword);
    updateSettingInSupabase('default_password', settingsDefaultPassword);
  }, [settingsDefaultPassword]);

  useEffect(() => {
    localStorage.setItem('raudhah_default_raudhah_slot', settingsDefaultRaudhahSlot);
    updateSettingInSupabase('default_raudhah_slot', settingsDefaultRaudhahSlot);
  }, [settingsDefaultRaudhahSlot]);

  useEffect(() => {
    localStorage.setItem('raudhah_export_columns', JSON.stringify(settingsExportColumns));
  }, [settingsExportColumns]);

  useEffect(() => {
    localStorage.setItem('raudhah_custom_fields', JSON.stringify(customFields));
  }, [customFields]);

  useEffect(() => {
    localStorage.setItem('raudhah_dark_mode', String(isDarkMode));
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [isDarkMode]);

  // Initialize dark mode on mount
  useEffect(() => {
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
    }
  }, []);

  const handleLoginSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError('');

    if (loginTargetId === 'admin') {
      if (loginPassword === adminPassword) {
        sessionStorage.setItem('raudhah_is_logged_in', 'true');
        sessionStorage.setItem('raudhah_active_operator_id', 'null');
        setIsLoggedIn(true);
        setActiveOperatorId(null);
        setLoginPassword('');
        setCurrentTab('dashboard');
      } else {
        setLoginError('Password Admin salah! (Default: admin123)');
      }
    } else {
      const op = operators.find(o => o.id === loginTargetId);
      const correctPassword = op?.password || '123456';
      if (loginPassword === correctPassword) {
        sessionStorage.setItem('raudhah_is_logged_in', 'true');
        sessionStorage.setItem('raudhah_active_operator_id', loginTargetId);
        setIsLoggedIn(true);
        setActiveOperatorId(loginTargetId);
        setLoginPassword('');
        setCurrentTab('dashboard');
      } else {
        setLoginError('Password Operator salah! (Default: 123456)');
      }
    }
  };

  // Jaga statistik import (total/valid/duplikat/tidak lengkap) selalu sinkron dgn previewRows.
  // WAJIB di atas early-return `if (!isLoggedIn)` agar urutan hooks konsisten saat login/logout.
  useEffect(() => {
    setImportStats({
      total: previewRows.length,
      valid: previewRows.filter(r => r.status === 'Valid').length,
      duplicate: previewRows.filter(r => r.status === 'Duplikat').length,
      incomplete: previewRows.filter(r => r.status === 'Tidak Lengkap').length,
    });
  }, [previewRows]);

  if (!isLoggedIn) {
    return (
      <div className="min-h-screen flex antialiased overflow-hidden bg-zinc-950 selection:bg-red-100 selection:text-red-900">
        {/* ── LEFT PANEL: Masjidil Haram Photo ───────────────────────────────── */}
        <div className="hidden lg:flex relative flex-1 overflow-hidden">
          <img
            src="https://images.unsplash.com/photo-1591604129939-f1efa4d9f7fa?q=90&w=1800&fit=crop"
            alt="Masjidil Haram, Makkah"
            className="absolute inset-0 w-full h-full object-cover object-center"
          />
          <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-black/30 to-black/10" />
          <div className="absolute inset-0 bg-gradient-to-r from-transparent to-black/25" />

          {/* Logo top-left */}
          <div className="relative z-10 p-8 self-start">
            <img src="/logo.png" alt="Logo" className="h-10 object-contain bg-white/90 px-3 py-1 rounded-xl shadow-lg" />
          </div>

          {/* Caption bottom */}
          <div className="relative z-10 mt-auto p-10 text-white">
            <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-red-500/20 border border-red-400/30 text-red-300 text-[11px] font-semibold tracking-wide mb-4">
              <span className="w-1.5 h-1.5 rounded-full bg-red-400 animate-pulse" />
              Makkah Al-Mukarramah
            </div>
            <h2 className="text-3xl font-bold leading-snug mb-2 drop-shadow-lg">
              Masjidil Haram<br />
              <span className="text-blue-300">Mekah, Arab Saudi</span>
            </h2>
            <p className="text-sm text-white/70 max-w-sm leading-relaxed">
              Sistem pengelolaan QR Code Nusuk Madinah untuk jemaah umrah &amp; haji — cepat, akurat, dan terpusat.
            </p>
            <div className="flex items-center gap-3 mt-6">
              {[{ icon: <BookOpen className="w-3 h-3" />, t: 'Manifes digital' }, { icon: <Sparkles className="w-3 h-3" />, t: 'Prioritas otomatis' }, { icon: <CheckCircle2 className="w-3 h-3" />, t: 'Multi-operator' }].map((f, i) => (
                <div key={i} className="flex items-center gap-1.5 text-[11px] text-white/60">
                  {f.icon}<span>{f.t}</span>
                </div>
              ))}
            </div>
            <p className="text-[10px] text-white/30 mt-5">Foto: Unsplash · © 2026 Raudhah Manager</p>
          </div>
        </div>

        {/* ── RIGHT PANEL: Login Form ─────────────────────────────────────────── */}
        <div className="relative w-full lg:w-[440px] xl:w-[480px] flex flex-col bg-white dark:bg-zinc-900 overflow-y-auto">

          {/* Dark mode toggle */}
          <button
            type="button"
            onClick={() => setIsDarkMode(!isDarkMode)}
            title={isDarkMode ? 'Mode Terang' : 'Mode Gelap'}
            className="absolute top-5 right-5 z-20 p-2.5 rounded-full bg-slate-100 dark:bg-zinc-800 border border-slate-200 dark:border-zinc-700 text-slate-500 dark:text-zinc-400 hover:text-red-600 dark:hover:text-red-400 shadow-sm transition-colors cursor-pointer"
          >
            {isDarkMode ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
          </button>

          {/* Mobile: top image strip */}
          <div className="lg:hidden h-40 relative overflow-hidden shrink-0">
            <img
              src="https://images.unsplash.com/photo-1591604129939-f1efa4d9f7fa?q=80&w=800&fit=crop"
              alt="Masjidil Haram"
              className="w-full h-full object-cover object-[center_40%]"
            />
            <div className="absolute inset-0 bg-gradient-to-b from-black/20 to-black/65" />
            <div className="absolute bottom-4 left-5 flex items-center gap-2.5">
              <img src="/logo.png" alt="Logo" className="h-8 object-contain bg-white/90 px-2 py-0.5 rounded-lg shadow" />
              <span className="text-white font-bold text-sm drop-shadow">Raudhah Manager</span>
            </div>
          </div>

          {/* Form content */}
          <div className="flex-1 flex flex-col justify-center px-8 sm:px-12 py-12">

            <div className="mb-8">
              <div className="w-12 h-12 rounded-2xl bg-gradient-to-br from-red-600 to-blue-700 flex items-center justify-center shadow-lg shadow-blue-700/25 mb-4">
                <span className="text-2xl leading-none">🕌</span>
              </div>
              <h1 className="text-2xl font-bold text-slate-900 dark:text-zinc-50 tracking-tight">Selamat Datang</h1>
              <p className="text-sm text-slate-500 dark:text-zinc-400 mt-1">Masuk ke dashboard pengelolaan jemaah Anda.</p>
            </div>

            <form onSubmit={handleLoginSubmit} className="space-y-5">
              <div className="space-y-1.5">
                <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Pilih Akun / Peran</label>
                <div className="relative">
                  <Building2 className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 dark:text-zinc-500 pointer-events-none" />
                  <select
                    value={loginTargetId}
                    onChange={(e) => setLoginTargetId(e.target.value)}
                    className="w-full text-sm border border-slate-200 dark:border-zinc-700 rounded-xl py-3 pl-10 pr-9 bg-slate-50 dark:bg-zinc-800 text-slate-800 dark:text-zinc-100 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none font-medium appearance-none cursor-pointer transition-all"
                  >
                    <option value="admin">Kantor Pusat (Admin)</option>
                    {operators.filter(op => op.isActive).map(op => (
                      <option key={op.id} value={op.id}>Operator: {op.name}</option>
                    ))}
                  </select>
                  <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 dark:text-zinc-500 pointer-events-none" />
                </div>
              </div>

              <div className="space-y-1.5">
                <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Password Akses</label>
                <div className="relative">
                  <Key className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 dark:text-zinc-500 pointer-events-none" />
                  <input
                    type={showLoginPassword ? 'text' : 'password'}
                    required
                    placeholder="Ketik password..."
                    value={loginPassword}
                    onChange={(e) => setLoginPassword(e.target.value)}
                    className="w-full text-sm border border-slate-200 dark:border-zinc-700 rounded-xl py-3 pl-10 pr-11 bg-slate-50 dark:bg-zinc-800 text-slate-800 dark:text-zinc-100 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none font-mono tracking-widest transition-all"
                  />
                  <button
                    type="button"
                    onClick={() => setShowLoginPassword(!showLoginPassword)}
                    className="absolute right-3 top-1/2 -translate-y-1/2 p-1.5 rounded-lg text-slate-400 dark:text-zinc-500 hover:text-blue-500 dark:hover:text-blue-400 transition-colors cursor-pointer"
                  >
                    {showLoginPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
              </div>

              {loginError && (
                <div className="flex items-center gap-2 text-[11px] text-red-700 dark:text-red-400 bg-red-50 dark:bg-red-500/10 border border-red-200 dark:border-red-500/20 px-3 py-2.5 rounded-xl font-medium animate-in fade-in slide-in-from-top-1 duration-200">
                  <AlertTriangle className="w-3.5 h-3.5 shrink-0" />
                  <span>{loginError}</span>
                </div>
              )}

              <button
                type="submit"
                className="group w-full py-3.5 bg-gradient-to-r from-red-600 to-blue-700 hover:from-red-700 hover:to-blue-800 text-white rounded-xl text-sm font-bold transition-all shadow-lg shadow-blue-700/25 hover:shadow-blue-700/40 cursor-pointer flex items-center justify-center gap-2 active:scale-[0.98]"
              >
                <span>Masuk Aplikasi</span>
                <ArrowRight className="w-4 h-4 transition-transform group-hover:translate-x-0.5" />
              </button>
            </form>

            {/* Credentials hint */}
            <div className="mt-8 p-4 rounded-2xl bg-slate-50 dark:bg-zinc-800/50 border border-slate-100 dark:border-zinc-700/50">
              <p className="text-[10px] font-semibold text-slate-400 dark:text-zinc-500 mb-2 uppercase tracking-wider">Kredensial Default</p>
              <div className="grid grid-cols-2 gap-3 text-[11px]">
                <div className="space-y-0.5">
                  <p className="text-slate-400 dark:text-zinc-500">Admin</p>
                  <p className="font-mono font-bold text-slate-700 dark:text-zinc-200">admin123</p>
                </div>
                <div className="space-y-0.5">
                  <p className="text-slate-400 dark:text-zinc-500">Operator</p>
                  <p className="font-mono font-bold text-slate-700 dark:text-zinc-200">123456</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // --- Priority counters for general dashboard ---
  const countStatus = (status: JamaahStatus) => {
    return jamaahs.filter(j => j.status === status).length;
  };

  const countOperatorStatus = (opId: string, status: JamaahStatus) => {
    return jamaahs.filter(j => j.operatorId === opId && j.status === status).length;
  };

  // Filter jamaah based on search & filter fields
  const filteredJamaah = jamaahs.filter(j => {
    // Search by Nama, Paspor, Visa, Email, Travel
    const matchesSearch =
      j.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      j.passport.toLowerCase().includes(searchQuery.toLowerCase()) ||
      j.visa.toLowerCase().includes(searchQuery.toLowerCase()) ||
      (j.email && j.email.toLowerCase().includes(searchQuery.toLowerCase())) ||
      (j.travel && j.travel.toLowerCase().includes(searchQuery.toLowerCase()));

    const matchesStatus = statusFilter === 'All' || j.status === statusFilter;
    const matchesOperator =
      operatorFilter === 'All' ||
      (operatorFilter === 'unassigned' ? j.operatorId === null : j.operatorId === operatorFilter);
    const matchesGender = genderFilter === 'All' || j.gender === genderFilter;
    const matchesEntry = !entryFilter || j.entryMadinah === entryFilter;
    const matchesTravel = travelFilter === 'All' || j.travel === travelFilter;
    const matchesPriority = priorityFilter === 'All' ||
      getPriorityInfo(j.entryMadinah, settingsReferenceDate).level === priorityFilter;
    const matchesQr = qrFilter === 'All' ||
      (qrFilter === 'uploaded' ? !!j.qrCodeUrl : !j.qrCodeUrl);

    // If simulating as specific operator, ONLY show records assigned to that operator
    const matchesSimulation = !activeOperatorId || j.operatorId === activeOperatorId;

    return matchesSearch && matchesStatus && matchesOperator && matchesGender && matchesEntry && matchesSimulation && matchesTravel && matchesPriority && matchesQr;
  });

  // Sort by priority first
  const sortedFilteredJamaah = [...filteredJamaah].sort((a, b) => sortJamaahByPriorityAndDate(a, b, settingsReferenceDate));

  // --- Deteksi duplikat di SELURUH data jemaah (paspor ATAU visa sama) ---
  // Union-find: dua jemaah terhubung jika berbagi paspor atau nomor visa (non-kosong).
  const { duplicateGroups, duplicateIds } = (() => {
    const parent: Record<string, string> = {};
    const find = (x: string): string => { while (parent[x] !== x) { parent[x] = parent[parent[x]]; x = parent[x]; } return x; };
    const union = (a: string, b: string) => { parent[find(a)] = find(b); };
    jamaahs.forEach(j => { parent[j.id] = j.id; });
    const keyOwner: Record<string, string> = {};
    jamaahs.forEach(j => {
      const keys: string[] = [];
      const p = (j.passport || '').trim().toUpperCase();
      const v = (j.visa || '').trim().toUpperCase();
      if (p) keys.push('P:' + p);
      if (v) keys.push('V:' + v);
      keys.forEach(k => {
        if (keyOwner[k]) union(j.id, keyOwner[k]);
        else keyOwner[k] = j.id;
      });
    });
    const groupsMap: Record<string, Jamaah[]> = {};
    jamaahs.forEach(j => {
      const root = find(j.id);
      (groupsMap[root] = groupsMap[root] || []).push(j);
    });
    const groups = Object.values(groupsMap).filter(g => g.length > 1);
    const ids = new Set<string>();
    groups.forEach(g => g.forEach(j => ids.add(j.id)));
    return { duplicateGroups: groups, duplicateIds: ids };
  })();

  // High priority list (H-1 and H-2 s/d H-4) for quick action panel on dashboard
  const highPriorityList = jamaahs
    .filter(j => {
      // If operator simulation, only their assigned items
      if (activeOperatorId && j.operatorId !== activeOperatorId) return false;
      const prio = getPriorityInfo(j.entryMadinah, settingsReferenceDate);
      return (prio.level === 'Tinggi' || prio.level === 'Sedang') && j.status !== 'QR Berhasil';
    })
    .sort((a, b) => sortJamaahByPriorityAndDate(a, b, settingsReferenceDate));

  // --- Functions ---
  const openAddModal = () => {
    setNewJamaah({
      name: '',
      passport: '',
      visa: '',
      gender: 'Laki-laki',
      phone: '',
      entryMadinah: '',
      exitMadinah: '',
      operatorId: '',
      notes: '',
      travel: '',
      email: '',
      password: settingsDefaultPassword,
      raudhahSlot: settingsDefaultRaudhahSlot,
    });
    setOcrError(null);
    setShowAddModal(true);
  };

  const handleAddJamaah = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newJamaah.name || !newJamaah.passport || !newJamaah.visa) {
      alert('Nama, Paspor, dan Visa harus diisi!');
      return;
    }

    // Check duplicate
    const isDuplicate = jamaahs.some(
      j => j.passport.toLowerCase() === newJamaah.passport.toLowerCase() ||
           j.visa.toLowerCase() === newJamaah.visa.toLowerCase()
    );

    if (isDuplicate) {
      alert('Nomor Paspor atau Visa ini sudah terdaftar di sistem!');
      return;
    }

    const generatedPassword = newJamaah.password.trim() || Math.floor(100000 + Math.random() * 900000).toString();

    const created: Jamaah = {
      id: 'jam-' + Date.now(),
      name: newJamaah.name,
      passport: newJamaah.passport.toUpperCase(),
      visa: newJamaah.visa,
      gender: newJamaah.gender,
      phone: newJamaah.phone || '-',
      entryMadinah: newJamaah.entryMadinah,
      exitMadinah: newJamaah.exitMadinah,
      operatorId: newJamaah.operatorId || null,
      status: 'Ready',
      notes: newJamaah.notes || '',
      qrCodeUrl: null,
      qrUploadedAt: null,
      createdAt: new Date().toISOString(),
      travel: newJamaah.travel || settingsTravelName,
      email: newJamaah.email || '',
      password: generatedPassword,
      raudhahSlot: newJamaah.raudhahSlot || null,
    };

    setJamaahs([created, ...jamaahs]);
    
    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      supabase.from('jamaahs').insert(mapJamaahToDb(created)).then(({ error }) => {
        if (error) console.error('Failed to save to Supabase:', error);
      });
    }

    setShowAddModal(false);
    // Reset form
    setNewJamaah({
      name: '',
      passport: '',
      visa: '',
      gender: 'Laki-laki',
      phone: '',
      entryMadinah: '',
      exitMadinah: '',
      operatorId: '',
      notes: '',
      travel: '',
      email: '',
      password: settingsDefaultPassword,
      raudhahSlot: settingsDefaultRaudhahSlot,
    });
  };

  const handleOpenEditModal = (j: Jamaah) => {
    setEditingJamaah({ ...j });
    setShowEditModal(true);
  };

  // Salin teks ke clipboard + feedback singkat (untuk tombol copy kredensial Nusuk).
  const copyToClipboard = async (text: string, key: string) => {
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      setCopiedField(key);
      setTimeout(() => setCopiedField(prev => (prev === key ? null : prev)), 1500);
    } catch (e) {
      console.error('Gagal menyalin:', e);
    }
  };

  const handleUpdateJamaah = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingJamaah) return;
    if (!editingJamaah.name || !editingJamaah.passport || !editingJamaah.visa) {
      alert('Nama, Paspor, dan Visa harus diisi!');
      return;
    }

    // Check duplicate (excluding itself)
    const isDuplicate = jamaahs.some(
      j => j.id !== editingJamaah.id && (
        j.passport.toLowerCase() === editingJamaah.passport.toLowerCase() ||
        j.visa.toLowerCase() === editingJamaah.visa.toLowerCase()
      )
    );

    if (isDuplicate) {
      alert('Nomor Paspor atau Visa ini sudah terdaftar di sistem!');
      return;
    }

    const normalized = { ...editingJamaah, entryMadinah: toDateValue(editingJamaah.entryMadinah) };
    setJamaahs(prev => prev.map(j => j.id === normalized.id ? normalized : j));
    
    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      supabase.from('jamaahs').update(mapJamaahToDb(normalized)).eq('id', normalized.id).then(({ error }) => {
        if (error) console.error('Failed to update in Supabase:', error);
      });
    }

    setShowEditModal(false);
    setEditingJamaah(null);
  };

  const handleDeleteJamaah = (id: string) => {
    if (window.confirm('Apakah Anda yakin ingin menghapus data jamaah ini?')) {
      setJamaahs(prev => prev.filter(j => j.id !== id));
      // also remove from selected list if present
      setSelectedJamaahIds(prev => {
        const copy = { ...prev };
        delete copy[id];
        return copy;
      });
      if (selectedJamaah?.id === id) {
        setSelectedJamaah(null);
      }

      const supabase = getSupabase();
      if (supabase && isSupabaseConnected) {
        supabase.from('jamaahs').delete().eq('id', id).then(({ error }) => {
          if (error) console.error('Failed to delete in Supabase:', error);
        });
      }
    }
  };

  const handleDeleteSelectedJamaahs = () => {
    const idsToDelete = Object.keys(selectedJamaahIds).filter(id => selectedJamaahIds[id]);
    if (idsToDelete.length === 0) {
      alert('Pilih setidaknya satu jamaah untuk dihapus!');
      return;
    }

    if (window.confirm(`Apakah Anda yakin ingin menghapus ${idsToDelete.length} data jamaah yang terpilih?`)) {
      setJamaahs(prev => prev.filter(j => !selectedJamaahIds[j.id]));
      setSelectedJamaahIds({});
      if (selectedJamaah && idsToDelete.includes(selectedJamaah.id)) {
        setSelectedJamaah(null);
      }

      const supabase = getSupabase();
      if (supabase && isSupabaseConnected) {
        supabase.from('jamaahs').delete().in('id', idsToDelete).then(({ error }) => {
          if (error) console.error('Failed to delete selected in Supabase:', error);
        });
      }
    }
  };

  const handleQuickStatusChange = (id: string, newStatus: JamaahStatus) => {
    setJamaahs(prev =>
      prev.map(j => {
        if (j.id === id) {
          const updated = { ...j, status: newStatus };
          // If selected in detail panel, keep synchronized
          if (selectedJamaah && selectedJamaah.id === id) {
            setSelectedJamaah(updated);
          }
          return updated;
        }
        return j;
      })
    );

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      supabase.from('jamaahs').update({ status: newStatus }).eq('id', id).then(({ error }) => {
        if (error) console.error('Failed to update status in Supabase:', error);
      });
    }
  };

  const handleAssignOperatorInDetail = (opId: string | null) => {
    if (!selectedJamaah) return;
    setJamaahs(prev =>
      prev.map(j => {
        if (j.id === selectedJamaah.id) {
          const updated = { ...j, operatorId: opId };
          setSelectedJamaah(updated);
          return updated;
        }
        return j;
      })
    );

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      supabase.from('jamaahs').update({ operator_id: opId }).eq('id', selectedJamaah.id).then(({ error }) => {
        if (error) console.error('Failed to assign operator in Supabase:', error);
      });
    }
  };

  const handleUpdateNotesInDetail = (text: string) => {
    if (!selectedJamaah) return;
    setJamaahs(prev =>
      prev.map(j => {
        if (j.id === selectedJamaah.id) {
          const updated = { ...j, notes: text };
          setSelectedJamaah(updated);
          return updated;
        }
        return j;
      })
    );

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      supabase.from('jamaahs').update({ notes: text }).eq('id', selectedJamaah.id).then(({ error }) => {
        if (error) console.error('Failed to update notes in Supabase:', error);
      });
    }
  };

  const handleUploadScreenshotInDetail = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !selectedJamaah) return;

    const reader = new FileReader();
    reader.onload = async (event) => {
      const rawDataUrl = event.target?.result as string;
      const compressedDataUrl = await compressImage(rawDataUrl);
      const nowStr = new Date().toISOString();
      
      setJamaahs(prev =>
        prev.map(j => {
          if (j.id === selectedJamaah.id) {
            const updated = {
              ...j,
              qrCodeUrl: compressedDataUrl,
              qrUploadedAt: nowStr,
              status: 'QR Berhasil' as JamaahStatus, // auto promote
            };
            setSelectedJamaah(updated);
            return updated;
          }
          return j;
        })
      );

      const supabase = getSupabase();
      if (supabase && isSupabaseConnected) {
        supabase.from('jamaahs').update({
          qr_code_url: compressedDataUrl,
          qr_uploaded_at: nowStr,
          status: 'QR Berhasil'
        }).eq('id', selectedJamaah.id).then(({ error }) => {
          if (error) console.error('Failed to save QR in Supabase:', error);
        });
      }
    };
    reader.readAsDataURL(file);
  };

  const handleAddOperator = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newOperatorName || !newOperatorPhone) {
      alert('Isi nama dan nomor HP operator!');
      return;
    }
    const created: Operator = {
      id: 'op-' + Date.now(),
      name: newOperatorName,
      phone: newOperatorPhone,
      password: newOperatorPassword || '123456',
      isActive: true,
    };
    setOperators([...operators, created]);

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      supabase.from('operators').insert(mapOperatorToDb(created)).then(({ error }) => {
        if (error) console.error('Failed to save operator in Supabase:', error);
      });
    }

    setNewOperatorName('');
    setNewOperatorPhone('');
    setNewOperatorPassword('123456');
  };

  const toggleOperatorStatus = (id: string) => {
    setOperators(prev =>
      prev.map(op => {
        if (op.id === id) {
          const updated = { ...op, isActive: !op.isActive };
          const supabase = getSupabase();
          if (supabase && isSupabaseConnected) {
            supabase.from('operators').update({ is_active: updated.isActive }).eq('id', id).then(({ error }) => {
              if (error) console.error('Failed to toggle operator status in Supabase:', error);
            });
          }
          return updated;
        }
        return op;
      })
    );
  };

  // Helper to parse Excel Serial dates with time support
  const formatExcelDateTime = (serial: number, hasTime: boolean = true) => {
    try {
      const excelEpoch = 25569;
      const msInDay = 86400000;
      const dateInfo = new Date(Math.round((serial - excelEpoch) * msInDay));
      
      const year = dateInfo.getFullYear();
      const month = String(dateInfo.getMonth() + 1).padStart(2, '0');
      const day = String(dateInfo.getDate()).padStart(2, '0');
      
      if (hasTime) {
        const hours = String(dateInfo.getHours()).padStart(2, '0');
        const minutes = String(dateInfo.getMinutes()).padStart(2, '0');
        return `${year}-${month}-${day}T${hours}:${minutes}`;
      }
      return `${year}-${month}-${day}`;
    } catch {
      return '';
    }
  };

  const padTime = (timeStr: string) => {
    if (!timeStr) return '08:00';
    const parts = timeStr.split(':');
    const h = (parts[0] || '0').padStart(2, '0');
    const m = (parts[1] || '0').padStart(2, '0');
    return `${h}:${m}`;
  };

  const cleanStringDateTime = (str: string) => {
    if (!str) return ''; // kosong -> kosong (jangan tanam tanggal default)
    let cleaned = str.replace(' ', 'T').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(cleaned)) {
      cleaned += 'T08:00';
    }
    // Handle formats with slashes (e.g., 27/06/2026 08:00 or DD/MM/YYYY HH:MM)
    if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}/.test(cleaned)) {
      try {
        const parts = cleaned.split('T');
        const datePart = parts[0];
        const timePart = parts[1] || '08:00';
        const dateSubparts = datePart.split(/[\/\-]/);
        if (dateSubparts.length === 3) {
          let dayStr = dateSubparts[0];
          let monthStr = dateSubparts[1];
          let yearStr = dateSubparts[2];
          // If year is first (YYYY-MM-DD or YYYY/MM/DD)
          if (dayStr.length === 4) {
            const tmp = dayStr;
            dayStr = yearStr;
            yearStr = tmp;
          }
          const formattedDate = `${yearStr.padStart(4, '20')}-${monthStr.padStart(2, '0')}-${dayStr.padStart(2, '0')}`;
          return `${formattedDate}T${padTime(timePart)}`;
        }
      } catch {
        // Fallback
      }
    }
    if (!cleaned.includes('T')) {
      return '2026-06-27T08:00';
    }
    const parts = cleaned.split('T');
    return `${parts[0]}T${padTime(parts[1])}`;
  };

  const cleanStringDateOnly = (str: string) => {
    if (!str) return ''; // kosong -> kosong (jangan tanam tanggal default)
    let cleaned = str.trim();
    if (cleaned.includes('T')) {
      cleaned = cleaned.split('T')[0];
    } else if (cleaned.includes(' ')) {
      cleaned = cleaned.split(' ')[0];
    }
    // Handle formats with slashes e.g. 02/07/2026 to 2026-07-02
    if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}/.test(cleaned)) {
      const parts = cleaned.split(/[\/\-]/);
      if (parts.length === 3) {
        let d = parts[0];
        let m = parts[1];
        let y = parts[2];
        if (d.length === 4) {
          const tmp = d;
          d = y;
          y = tmp;
        }
        return `${y.padStart(4, '20')}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
      }
    }
    return cleaned;
  };

  // Helper to download real, pre-filled Excel template matching the required schema
  const downloadExcelTemplate = () => {
    const headers = [
      'Nama Lengkap',
      'Nomor Paspor',
      'Nomor Visa',
      'Jenis Kelamin (Laki-laki / Perempuan)',
      'No WhatsApp',
      'Email',
      'Tanggal Masuk Madinah (YYYY-MM-DD)',
      'Tanggal Keluar Madinah (YYYY-MM-DD)',
      'Nama Travel / Rombongan',
      'Password Akses Jemaah'
    ];

    // Append custom fields to template headers
    customFields.forEach(cf => {
      headers.push(cf.label);
    });

    // Example rows (only placed on the "Contoh Pengisian" sheet)
    const exampleRows = [
      [
        'Mochammad Rizky', 'A9122384', 'V445588111', 'Laki-laki', '+62811223344',
        'rizky@annahl-travel.com', '2026-06-27 08:00', '2026-07-02', 'An-Nahl Umrah & Haji', 'rizky123',
        ...customFields.map((_, i) => i === 0 ? 'Kamar 402' : '-')
      ],
      [
        'Siti Aminah', 'A5566778', 'V445566778', 'Perempuan', '+628123456789',
        'siti.aminah@annahl-travel.com', '2026-06-28 14:00', '2026-07-04', 'An-Nahl Umrah & Haji', 'siti990',
        ...customFields.map((_, i) => i === 0 ? 'Kamar 403' : '-')
      ]
    ];

    // Beautiful, consistent column widths shared by both sheets
    const baseCols = [
      { wch: 24 }, // Nama Lengkap
      { wch: 15 }, // Nomor Paspor
      { wch: 15 }, // Nomor Visa
      { wch: 30 }, // Jenis Kelamin
      { wch: 18 }, // No WhatsApp
      { wch: 30 }, // Email
      { wch: 34 }, // Tanggal Masuk Madinah
      { wch: 32 }, // Tanggal Keluar Madinah
      { wch: 26 }, // Nama Travel / Rombongan
      { wch: 20 }  // Password Akses Jemaah
    ];
    const customCols = customFields.map(cf => ({ wch: Math.max(16, cf.label.length + 4) }));
    const cols = [...baseCols, ...customCols];
    const lastColRef = `${XLSX.utils.encode_col(headers.length - 1)}1`;

    // Helper to format a sheet nicely (widths, header row height, filter dropdowns, freeze header)
    const decorateSheet = (ws: any) => {
      ws['!cols'] = cols;
      ws['!rows'] = [{ hpt: 24 }]; // taller header row
      ws['!autofilter'] = { ref: `A1:${lastColRef}` };
      ws['!freeze'] = { xSplit: 0, ySplit: 1 }; // keep header visible when scrolling
    };

    // Sheet 1: empty manifest — only the header table, ready to fill
    const wsManifest = XLSX.utils.aoa_to_sheet([headers]);
    decorateSheet(wsManifest);

    // Sheet 2: example data showing the correct format
    const wsExample = XLSX.utils.aoa_to_sheet([headers, ...exampleRows]);
    decorateSheet(wsExample);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsManifest, 'Manifest Jemaah');
    XLSX.utils.book_append_sheet(wb, wsExample, 'Contoh Pengisian');
    XLSX.writeFile(wb, 'template_manifest_jamaah.xlsx');
  };



  // Export current (filtered) jamaah data to a real Excel file
  const exportJamaahToExcel = () => {
    const dataToExport = sortedFilteredJamaah;
    if (dataToExport.length === 0) {
      alert('Tidak ada data jamaah untuk diekspor.');
      return;
    }

    // Mapping columns configuration to actual labels, value getters, and widths
    const columnDefinitions = [
      { key: 'name', label: 'Nama Lengkap', getValue: (j: Jamaah) => j.name, width: 24 },
      { key: 'passport', label: 'Nomor Paspor', getValue: (j: Jamaah) => j.passport, width: 15 },
      { key: 'visa', label: 'Nomor Visa', getValue: (j: Jamaah) => j.visa, width: 15 },
      { key: 'gender', label: 'Jenis Kelamin', getValue: (j: Jamaah) => j.gender, width: 14 },
      { key: 'phone', label: 'No WhatsApp', getValue: (j: Jamaah) => j.phone || '-', width: 18 },
      { key: 'email', label: 'Email', getValue: (j: Jamaah) => j.email || '-', width: 28 },
      { key: 'entryMadinah', label: 'Tanggal Masuk Madinah', getValue: (j: Jamaah) => (j.entryMadinah || '').replace('T', ' '), width: 22 },
      { key: 'exitMadinah', label: 'Tanggal Keluar Madinah', getValue: (j: Jamaah) => j.exitMadinah || '', width: 20 },
      { key: 'travel', label: 'Nama Travel / Rombongan', getValue: (j: Jamaah) => j.travel || settingsTravelName, width: 26 },
      { key: 'password', label: 'Password Akses Jemaah', getValue: (j: Jamaah) => j.password || '', width: 18 },
      { key: 'status', label: 'Status Booking', getValue: (j: Jamaah) => j.status, width: 16 },
      { 
        key: 'operator', 
        label: 'Operator Penanggung Jawab', 
        getValue: (j: Jamaah) => {
          const op = operators.find(o => o.id === j.operatorId);
          return op ? op.name : 'Belum Ditugaskan';
        }, 
        width: 22 
      },
      { key: 'notes', label: 'Catatan', getValue: (j: Jamaah) => j.notes || '', width: 30 }
    ];

    // Filter dynamic column definitions based on settings
    const activeColumns = columnDefinitions.filter(col => (settingsExportColumns as any)[col.key] !== false);

    if (activeColumns.length === 0) {
      alert('Anda harus memilih setidaknya satu kolom untuk diekspor di menu Settings!');
      return;
    }

    const headers = activeColumns.map(col => col.label);
    const rows = dataToExport.map(j => activeColumns.map(col => col.getValue(j)));

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = activeColumns.map(col => ({ wch: col.width }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data Jamaah');

    const today = new Date();
    const stamp = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
    XLSX.writeFile(wb, `data_jamaah_${stamp}.xlsx`);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportFile(file);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        // Read workbook
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        // Convert to array of arrays
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];

        let colNameIdx = -1;
        let colPassportIdx = -1;
        let colVisaIdx = -1;
        let colGenderIdx = -1;
        let colPhoneIdx = -1;
        let colEntryIdx = -1;
        let colExitIdx = -1;
        let colTravelIdx = -1;
        let colEmailIdx = -1;
        let colPasswordIdx = -1;

        if (data.length > 0) {
          const headers = data[0].map(h => String(h || '').toLowerCase().trim());
          
          headers.forEach((h, index) => {
            if (h.includes('travel') || h.includes('rombongan') || h.includes('grup') || h.includes('group') || h.includes('biro') || h.includes('agen')) {
              colTravelIdx = index;
            } else if (h.includes('paspor') || h.includes('passport') || h.includes('no paspor') || h.includes('nomor paspor')) {
              colPassportIdx = index;
            } else if (h.includes('visa') || h.includes('no visa') || h.includes('nomor visa')) {
              colVisaIdx = index;
            } else if (h.includes('gender') || h.includes('kelamin') || h.includes('sex') || h.includes('jenis kelamin')) {
              colGenderIdx = index;
            } else if (h.includes('phone') || h.includes('hp') || h.includes('telepon') || h.includes('no hp') || h.includes('wa') || h.includes('whatsapp')) {
              colPhoneIdx = index;
            } else if (h.includes('masuk') || h.includes('entry') || h.includes('datang') || h.includes('checkin') || h.includes('check-in')) {
              colEntryIdx = index;
            } else if (h.includes('keluar') || h.includes('exit') || h.includes('pulang') || h.includes('checkout') || h.includes('check-out')) {
              colExitIdx = index;
            } else if (h.includes('email') || h.includes('surel') || h.includes('mail')) {
              colEmailIdx = index;
            } else if (h.includes('password') || h.includes('sandi') || h.includes('pass')) {
              colPasswordIdx = index;
            } else if (h.includes('nama') || h.includes('name') || h.includes('lengkap')) {
              colNameIdx = index;
            }
          });
        }
        
        // Fallbacks for header indices if not found by name
        if (colNameIdx === -1) colNameIdx = 0;
        if (colPassportIdx === -1) colPassportIdx = 1;
        if (colVisaIdx === -1) colVisaIdx = 2;
        if (colGenderIdx === -1) colGenderIdx = 3;
        if (colPhoneIdx === -1) colPhoneIdx = 4;
        if (colEmailIdx === -1) colEmailIdx = 5;
        if (colEntryIdx === -1) colEntryIdx = 6;
        if (colExitIdx === -1) colExitIdx = 7;
        if (colTravelIdx === -1) colTravelIdx = 8;
        if (colPasswordIdx === -1) colPasswordIdx = 9;

        // Map dynamic custom fields headers
        const customFieldsMapping = customFields.map(cf => {
          let matchedIdx = -1;
          if (data.length > 0) {
            const headers = data[0].map(h => String(h || '').toLowerCase().trim());
            const cleanCfLabel = cf.label.toLowerCase().trim();
            headers.forEach((h, idx) => {
              if (
                h === cleanCfLabel || 
                h.includes(cleanCfLabel) || 
                cleanCfLabel.includes(h) || 
                h.replace(/\s+/g, '') === cleanCfLabel.replace(/\s+/g, '')
              ) {
                matchedIdx = idx;
              }
            });
          }
          return { id: cf.id, index: matchedIdx };
        });

        const parsedRows: any[] = [];
        for (let i = 1; i < data.length; i++) {
          const r = data[i];
          if (!r || r.length === 0 || !r[colNameIdx]) continue; // skip empty rows

          const name = String(r[colNameIdx] || '').trim();
          const passport = String(r[colPassportIdx] || '').trim().toUpperCase();
          const visa = String(r[colVisaIdx] || '').trim();
          
          const genderVal = String(r[colGenderIdx] || '').trim().toLowerCase();
          // Gender DIKOSONGKAN jika tidak ada datanya / tidak dikenali (diisi manual oleh user).
          let gender: Gender | '' = '';
          if (genderVal) {
            if (
              genderVal.startsWith('p') ||
              genderVal.includes('perempuan') ||
              genderVal === 'f' ||
              genderVal === 'female' ||
              genderVal.includes('wanita') ||
              genderVal === 'w'
            ) {
              gender = 'Perempuan';
            } else if (
              genderVal.startsWith('l') ||
              genderVal.includes('laki') ||
              genderVal === 'm' ||
              genderVal === 'male' ||
              genderVal === 'pria'
            ) {
              gender = 'Laki-laki';
            }
            // selain itu: tetap '' (tidak dikenali)
          }

          const phone = String(r[colPhoneIdx] || '').trim();
          
          let entryMadinah = String(r[colEntryIdx] || '').trim();
          let exitMadinah = String(r[colExitIdx] || '').trim();
          
          // Excel serial dates can be read sometimes. If it's a number (with optional decimal), convert it.
          // Entry is date-only (no hour needed).
          if (/^\d{5}(\.\d+)?$/.test(entryMadinah)) {
            entryMadinah = formatExcelDateTime(Number(entryMadinah), false);
          } else {
            entryMadinah = cleanStringDateOnly(entryMadinah);
          }

          if (/^\d{5}(\.\d+)?$/.test(exitMadinah)) {
            exitMadinah = formatExcelDateTime(Number(exitMadinah), false);
          } else {
            exitMadinah = cleanStringDateOnly(exitMadinah);
          }

          // Tanggal DIKOSONGKAN jika tidak ada datanya (diisi manual oleh user).
          if (entryMadinah === 'undefined') entryMadinah = '';
          if (exitMadinah === 'undefined') exitMadinah = '';

          // Read Travel column. If empty or missing, fallback to settingsTravelName
          let travel = settingsTravelName;
          if (colTravelIdx !== -1 && r[colTravelIdx]) {
            travel = String(r[colTravelIdx]).trim();
          }
          travel = travel.toUpperCase().trim();

          // Read Email column. Jika kosong, BUAT OTOMATIS dari nama:
          // huruf kecil semua, tanpa spasi/simbol, + @mailnesia.com
          let email = '';
          if (colEmailIdx !== -1 && r[colEmailIdx]) {
            email = String(r[colEmailIdx]).trim();
          }
          if (!email && name) {
            const slug = name.toLowerCase().replace(/[^a-z0-9]/g, '');
            if (slug) email = `${slug}@mailnesia.com`;
          }

          // Read Password column
          let password = '';
          if (colPasswordIdx !== -1 && r[colPasswordIdx]) {
            password = String(r[colPasswordIdx]).trim();
          }

          // Read dynamic custom values
          const customValues: Record<string, string> = {};
          customFieldsMapping.forEach(m => {
            if (m.index !== -1 && r[m.index] !== undefined && r[m.index] !== null) {
              customValues[m.id] = String(r[m.index]).trim();
            }
          });

          // Validate duplicates
          const isDuplicateInState = jamaahs.some(j => j.passport === passport || j.visa === visa);
          const isDuplicateInPreview = parsedRows.some(p => p.passport === passport || p.visa === visa);
          
          let status = 'Valid';
          let reason = 'Siap di-import';

          if (!name || !passport || !visa) {
            status = 'Tidak Lengkap';
            reason = !name ? 'Nama kosong' : !passport ? 'Paspor kosong' : 'Visa kosong';
          } else if (isDuplicateInState || isDuplicateInPreview) {
            status = 'Duplikat';
            reason = 'Nomor Paspor/Visa sudah terdaftar di sistem';
          }

          parsedRows.push({
            name,
            passport,
            visa,
            gender,
            phone,
            email,
            entryMadinah,
            exitMadinah,
            travel,
            status,
            reason,
            password,
            customValues
          });
        }

        setPreviewRows(parsedRows);
        setImportStats({
          total: parsedRows.length,
          valid: parsedRows.filter(r => r.status === 'Valid').length,
          duplicate: parsedRows.filter(r => r.status === 'Duplikat').length,
          incomplete: parsedRows.filter(r => r.status === 'Tidak Lengkap').length,
        });

      } catch (err) {
        console.error(err);
        alert('Gagal membaca file Excel. Pastikan format file sesuai.');
      }
    };

    reader.readAsBinaryString(file);
  };

  // Validasi ulang satu baris preview (status & alasan) berdasarkan kelengkapan & duplikat.
  const revalidatePreviewRow = (row: any, allRows: any[]): { status: string; reason: string } => {
    const name = (row.name || '').trim();
    const passport = (row.passport || '').trim().toUpperCase();
    const visa = (row.visa || '').trim();
    if (!name || !passport || !visa) {
      return { status: 'Tidak Lengkap', reason: !name ? 'Nama kosong' : !passport ? 'Paspor kosong' : 'Visa kosong' };
    }
    const dupInState = jamaahs.some(j => j.passport === passport || (visa && j.visa === visa));
    const dupInPreview = allRows.some(p => p !== row && ((p.passport || '').trim().toUpperCase() === passport || (visa && (p.visa || '').trim() === visa)));
    if (dupInState || dupInPreview) {
      return { status: 'Duplikat', reason: 'Nomor Paspor/Visa sudah terdaftar di sistem' };
    }
    return { status: 'Valid', reason: 'Siap di-import' };
  };

  // Edit satu sel pada baris preview sebelum import; status & alasan dihitung ulang.
  // Statistik diperbarui otomatis via useEffect di bawah.
  const updatePreviewRow = (targetRow: any, field: string, value: string) => {
    setPreviewRows(prev => {
      const editedIdx = prev.indexOf(targetRow);
      const next = prev.map(r => (r === targetRow ? { ...r, [field]: value } : r));
      if (editedIdx !== -1) {
        const { status, reason } = revalidatePreviewRow(next[editedIdx], next);
        next[editedIdx] = { ...next[editedIdx], status, reason };
      }
      return next;
    });
  };

  // Ganti nama travel/rombongan untuk semua baris preview dalam satu grup (sebelum import).
  const renamePreviewTravel = (oldName: string, newNameRaw: string) => {
    const newName = newNameRaw.trim().toUpperCase();
    const oldNameUpper = oldName.trim().toUpperCase();
    if (newName === oldNameUpper) return;
    setPreviewRows(prev =>
      prev.map(r => {
        const currentTravel = (r.travel || settingsTravelName || '').trim().toUpperCase();
        return currentTravel === oldNameUpper ? { ...r, travel: newName } : r;
      })
    );
  };

  const executeImport = (specificTravel?: string) => {
    // Only import 'Valid' rows. 'Duplikat' and 'Tidak Lengkap' rows are strictly BLOCKED.
    const rowsToImport = previewRows.filter(r => r.status === 'Valid' && (!specificTravel || r.travel === specificTravel));
    
    if (rowsToImport.length === 0) {
      alert(specificTravel ? `Tidak ada data baru/valid untuk travel "${specificTravel}" yang dapat di-import.` : 'Tidak ada data baru/valid yang dapat di-import.');
      return;
    }

    let newCount = 0;

    setJamaahs(prev => {
      let updated = [...prev];
      rowsToImport.forEach((r, index) => {
        const generatedPassword = r.password || Math.floor(100000 + Math.random() * 900000).toString();
        
        const importedItem: Jamaah = {
          id: 'jam-imported-' + Date.now() + index,
          name: r.name,
          passport: r.passport.toUpperCase(),
          visa: r.visa,
          gender: r.gender,
          phone: r.phone || '-',
          email: r.email || '',
          entryMadinah: r.entryMadinah,
          exitMadinah: r.exitMadinah,
          operatorId: null,
          status: 'Ready' as JamaahStatus,
          notes: `Di-import via Excel (${r.travel}).`,
          qrCodeUrl: null,
          qrUploadedAt: null,
          createdAt: new Date().toISOString(),
          travel: r.travel || settingsTravelName,
          password: generatedPassword,
          raudhahSlot: settingsDefaultRaudhahSlot || null,
          customValues: r.customValues || {}
        };

        updated.unshift(importedItem);
        newCount++;
      });
      return updated;
    });

    const finalImportedList: Jamaah[] = [];
    rowsToImport.forEach((r, index) => {
      const generatedPassword = r.password || Math.floor(100000 + Math.random() * 900000).toString();
      const importedItem: Jamaah = {
        id: 'jam-imported-' + Date.now() + index,
        name: r.name,
        passport: r.passport.toUpperCase(),
        visa: r.visa,
        gender: r.gender,
        phone: r.phone || '-',
        email: r.email || '',
        entryMadinah: r.entryMadinah,
        exitMadinah: r.exitMadinah,
        operatorId: null,
        status: 'Ready' as JamaahStatus,
        notes: `Di-import via Excel (${r.travel}).`,
        qrCodeUrl: null,
        qrUploadedAt: null,
        createdAt: new Date().toISOString(),
        travel: r.travel || settingsTravelName,
        password: generatedPassword,
        raudhahSlot: settingsDefaultRaudhahSlot || null,
        customValues: r.customValues || {}
      };
      finalImportedList.push(importedItem);
    });

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected && finalImportedList.length > 0) {
      supabase.from('jamaahs').insert(finalImportedList.map(mapJamaahToDb)).then(({ error }) => {
        if (error) console.error('Failed to insert imported jamaahs in Supabase:', error);
      });
    }

    saveJamaahsToDB([...finalImportedList, ...jamaahs]);

    alert(`Berhasil mengimpor ${newCount} jemaah baru. Data duplikat otomatis dilewati.`);
    
    // Clear preview rows that were imported
    const remainingRows = previewRows.filter(r => !rowsToImport.includes(r));
    setPreviewRows(remainingRows);
    setImportStats({
      total: remainingRows.length,
      valid: remainingRows.filter(r => r.status === 'Valid').length,
      duplicate: remainingRows.filter(r => r.status === 'Duplikat').length,
      incomplete: remainingRows.filter(r => r.status === 'Tidak Lengkap').length,
    });

    if (remainingRows.length === 0) {
      setImportFile(null);
      setCurrentTab('jamaah');
    }
  };

  const cancelImport = () => {
    setImportFile(null);
    setPreviewRows([]);
  };

  const handleScanVisaImage = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!settingsGeminiApiKey) {
      alert('Harap masukkan Google Gemini API Key terlebih dahulu di menu Settings (Pengaturan)!');
      return;
    }

    setOcrError(null);
    setIsScanningVisa(true);
    setScanVisaStatus('Membaca berkas gambar visa...');

    try {
      recordGeminiApiCall();
      const reader = new FileReader();
      reader.onloadend = async () => {
        try {
          const rawBase64 = reader.result as string;
          setScanVisaStatus('Mengompresi gambar (menghemat token & kuota)...');
          const compressedBase64 = await compressImage(rawBase64, 800, 0.7);
          
          const base64Data = compressedBase64.split(',')[1];
          const mimeType = compressedBase64.split(';')[0].split(':')[1];

          setScanVisaStatus('Mengirim dokumen ke Gemini AI...');

          const response = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/${settingsGeminiModel}:generateContent?key=${settingsGeminiApiKey}`,
            {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
              },
              body: JSON.stringify({
                contents: [
                  {
                    parts: [
                      {
                        text: "Ekstrak data penting dari gambar visa umrah/haji ini ke dalam format JSON dengan kunci berikut:\n- name (Nama lengkap jemaah)\n- passport (Nomor Paspor)\n- visa (Nomor Visa)\n- gender (Isi dengan \"Laki-laki\" atau \"Perempuan\" saja. Tebak secara cerdas dari nama jemaah jika kolom tidak tertulis eksplisit)\n- travel (Nama travel/agen umrah yang tertera pada visa, cari di kolom 'External Agent' atau 'Umrah Operator'. Jika tertulis dalam bahasa/huruf Arab, terjemahkan atau transliterasikan ke huruf Latin Indonesia/Inggris secara cerdas, contoh: 'مجموعة بي تي رحمة الدولية' menjadi 'PT Rahma Internasional')\n\nHarap kembalikan HANYA string JSON mentah. Jangan gunakan blok format markdown (seperti ```json) atau teks penjelasan lainnya agar data langsung dapat di-parse oleh sistem. Jika informasi tertentu tidak ditemukan, berikan nilai string kosong \"\"."
                      },
                      {
                        inlineData: {
                          mimeType: mimeType,
                          data: base64Data
                        }
                      }
                    ]
                  }
                ],
                generationConfig: {
                  responseMimeType: 'application/json',
                  temperature: 0,
                  // Matikan "thinking" pada model 2.5/3.x agar respons jauh lebih cepat.
                  ...((settingsGeminiModel.includes('2.5') || settingsGeminiModel.includes('3.'))
                    ? { thinkingConfig: { thinkingBudget: 0 } }
                    : {})
                }
              })
            }
          );

          if (!response.ok) {
            let errorMsg = `API error status ${response.status}`;
            try {
              const errData = await response.json();
              if (errData?.error?.message) {
                errorMsg = errData.error.message;
              }
            } catch (e) {
              // ignore body parsing error
            }
            throw new Error(errorMsg);
          }

          setScanVisaStatus('Memproses hasil pembacaan...');
          const data = await response.json();
          const text = data.candidates?.[0]?.content?.parts?.[0]?.text;

          if (!text) {
            throw new Error('Gemini API returned no content parts. The image might be blank or blocked by safety filters.');
          }

          let cleanJson = text.trim();
          const jsonMatch = cleanJson.match(/\{[\s\S]*\}/);
          if (jsonMatch) {
            cleanJson = jsonMatch[0];
          }

          const extracted = JSON.parse(cleanJson);
          
          setNewJamaah(prev => ({
            ...prev,
            name: extracted.name || prev.name,
            passport: (extracted.passport || '').toUpperCase() || prev.passport,
            visa: extracted.visa || prev.visa,
            gender: (extracted.gender === 'Laki-laki' || extracted.gender === 'Perempuan') ? extracted.gender : prev.gender,
            // Tanggal Masuk/Keluar Madinah sengaja DIKOSONGKAN agar diisi manual oleh user,
            // bukan auto-isi dari hasil scan.
            entryMadinah: '',
            exitMadinah: '',
            travel: extracted.travel || prev.travel
          }));

          alert('AI Sukses! Data visa berhasil dideteksi dan diisi otomatis.');
        } catch (err: any) {
          console.error(err);
          setOcrError(err.message || String(err));
          alert(`Gagal mendeteksi data visa. Silakan periksa detail error di bawah deskripsi scan visa untuk menyalinnya.`);
        } finally {
          setIsScanningVisa(false);
          setScanVisaStatus('');
        }
      };
      reader.readAsDataURL(file);
    } catch (err) {
      console.error(err);
      alert('Gagal memuat berkas gambar.');
      setIsScanningVisa(false);
      setScanVisaStatus('');
    }
  };

  const handleTestGeminiConnection = async () => {
    if (!settingsGeminiApiKey) {
      alert('Harap masukkan API Key Gemini terlebih dahulu!');
      return;
    }

    setIsTestingApi(true);
    setApiTestResult(null);
    recordGeminiApiCall();

    const diagnostics: {
      status: 'success' | 'error';
      message: string;
      availableModels: string[];
      selectedModelStatus: 'working' | 'quota_exceeded' | 'not_found' | 'error';
      selectedModelDetails: string;
    } = {
      status: 'success',
      message: '',
      availableModels: [],
      selectedModelStatus: 'working',
      selectedModelDetails: '',
    };

    try {
      // 1. Fetch available models
      const listResponse = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models?key=${settingsGeminiApiKey}`
      );

      if (!listResponse.ok) {
        let errorMsg = `HTTP Error ${listResponse.status}`;
        try {
          const errData = await listResponse.json();
          if (errData?.error?.message) {
            errorMsg = errData.error.message;
          }
        } catch (e) {}
        throw new Error(`Gagal memuat daftar model: ${errorMsg}`);
      }

      const listData = await listResponse.json();
      const models = listData.models || [];
      const modelNames = models.map((m: any) => m.name.replace('models/', ''));
      diagnostics.availableModels = modelNames;

      // Check if selected model is in list
      const isSelectedModelAvailable = modelNames.includes(settingsGeminiModel);

      if (!isSelectedModelAvailable) {
        diagnostics.selectedModelStatus = 'not_found';
        diagnostics.selectedModelDetails = `Model "${settingsGeminiModel}" tidak ditemukan untuk API Key Anda. Silakan pilih model lain di dropdown.`;
      } else {
        // 2. Perform dry-run test generateContent using selected model
        try {
          const testResponse = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/${settingsGeminiModel}:generateContent?key=${settingsGeminiApiKey}`,
            {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
              },
              body: JSON.stringify({
                contents: [{ parts: [{ text: 'Hi' }] }]
              })
            }
          );

          if (!testResponse.ok) {
            let errorMsg = `HTTP Error ${testResponse.status}`;
            try {
              const errData = await testResponse.json();
              if (errData?.error?.message) {
                errorMsg = errData.error.message;
              }
            } catch (e) {}

            if (testResponse.status === 429 || errorMsg.toLowerCase().includes('quota') || errorMsg.toLowerCase().includes('limit')) {
              diagnostics.selectedModelStatus = 'quota_exceeded';
              diagnostics.selectedModelDetails = `Kuota terlampaui / Limit 0: ${errorMsg}`;
              // Tandai rate-limit NYATA agar tampil di panel Settings (cooldown default).
              markGeminiRateLimited();
            } else {
              diagnostics.selectedModelStatus = 'error';
              diagnostics.selectedModelDetails = `Gagal uji coba konten: ${errorMsg}`;
            }
          } else {
            diagnostics.selectedModelStatus = 'working';
            diagnostics.selectedModelDetails = `Model berjalan normal dan siap digunakan.`;
            // Model merespons normal -> pasti tidak sedang kena rate limit. Bersihkan status.
            clearGeminiRateLimited();
          }
        } catch (testErr: any) {
          diagnostics.selectedModelStatus = 'error';
          diagnostics.selectedModelDetails = `Error koneksi uji coba: ${testErr.message || testErr}`;
        }
      }

      diagnostics.status = 'success';
      diagnostics.message = 'Koneksi ke Google API Server berhasil terhubung!';
    } catch (err: any) {
      diagnostics.status = 'error';
      diagnostics.message = err.message || String(err);
      diagnostics.selectedModelStatus = 'error';
      diagnostics.selectedModelDetails = 'Tidak dapat melakukan tes model karena inisialisasi API key gagal.';
    } finally {
      setIsTestingApi(false);
      setApiTestResult(diagnostics);
    }
  };

  const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

  const processBatchVisaScan = async (filesList: FileList | File[]) => {
    const files = Array.from(filesList);
    if (files.length === 0) return;

    if (!settingsGeminiApiKey) {
      alert('Harap masukkan Google Gemini API Key terlebih dahulu di menu Settings (Pengaturan)!');
      return;
    }

    if (files.length > 50) {
      alert('Batas maksimal scan massal adalah 50 file sekaligus! Harap kurangi jumlah file yang Anda pilih.');
      return;
    }

    // Filter berkas (Hanya mendukung PDF dengan teks digital di bawah 20MB)
    const validFiles = files.filter(f => {
      if (f.type !== 'application/pdf') {
        setBatchScanErrors(prev => [...prev, { fileName: f.name, error: 'Hanya mendukung berkas PDF dengan teks digital.' }]);
        return false;
      }
      if (f.size > 20 * 1024 * 1024) {
        setBatchScanErrors(prev => [...prev, { fileName: f.name, error: 'File terlalu besar (melebihi 20MB)' }]);
        return false;
      }
      return true;
    });

    if (validFiles.length === 0) {
      alert('Tidak ada berkas PDF valid yang terpilih untuk diproses.');
      return;
    }

    setIsBatchScanning(true);
    setIsBatchScanPaused(false);
    setBatchScanResults([]);
    setBatchScanErrors([]);
    setBatchScanFileMeta(null);
    setBatchScanSuccessFilesCount(0);
    setBatchScanFailedFilesCount(0);
    setBatchScanTotalFilesCount(validFiles.length);
    cancelBatchRef.current = false;
    pauseBatchRef.current = false;

    // Start stopwatch timer
    setBatchScanTimeElapsed(0);
    const startTimestamp = Date.now();
    const intervalId = setInterval(() => {
      if (!pauseBatchRef.current) {
        setBatchScanTimeElapsed(Math.floor((Date.now() - startTimestamp) / 1000));
      }
    }, 1000);

    let successVisaCount = 0;
    let failedFileCount = 0;
    const seenPassportsInBatch = new Set<string>();

    // Newer Flash models (2.5+/3.x) allow up to 65k output tokens; older ones cap at 8192.
    const isThinkingModel = settingsGeminiModel.includes('2.5') || settingsGeminiModel.includes('3.');
    const maxOut = isThinkingModel ? 65536 : 8192;
    // Thinking models (2.5/3.x) "berpikir" dulu sebelum menjawab -> sangat lambat untuk
    // ekstraksi sederhana. Matikan dengan thinkingBudget 0. Model 2.0 tidak mendukung param ini.
    const thinkingConfig = isThinkingModel ? { thinkingConfig: { thinkingBudget: 0 } } : {};

    try {
      for (let i = 0; i < validFiles.length; i++) {
        // Check for cancel
        if (cancelBatchRef.current) {
          setBatchScanProgress(prev => ({ ...prev, status: 'Proses scan massal dibatalkan oleh pengguna.' }));
          break;
        }

        // Check for pause
        while (pauseBatchRef.current && !cancelBatchRef.current) {
          setBatchScanProgress(prev => ({ ...prev, status: 'Scan dijeda. Menunggu untuk dilanjutkan...' }));
          await sleep(500);
        }
        if (cancelBatchRef.current) break;

        const file = validFiles[i];
        
        // Initialize file meta
        let pageCount = 1;
        let charCount = 0;
        let docType: 'PDF (Teks Digital)' | 'PDF (Gambar/Scan)' | 'Gambar (OCR Visual)' | 'Format Lain' = 
          file.type.startsWith('image/') ? 'Gambar (OCR Visual)' : 'Format Lain';

        setBatchScanFileMeta({
          name: file.name,
          size: (file.size / 1024 / 1024).toFixed(2) + ' MB',
          pages: pageCount,
          charCount: charCount,
          type: docType
        });

        setBatchScanProgress({
          current: i + 1,
          total: validFiles.length,
          status: `Sedang membaca berkas: ${file.name}...`
        });

        let attempt = 0;
        let success = false;

        while (attempt < 3 && !success) {
          if (cancelBatchRef.current) break;
          attempt++;
          try {
            let pdfText = '';
            setBatchActiveFileProgress(10);

            // Mengekstrak teks digital dari PDF
            setBatchScanProgress(prev => ({ ...prev, status: `Mengekstrak teks digital dari PDF: ${file.name}...` }));
            const arrayBuffer = await new Promise<ArrayBuffer>((resolve, reject) => {
              const reader = new FileReader();
              reader.onload = () => resolve(reader.result as ArrayBuffer);
              reader.onerror = () => reject(new Error('Gagal membaca ArrayBuffer'));
              reader.readAsArrayBuffer(file);
            });

            // Count pages using PDFJS
            // @ts-ignore
            const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer.slice(0)) });
            const pdfDoc = await loadingTask.promise;
            pageCount = pdfDoc.numPages;

            pdfText = await extractTextFromPdf(arrayBuffer, (curr, tot) => {
              setBatchActiveFileProgress(Math.floor(10 + (curr / tot) * 30));
            });
            const cleanedText = pdfText.replace(/--- HALAMAN \d+ ---/g, '').trim();
            charCount = cleanedText.length;
            docType = 'PDF (Teks Digital)';

            // Update metadata
            setBatchScanFileMeta({
              name: file.name,
              size: (file.size / 1024 / 1024).toFixed(2) + ' MB',
              pages: pageCount,
              charCount: charCount,
              type: docType
            });

            if (cleanedText.length < 15) {
              throw new Error('PDF tidak memiliki lapisan teks digital (hasil scan/foto tidak didukung).');
            }

            setBatchActiveFileProgress(55);

            // Bungkus tipis di sekitar modul visaScan: suntik kredensial & config dari state.
            const requestVisa = (parts: any[]): Promise<any[]> =>
              requestVisaExtraction({
                apiKey: settingsGeminiApiKey,
                model: settingsGeminiModel,
                parts,
                maxOut,
                thinkingConfig,
              });

            // PDF teks digital: kirim teks saja (cepat & hemat).
            setBatchScanProgress(prev => ({ ...prev, status: `Mengirim teks visa ke Gemini AI (Hemat & Cepat): ${file.name}...` }));
            recordGeminiApiCall(); // hitung kuota hanya saat benar-benar memanggil API
            const extracted = await requestVisa([{ text: buildVisaTextPrompt(pdfText) }]);

            const items = Array.isArray(extracted) ? extracted : [extracted];
            const validItems = items.filter(item => {
              const passport = (item.passport || '').toUpperCase().trim();
              item.passport = passport;
              item.visa = (item.visa || '').trim();
              item.customValues = item.customValues || {};
              item.fileName = file.name;
              if (!passport || passport.length < 4) return false;
              if (seenPassportsInBatch.has(passport)) return false;
              seenPassportsInBatch.add(passport);
              return true;
            });
            if (validItems.length > 0) {
              setBatchScanResults(prev => [...prev, ...validItems]);
            }
            if (items.length !== validItems.length) {
              const skipped = items.length - validItems.length;
              console.warn(`Batch scan: ${skipped} item(s) skipped (no passport/duplicate) in file: ${file.name}`);
            }
            successVisaCount += validItems.length;
            setBatchActiveFileProgress(100);
            success = true;
            setBatchScanSuccessFilesCount(prev => prev + 1);
          } catch (err: any) {
            console.error(err);
            const msg = String(err?.message || err);
            // Hanya rate-limit & gangguan jaringan yang layak diulang. Error lain
            // (PDF tanpa teks, output terpotong, JSON invalid) pasti berulang -> gagal cepat.
            const isNetwork = /failed to fetch|networkerror|load failed/i.test(msg);
            const retryable = Boolean(err?.isRateLimit) || isNetwork;
            const waitSec = err?.isRateLimit ? Math.max(err.retryDelaySec || 0, 30) : 2;
            if (err?.isRateLimit) {
              // Tandai status rate-limit NYATA agar tampil di panel Settings.
              markGeminiRateLimited(err.retryDelaySec);
              setBatchScanProgress(prev => ({
                ...prev,
                status: `Rate limit terlampaui. Menjeda cooldown selama ${waitSec} detik sesuai saran Google...`
              }));
            }
            if (!retryable || attempt === 3) {
              failedFileCount++;
              setBatchScanFailedFilesCount(prev => prev + 1);
              setBatchScanErrors(prev => [...prev, { fileName: file.name, error: msg }]);
              break; // berhenti mengulang berkas ini
            } else {
              await sleep(waitSec * 1000); // cooldown sebelum percobaan berikutnya
            }
          }
      }

      // Add a minor gap sleep (1.5 seconds) to avoid rate limit spamming on success
      if (success && i < validFiles.length - 1) {
        await sleep(1500);
      }
    }
    } finally {
      clearInterval(intervalId);
      setIsBatchScanning(false);
      setBatchScanProgress(prev => ({
        ...prev,
        status: cancelBatchRef.current
          ? `Scan dibatalkan. ${successVisaCount} visa berhasil dibaca, ${failedFileCount} berkas gagal.`
          : `Scan selesai! Berhasil membaca ${successVisaCount} visa dari ${validFiles.length} berkas, ${failedFileCount} berkas gagal.`
      }));
    }
  };

  const saveBatchScanResults = () => {
    if (batchScanResults.length === 0) return;

    // Filter out entries without valid passport numbers
    const validResults = batchScanResults.filter(r => {
      const passport = (r.passport || '').toUpperCase().trim();
      return passport.length >= 4;
    });

    if (validResults.length !== batchScanResults.length) {
      const skipped = batchScanResults.length - validResults.length;
      console.warn(`saveBatchScanResults: ${skipped} item(s) skipped (no valid passport)`);
    }

    if (validResults.length === 0) {
      alert('Tidak ada data jemaah valid yang bisa disimpan. Pastikan hasil scan memiliki nomor paspor.');
      return;
    }

    // Deduplicate within validResults itself (by passport first, then visa)
    const seenKeys = new Set<string>();
    const deduped = validResults.filter(r => {
      const passport = (r.passport || '').toUpperCase().trim();
      const visa = (r.visa || '').trim();
      if (seenKeys.has(passport) || seenKeys.has(visa)) return false;
      if (passport) seenKeys.add(passport);
      if (visa) seenKeys.add(visa);
      return true;
    });

    if (deduped.length !== validResults.length) {
      console.warn(`saveBatchScanResults: ${validResults.length - deduped.length} item(s) removed as internal duplicates`);
    }

    if (deduped.length === 0) {
      alert('Tidak ada data unik yang bisa disimpan (semua data sudah ada di sistem).');
      return;
    }

    const batchTimestamp = Date.now();
    let duplicateCount = 0;
    let newCount = 0;

    const newItems: Jamaah[] = [];

    deduped.forEach((r, index) => {
      const passport = (r.passport || '').toUpperCase().trim();
      const visa = (r.visa || '').trim();
      const isDuplicate = jamaahs.some(
        j => j.passport.toLowerCase() === passport.toLowerCase() ||
             (visa && j.visa.toLowerCase() === visa.toLowerCase())
      );

      if (isDuplicate) {
        duplicateCount++;
        return; // SKIP duplicate completely!
      }

      const generatedPassword = settingsDefaultPassword || Math.floor(100000 + Math.random() * 900000).toString();

      let email = '';
      const name = r.name || 'Jemaah Tanpa Nama';
      const slug = name.toLowerCase().replace(/[^a-z0-9]/g, '');
      if (slug) {
        email = `${slug}@mailnesia.com`;
      }

      const newItem: Jamaah = {
        id: 'jam-scan-' + batchTimestamp + '-' + index + '-' + passport.slice(-4),
        name: name,
        passport,
        visa: visa,
        gender: (r.gender === 'Laki-laki' || r.gender === 'Perempuan') ? r.gender : 'Laki-laki',
        phone: '-',
        email: email,
        entryMadinah: '',
        exitMadinah: '',
        operatorId: null,
        status: 'Ready',
        notes: `Di-import via Scan Massal Visa (${r.fileName}).`,
        qrCodeUrl: null,
        qrUploadedAt: null,
        createdAt: new Date().toISOString(),
        travel: r.travel || settingsTravelName,
        password: generatedPassword,
        customValues: r.customValues || {}
      };

      newItems.push(newItem);
      newCount++;
    });

    if (newItems.length > 0) {
      const finalList = [...newItems, ...jamaahs];
      // Update state
      setJamaahs(finalList);

      // Save to IndexedDB
      saveJamaahsToDB(finalList);

      // Save to Supabase
      const supabase = getSupabase();
      if (supabase && isSupabaseConnected) {
        supabase.from('jamaahs').insert(newItems.map(mapJamaahToDb)).then(({ error }) => {
          if (error) console.error('Failed to insert scanned jamaahs to Supabase:', error);
        });
      }
    }

    alert(`Sukses mengimpor jemaah dari berkas visa!\n- Jemaah Baru Berhasil Masuk: ${newCount}\n- Data Duplikat Ditolak (Skip): ${duplicateCount}`);
    
    // Reset batch state
    setShowBatchScanModal(false);
    setBatchScanResults([]);
    setBatchScanErrors([]);
    setBatchScanFiles([]);
  };

  const handleAddCustomField = () => {
    if (!newCustomFieldLabel.trim()) return;
    const label = newCustomFieldLabel.trim();
    if (customFields.some(cf => cf.label.toLowerCase() === label.toLowerCase())) {
      alert('Nama kolom kustom sudah ada!');
      return;
    }
    const newField: CustomField = {
      id: 'cf-' + Date.now(),
      label
    };
    setCustomFields(prev => [...prev, newField]);

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      supabase.from('custom_fields').insert(mapCustomFieldToDb(newField)).then(({ error }) => {
        if (error) console.error('Failed to save custom field to Supabase:', error);
      });
    }

    setSettingsExportColumns(prev => ({
      ...prev,
      [newField.id]: true
    }));
    setNewCustomFieldLabel('');
  };

  const handleRenameCustomField = (id: string) => {
    if (!editingCustomFieldLabel.trim()) return;
    const label = editingCustomFieldLabel.trim();
    if (customFields.some(cf => cf.id !== id && cf.label.toLowerCase() === label.toLowerCase())) {
      alert('Nama kolom kustom sudah digunakan!');
      return;
    }
    setCustomFields(prev => prev.map(cf => cf.id === id ? { ...cf, label } : cf));

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      supabase.from('custom_fields').update({ label }).eq('id', id).then(({ error }) => {
        if (error) console.error('Failed to update custom field in Supabase:', error);
      });
    }

    setEditingCustomFieldId(null);
    setEditingCustomFieldLabel('');
  };

  const handleDeleteCustomField = (id: string) => {
    if (window.confirm('Apakah Anda yakin ingin menghapus kolom kustom ini? Data jemaah yang sudah tersimpan untuk kolom ini tidak akan terhapus dari database, tetapi kolom tidak akan ditampilkan lagi.')) {
      setCustomFields(prev => prev.filter(cf => cf.id !== id));

      const supabase = getSupabase();
      if (supabase && isSupabaseConnected) {
        supabase.from('custom_fields').delete().eq('id', id).then(({ error }) => {
          if (error) console.error('Failed to delete custom field in Supabase:', error);
        });
      }

      setSettingsExportColumns(prev => {
        const copy = { ...prev };
        delete (copy as any)[id];
        return copy;
      });
    }
  };

  const clearAllJamaahs = () => {
    if (window.confirm('Apakah Anda yakin ingin menghapus SELURUH data jemaah aktif yang ada di sistem? Semua data pendaftaran jemaah akan dihapus secara permanen untuk memulai dari nol.')) {
      setJamaahs([]);
      saveJamaahsToDB([]);

      const supabase = getSupabase();
      if (supabase && isSupabaseConnected) {
        supabase.from('jamaahs').delete().neq('id', '').then(({ error }) => {
          if (error) console.error('Failed to clear jamaahs in Supabase:', error);
        });
      }

      alert('Semua data jemaah telah berhasil dibersihkan! Sistem sekarang dalam keadaan kosong dan siap menerima import manifest baru.');
    }
  };

  // Build a default email from a name: lowercase, strip accents & non-alphanumeric, then @mailnesia.com
  const buildDefaultEmail = (name: string) => {
    const local = (name || '')
      .toLowerCase()
      .normalize('NFD')                 // split accented letters into base + combining mark
      .replace(/[^a-z0-9]/g, '');       // strip combining marks, spaces & symbols
    return local ? `${local}@mailnesia.com` : '';
  };

  // Apply default login email + uniform password to ALL jamaah, overwriting.
  const applyDefaultAccounts = () => {
    if (jamaahs.length === 0) {
      alert('Belum ada data jemaah untuk diproses.');
      return;
    }
    const pwd = settingsDefaultPassword || DEFAULT_JAMAAH_PASSWORD;
    if (!window.confirm(
      `Terapkan ke SEMUA ${jamaahs.length} jemaah?\n\n` +
      `• Email login dibuat ulang dari nama (huruf kecil) @mailnesia.com\n` +
      `• Password diseragamkan menjadi "${pwd}"\n\n` +
      `Tindakan ini menimpa email & password yang ada.`
    )) return;

    const updatedList = jamaahs.map(j => {
      const email = buildDefaultEmail(j.name) || j.email;
      const updated = { ...j, email, password: pwd };
      if (selectedJamaah && selectedJamaah.id === j.id) setSelectedJamaah(updated);
      return updated;
    });

    setJamaahs(updatedList);

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      supabase.from('jamaahs').upsert(updatedList.map(mapJamaahToDb)).then(({ error }) => {
        if (error) console.error('Failed to upsert jamaahs in Supabase:', error);
      });
    }

    alert(`Berhasil diterapkan ke ${jamaahs.length} jemaah: email default @mailnesia.com & password seragam "${pwd}".`);
  };

  // Set ONLY the password (global) to all jamaah, using the editable default password
  const applyGlobalPassword = () => {
    const pwd = (settingsDefaultPassword || '').trim();
    if (!pwd) {
      alert('Isi dulu Password Global Jemaah.');
      return;
    }
    if (jamaahs.length === 0) {
      alert('Belum ada data jemaah untuk diproses.');
      return;
    }
    if (!window.confirm(`Set password "${pwd}" ke SEMUA ${jamaahs.length} jemaah? Password lama akan ditimpa.`)) return;

    const updatedList = jamaahs.map(j => {
      const updated = { ...j, password: pwd };
      if (selectedJamaah && selectedJamaah.id === j.id) setSelectedJamaah(updated);
      return updated;
    });

    setJamaahs(updatedList);

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      supabase.from('jamaahs').upsert(updatedList.map(mapJamaahToDb)).then(({ error }) => {
        if (error) console.error('Failed to upsert jamaahs in Supabase:', error);
      });
    }

    alert(`Password "${pwd}" berhasil diterapkan ke ${jamaahs.length} jemaah.`);
  };

  // Rename a travel/rombongan group — updates every jamaah currently in that group
  const renameTravelGroup = (oldName: string) => {
    const input = window.prompt(`Ubah nama Travel / Rombongan "${oldName}" menjadi:`, oldName);
    if (input === null) return; // cancelled
    const newName = input.trim();
    if (!newName || newName === oldName) return;

    const updatedList = jamaahs.map(j => {
      const current = j.travel || settingsTravelName;
      if (current === oldName) {
        const updated = { ...j, travel: newName };
        if (selectedJamaah && selectedJamaah.id === j.id) setSelectedJamaah(updated);
        return updated;
      }
      return j;
    });

    setJamaahs(updatedList);

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      const changedOnly = updatedList.filter(j => j.travel === newName);
      supabase.from('jamaahs').upsert(changedOnly.map(mapJamaahToDb)).then(({ error }) => {
        if (error) console.error('Failed to update travel name in Supabase:', error);
      });
    }

    // Keep active filter / expanded state consistent with the new name
    setTravelFilter(prevFilter => (prevFilter === oldName ? newName : prevFilter));
    setExpandedTravels(prev => {
      if (!(oldName in prev)) return prev;
      const copy = { ...prev };
      copy[newName] = copy[oldName];
      delete copy[oldName];
      return copy;
    });
  };

  const handleMigrateToSupabase = async () => {
    const supabase = getSupabase();
    if (!supabase || !isSupabaseConnected) {
      alert('Supabase tidak terhubung! Silakan isi URL dan Key dan hubungkan terlebih dahulu.');
      return;
    }
    if (!window.confirm('Tindakan ini akan mengunggah seluruh data Jemaah, Operator, dan Kolom Kustom lokal Anda saat ini ke database online Supabase. Data yang sudah ada di Supabase dengan ID yang sama akan diperbarui.')) return;
    
    setIsSupabaseLoading(true);
    try {
      // 1. Migrate settings
      await supabase.from('settings').upsert([
        { key: 'travel_name', value: settingsTravelName },
        { key: 'nusuk_limit', value: String(settingsNusukLimit) },
        { key: 'reference_date', value: settingsReferenceDate },
        { key: 'default_password', value: settingsDefaultPassword },
        { key: 'default_raudhah_slot', value: settingsDefaultRaudhahSlot },
        { key: 'admin_password', value: adminPassword },
      ]);

      // 2. Migrate operators
      if (operators.length > 0) {
        const mappedOps = operators.map(mapOperatorToDb);
        const { error: opErr } = await supabase.from('operators').upsert(mappedOps);
        if (opErr) throw opErr;
      }

      // 3. Migrate custom fields
      if (customFields.length > 0) {
        const mappedCfs = customFields.map(mapCustomFieldToDb);
        const { error: cfErr } = await supabase.from('custom_fields').upsert(mappedCfs);
        if (cfErr) throw cfErr;
      }

      // 4. Migrate jamaahs
      if (jamaahs.length > 0) {
        const mappedJams = jamaahs.map(mapJamaahToDb);
        const { error: jamErr } = await supabase.from('jamaahs').upsert(mappedJams);
        if (jamErr) throw jamErr;
      }

      alert('Migrasi data lokal ke Supabase berhasil! Seluruh data Anda kini tersimpan di database online.');
    } catch (err: any) {
      console.error(err);
      alert('Gagal melakukan migrasi data: ' + (err.message || String(err)));
    } finally {
      setIsSupabaseLoading(false);
    }
  };

  const openResetModal = () => {
    setResetScopes({ jamaah: true, operators: true, settings: true, credentials: true });
    setResetConfirmed(false);
    setShowResetModal(true);
  };

  const handleClearDatabase = () => {
    const { jamaah, operators: opScope, settings, credentials } = resetScopes;

    // 1. Data Jamaah (IndexedDB)
    if (jamaah) {
      setJamaahs(INITIAL_JAMAAH);
      saveJamaahsToDB(INITIAL_JAMAAH);
      setSelectedJamaah(null);
      setSelectedJamaahIds({});
      setSearchQuery('');
      setStatusFilter('All');
      setOperatorFilter('All');
      setGenderFilter('All');
      setTravelFilter('All');
      setEntryFilter('');
      setExpandedTravels({});
    }

    // 2. Data Operator
    if (opScope) {
      localStorage.removeItem('raudhah_operators');
      setOperators(INITIAL_OPERATORS);
      setActiveOperatorId(null);
    }

    // 3. Pengaturan aplikasi (travel, kuota, tanggal, slot, kolom ekspor, field kustom)
    if (settings) {
      localStorage.removeItem('raudhah_travel_name');
      localStorage.removeItem('raudhah_nusuk_limit');
      localStorage.removeItem('raudhah_qr_lead_hours');
      localStorage.removeItem('raudhah_reference_date');
      localStorage.removeItem('raudhah_export_columns');
      localStorage.removeItem('raudhah_custom_fields');
      localStorage.removeItem('raudhah_default_password');
      localStorage.removeItem('raudhah_default_raudhah_slot');
      localStorage.removeItem('raudhah_enable_sound');

      setSettingsDefaultPassword(DEFAULT_JAMAAH_PASSWORD);
      setSettingsDefaultRaudhahSlot('');
      setSettingsEnableSound(true);
      setCustomFields([]);
      setSettingsExportColumns({
        name: true,
        passport: true,
        visa: true,
        gender: true,
        phone: true,
        email: true,
        entryMadinah: true,
        exitMadinah: true,
        travel: true,
        password: true,
        status: true,
        operator: true,
        notes: true,
      });
      setSettingsTravelName('Raudhah Al-Haramain Travel');
      setSettingsNusukLimit(15);
      setSettingsQrLeadHours(2);
      const today = new Date();
      const yyyy = today.getFullYear();
      const mm = String(today.getMonth() + 1).padStart(2, '0');
      const dd = String(today.getDate()).padStart(2, '0');
      setSettingsReferenceDate(`${yyyy}-${mm}-${dd}`);
    }

    // 4. Kredensial & API (password admin, API key Gemini)
    if (credentials) {
      localStorage.removeItem('raudhah_admin_password');
      localStorage.removeItem('raudhah_gemini_api_key');
      localStorage.removeItem('raudhah_gemini_model');
      setAdminPassword('admin123');
      setSettingsGeminiApiKey('');
      setSettingsGeminiModel('gemini-2.0-flash');
    }

    const supabase = getSupabase();
    if (supabase && isSupabaseConnected) {
      const runSupabaseReset = async () => {
        if (jamaah) {
          await supabase.from('jamaahs').delete().neq('id', '');
          await supabase.from('jamaahs').insert(INITIAL_JAMAAH.map(mapJamaahToDb));
        }
        if (opScope) {
          await supabase.from('operators').delete().neq('id', '');
          await supabase.from('operators').insert(INITIAL_OPERATORS.map(mapOperatorToDb));
        }
        if (settings) {
          await supabase.from('settings').delete().neq('key', '');
        }
      };
      runSupabaseReset().catch(e => console.error('Failed to reset Supabase:', e));
    }

    setShowResetModal(false);
    setResetConfirmed(false);

    const labels: string[] = [];
    if (jamaah) labels.push('Data Jamaah');
    if (opScope) labels.push('Data Operator');
    if (settings) labels.push('Pengaturan Aplikasi');
    if (credentials) labels.push('Kredensial & API');
    alert(`Pembersihan selesai. Bagian yang disetel ulang: ${labels.join(', ')}.`);
  };

  // Simulated automatic official Raudhah barcode generator
  return (
    <div className="min-h-screen bg-slate-50 dark:bg-zinc-800/50 dark:bg-zinc-950 flex flex-col md:flex-row text-slate-900 dark:text-zinc-100 font-sans antialiased selection:bg-red-100 selection:text-red-900 overflow-x-hidden">
      
      {/* Mobile Sidebar Backdrop */}
      {isMobileSidebarOpen && (
        <div
          className="fixed inset-0 z-40 bg-zinc-950/60 md:hidden transition-opacity duration-300"
          onClick={() => setIsMobileSidebarOpen(false)}
        />
      )}

      {/* SIDEBAR */}
      <aside className={`fixed inset-y-0 left-0 z-50 w-64 bg-zinc-900 text-zinc-100 flex flex-col border-r border-zinc-800 shrink-0 transform transition-transform duration-300 ease-in-out md:relative md:translate-x-0 ${
        isMobileSidebarOpen ? 'translate-x-0' : '-translate-x-full'
      }`}>
        <div className="p-6 border-b border-zinc-800">
          <div className="flex items-center gap-3">
            <div className="h-10 flex items-center shrink-0">
              <img src="/logo.png" alt="24 Visa Logo" className="h-8 object-contain bg-white px-2 py-0.5 rounded-lg shadow-xs" />
            </div>
            <div className="min-w-0">
              <h1 className="font-sans font-bold text-xs tracking-tight text-white leading-none truncate">Raudhah Barcode</h1>
              <p className="text-[9px] text-zinc-400 font-mono tracking-widest mt-0.5 uppercase">Manager</p>
            </div>
          </div>
        </div>

        {/* Navigation */}
        <nav className="flex-1 p-4 space-y-1">
          <button
            onClick={() => { setCurrentTab('dashboard'); setSelectedJamaah(null); setIsMobileSidebarOpen(false); }}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
              currentTab === 'dashboard' ? 'bg-zinc-800 text-white' : 'text-zinc-400 hover:text-white hover:bg-zinc-800/50'
            }`}
          >
            <LayoutDashboard className="w-4 h-4 shrink-0 text-zinc-400" />
            <span>Dashboard</span>
          </button>

          <button
            onClick={() => { setCurrentTab('jamaah'); setSelectedJamaah(null); setIsMobileSidebarOpen(false); }}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
              currentTab === 'jamaah' ? 'bg-zinc-800 text-white' : 'text-zinc-400 hover:text-white hover:bg-zinc-800/50'
            }`}
          >
            <Users className="w-4 h-4 shrink-0 text-zinc-400" />
            <span>Data Jamaah</span>
            <span className="ml-auto text-[10px] px-1.5 py-0.5 bg-zinc-800 text-zinc-400 rounded-md font-mono">
              {activeOperatorId === null ? jamaahs.length : jamaahs.filter(j => j.operatorId === activeOperatorId).length}
            </span>
          </button>

          {activeOperatorId === null && (
            <>
              <button
                onClick={() => { setCurrentTab('import'); setSelectedJamaah(null); setIsMobileSidebarOpen(false); }}
                className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
                  currentTab === 'import' ? 'bg-zinc-800 text-white' : 'text-zinc-400 hover:text-white hover:bg-zinc-800/50'
                }`}
              >
                <FileSpreadsheet className="w-4 h-4 shrink-0 text-zinc-400" />
                <span>Import Excel</span>
                <span className="ml-auto w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
              </button>

              <button
                onClick={() => { setCurrentTab('operator'); setSelectedJamaah(null); setIsMobileSidebarOpen(false); }}
                className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
                  currentTab === 'operator' ? 'bg-zinc-800 text-white' : 'text-zinc-400 hover:text-white hover:bg-zinc-800/50'
                }`}
              >
                <UserCheck className="w-4 h-4 shrink-0 text-zinc-400" />
                <span>Operator</span>
              </button>

              <button
                onClick={() => { setCurrentTab('settings'); setSelectedJamaah(null); setIsMobileSidebarOpen(false); }}
                className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
                  currentTab === 'settings' ? 'bg-zinc-800 text-white' : 'text-zinc-400 hover:text-white hover:bg-zinc-800/50'
                }`}
              >
                <Settings className="w-4 h-4 shrink-0 text-zinc-400" />
                <span>Settings</span>
              </button>
            </>
          )}

          <button
            onClick={() => { setCurrentTab('guide'); setSelectedJamaah(null); setIsMobileSidebarOpen(false); }}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
              currentTab === 'guide' ? 'bg-zinc-800 text-white' : 'text-zinc-400 hover:text-white hover:bg-zinc-800/50'
            }`}
          >
            <BookOpen className="w-4 h-4 shrink-0 text-zinc-400" />
            <span>Panduan Sistem</span>
          </button>
        </nav>

        {/* Active User & Logout */}
        <div className="p-4 border-t border-zinc-800 bg-zinc-950/40 space-y-3 shrink-0">
          {/* Dark Mode Toggle */}
          <button
            onClick={() => setIsDarkMode(!isDarkMode)}
            className="w-full flex items-center justify-between px-3 py-2 rounded-lg bg-zinc-800/60 hover:bg-zinc-700/60 border border-zinc-700/40 text-xs font-medium text-zinc-300 transition-all cursor-pointer"
          >
            <span className="flex items-center gap-2">
              {isDarkMode ? <Moon className="w-3.5 h-3.5 text-blue-400" /> : <Sun className="w-3.5 h-3.5 text-amber-400" />}
              <span>{isDarkMode ? 'Mode Gelap' : 'Mode Terang'}</span>
            </span>
            <span className="text-[9px] font-mono text-zinc-500">{isDarkMode ? '🌙' : '☀️'}</span>
          </button>

          <div>
            <span className="block text-[9px] font-semibold text-zinc-500 uppercase tracking-wider">
              Akun Aktif
            </span>
            <span className="text-xs font-semibold text-white mt-1 flex items-center gap-1.5 truncate">
              {activeOperatorId === null ? (
                <>
                  <Building2 className="w-3.5 h-3.5 text-zinc-400 shrink-0" />
                  <span>Kantor Pusat (Admin)</span>
                </>
              ) : (
                <>
                  <Laptop className="w-3.5 h-3.5 text-zinc-400 shrink-0" />
                  <span>{operators.find(o => o.id === activeOperatorId)?.name}</span>
                </>
              )}
            </span>
          </div>
          <button
            onClick={() => {
              sessionStorage.removeItem('raudhah_is_logged_in');
              sessionStorage.removeItem('raudhah_active_operator_id');
              setIsLoggedIn(false);
              setActiveOperatorId(null);
              setCurrentTab('dashboard');
            }}
            className="w-full py-1.5 bg-zinc-800 hover:bg-red-900/40 hover:text-red-200 border border-zinc-700/60 rounded text-center text-xs font-semibold transition-all text-zinc-300 cursor-pointer"
          >
            Keluar Akun (Logout)
          </button>
        </div>
      </aside>

      {/* MAIN CONTAINER */}
      <main className="flex-1 flex flex-col min-w-0">
        
        {/* TOP BAR */}
        <header className="h-16 bg-white dark:bg-zinc-900 border-b border-slate-100 dark:border-zinc-800 px-4 md:px-6 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            {/* Hamburger Button for Mobile */}
            <button
              onClick={() => setIsMobileSidebarOpen(true)}
              className="p-2 -ml-2 rounded-lg text-slate-500 dark:text-zinc-400 hover:bg-slate-100 dark:bg-zinc-700 block md:hidden cursor-pointer"
            >
              <Menu className="w-5 h-5" />
            </button>

            <h2 className="text-base font-semibold text-slate-800 dark:text-zinc-100 tracking-tight capitalize">
              {currentTab}
            </h2>
            <span className="text-slate-300">/</span>
            <div className="flex items-center gap-2">
              <span className="text-[10px] sm:text-xs font-medium px-2 py-0.5 rounded-full bg-slate-100 dark:bg-zinc-700 text-slate-600 dark:text-zinc-300 truncate max-w-[120px] sm:max-w-none flex items-center gap-1">
                {activeOperatorId === null ? (
                  <>
                    <Building2 className="w-3 h-3 text-slate-500 dark:text-zinc-400 shrink-0" />
                    <span>Kantor Pusat (Admin)</span>
                  </>
                ) : (
                  <>
                    <Laptop className="w-3 h-3 text-slate-500 dark:text-zinc-400 shrink-0" />
                    <span>Operator: {operators.find(o => o.id === activeOperatorId)?.name}</span>
                  </>
                )}
              </span>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="hidden sm:flex items-center gap-1.5 text-slate-600 dark:text-zinc-300 text-xs font-medium bg-slate-100 dark:bg-zinc-700 px-3 py-1 rounded-full border border-slate-200 dark:border-zinc-600/40">
              <Calendar className="w-3.5 h-3.5 text-emerald-600 animate-pulse" />
              <span>Tanggal Acuan: <strong className="text-slate-800 dark:text-zinc-100 font-semibold">{settingsReferenceDate}</strong></span>
            </div>
            
            <div className="flex items-center gap-2">
              {activeOperatorId === null && (
                <>
                  <button
                    onClick={openAddModal}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-zinc-900 text-white text-xs font-medium hover:bg-zinc-800 transition-all shadow-xs shrink-0"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Tambah Manual</span>
                  </button>

                  <button
                    onClick={() => { setShowBatchScanModal(true); setBatchScanFiles([]); setBatchScanResults([]); setBatchScanErrors([]); }}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-gradient-to-r from-blue-600 to-indigo-600 text-white text-xs font-semibold hover:from-blue-700 hover:to-indigo-700 transition-all shadow-xs shrink-0"
                  >
                    <Sparkles className="w-3.5 h-3.5" />
                    <span>Scan Massal Visa (PDF)</span>
                  </button>

                  {duplicateIds.size > 0 && (
                    <button
                      onClick={() => setShowDuplicateModal(true)}
                      title="Ada data jemaah dengan paspor/visa sama"
                      className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-amber-500 hover:bg-amber-600 text-white text-xs font-semibold transition-all shadow-xs shrink-0 animate-pulse"
                    >
                      <AlertTriangle className="w-3.5 h-3.5" />
                      <span>Duplikat ({duplicateGroups.length})</span>
                    </button>
                  )}
                </>
              )}
            </div>
          </div>
        </header>

        {/* CONTENT AREA */}
        <div className="flex-1 overflow-y-auto p-6 space-y-6 dark:text-zinc-200">

          {/* TAB 1: DASHBOARD */}
          {currentTab === 'dashboard' && (
            <div className="space-y-6">
              
              {/* OPERATOR HEADER FOR SPECIFIC ROLE */}
              {activeOperatorId !== null && (
                <div className="bg-emerald-50 border border-emerald-100 rounded-xl p-5 flex flex-col md:flex-row items-start md:items-center justify-between gap-4 dark:bg-emerald-950/20 dark:border-emerald-800/30 select-none">
                  <div>
                    <h3 className="font-semibold text-emerald-900 dark:text-emerald-200 text-base">Assalamualaikum, {operators.find(o => o.id === activeOperatorId)?.name}!</h3>
                    <p className="text-xs text-emerald-700/90 dark:text-emerald-400 mt-0.5">Berikut adalah ringkasan tanggung jawab Anda hari ini untuk penanganan Nusuk.</p>
                  </div>
                  <div className="flex items-center gap-2 bg-white/80 dark:bg-zinc-800/80 border border-emerald-200/50 dark:border-emerald-800/50 rounded-lg px-4 py-2 text-xs text-emerald-800 dark:text-emerald-300 font-medium">
                    <span>Hari ini saya memiliki:</span>
                    <strong className="text-slate-800 dark:text-zinc-200">{countOperatorStatus(activeOperatorId, 'Ready')} Ready</strong>
                    <span className="text-emerald-200 dark:text-emerald-800">|</span>
                    <strong className="text-amber-700">{countOperatorStatus(activeOperatorId, 'Sedang War')} War</strong>
                    <span className="text-emerald-200 dark:text-emerald-800">|</span>
                    <strong className="text-emerald-700 dark:text-emerald-400">{countOperatorStatus(activeOperatorId, 'QR Berhasil')} Sukses</strong>
                  </div>
                </div>
              )}

              {/* STATS CARDS */}
              <div className="grid grid-cols-2 lg:grid-cols-5 gap-4">
                
                {/* Stat 1 */}
                <div className="bg-white dark:bg-zinc-800 p-5 rounded-xl border border-slate-100 dark:border-zinc-700 shadow-xs flex flex-col justify-between">
                  <span className="text-xs font-medium text-slate-500 dark:text-zinc-400">Total Jamaah</span>
                  <div className="flex items-baseline justify-between mt-2">
                    <span className="text-2xl font-bold tracking-tight text-slate-800 dark:text-zinc-100">
                      {activeOperatorId ? jamaahs.filter(j => j.operatorId === activeOperatorId).length : jamaahs.length}
                    </span>
                    <span className="text-[10px] font-mono bg-slate-50 dark:bg-zinc-800/50 text-slate-500 dark:text-zinc-400 px-1.5 py-0.5 rounded border border-slate-100">Pax</span>
                  </div>
                </div>

                {/* Stat 2 - Ready */}
                <div className="bg-white dark:bg-zinc-800 p-5 rounded-xl border border-slate-100 dark:border-zinc-700 shadow-xs flex flex-col justify-between border-l-4 border-l-slate-400">
                  <span className="text-xs font-medium text-slate-500 dark:text-zinc-400 flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-slate-400"></span>
                    <span>Ready</span>
                  </span>
                  <div className="flex items-baseline justify-between mt-2">
                    <span className="text-2xl font-bold tracking-tight text-slate-800 dark:text-zinc-100">
                      {activeOperatorId ? countOperatorStatus(activeOperatorId, 'Ready') : countStatus('Ready')}
                    </span>
                    <span className="text-[10px] text-slate-500 dark:text-zinc-400 font-mono">Antrean</span>
                  </div>
                </div>

                {/* Stat 3 - Sedang War */}
                <div className="bg-white dark:bg-zinc-800 p-5 rounded-xl border border-slate-100 dark:border-zinc-700 shadow-xs flex flex-col justify-between border-l-4 border-l-amber-400">
                  <span className="text-xs font-medium text-slate-500 dark:text-zinc-400 flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-amber-500 animate-pulse"></span>
                    <span>Sedang War</span>
                  </span>
                  <div className="flex items-baseline justify-between mt-2">
                    <span className="text-2xl font-bold tracking-tight text-slate-800 dark:text-zinc-100">
                      {activeOperatorId ? countOperatorStatus(activeOperatorId, 'Sedang War') : countStatus('Sedang War')}
                    </span>
                    <span className="text-[10px] text-amber-600 font-mono">War Nusuk</span>
                  </div>
                </div>

                {/* Stat 4 - QR Berhasil */}
                <div className="bg-white dark:bg-zinc-800 p-5 rounded-xl border border-slate-100 dark:border-zinc-700 shadow-xs flex flex-col justify-between border-l-4 border-l-emerald-500">
                  <span className="text-xs font-medium text-slate-500 dark:text-zinc-400 flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-500"></span>
                    <span>QR Berhasil</span>
                  </span>
                  <div className="flex items-baseline justify-between mt-2">
                    <span className="text-2xl font-bold tracking-tight text-emerald-700 dark:text-emerald-400">
                      {activeOperatorId ? countOperatorStatus(activeOperatorId, 'QR Berhasil') : countStatus('QR Berhasil')}
                    </span>
                    <span className="text-[10px] text-emerald-600 font-mono">Sukses</span>
                  </div>
                </div>

                {/* Stat 5 - Belum Berhasil */}
                <div className="bg-white dark:bg-zinc-800 p-5 rounded-xl border border-slate-100 dark:border-zinc-700 shadow-xs flex flex-col justify-between border-l-4 border-l-rose-500">
                  <span className="text-xs font-medium text-slate-500 dark:text-zinc-400 flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-rose-500"></span>
                    <span>Belum Berhasil</span>
                  </span>
                  <div className="flex items-baseline justify-between mt-2">
                    <span className="text-2xl font-bold tracking-tight text-rose-600 dark:text-rose-400">
                      {activeOperatorId ? countOperatorStatus(activeOperatorId, 'Belum Berhasil') : countStatus('Belum Berhasil')}
                    </span>
                    <span className="text-[10px] text-rose-500 font-mono">Tertunda</span>
                  </div>
                </div>

              </div>

              {/* TWO COLUMN SUMMARY */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                
                {/* Left: Priority Action Items (Main Focus) */}
                <div className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl shadow-xs p-5 lg:col-span-8 flex flex-col">
                  <div className="flex items-center justify-between pb-4 border-b border-slate-100">
                    <div>
                      <h3 className="font-semibold text-slate-800 dark:text-zinc-100 text-sm">Prioritas Tindakan Tertinggi</h3>
                      <p className="text-xs text-slate-500 dark:text-zinc-400 mt-0.5">Jamaah dengan jadwal masuk Madinah terdekat yang belum memiliki QR Code.</p>
                    </div>
                    <span className="text-xs font-mono font-medium px-2 py-0.5 rounded bg-red-50 text-red-700 border border-red-100">Urut Prioritas Terdekat</span>
                  </div>

                  {highPriorityList.length === 0 ? (
                    <div className="flex-1 flex flex-col items-center justify-center py-12 text-center">
                      <div className="w-12 h-12 rounded-full bg-blue-50 text-red-600 flex items-center justify-center mb-3">
                        <Check className="w-6 h-6" />
                      </div>
                      <h4 className="font-semibold text-slate-800 dark:text-zinc-100 text-sm">Semua Aman Terkendali!</h4>
                      <p className="text-xs text-slate-500 dark:text-zinc-400 mt-1 max-w-sm">Seluruh jamaah prioritas terdekat sudah memiliki QR Code Raudhah.</p>
                    </div>
                  ) : (
                    <div className="flex-1 overflow-y-auto max-h-[500px] pr-1 space-y-6 mt-3">
                      {(() => {
                        // Group by travel
                        const grouped = highPriorityList.reduce<Record<string, Jamaah[]>>((acc, j) => {
                          const trv = j.travel || 'Lainnya';
                          if (!acc[trv]) acc[trv] = [];
                          acc[trv].push(j);
                          return acc;
                        }, {});

                        // Sort travel groups by earliest entry date
                        const sortedGroups = (Object.entries(grouped) as [string, Jamaah[]][]).sort((a, b) => {
                          const minA = a[1].reduce((min, item) => item.entryMadinah < min ? item.entryMadinah : min, '9999-99-99');
                          const minB = b[1].reduce((min, item) => item.entryMadinah < min ? item.entryMadinah : min, '9999-99-99');
                          return minA.localeCompare(minB);
                        });

                        return sortedGroups.map(([travelName, list]) => {
                          // Find earliest entry & priority level for the travel group display
                          const earliestJamaah = list.reduce((earliest, item) => item.entryMadinah < earliest.entryMadinah ? item : earliest, list[0]);
                          const groupPrio = getPriorityInfo(earliestJamaah.entryMadinah, settingsReferenceDate);

                          return (
                            <div key={travelName} className="space-y-2.5">
                              {/* Travel Group Section Header */}
                              <div className="flex items-center justify-between bg-slate-100 dark:bg-zinc-700 px-3 py-1.5 rounded-lg border border-slate-200 dark:border-zinc-600/50">
                                <div className="flex items-center gap-2">
                                  <span className="w-2 h-2 rounded-full bg-red-600"></span>
                                  <h4 className="text-xs font-bold text-slate-800 dark:text-zinc-100">{travelName}</h4>
                                  <span className="text-[10px] bg-emerald-50 text-emerald-700 border border-emerald-100 px-1.5 py-0.2 rounded font-semibold select-none">
                                    {list.length} Jamaah Prioritas
                                  </span>
                                </div>
                                <span className={`text-[9px] font-bold border px-2 py-0.5 rounded-full ${groupPrio.badgeColor}`}>
                                  {groupPrio.level} (H{groupPrio.daysRemaining >= 0 ? `-${groupPrio.daysRemaining}` : `+${Math.abs(groupPrio.daysRemaining)}`})
                                </span>
                              </div>

                              {/* Table inside Travel Group on Dashboard */}
                              <div className="overflow-x-auto border border-slate-200 dark:border-zinc-700 rounded-lg shadow-3xs bg-white dark:bg-zinc-800">
                                <table className="w-full text-left border-collapse text-[11px]">
                                  <thead>
                                    <tr className="bg-slate-50 dark:bg-zinc-800/50 dark:bg-zinc-700 border-b border-slate-200 dark:border-zinc-600 dark:border-zinc-600 text-slate-500 dark:text-zinc-400 font-medium select-none">
                                      <th className="py-2 px-3">Nama Jamaah</th>
                                      <th className="py-2 px-3">Nomor Paspor</th>
                                      <th className="py-2 px-3">Tanggal Masuk (GMT+3)</th>
                                      <th className="py-2 px-3 text-center">Status</th>
                                      <th className="py-2 px-3">Operator</th>
                                      <th className="py-2 px-3 text-right">Aksi</th>
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-slate-100 dark:divide-zinc-700 text-slate-700 dark:text-zinc-200 dark:text-zinc-200">
                                    {list.map(j => {
                                      const isUrgent = (() => {
                                        if (!j.raudhahSlot || j.status === 'QR Berhasil') return false;
                                        const dist = getDistributionInstant(j.raudhahSlot, settingsQrLeadHours);
                                        const slot = getDistributionInstant(j.raudhahSlot, 0);
                                        if (!dist || !slot) return false;
                                        const t = now.getTime();
                                        return (slot.getTime() - t) > 0 && (dist.getTime() - t) / 60000 <= 60;
                                      })();

                                      return (
                                        <tr key={j.id} className={`transition-colors cursor-pointer ${
                                          isUrgent 
                                            ? 'bg-rose-50/80 hover:bg-rose-100/70 border-l-2 border-l-rose-600 dark:bg-rose-950/30 dark:hover:bg-rose-900/40 dark:border-l-rose-800' 
                                            : 'hover:bg-slate-50 dark:bg-zinc-800/50 dark:hover:bg-zinc-700/30'
                                        }`} onClick={() => setSelectedJamaah(j)}>
                                          <td className="py-2 px-3 font-semibold text-slate-800 dark:text-zinc-100">
                                            <div className="flex items-center gap-1.5 flex-wrap">
                                              <span>{j.name}</span>
                                              {isUrgent && (
                                                <span className="text-[8.5px] text-rose-800 dark:text-rose-300 bg-rose-100 dark:bg-rose-500/15 border border-rose-200 dark:border-rose-500/25 px-1.5 py-0.5 rounded font-bold animate-pulse flex items-center gap-0.5">
                                                  <AlertTriangle className="w-3 h-3 text-rose-600 shrink-0" />
                                                  <span>Raudhah &lt; 1 Jam!</span>
                                                </span>
                                              )}
                                              {!isUrgent && j.status !== 'QR Berhasil' && (
                                                <span className="text-[8px] text-rose-700 bg-rose-50 border border-rose-100 px-1 py-0.2 rounded font-bold animate-pulse flex items-center gap-0.5">
                                                  <AlertTriangle className="w-2.5 h-2.5 text-rose-600" />
                                                  <span>Belum QR</span>
                                                </span>
                                              )}
                                            </div>
                                          </td>
                                          <td className="py-2 px-3">
                                            <span className="font-mono font-bold text-slate-700 dark:text-zinc-200 text-[10px] bg-slate-100 dark:bg-zinc-700 px-1.5 py-0.5 rounded border border-slate-200 dark:border-zinc-600">{j.passport}</span>
                                          </td>
                                          <td className="py-2 px-3">
                                            <div className="space-y-0.5">
                                              <div className="text-slate-700 dark:text-zinc-200 font-medium">{formatDateLabel(j.entryMadinah)}</div>
                                              {j.raudhahSlot && (() => {
                                                const r = getQrReminder(j.raudhahSlot, j.status, now.getTime(), settingsQrLeadHours);
                                                if (!r) return null;
                                                return (
                                                  <div className="space-y-0.5 mt-0.5">
                                                    <div className="text-[9px] text-sky-800 dark:text-sky-300 bg-sky-50/70 dark:bg-sky-900/15 border border-sky-100 dark:border-sky-900/40 rounded px-1.5 py-0.5 inline-flex items-center gap-0.5 select-none">
                                                      <Download className="w-2.5 h-2.5 text-sky-600" />
                                                      <span>QR tersedia: {formatTimeColon(r.dist, TZ_WITA)} WITA</span>
                                                    </div>
                                                    <div className={`text-[9px] font-bold border rounded px-1.5 py-0.5 inline-flex items-center gap-0.5 select-none ${QR_REMINDER_BADGE[r.tone]}`}>
                                                      <Bell className="w-2.5 h-2.5" />
                                                      <span>{r.countdownLabel}</span>
                                                    </div>
                                                    <div className="text-[9px] text-slate-500 dark:text-zinc-400 flex items-center gap-0.5">
                                                      <Clock className="w-2.5 h-2.5 text-emerald-600" />
                                                      <span>Slot Raudhah: {formatTimeColon(r.slot, TZ_WITA)} WITA · {formatTimeColon(r.slot, TZ_MADINAH)} Mad</span>
                                                    </div>
                                                  </div>
                                                );
                                              })()}
                                            </div>
                                          </td>
                                          <td className="py-2 px-3 text-center" onClick={(e) => e.stopPropagation()}>
                                            <select
                                              value={j.status}
                                              onChange={(e) => handleQuickStatusChange(j.id, e.target.value as JamaahStatus)}
                                              className={`text-[10px] rounded border py-0.5 px-1 font-semibold shadow-3xs outline-hidden focus:ring-1 focus:ring-emerald-500 cursor-pointer ${
                                                j.status === 'Ready' ? 'bg-slate-100 dark:bg-zinc-700 border-slate-200 dark:border-zinc-600 text-slate-700 dark:text-zinc-200' :
                                                j.status === 'Sedang War' ? 'bg-amber-50 border-amber-200 text-amber-700' :
                                                j.status === 'QR Berhasil' ? 'bg-emerald-50 border-emerald-200 text-emerald-700' :
                                                'bg-rose-50 border-rose-200 text-rose-700'
                                              }`}
                                            >
                                              <option value="Ready">Ready</option>
                                              <option value="Sedang War">Sedang War</option>
                                              <option value="QR Berhasil">QR Berhasil</option>
                                              <option value="Belum Berhasil">Belum Berhasil</option>
                                            </select>
                                          </td>
                                          <td className="py-2 px-3" onClick={(e) => e.stopPropagation()}>
                                            {activeOperatorId === null ? (
                                              <select
                                                value={j.operatorId || ''}
                                                onChange={(e) => {
                                                  const opId = e.target.value || null;
                                                  setJamaahs(prev =>
                                                    prev.map(item => item.id === j.id ? { ...item, operatorId: opId } : item)
                                                  );
                                                  const supabase = getSupabase();
                                                  if (supabase && isSupabaseConnected) {
                                                    supabase.from('jamaahs').update({ operator_id: opId }).eq('id', j.id).then(({ error }) => {
                                                      if (error) console.error('Failed to update operator in Supabase:', error);
                                                    });
                                                  }
                                                }}
                                                className={`text-[10px] rounded border p-0.5 outline-hidden focus:border-red-500 font-semibold cursor-pointer ${
                                                  !j.operatorId ? 'bg-red-50 dark:bg-red-500/10 border-red-200 dark:border-red-500/30 text-red-700 dark:text-red-400' : 'bg-white dark:bg-zinc-700 border-slate-200 dark:border-zinc-600 text-slate-700 dark:text-zinc-200'
                                                }`}
                                              >
                                                <option value="" className="text-red-700 dark:text-red-400 font-semibold bg-white dark:bg-zinc-800">Belum Ditugaskan</option>
                                                {operators.filter(o => o.isActive).map(o => (
                                                  <option key={o.id} value={o.id} className="text-slate-800 dark:text-zinc-100 font-normal bg-white dark:bg-zinc-800">{o.name}</option>
                                                ))}
                                              </select>
                                            ) : (
                                              <span className={`text-[10px] font-semibold ${!j.operatorId ? 'text-red-400' : 'text-zinc-200'}`}>
                                                {j.operatorId ? operators.find(o => o.id === j.operatorId)?.name || '-' : '-'}
                                              </span>
                                            )}
                                          </td>
                                          <td className="py-2 px-3 text-right" onClick={(e) => e.stopPropagation()}>
                                            <button
                                              onClick={() => setSelectedJamaah(j)}
                                              className="text-[10px] font-bold text-red-600 hover:text-red-700 hover:bg-red-50 px-2 py-0.5 rounded transition-colors inline-flex items-center gap-0.5 border border-red-600/20 hover:border-red-600 cursor-pointer"
                                            >
                                              <span>Detail</span>
                                              <ChevronRight className="w-2.5 h-2.5" />
                                            </button>
                                          </td>
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>
                              </div>
                            </div>
                          );
                        });
                      })()}
                    </div>
                  )}
                </div>

                {/* Right: Quick Stats & Operasional Tips */}
                <div className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl shadow-xs p-5 lg:col-span-4 space-y-5">
                  <div>
                    <h3 className="font-semibold text-slate-800 dark:text-zinc-100 text-sm">Informasi & Panduan Nusuk</h3>
                    <p className="text-xs text-slate-500 dark:text-zinc-400 mt-0.5">Ringkasan rilis slot dan penanganan darurat.</p>
                  </div>

                  <div className="space-y-3.5">
                    <div className="p-3 bg-zinc-50 dark:bg-zinc-700/40 border border-zinc-100 dark:border-zinc-700 rounded-lg">
                      <div className="flex items-center gap-2">
                        <Clock className="w-4 h-4 text-emerald-600" />
                        <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Waktu Sekarang (Sinkron)</span>
                      </div>
                      <div className="grid grid-cols-2 gap-2 mt-2">
                        <div className="bg-white dark:bg-zinc-800 border border-zinc-100 dark:border-zinc-700 rounded-md p-2 text-center">
                          <div className="text-[10px] font-semibold text-emerald-700 dark:text-emerald-400 uppercase tracking-wide">Madinah · GMT+3</div>
                          <div className="font-mono font-bold text-base text-slate-800 dark:text-zinc-100 tabular-nums">
                            {now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false, timeZone: 'Asia/Riyadh' })}
                          </div>
                          <div className="text-[9px] text-slate-400 dark:text-zinc-500">
                            {now.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', timeZone: 'Asia/Riyadh' })}
                          </div>
                        </div>
                        <div className="bg-white dark:bg-zinc-800 border border-zinc-100 dark:border-zinc-700 rounded-md p-2 text-center">
                          <div className="text-[10px] font-semibold text-sky-700 dark:text-sky-400 uppercase tracking-wide">Indonesia · WITA</div>
                          <div className="font-mono font-bold text-base text-slate-800 dark:text-zinc-100 tabular-nums">
                            {now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false, timeZone: 'Asia/Makassar' })}
                          </div>
                          <div className="text-[9px] text-slate-400 dark:text-zinc-500">
                            {now.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', timeZone: 'Asia/Makassar' })}
                          </div>
                        </div>
                      </div>
                      <p className="text-[10px] text-slate-500 dark:text-zinc-400 mt-2 leading-relaxed">Selisih tetap <strong>5 jam</strong> (WITA lebih cepat). Raudhah pakai waktu Madinah, war slot Nusuk ikut waktu Indonesia — pantau keduanya sekaligus. Booking manual tepat setelah rilis slot umum.</p>
                    </div>

                    <div className="p-3 bg-amber-50/50 dark:bg-amber-900/15 border border-amber-100 dark:border-amber-900/40 rounded-lg">
                      <div className="flex items-center gap-2">
                        <AlertTriangle className="w-4 h-4 text-amber-600" />
                        <span className="text-xs font-semibold text-amber-900 dark:text-amber-300">Pembatasan Kuota</span>
                      </div>
                      <p className="text-xs text-amber-800/90 dark:text-amber-200/80 mt-1">Rata-rata limit harian Nusuk: <strong>{settingsNusukLimit} Jamaah</strong> per akun operator. Pastikan membagi merata jamaah ke operator aktif.</p>
                    </div>

                    <div className="pt-2">
                      <h4 className="text-[11px] font-semibold text-slate-500 dark:text-zinc-400 uppercase tracking-wider mb-2">Statistik Penugasan</h4>
                      <div className="space-y-1.5">
                        {operators.map(op => {
                          const assigned = jamaahs.filter(j => j.operatorId === op.id);
                          const successCount = assigned.filter(j => j.status === 'QR Berhasil').length;
                          const progressPercent = assigned.length > 0 ? Math.round((successCount / assigned.length) * 100) : 0;
                          return (
                            <div key={op.id} className="text-xs flex items-center justify-between py-1 border-b border-slate-50 last:border-0">
                              <span className="text-slate-600 dark:text-zinc-300 font-medium">{op.name}</span>
                              <div className="flex items-center gap-2">
                                <span className="font-mono text-slate-500 dark:text-zinc-400">{successCount}/{assigned.length}</span>
                                <span className={`text-[10px] font-semibold px-1 py-0.5 rounded ${op.isActive ? 'bg-emerald-50 text-emerald-700' : 'bg-slate-100 dark:bg-zinc-700 text-slate-500 dark:text-zinc-400'}`}>
                                  {progressPercent}%
                                </span>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                </div>

              </div>

            </div>
          )}

          {/* TAB 2: DATA JAMAAH */}
          {currentTab === 'jamaah' && (
            <div className="space-y-4">
              
              {/* SEARCH & FILTER MODULE */}
              <div className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl shadow-xs p-4 space-y-3.5">
                
                {/* Line 1: Search and simple tab toggles */}
                <div className="flex flex-col md:flex-row gap-3">
                  <div className="relative flex-1">
                    <Search className="absolute left-3 top-2.5 h-4 w-4 text-slate-400" />
                    <input
                      type="text"
                      placeholder="Cari berdasarkan Nama Jamaah, Nomor Paspor, atau Nomor Visa..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="w-full pl-9 pr-4 py-2 text-sm border border-slate-200 dark:border-zinc-600 rounded-lg outline-hidden focus:border-red-500 focus:ring-1 focus:ring-red-500/50 placeholder-slate-400 bg-slate-50 dark:bg-zinc-800/50"
                    />
                  </div>

                  {activeOperatorId === null && (
                    <button
                      onClick={exportJamaahToExcel}
                      title="Ekspor data jamaah sesuai filter aktif ke file Excel"
                      className="flex items-center justify-center gap-2 px-4 py-2 rounded-lg bg-emerald-600 text-white text-sm font-medium hover:bg-emerald-700 transition-colors shadow-xs shrink-0"
                    >
                      <FileSpreadsheet className="w-4 h-4" />
                      <span>Export Excel</span>
                    </button>
                  )}

                  {activeOperatorId === null && (
                    <>
                      {/* Batch OCR Trigger */}
                      <button
                        onClick={() => { setShowBatchScanModal(true); setBatchScanFiles([]); setBatchScanResults([]); setBatchScanErrors([]); }}
                        className="flex items-center justify-center gap-2 px-4 py-2 rounded-lg bg-gradient-to-r from-blue-600 to-indigo-600 text-white text-sm font-semibold hover:from-blue-700 hover:to-indigo-700 transition-colors shadow-xs shrink-0"
                      >
                        <Sparkles className="w-4 h-4" />
                        <span>Scan Massal Visa (PDF)</span>
                      </button>

                      {/* Add Manual Trigger */}
                      <button
                        onClick={openAddModal}
                        className="flex items-center justify-center gap-2 px-4 py-2 rounded-lg bg-zinc-900 text-white text-sm font-medium hover:bg-zinc-800 transition-colors shadow-xs shrink-0"
                      >
                        <Plus className="w-4 h-4" />
                        <span>Tambah Jamaah Manual</span>
                      </button>
                    </>
                  )}
                </div>

                {/* Line 2: Advanced filters */}
                <div className="grid grid-cols-2 md:grid-cols-4 gap-3 pt-2 border-t border-slate-100">
                  
                  {/* Filter Status */}
                  <div className="space-y-1">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">Status</label>
                    <select
                      value={statusFilter}
                      onChange={(e) => setStatusFilter(e.target.value)}
                      className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-md py-1.5 px-2 bg-white dark:bg-zinc-700 text-slate-700 dark:text-zinc-200"
                    >
                      <option value="All">Semua Status</option>
                      <option value="Ready">Ready</option>
                      <option value="Sedang War">Sedang War</option>
                      <option value="QR Berhasil">QR Berhasil</option>
                      <option value="Belum Berhasil">Belum Berhasil</option>
                    </select>
                  </div>

                  {/* Filter Operator */}
                  <div className="space-y-1">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">Operator</label>
                    <select
                      value={operatorFilter}
                      onChange={(e) => setOperatorFilter(e.target.value)}
                      className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-md py-1.5 px-2 bg-white dark:bg-zinc-700 text-slate-700 dark:text-zinc-200"
                    >
                      <option value="All">Semua Operator</option>
                      <option value="unassigned">Belum Ditugaskan</option>
                      {operators.map(op => (
                        <option key={op.id} value={op.id}>{op.name}</option>
                      ))}
                    </select>
                  </div>

                  {/* Filter Travel */}
                  <div className="space-y-1">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">Nama Travel</label>
                    <select
                      value={travelFilter}
                      onChange={(e) => setTravelFilter(e.target.value)}
                      className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-md py-1.5 px-2 bg-white dark:bg-zinc-700 text-slate-700 dark:text-zinc-200"
                    >
                      <option value="All">Semua Travel</option>
                      {Array.from(new Set(jamaahs.map(j => j.travel || settingsTravelName))).filter(Boolean).map(trv => (
                        <option key={trv} value={trv}>{trv}</option>
                      ))}
                    </select>
                  </div>

                  {/* Filter Gender */}
                  <div className="space-y-1">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">Gender</label>
                    <select
                      value={genderFilter}
                      onChange={(e) => setGenderFilter(e.target.value)}
                      className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-md py-1.5 px-2 bg-white dark:bg-zinc-700 text-slate-700 dark:text-zinc-200"
                    >
                      <option value="All">Semua Gender</option>
                      <option value="Laki-laki">Laki-laki</option>
                      <option value="Perempuan">Perempuan</option>
                    </select>
                  </div>

                  {/* Filter Prioritas */}
                  <div className="space-y-1">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">Prioritas</label>
                    <select
                      value={priorityFilter}
                      onChange={(e) => setPriorityFilter(e.target.value)}
                      className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-md py-1.5 px-2 bg-white dark:bg-zinc-700 text-slate-700 dark:text-zinc-200"
                    >
                      <option value="All">Semua Prioritas</option>
                      <option value="Tinggi">Tinggi</option>
                      <option value="Sedang">Sedang</option>
                      <option value="Rendah">Rendah</option>
                      <option value="Belum Ada">Belum Ada</option>
                    </select>
                  </div>

                  {/* Filter File QR */}
                  <div className="space-y-1">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">File QR</label>
                    <select
                      value={qrFilter}
                      onChange={(e) => setQrFilter(e.target.value)}
                      className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-md py-1.5 px-2 bg-white dark:bg-zinc-700 text-slate-700 dark:text-zinc-200"
                    >
                      <option value="All">Semua</option>
                      <option value="uploaded">Sudah Upload</option>
                      <option value="pending">Belum Upload</option>
                    </select>
                  </div>

                  {/* Filter Tanggal Masuk */}
                  <div className="space-y-1">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">Tanggal Masuk</label>
                    <input
                      type="date"
                      value={entryFilter}
                      onChange={(e) => setEntryFilter(e.target.value)}
                      className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-md py-1.5 px-2 bg-white dark:bg-zinc-700 text-slate-700 dark:text-zinc-200"
                    />
                  </div>

                  {/* Clear filter button */}
                  <div className="flex items-end">
                    <button
                      onClick={() => {
                        setSearchQuery('');
                        setStatusFilter('All');
                        setOperatorFilter('All');
                        setGenderFilter('All');
                        setEntryFilter('');
                        setTravelFilter('All');
                        setPriorityFilter('All');
                        setQrFilter('All');
                      }}
                      className="w-full text-xs font-medium text-slate-500 dark:text-zinc-400 bg-slate-50 dark:bg-zinc-800/50 border border-slate-200 dark:border-zinc-600 hover:bg-slate-100 dark:bg-zinc-700 hover:text-slate-700 dark:text-zinc-200 rounded-md py-1.5 px-2 transition-colors flex items-center justify-center gap-1 animate-none cursor-pointer"
                    >
                      <span>Hapus Filter</span>
                    </button>
                  </div>

                </div>

              </div>

              {/* GROUPED TRAVEL ACCORDION VIEW */}
              {(() => {
                // --- Pagination over the sorted+filtered jamaah list ---
                const totalItems = sortedFilteredJamaah.length;
                const totalPages = Math.max(1, Math.ceil(totalItems / jamaahPageSize));
                const currentPage = Math.min(jamaahPage, totalPages);
                const pageStart = (currentPage - 1) * jamaahPageSize;
                const pageEnd = Math.min(pageStart + jamaahPageSize, totalItems);
                const pageJamaah = sortedFilteredJamaah.slice(pageStart, pageEnd);

                // === TAMPILAN OPERATOR: kartu mobile-friendly (tanpa tabel) ===
                // Operator tidak butuh edit travel / tabel — cukup daftar kartu yang
                // bisa diklik untuk membuka halaman detail jemaah.
                if (activeOperatorId !== null) {
                  // Operator: tampilkan SEMUA jemaah yang ditugaskan (tanpa paginasi tabel admin).
                  const operatorJamaah = sortedFilteredJamaah;
                  if (operatorJamaah.length === 0) {
                    return (
                      <div className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl shadow-xs p-12 text-center text-slate-400">
                        Tidak ada data jamaah yang ditugaskan kepada Anda.
                      </div>
                    );
                  }
                  const statusStyle: Record<string, string> = {
                    'Ready': 'bg-slate-100 text-slate-600 dark:bg-zinc-700 dark:text-zinc-300',
                    'Sedang War': 'bg-amber-100 text-amber-700 dark:bg-amber-500/15 dark:text-amber-300',
                    'QR Berhasil': 'bg-emerald-100 text-emerald-700 dark:bg-emerald-500/15 dark:text-emerald-300',
                    'Belum Berhasil': 'bg-rose-100 text-rose-700 dark:bg-rose-500/15 dark:text-rose-300',
                  };
                  return (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      {operatorJamaah.map(j => {
                        const prio = getPriorityInfo(j.entryMadinah, settingsReferenceDate);
                        return (
                          <button
                            key={j.id}
                            onClick={() => setSelectedJamaah(j)}
                            className="text-left bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-700 rounded-xl p-4 shadow-3xs hover:shadow-md hover:border-blue-300 dark:hover:border-blue-700/50 active:scale-[0.99] transition-all cursor-pointer flex flex-col gap-2.5"
                          >
                            <div className="flex items-start justify-between gap-2">
                              <span className="font-bold text-sm text-slate-800 dark:text-zinc-100 leading-tight">{j.name || '[Tanpa Nama]'}</span>
                              <span className={`shrink-0 text-[10px] px-2 py-0.5 rounded-full font-semibold ${statusStyle[j.status] || statusStyle['Ready']}`}>
                                {j.status}
                              </span>
                            </div>
                            <div className="flex flex-wrap items-center gap-1.5 text-[11px]">
                              {j.gender && (
                                <span className={`px-1.5 py-0.5 rounded font-medium ${j.gender === 'Perempuan' ? 'bg-pink-50 text-pink-700 dark:bg-pink-500/15 dark:text-pink-300' : 'bg-blue-50 text-blue-700 dark:bg-blue-500/15 dark:text-blue-300'}`}>
                                  {j.gender === 'Perempuan' ? '♀' : '♂'} {j.gender}
                                </span>
                              )}
                              <span className={`px-1.5 py-0.5 rounded font-semibold ${prio.badgeColor}`}>Prioritas {prio.level}</span>
                              {j.qrCodeUrl
                                ? <span className="px-1.5 py-0.5 rounded font-medium bg-emerald-50 text-emerald-700 dark:bg-emerald-500/15 dark:text-emerald-300">QR ✓</span>
                                : <span className="px-1.5 py-0.5 rounded font-medium bg-slate-100 text-slate-500 dark:bg-zinc-700 dark:text-zinc-400">Belum QR</span>}
                            </div>
                            <div className="flex items-center justify-between text-[11px] text-slate-500 dark:text-zinc-400 pt-1 border-t border-slate-100 dark:border-zinc-700/60">
                              <span className="font-mono">{j.passport || '—'}</span>
                              <span>Masuk: {j.entryMadinah ? formatDateLabel(j.entryMadinah) : '—'}</span>
                            </div>
                            <span className="text-[10px] font-semibold text-blue-600 dark:text-blue-400 flex items-center gap-1 mt-0.5">
                              Lihat detail →
                            </span>
                          </button>
                        );
                      })}
                    </div>
                  );
                }

                const jamaahsByTravel = pageJamaah.reduce((acc: { [key: string]: Jamaah[] }, j) => {
                  const travelName = j.travel || settingsTravelName;
                  if (!acc[travelName]) acc[travelName] = [];
                  acc[travelName].push(j);
                  return acc;
                }, {});

                const travelKeys = Object.keys(jamaahsByTravel).sort();
                const isFiltering = searchQuery.trim() !== '' ||
                                    statusFilter !== 'All' ||
                                    operatorFilter !== 'All' ||
                                    genderFilter !== 'All' ||
                                    travelFilter !== 'All' ||
                                    priorityFilter !== 'All' ||
                                    qrFilter !== 'All' ||
                                    entryFilter !== '';

                const getExpandedState = (travelName: string) => {
                  if (isFiltering) {
                    return expandedTravels[travelName] !== false; // auto-expand if filtering
                  }
                  return !!expandedTravels[travelName];
                };

                const toggleTravel = (travelName: string) => {
                  const currentVal = getExpandedState(travelName);
                  setExpandedTravels(prev => ({
                    ...prev,
                    [travelName]: !currentVal
                  }));
                };

                if (travelKeys.length === 0) {
                  return (
                    <div className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl shadow-xs p-12 text-center text-slate-400">
                      Tidak ada data jamaah yang memenuhi filter pencarian.
                    </div>
                  );
                }

                return (
                  <div className="space-y-4">
                    
                    {/* Header Toolbar */}
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between bg-slate-50 dark:bg-zinc-800/50 border border-slate-100 dark:border-zinc-700 rounded-lg px-4 py-2.5 gap-2 text-xs">
                      <span className="text-slate-600 dark:text-zinc-300 font-medium">
                        Menampilkan data terkelompok dalam <strong className="text-zinc-900 dark:text-zinc-100">{travelKeys.length} Agen Travel / Rombongan</strong>
                      </span>
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => {
                            const expanded: { [key: string]: boolean } = {};
                            travelKeys.forEach(k => { expanded[k] = true; });
                            setExpandedTravels(expanded);
                          }}
                          className="px-2.5 py-1 bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-600 hover:bg-slate-100 dark:bg-zinc-700 dark:hover:bg-zinc-700 text-slate-700 dark:text-zinc-200 hover:text-slate-900 rounded font-semibold transition-colors cursor-pointer"
                        >
                          Buka Semua
                        </button>
                        <button
                          onClick={() => {
                            const expanded: { [key: string]: boolean } = {};
                            travelKeys.forEach(k => { expanded[k] = false; });
                            setExpandedTravels(expanded);
                          }}
                          className="px-2.5 py-1 bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-600 hover:bg-slate-100 dark:bg-zinc-700 dark:hover:bg-zinc-700 text-slate-700 dark:text-zinc-200 hover:text-slate-900 rounded font-semibold transition-colors cursor-pointer"
                        >
                          Tutup Semua
                        </button>
                      </div>
                    </div>

                    {/* BULK ACTIONS TOOLBAR */}
                    {Object.values(selectedJamaahIds).filter(Boolean).length > 0 && (
                      <div className="bg-rose-50 dark:bg-rose-500/10 border border-rose-100 dark:border-rose-500/20 rounded-xl p-3 px-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-xs animate-in slide-in-from-top-2 duration-200">
                        <div className="flex items-center gap-2">
                          <span className="flex h-2 w-2 relative">
                            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span>
                            <span className="relative inline-flex rounded-full h-2 w-2 bg-rose-500"></span>
                          </span>
                          <span className="text-xs font-semibold text-rose-800 dark:text-rose-300">
                            Terpilih {Object.values(selectedJamaahIds).filter(Boolean).length} Jamaah dari Agen Travel
                          </span>
                        </div>
                        <div className="flex items-center justify-between sm:justify-end gap-3 w-full sm:w-auto">
                          <button
                            onClick={() => setSelectedJamaahIds({})}
                            className="text-xs text-rose-700 dark:text-rose-400 font-medium hover:underline cursor-pointer"
                          >
                            Batal Pilihan
                          </button>
                          {activeOperatorId === null && (
                            <button
                              onClick={handleDeleteSelectedJamaahs}
                              className="px-3.5 py-1.5 bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer shadow-3xs"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                              <span>Hapus Terpilih ({Object.values(selectedJamaahIds).filter(Boolean).length})</span>
                            </button>
                          )}
                        </div>
                      </div>
                    )}

                    {/* Collapsible Accordion Cards */}
                    <div className="space-y-2">
                      {travelKeys.map(travelName => {
                        const list = jamaahsByTravel[travelName];
                        const isExpanded = getExpandedState(travelName);

                        // Gender stats
                        const maleCount = list.filter(j => j.gender === 'Laki-laki').length;
                        const femaleCount = list.filter(j => j.gender === 'Perempuan').length;

                        // Status stats
                        const readyCount = list.filter(j => j.status === 'Ready').length;
                        const warCount = list.filter(j => j.status === 'Sedang War').length;
                        const successCount = list.filter(j => j.status === 'QR Berhasil').length;
                        const failCount = list.filter(j => j.status === 'Belum Berhasil').length;

                        return (
                          <div key={travelName} className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl shadow-3xs overflow-hidden transition-all duration-200 hover:shadow-2xs">
                            
                            {/* Accordion Header Button */}
                            <button
                              onClick={() => toggleTravel(travelName)}
                              className="w-full flex items-center justify-between p-3 bg-slate-50 dark:bg-zinc-800/50 hover:bg-slate-50 dark:bg-zinc-800/50 dark:bg-zinc-800/30 dark:hover:bg-zinc-700/40 text-left transition-colors border-b border-slate-100/60 dark:border-zinc-700/60 cursor-pointer"
                            >
                              <div className="flex items-center gap-3 flex-1 min-w-0">
                                <div className="w-8 h-8 rounded-lg bg-red-600 text-white font-bold flex items-center justify-center text-xs shrink-0 shadow-3xs uppercase">
                                  {travelName.charAt(0)}
                                </div>
                                <div className="min-w-0">
                                  <h4 className="font-bold text-slate-800 dark:text-zinc-100 text-sm flex items-center gap-2 flex-wrap">
                                    <span className="truncate">{travelName}</span>
                                    <span
                                      role="button"
                                      tabIndex={0}
                                      title="Ubah nama Travel / Rombongan (mengubah semua jemaah di grup ini)"
                                      onClick={(e) => { e.stopPropagation(); renameTravelGroup(travelName); }}
                                      className="p-1 rounded-md text-slate-400 hover:text-red-600 dark:hover:text-red-400 hover:bg-slate-200/70 dark:hover:bg-zinc-700 transition-colors cursor-pointer shrink-0"
                                    >
                                      <Edit className="w-3.5 h-3.5" />
                                    </span>
                                    <span className="px-2 py-0.5 rounded bg-emerald-50 dark:bg-emerald-500/10 text-emerald-700 dark:text-emerald-400 border border-emerald-100 dark:border-emerald-500/20 text-[10px] font-bold shrink-0">
                                      {list.length} Pax
                                    </span>
                                  </h4>
                                  <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 text-[10px] text-slate-400 mt-1">
                                    <span className="font-medium text-slate-500 dark:text-zinc-400">
                                      Laki-laki: {maleCount}
                                    </span>
                                    <span>•</span>
                                    <span className="font-medium text-slate-500 dark:text-zinc-400">
                                      Perempuan: {femaleCount}
                                    </span>
                                  </div>
                                </div>
                              </div>

                              {/* Status Badges Summary */}
                              <div className="hidden md:flex items-center gap-2.5 px-4 shrink-0">
                                {readyCount > 0 && (
                                  <span className="text-[10px] bg-slate-50 dark:bg-zinc-800/50 text-slate-600 dark:text-zinc-300 border border-slate-100 dark:border-zinc-700 px-2 py-0.5 rounded-full font-medium flex items-center gap-1">
                                    <span className="w-1.5 h-1.5 rounded-full bg-slate-400"></span>
                                    <span>{readyCount} Ready</span>
                                  </span>
                                )}
                                {warCount > 0 && (
                                  <span className="text-[10px] bg-amber-50 dark:bg-amber-500/10 text-amber-800 dark:text-amber-400 border border-amber-100 dark:border-amber-500/20 px-2 py-0.5 rounded-full font-medium flex items-center gap-1">
                                    <span className="w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse"></span>
                                    <span>{warCount} War</span>
                                  </span>
                                )}
                                {successCount > 0 && (
                                  <span className="text-[10px] bg-emerald-50 dark:bg-emerald-500/10 text-emerald-800 dark:text-emerald-400 border border-emerald-100 dark:border-emerald-500/20 px-2 py-0.5 rounded-full font-medium flex items-center gap-1">
                                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                                    <span>{successCount} Sukses</span>
                                  </span>
                                )}
                                {failCount > 0 && (
                                  <span className="text-[10px] bg-red-50 dark:bg-red-500/10 text-red-800 dark:text-red-400 border border-red-100 dark:border-red-500/20 px-2 py-0.5 rounded-full font-medium flex items-center gap-1">
                                    <span className="w-1.5 h-1.5 rounded-full bg-rose-500"></span>
                                    <span>{failCount} Gagal</span>
                                  </span>
                                )}
                              </div>

                              {/* Chevron Trigger */}
                              <div className="text-slate-400 hover:text-slate-600 dark:text-zinc-300 p-1 shrink-0">
                                {isExpanded ? (
                                  <ChevronDown className="w-4 h-4" />
                                ) : (
                                  <ChevronRight className="w-4 h-4" />
                                )}
                              </div>
                            </button>

                            {/* Collapsible Table Content */}
                            {isExpanded && (
                              <div className="overflow-x-auto">
                                <table className="w-full text-left border-collapse text-xs">
                                  <thead>
                                    <tr className="bg-slate-50 dark:bg-zinc-800/50 border-b border-slate-100 dark:border-zinc-700 text-slate-500 dark:text-zinc-400 font-medium select-none text-xs">
                                      <th className="py-2.5 px-2 text-center w-8">
                                        <input
                                          type="checkbox"
                                          checked={list.length > 0 && list.every(item => selectedJamaahIds[item.id])}
                                          onChange={(e) => {
                                            const checked = e.target.checked;
                                            setSelectedJamaahIds(prev => {
                                              const copy = { ...prev };
                                              list.forEach(item => {
                                                copy[item.id] = checked;
                                              });
                                              return copy;
                                            });
                                          }}
                                          className="rounded border-slate-300 text-red-600 focus:ring-red-500 cursor-pointer h-3.5 w-3.5"
                                        />
                                      </th>
                                      <th className="py-2.5 px-2 whitespace-nowrap">Nama Jamaah</th>
                                      <th className="py-2.5 px-2 whitespace-nowrap">Email</th>
                                      <th className="py-2.5 px-2 whitespace-nowrap">Paspor</th>
                                      <th className="py-2.5 px-2 whitespace-nowrap">Visa</th>
                                      <th className="py-2.5 px-2 whitespace-nowrap">Password</th>
                                      <th className="py-2.5 px-2 whitespace-nowrap">Jadwal</th>
                                      <th className="py-2.5 px-2 whitespace-nowrap">Prioritas</th>
                                      <th className="py-2.5 px-2 whitespace-nowrap">Status</th>
                                      <th className="py-2.5 px-2 whitespace-nowrap text-center">QR</th>
                                      <th className="py-2.5 px-2 whitespace-nowrap">Operator</th>
                                      <th className="py-2.5 px-2 whitespace-nowrap text-right">Aksi</th>
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-slate-100 dark:divide-zinc-700 text-slate-700 dark:text-zinc-200 dark:text-zinc-200 bg-white dark:bg-zinc-800">
                                    {list.map(j => {
                                       const prio = getPriorityInfo(j.entryMadinah, settingsReferenceDate);
                                       const isUrgent = (() => {
                                         if (!j.raudhahSlot || j.status === 'QR Berhasil') return false;
                                         const dist = getDistributionInstant(j.raudhahSlot, settingsQrLeadHours);
                                         const slot = getDistributionInstant(j.raudhahSlot, 0);
                                         if (!dist || !slot) return false;
                                         const t = now.getTime();
                                         return (slot.getTime() - t) > 0 && (dist.getTime() - t) / 60000 <= 60;
                                       })();

                                       return (
                                         <tr key={j.id} className={`transition-colors ${
                                          selectedJamaahIds[j.id] ? 'bg-blue-50/10 dark:bg-blue-900/20' : ''
                                        } ${
                                          isUrgent 
                                            ? 'bg-rose-50/50 hover:bg-rose-100/40 border-l-2 border-l-rose-600 dark:bg-rose-950/30 dark:hover:bg-rose-900/40 dark:border-l-rose-800' 
                                            : 'hover:bg-slate-50 dark:hover:bg-zinc-700/30'
                                         }`}>
                                          <td className="py-2.5 px-2 text-center">
                                            <input
                                              type="checkbox"
                                              checked={!!selectedJamaahIds[j.id]}
                                              onChange={(e) => {
                                                const checked = e.target.checked;
                                                setSelectedJamaahIds(prev => ({
                                                  ...prev,
                                                  [j.id]: checked
                                                }));
                                              }}
                                              className="rounded border-slate-300 text-red-600 focus:ring-red-500 cursor-pointer h-3.5 w-3.5"
                                            />
                                          </td>
                                          <td className="py-2.5 px-2">
                                            <div className="font-semibold text-slate-800 dark:text-zinc-100 flex items-center gap-1.5 truncate text-sm">
                                               <span className="truncate">{j.name}</span>
                                               {j.gender && (
                                                 <span className={`text-[10px] px-1.5 py-0.5 rounded-full font-bold border shrink-0 leading-none ${
                                                   j.gender === 'Perempuan'
                                                     ? 'bg-pink-50 dark:bg-pink-500/15 text-pink-700 dark:text-pink-300 border-pink-200 dark:border-pink-500/30'
                                                     : 'bg-blue-50 dark:bg-blue-500/15 text-blue-700 dark:text-blue-300 border-blue-200 dark:border-blue-500/30'
                                                 }`}>
                                                    {j.gender === 'Perempuan' ? '♀ Perempuan' : '♂ Laki-laki'}
                                                 </span>
                                               )}
                                               {duplicateIds.has(j.id) && (
                                                 <button
                                                   type="button"
                                                   onClick={(e) => { e.stopPropagation(); setShowDuplicateModal(true); }}
                                                   title="Paspor/Visa sama dengan jemaah lain — klik untuk membersihkan"
                                                   className="text-[8px] text-amber-800 dark:text-amber-300 bg-amber-100 dark:bg-amber-500/15 border border-amber-200 dark:border-amber-500/25 px-1.5 py-0.5 rounded font-bold flex items-center gap-0.5 shrink-0 leading-none hover:bg-amber-200 transition-colors"
                                                 >
                                                   <AlertTriangle className="w-2.5 h-2.5 shrink-0" />
                                                   <span>Duplikat</span>
                                                 </button>
                                               )}
                                               {isUrgent && (
                                                 <span className="text-[8px] text-rose-800 dark:text-rose-300 bg-rose-100 dark:bg-rose-500/15 border border-rose-200/50 dark:border-rose-500/25 px-1.5 py-0.5 rounded font-bold animate-pulse flex items-center gap-0.5 shrink-0 leading-none">
                                                   <AlertTriangle className="w-2.5 h-2.5 text-rose-600 shrink-0" />
                                                   <span>&lt;1 Jam!</span>
                                                 </span>
                                               )}
                                            </div>

                                          </td>
                                          <td className="py-2.5 px-2">
                                            {j.email ? (
                                              <span className="font-mono text-xs text-blue-700 dark:text-blue-300 bg-blue-50 dark:bg-blue-500/10 px-1.5 py-0.5 rounded border border-blue-200/60 dark:border-blue-500/20 inline-block truncate w-full align-middle leading-tight">{j.email}</span>
                                            ) : (
                                              <span className="text-xs text-slate-400 italic">-</span>
                                            )}
                                          </td>
                                          <td className="py-2.5 px-2">
                                            <span className="font-mono font-bold text-slate-800 dark:text-zinc-100 text-xs bg-slate-100 dark:bg-zinc-700 px-1.5 py-0.5 rounded border border-slate-200 dark:border-zinc-600 inline-block truncate w-full align-middle leading-tight">{j.passport}</span>
                                          </td>
                                          <td className="py-2.5 px-2">
                                            <span className="font-mono font-medium text-slate-700 dark:text-zinc-200 text-xs bg-slate-50 dark:bg-zinc-800/50 px-1.5 py-0.5 rounded border border-slate-200 dark:border-zinc-600/50 inline-block truncate w-full align-middle leading-tight">{j.visa}</span>
                                          </td>
                                          <td className="py-2.5 px-2">
                                            <span className="font-mono text-xs font-semibold bg-amber-50 dark:bg-amber-500/10 text-amber-800 dark:text-amber-300 px-1.5 py-0.5 rounded border border-dashed border-amber-200 dark:border-amber-500/30 inline-flex items-center gap-1 truncate w-full align-middle leading-tight">
                                              <Key className="w-3 h-3 text-amber-600 shrink-0" />
                                              <span className="truncate">{j.password || '123456'}</span>
                                            </span>
                                          </td>
                                          <td className="py-2.5 px-2 align-top">
                                            <div className="text-xs font-medium text-slate-700 dark:text-zinc-200 leading-tight">{formatDateLabel(j.entryMadinah)}</div>
                                            {j.raudhahSlot ? (() => {
                                              const r = getQrReminder(j.raudhahSlot, j.status, now.getTime(), settingsQrLeadHours);
                                              if (!r) return null;
                                              return (
                                                <div className="mt-1 space-y-0.5">
                                                  <div className="text-[10px] font-bold text-sky-800 dark:text-sky-300 bg-sky-50 dark:bg-sky-900/15 border border-sky-200 dark:border-sky-900/40 rounded px-1 py-0.5 inline-flex items-center gap-0.5 select-none leading-tight">
                                                    <Download className="w-3 h-3 text-sky-600 shrink-0" />
                                                    <span className="truncate">{formatTimeColon(r.dist, TZ_WITA)}</span>
                                                  </div>
                                                  <div className={`text-[10px] font-bold border rounded px-1 py-0.5 inline-flex items-center gap-0.5 select-none leading-tight ${QR_REMINDER_BADGE[r.tone]}`}>
                                                    <Bell className="w-3 h-3 shrink-0" />
                                                    <span>{r.countdownLabel}</span>
                                                  </div>
                                                  <div className="text-[10px] text-slate-500 dark:text-zinc-400 leading-tight truncate">
                                                    Slot: {formatTimeColon(r.slot, TZ_WITA)}
                                                  </div>
                                                </div>
                                              );
                                            })() : (
                                              <div className="text-xs text-slate-400 italic leading-tight">-</div>
                                            )}
                                          </td>
                                          <td className="py-2.5 px-2">
                                            <span className={`inline-flex items-center gap-1 border px-1.5 py-0.5 rounded-full text-xs font-medium ${prio.badgeColor}`}>
                                              <span className={`w-1.5 h-1.5 rounded-full ${prio.dotColor}`}></span>
                                              {prio.level} (H{prio.daysRemaining >= 0 ? `-${prio.daysRemaining}` : `+${Math.abs(prio.daysRemaining)}`})
                                            </span>
                                          </td>
                                          <td className="py-2.5 px-2">
                                            <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-xs font-semibold uppercase whitespace-nowrap ${
                                              j.status === 'Ready' ? 'bg-slate-100 dark:bg-zinc-700 text-slate-700 dark:text-zinc-200 border border-slate-200 dark:border-zinc-600/50' :
                                              j.status === 'Sedang War' ? 'bg-amber-100 dark:bg-amber-500/15 text-amber-800 dark:text-amber-300 border border-amber-200/30 dark:border-amber-500/20' :
                                              j.status === 'QR Berhasil' ? 'bg-emerald-100 dark:bg-emerald-500/15 text-emerald-800 dark:text-emerald-300 border border-emerald-200/30 dark:border-emerald-500/20' :
                                              'bg-rose-100 dark:bg-rose-500/15 text-rose-800 dark:text-rose-300 border border-rose-200/30 dark:border-rose-500/20'
                                            }`}>
                                              {j.status === 'Ready' && (
                                                <><span className="w-1.5 h-1.5 rounded-full bg-slate-400"></span> Ready</>
                                              )}
                                              {j.status === 'Sedang War' && (
                                                <><span className="w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse"></span> War</>
                                              )}
                                              {j.status === 'QR Berhasil' && (
                                                <><span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span> OK</>
                                              )}
                                              {j.status === 'Belum Berhasil' && (
                                                <><span className="w-1.5 h-1.5 rounded-full bg-rose-500"></span> Gagal</>
                                              )}
                                            </span>
                                          </td>
                                          <td className="py-2.5 px-2 text-center">
                                            {j.qrCodeUrl ? (
                                              <span className="inline-flex items-center gap-1 text-xs text-emerald-700 dark:text-emerald-400 bg-emerald-50 dark:bg-emerald-500/10 border border-emerald-100 dark:border-emerald-500/20 px-2 py-0.5 rounded-full font-bold whitespace-nowrap">
                                                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                                                QR
                                              </span>
                                            ) : (
                                              <span className="inline-flex items-center gap-1 text-xs text-rose-700 dark:text-rose-400 bg-rose-50 dark:bg-rose-500/10 border border-rose-100 dark:border-rose-500/20 px-2 py-0.5 rounded-full font-bold animate-pulse whitespace-nowrap">
                                                <AlertTriangle className="w-3 h-3 text-rose-600" />
                                                -
                                              </span>
                                            )}
                                          </td>
                                          <td className="py-2.5 px-2">
                                            {activeOperatorId === null ? (
                                              <select
                                                value={j.operatorId || ''}
                                                onChange={(e) => {
                                                  const val = e.target.value || null;
                                                  setJamaahs(prev => prev.map(item => item.id === j.id ? { ...item, operatorId: val } : item));
                                                  const supabase = getSupabase();
                                                  if (supabase && isSupabaseConnected) {
                                                    supabase.from('jamaahs').update({ operator_id: val }).eq('id', j.id).then(({ error }) => {
                                                      if (error) console.error('Failed to update operator in Supabase:', error);
                                                    });
                                                  }
                                                }}
                                                className={`text-xs rounded border p-0.5 text-slate-700 dark:text-zinc-200 outline-none max-w-[110px] cursor-pointer font-semibold ${
                                                  !j.operatorId ? 'bg-red-50 dark:bg-red-500/10 border-red-200 dark:border-red-500/30 text-red-700 dark:text-red-400' : 'bg-white dark:bg-zinc-700 border-slate-200 dark:border-zinc-600 text-slate-700 dark:text-zinc-200'
                                                }`}
                                              >
                                                <option value="" className="text-red-700 dark:text-red-400 font-semibold bg-white dark:bg-zinc-800">-</option>
                                                {operators.filter(o => o.isActive).map(o => (
                                                  <option key={o.id} value={o.id} className="text-slate-800 dark:text-zinc-100 font-normal bg-white dark:bg-zinc-800">{o.name}</option>
                                                ))}
                                              </select>
                                            ) : (
                                              <span className={`text-xs font-semibold ${!j.operatorId ? 'text-red-400' : 'text-zinc-200'}`}>
                                                {j.operatorId ? operators.find(o => o.id === j.operatorId)?.name || '-' : '-'}
                                              </span>
                                            )}
                                          </td>
                                          <td className="py-2.5 px-2 text-right">
                                            <div className="flex items-center justify-end gap-1">
                                              <button
                                                onClick={() => setSelectedJamaah(j)}
                                                title="Lihat Detail"
                                                className="p-1 text-slate-600 dark:text-zinc-300 hover:text-slate-900 bg-slate-100 dark:bg-zinc-700 hover:bg-slate-200/80 rounded transition-all cursor-pointer"
                                              >
                                                <Eye className="w-3.5 h-3.5" />
                                              </button>
                                              {activeOperatorId === null && (
                                                <>
                                                  <button
                                                    onClick={() => handleOpenEditModal(j)}
                                                    title="Edit Data"
                                                    className="p-1 text-red-600 dark:text-red-400 hover:text-red-900 bg-red-50 dark:bg-red-500/10 hover:bg-red-100/80 dark:hover:bg-red-500/20 rounded transition-all cursor-pointer"
                                                  >
                                                    <Edit className="w-3.5 h-3.5" />
                                                  </button>
                                                  <button
                                                    onClick={() => handleDeleteJamaah(j.id)}
                                                    title="Hapus Data"
                                                    className="p-1 text-rose-600 dark:text-rose-400 hover:text-rose-950 bg-rose-50 dark:bg-rose-500/10 hover:bg-rose-100/80 dark:hover:bg-rose-500/20 rounded transition-all cursor-pointer"
                                                  >
                                                    <Trash2 className="w-3.5 h-3.5" />
                                                  </button>
                                                </>
                                              )}
                                            </div>
                                          </td>
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>
                              </div>
                            )}

                          </div>
                        );
                      })}
                    </div>

                    {/* Pagination Controls */}
                    {totalItems > 0 && (
                      <div className="flex flex-col sm:flex-row items-center justify-between gap-3 bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl px-4 py-3 text-xs">
                        <div className="flex items-center gap-2 text-slate-600 dark:text-zinc-300">
                          <span>Menampilkan <strong className="text-slate-800 dark:text-zinc-100">{pageStart + 1}–{pageEnd}</strong> dari <strong className="text-slate-800 dark:text-zinc-100">{totalItems}</strong> jemaah</span>
                          <select
                            value={jamaahPageSize}
                            onChange={(e) => setJamaahPageSize(Number(e.target.value))}
                            className="ml-1 border border-slate-200 dark:border-zinc-600 rounded-md py-1 px-1.5 bg-slate-50 dark:bg-zinc-700 text-slate-700 dark:text-zinc-200 cursor-pointer font-medium"
                          >
                            {[10, 25, 50, 100].map(n => <option key={n} value={n}>{n}/halaman</option>)}
                          </select>
                        </div>
                        <div className="flex items-center gap-1">
                          <button
                            onClick={() => setJamaahPage(1)}
                            disabled={currentPage === 1}
                            className="px-2.5 py-1 rounded-md border border-slate-200 dark:border-zinc-600 bg-white dark:bg-zinc-800 hover:bg-slate-100 dark:hover:bg-zinc-700 text-slate-700 dark:text-zinc-200 font-semibold disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer transition-colors"
                          >«</button>
                          <button
                            onClick={() => setJamaahPage(Math.max(1, currentPage - 1))}
                            disabled={currentPage === 1}
                            className="px-2.5 py-1 rounded-md border border-slate-200 dark:border-zinc-600 bg-white dark:bg-zinc-800 hover:bg-slate-100 dark:hover:bg-zinc-700 text-slate-700 dark:text-zinc-200 font-semibold disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer transition-colors"
                          >Sebelumnya</button>
                          <span className="px-2.5 py-1 font-semibold text-slate-700 dark:text-zinc-200">Hal {currentPage} / {totalPages}</span>
                          <button
                            onClick={() => setJamaahPage(Math.min(totalPages, currentPage + 1))}
                            disabled={currentPage === totalPages}
                            className="px-2.5 py-1 rounded-md border border-slate-200 dark:border-zinc-600 bg-white dark:bg-zinc-800 hover:bg-slate-100 dark:hover:bg-zinc-700 text-slate-700 dark:text-zinc-200 font-semibold disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer transition-colors"
                          >Berikutnya</button>
                          <button
                            onClick={() => setJamaahPage(totalPages)}
                            disabled={currentPage === totalPages}
                            className="px-2.5 py-1 rounded-md border border-slate-200 dark:border-zinc-600 bg-white dark:bg-zinc-800 hover:bg-slate-100 dark:hover:bg-zinc-700 text-slate-700 dark:text-zinc-200 font-semibold disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer transition-colors"
                          >»</button>
                        </div>
                      </div>
                    )}

                  </div>
                );
              })()}

            </div>
          )}

          {/* TAB 3: IMPORT EXCEL */}
          {currentTab === 'import' && (
            <div className="space-y-6">
              
              {previewRows.length === 0 ? (
                <div className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl shadow-xs p-6 max-w-2xl mx-auto space-y-6">
                  <div className="text-center space-y-2">
                    <h3 className="font-semibold text-slate-800 dark:text-zinc-100 text-base">Import Data Jamaah dari File Excel / CSV</h3>
                    <p className="text-xs text-slate-500 dark:text-zinc-400 max-w-md mx-auto">
                      Unggah daftar manifest jamaah yang visanya sudah terbit untuk langsung di-input ke antrean booking Raudhah.
                    </p>
                  </div>

                  {/* Drag and Drop Zone */}
                  <div
                    onClick={() => fileInputRef.current?.click()}
                    className="border-2 border-dashed border-slate-200 dark:border-zinc-600 hover:border-blue-500 hover:bg-blue-50/10 rounded-xl p-8 flex flex-col items-center justify-center gap-3 cursor-pointer transition-all group"
                  >
                    <div className="w-12 h-12 rounded-full bg-slate-50 dark:bg-zinc-800/50 group-hover:bg-red-100 text-slate-400 group-hover:text-red-700 flex items-center justify-center transition-all">
                      <Upload className="w-6 h-6" />
                    </div>
                    <div className="text-center">
                      <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Klik untuk mengunggah file</span>
                      <p className="text-[11px] text-slate-400 mt-0.5">Mendukung format .XLSX, .XLS, atau .CSV (Maks. 10MB)</p>
                    </div>
                    <input
                      type="file"
                      ref={fileInputRef}
                      onChange={handleFileUpload}
                      className="hidden"
                      accept=".csv, .xlsx, .xls"
                    />
                  </div>

                  {/* Standard Column Specification Template */}
                  <div className="p-4 bg-slate-50 dark:bg-zinc-800/50 rounded-lg space-y-3 text-xs">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-200 dark:border-zinc-600/60 pb-2">
                      <div>
                        <span className="font-bold text-slate-800 dark:text-zinc-100 text-xs">Format & Template Excel Sinkron</span>
                        <p className="text-[11px] text-slate-500 dark:text-zinc-400 mt-0.5">Semua data terintegrasi termasuk Password Akses Jemaah & Jam Spesifik.</p>
                      </div>
                      <button
                        type="button"
                        onClick={downloadExcelTemplate}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-white bg-red-600 hover:bg-red-700 rounded-lg transition-colors cursor-pointer shadow-3xs"
                      >
                        <FileSpreadsheet className="w-3.5 h-3.5" />
                        <span>Unduh Template Excel (.xlsx)</span>
                      </button>
                    </div>

                    <div className="space-y-1">
                      <p className="text-[11px] font-semibold text-slate-600 dark:text-zinc-300">Urutan Kolom yang Didukung (10 Kolom):</p>
                      <div className="grid grid-cols-2 md:grid-cols-5 gap-2 text-[10px] font-mono text-slate-600 dark:text-zinc-300 bg-white dark:bg-zinc-800 p-2.5 rounded border border-slate-200 dark:border-zinc-600/50">
                        <div>1. Nama Lengkap</div>
                        <div>2. Nomor Paspor</div>
                        <div>3. Nomor Visa</div>
                        <div>4. Jenis Kelamin</div>
                        <div>5. No WhatsApp</div>
                        <div>6. Email</div>
                        <div>7. Tgl Masuk (Spesifik Jam)</div>
                        <div>8. Tgl Keluar Madinah</div>
                        <div>9. Nama Travel</div>
                        <div>10. Password Akses</div>
                      </div>
                    </div>

                  </div>
                </div>
              ) : (
                // PREVIEW PRE-IMPORT STATE (GROUPED BY PAX)
                <div className="space-y-6">
                  <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-white dark:bg-zinc-800 p-4 rounded-xl border border-slate-100 dark:border-zinc-700 shadow-xs">
                    <div className="flex flex-col sm:flex-row sm:items-center gap-4">
                      <div>
                        <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">Pratinjau Unggahan Excel (Terbagi per Rombongan)</h3>
                        <p className="text-xs text-slate-500 dark:text-zinc-400 mt-0.5">Sistem mengelompokkan jamaah otomatis berdasarkan kolom "Grup" di file Excel Anda.</p>
                      </div>
                    </div>
                    
                    <div className="flex items-center gap-2">
                      <button
                        onClick={cancelImport}
                        className="px-3.5 py-1.5 rounded-lg border border-red-200 text-xs font-semibold bg-red-50 hover:bg-red-100 text-red-700 cursor-pointer flex items-center gap-1 transition-all"
                      >
                        <X className="w-3.5 h-3.5" />
                        <span>Bersihkan Pratinjau / Hapus File</span>
                      </button>
                      <button
                        onClick={() => executeImport()}
                        className="px-4 py-1.5 rounded-lg bg-red-600 text-white text-xs font-semibold hover:bg-red-700 transition-all shadow-xs cursor-pointer flex items-center gap-1"
                      >
                        <Check className="w-3.5 h-3.5" />
                        <span>Impor Semua ({importStats.valid} Data)</span>
                      </button>
                    </div>
                  </div>

                  {/* PREVIEW STATS */}
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <div className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 p-4 rounded-xl shadow-xs">
                      <span className="text-[10px] font-semibold uppercase text-slate-400 tracking-wider">Total Baris</span>
                      <div className="text-xl font-bold text-slate-800 dark:text-zinc-100 mt-1">{importStats.total}</div>
                    </div>
                    <div className="bg-white dark:bg-zinc-800 border border-blue-100 dark:border-blue-900/50 p-4 rounded-xl shadow-xs border-l-4 border-l-blue-500">
                      <span className="text-[10px] font-semibold uppercase text-red-600 tracking-wider">Total Pax Valid</span>
                      <div className="text-xl font-bold text-red-700 mt-1">{importStats.valid} Pax</div>
                    </div>
                    <div className="bg-white dark:bg-zinc-800 border border-amber-100 dark:border-amber-900/50 p-4 rounded-xl shadow-xs border-l-4 border-l-amber-500">
                      <span className="text-[10px] font-semibold uppercase text-amber-600 tracking-wider">Duplikat Terdeteksi</span>
                      <div className="text-xl font-bold text-amber-700 mt-1">{importStats.duplicate}</div>
                    </div>
                    <div className="bg-white dark:bg-zinc-800 border border-red-100 dark:border-red-900/50 p-4 rounded-xl shadow-xs border-l-4 border-l-red-500">
                      <span className="text-[10px] font-semibold uppercase text-red-600 tracking-wider">Tidak Lengkap</span>
                      <div className="text-xl font-bold text-red-700 mt-1">{importStats.incomplete}</div>
                    </div>
                  </div>

                  {/* GROUPED TABLES */}
                  <div className="space-y-6">
                    {Object.entries(
                      previewRows.reduce((acc: { [key: string]: any[] }, r) => {
                        const trv = r.travel || settingsTravelName;
                        if (!acc[trv]) acc[trv] = [];
                        acc[trv].push(r);
                        return acc;
                      }, {})
                    ).map(([travelName, rawRows]) => {
                      const rows = rawRows as any[];
                      const validCount = rows.filter(r => r.status === 'Valid').length;
                      return (
                        <div key={travelName} className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl shadow-xs overflow-hidden">
                          {/* Travel Header with Pax info and targeted import */}
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between p-4 bg-slate-50 dark:bg-zinc-800/50 dark:bg-zinc-800/50 border-b border-slate-100 dark:border-zinc-700 gap-3">
                            <div className="flex items-center gap-3">
                              <div className={`w-8 h-8 rounded-lg font-bold flex items-center justify-center text-xs uppercase ${
                                !travelName.trim() ? 'bg-rose-600 text-white animate-pulse' : 'bg-zinc-900 text-white'
                              }`}>
                                {travelName ? travelName.charAt(0).toUpperCase() : '?'}
                              </div>
                              <div>
                                <h4 className="font-bold text-slate-800 dark:text-zinc-100 text-sm flex items-center gap-2">
                                  <input
                                    defaultValue={travelName}
                                    placeholder="KETIK NAMA AGEN TRAVEL DI SINI..."
                                    onBlur={(e) => renamePreviewTravel(travelName, e.target.value)}
                                    onKeyDown={(e) => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur(); }}
                                    title="Klik untuk mengubah nama travel/rombongan grup ini"
                                    className={`font-bold text-sm bg-transparent border rounded px-2 py-0.5 outline-none transition-all min-w-[280px] ${
                                      !travelName.trim()
                                        ? 'border-dashed border-rose-400 dark:border-rose-700 bg-rose-50/20 dark:bg-rose-950/10 text-rose-600 placeholder:text-rose-400'
                                        : 'border-transparent text-slate-800 dark:text-zinc-100 hover:border-slate-300 dark:hover:border-zinc-600 focus:border-blue-400 focus:bg-white dark:focus:bg-zinc-900'
                                    }`}
                                  />
                                  <span className="px-2 py-0.5 rounded bg-blue-50 text-red-700 border border-blue-100 text-[10px] font-semibold shrink-0">
                                    {rows.length} Pax
                                  </span>
                                  {!travelName.trim() && (
                                    <span className="px-2 py-0.5 rounded-full bg-rose-50 text-rose-700 border border-rose-100 text-[9px] font-bold animate-pulse shrink-0">
                                      ⚠️ NAMA TRAVEL KOSONG
                                    </span>
                                  )}
                                </h4>
                                <p className="text-[10px] text-slate-500 dark:text-zinc-400 mt-0.5">Rincian manifest travel ini ({validCount} Pax valid siap dimasukkan) — klik nama travel di atas untuk mengubahnya.</p>
                              </div>
                            </div>
                            
                            <div>
                              <button
                                onClick={() => executeImport(travelName)}
                                disabled={validCount === 0}
                                className={`px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-xs flex items-center gap-1.5 transition-all ${
                                  validCount > 0 
                                    ? 'bg-red-600 hover:bg-red-700 text-white cursor-pointer' 
                                    : 'bg-slate-100 dark:bg-zinc-700 text-slate-400 cursor-not-allowed'
                                }`}
                              >
                                <Upload className="w-3.5 h-3.5" />
                                <span>Impor Travel Ini Saja ({validCount} Pax)</span>
                              </button>
                            </div>
                          </div>

                          {/* Group Table */}
                          <div className="overflow-x-auto">
                            <table className="w-full text-left text-xs border-collapse">
                              <thead>
                                <tr className="bg-white dark:bg-zinc-800 border-b border-slate-100 dark:border-zinc-700 text-slate-500 dark:text-zinc-400 font-medium">
                                  <th className="py-2.5 px-4">Nama Jamaah</th>
                                  <th className="py-2.5 px-3">Email</th>
                                  <th className="py-2.5 px-3">No Paspor</th>
                                  <th className="py-2.5 px-3">No Visa</th>
                                  <th className="py-2.5 px-3">Password</th>
                                  <th className="py-2.5 px-3">Gender</th>
                                  <th className="py-2.5 px-3">Masuk Madinah</th>
                                  {customFields.map(cf => (
                                    <th key={cf.id} className="py-2.5 px-3">{cf.label}</th>
                                  ))}
                                  <th className="py-2.5 px-3">Status Validasi</th>
                                  <th className="py-2.5 px-4 text-right">Keterangan</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100 dark:divide-zinc-700">
                                {rows.map((r, idx) => (
                                  <tr key={idx} className="hover:bg-slate-50 dark:bg-zinc-800/50 dark:hover:bg-zinc-700/30">
                                    <td className="py-1.5 px-2 font-medium text-slate-800 dark:text-zinc-100">
                                      <input value={r.name || ''} onChange={(e) => updatePreviewRow(r, 'name', e.target.value)} placeholder="[Kosong]"
                                        className="w-full min-w-[140px] bg-transparent border border-transparent hover:border-slate-200 dark:hover:border-zinc-600 focus:border-blue-400 focus:bg-white dark:focus:bg-zinc-900 rounded px-1.5 py-1 outline-none transition-colors placeholder:text-red-400 placeholder:italic" />
                                    </td>
                                    <td className="py-1.5 px-2 text-slate-600 dark:text-zinc-300">
                                      <input value={r.email || ''} onChange={(e) => updatePreviewRow(r, 'email', e.target.value)} placeholder="-"
                                        className="w-full min-w-[160px] bg-transparent border border-transparent hover:border-slate-200 dark:hover:border-zinc-600 focus:border-blue-400 focus:bg-white dark:focus:bg-zinc-900 rounded px-1.5 py-1 outline-none transition-colors" />
                                    </td>
                                    <td className="py-1.5 px-2 font-mono text-slate-600 dark:text-zinc-300">
                                      <input value={r.passport || ''} onChange={(e) => updatePreviewRow(r, 'passport', e.target.value.toUpperCase())} placeholder="[Kosong]"
                                        className="w-full min-w-[110px] bg-transparent border border-transparent hover:border-slate-200 dark:hover:border-zinc-600 focus:border-blue-400 focus:bg-white dark:focus:bg-zinc-900 rounded px-1.5 py-1 outline-none transition-colors placeholder:text-red-400 placeholder:italic" />
                                    </td>
                                    <td className="py-1.5 px-2 font-mono text-slate-600 dark:text-zinc-300">
                                      <input value={r.visa || ''} onChange={(e) => updatePreviewRow(r, 'visa', e.target.value)} placeholder="-"
                                        className="w-full min-w-[110px] bg-transparent border border-transparent hover:border-slate-200 dark:hover:border-zinc-600 focus:border-blue-400 focus:bg-white dark:focus:bg-zinc-900 rounded px-1.5 py-1 outline-none transition-colors" />
                                    </td>
                                    <td className="py-1.5 px-2">
                                      <input value={r.password || ''} onChange={(e) => updatePreviewRow(r, 'password', e.target.value)} placeholder="Otomatis"
                                        className="w-full min-w-[90px] font-mono text-[11px] bg-transparent border border-transparent hover:border-slate-200 dark:hover:border-zinc-600 focus:border-blue-400 focus:bg-white dark:focus:bg-zinc-900 rounded px-1.5 py-1 outline-none transition-colors placeholder:text-slate-400 placeholder:italic" />
                                    </td>
                                    <td className="py-1.5 px-2 text-slate-500 dark:text-zinc-400">
                                      <select value={r.gender || ''} onChange={(e) => updatePreviewRow(r, 'gender', e.target.value)}
                                        className="w-full min-w-[90px] bg-transparent border border-transparent hover:border-slate-200 dark:hover:border-zinc-600 focus:border-blue-400 focus:bg-white dark:focus:bg-zinc-900 rounded px-1 py-1 outline-none transition-colors cursor-pointer">
                                        <option value="">—</option>
                                        <option value="Laki-laki">Laki-laki</option>
                                        <option value="Perempuan">Perempuan</option>
                                      </select>
                                    </td>
                                    <td className="py-1.5 px-2 font-mono text-slate-500 dark:text-zinc-400">
                                      <input type="date" value={(r.entryMadinah || '').split('T')[0]} onChange={(e) => updatePreviewRow(r, 'entryMadinah', e.target.value)}
                                        className="w-full min-w-[130px] bg-transparent border border-transparent hover:border-slate-200 dark:hover:border-zinc-600 focus:border-blue-400 focus:bg-white dark:focus:bg-zinc-900 rounded px-1.5 py-1 outline-none transition-colors" />
                                    </td>
                                    {customFields.map(cf => {
                                      const cVal = r.customValues?.[cf.id] || '-';
                                      return (
                                        <td key={cf.id} className="py-2.5 px-3 text-slate-600 dark:text-zinc-300 font-medium truncate max-w-[120px]" title={cVal}>
                                          {cVal}
                                        </td>
                                      );
                                    })}
                                    <td className="py-2.5 px-3">
                                      <span className={`inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-semibold ${
                                        r.status === 'Valid' ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' :
                                        r.status === 'Duplikat' ? 'bg-amber-50 text-amber-700 border border-amber-100' :
                                        'bg-rose-50 text-rose-700 border border-rose-100'
                                      }`}>
                                        {r.status === 'Valid' && (
                                          <>
                                            <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                                            <span>Valid</span>
                                          </>
                                        )}
                                        {r.status === 'Duplikat' && (
                                          <>
                                            <span className="w-1.5 h-1.5 rounded-full bg-amber-500"></span>
                                            <span>Duplikat</span>
                                          </>
                                        )}
                                        {r.status === 'Tidak Lengkap' && (
                                          <>
                                            <span className="w-1.5 h-1.5 rounded-full bg-rose-500"></span>
                                            <span>Tidak Lengkap</span>
                                          </>
                                        )}
                                      </span>
                                    </td>
                                    <td className="py-2.5 px-4 text-right text-[11px] text-slate-500 dark:text-zinc-400">{r.reason}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* SECTION PANDUAN IMPORT SINKRON (EXCEL MANIFEST GUIDE) */}
              <div className="bg-slate-50 dark:bg-zinc-800/50 border border-slate-200 dark:border-zinc-600/60 rounded-2xl p-6 mt-8 space-y-6">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200 dark:border-zinc-600 pb-4">
                  <div className="flex items-center gap-2.5">
                    <div className="w-9 h-9 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center">
                      <BookOpen className="w-5 h-5" />
                    </div>
                    <div>
                      <h4 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">Panduan & Instruksi Sinkronisasi File Excel</h4>
                      <p className="text-xs text-slate-500 dark:text-zinc-400 mt-0.5">Instruksi pengisian format manifest jemaah dan sinkronisasi sistem.</p>
                    </div>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-[11px] text-slate-400 font-medium">Mulai dengan database kosong?</span>
                    <button
                      type="button"
                      onClick={clearAllJamaahs}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 border border-red-100 text-xs font-bold transition-all cursor-pointer shadow-3xs"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      <span>Kosongkan Semua Data Jemaah</span>
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                  {/* Card 1: Template Bersih */}
                  <div className="bg-white dark:bg-zinc-800 p-4 rounded-xl border border-slate-100 dark:border-zinc-700 shadow-3xs space-y-2.5">
                    <FileText className="w-5 h-5 text-blue-600" />
                    <h5 className="font-semibold text-slate-800 dark:text-zinc-100 text-xs">1. Unduh & Gunakan Template</h5>
                    <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                      Unduh template file excel resmi di atas. Kami telah membersihkan datanya dan hanya menyisakan format siap pakai dengan kolom yang telah disinkronkan sepenuhnya dengan sistem database utama.
                    </p>
                  </div>

                  {/* Card 2: Validasi Duplikat */}
                  <div className="bg-white dark:bg-zinc-800 p-4 rounded-xl border border-slate-100 dark:border-zinc-700 shadow-3xs space-y-2.5">
                    <Search className="w-5 h-5 text-blue-600" />
                    <h5 className="font-semibold text-slate-800 dark:text-zinc-100 text-xs">2. Validasi & Deteksi Duplikat</h5>
                    <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                      Sistem melakukan pencocokan otomatis pada nomor <strong>Paspor</strong> dan <strong>Visa</strong>. Jika nomor tersebut sudah ada di database, baris tersebut ditandai <span className="text-amber-600 font-semibold">Duplikat</span> untuk mencegah pendaftaran ganda.
                    </p>
                  </div>

                  {/* Card 3: Presisi Jam Masuk */}
                  <div className="bg-white dark:bg-zinc-800 p-4 rounded-xl border border-slate-100 dark:border-zinc-700 shadow-3xs space-y-2.5">
                    <Clock className="w-5 h-5 text-emerald-600" />
                    <h5 className="font-semibold text-slate-800 dark:text-zinc-100 text-xs">3. Presisi Jam Masuk Madinah</h5>
                    <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                      Gunakan format <code>YYYY-MM-DD HH:MM</code> (contoh: <code>2026-06-27 08:00</code>) di kolom Masuk. Jam kedatangan sangat krusial agar operator dapat menentukan jadwal booking Raudhah di aplikasi Nusuk dengan tepat.
                    </p>
                  </div>

                  {/* Card 4: Sandi & Travel Grup */}
                  <div className="bg-white dark:bg-zinc-800 p-4 rounded-xl border border-slate-100 dark:border-zinc-700 shadow-3xs space-y-2.5">
                    <Settings className="w-5 h-5 text-blue-600" />
                    <h5 className="font-semibold text-slate-800 dark:text-zinc-100 text-xs">4. Password Jemaah & Travel Grup</h5>
                    <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                      Grup travel dideteksi dari kolom <strong>Nama Travel</strong> secara otomatis. Sandi akses jemaah bisa diatur manual di kolom Excel, atau dikosongkan agar sistem men-generate PIN 6-digit otomatis.
                    </p>
                  </div>
                </div>

                <div className="bg-amber-50/50 rounded-xl p-4 border border-amber-100 text-[11px] text-amber-800 flex items-start gap-2.5">
                  <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                  <div className="space-y-1">
                    <span className="font-bold">Tips Sinkronisasi Sukses:</span>
                    <p className="leading-relaxed">
                      Anda bisa mengedit data jemaah yang memiliki status <strong>"Tidak Lengkap"</strong> atau <strong>"Duplikat"</strong> langsung di tabel pratinjau sebelum menekan tombol Impor. Tekan tombol <strong>"Impor Semua"</strong> untuk memasukkan seluruh jemaah berstatus Valid ke dalam antrean booking aktif.
                    </p>
                  </div>
                </div>
              </div>

            </div>
          )}

          {/* TAB 4: OPERATOR MANAGEMENT */}
          {currentTab === 'operator' && (
            <div className="space-y-6">
              
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                
                {/* Add Operator (Form) */}
                <div className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl shadow-xs p-5 lg:col-span-4 self-start">
                  <h3 className="font-semibold text-slate-800 dark:text-zinc-100 text-sm pb-3 border-b border-slate-100">Registrasi Operator Baru</h3>
                  
                  <form onSubmit={handleAddOperator} className="space-y-4 pt-4">
                    <div className="space-y-1">
                      <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Nama Lengkap</label>
                      <input
                        type="text"
                        required
                        value={newOperatorName}
                        onChange={(e) => setNewOperatorName(e.target.value)}
                        placeholder="Contoh: Ahmad Fauzi"
                        className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2 bg-slate-50 dark:bg-zinc-800/50 outline-hidden focus:border-red-500"
                      />
                    </div>

                    <div className="space-y-1">
                      <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Nomor HP / WhatsApp</label>
                      <input
                        type="text"
                        required
                        value={newOperatorPhone}
                        onChange={(e) => setNewOperatorPhone(e.target.value)}
                        placeholder="Contoh: +6281299998888"
                        className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2 bg-slate-50 dark:bg-zinc-800/50 outline-hidden focus:border-red-500"
                      />
                    </div>

                    <div className="space-y-1">
                      <div className="flex items-center justify-between">
                        <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Password Akses (Default: 123456)</label>
                        <button
                          type="button"
                          onClick={() => setShowNewOperatorPassword(!showNewOperatorPassword)}
                          className="text-[10px] text-red-600 font-semibold hover:text-red-700 transition-colors"
                        >
                          {showNewOperatorPassword ? 'Sembunyikan' : 'Lihat'}
                        </button>
                      </div>
                      <input
                        type={showNewOperatorPassword ? 'text' : 'password'}
                        required
                        value={newOperatorPassword}
                        onChange={(e) => setNewOperatorPassword(e.target.value)}
                        placeholder="Contoh: op123"
                        className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2 bg-slate-50 dark:bg-zinc-800/50 outline-hidden focus:border-red-500 font-mono"
                      />
                    </div>

                    <button
                      type="submit"
                      className="w-full py-2 bg-zinc-900 text-white rounded-lg text-xs font-semibold hover:bg-zinc-800 transition-colors shadow-xs"
                    >
                      Daftarkan Operator
                    </button>
                  </form>
                </div>

                {/* Operator List */}
                <div className="bg-white dark:bg-zinc-800 border border-slate-100 dark:border-zinc-700 rounded-xl shadow-xs p-5 lg:col-span-8">
                  <h3 className="font-semibold text-slate-800 dark:text-zinc-100 text-sm pb-3 border-b border-slate-100">Daftar Operator Aktif</h3>
                  
                  <div className="divide-y divide-slate-100 mt-2">
                    {operators.map(op => {
                      const assignedJamaahs = jamaahs.filter(j => j.operatorId === op.id);
                      const pendingCount = assignedJamaahs.filter(j => j.status !== 'QR Berhasil').length;
                      return (
                        <div key={op.id} className="py-4 flex items-center justify-between gap-4">
                          <div className="space-y-1">
                            <div className="flex items-center gap-2">
                              <span className="font-semibold text-slate-800 dark:text-zinc-100 text-sm">{op.name}</span>
                              <span className={`text-[10px] px-2 py-0.5 rounded-full font-medium ${
                                op.isActive ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' : 'bg-slate-100 dark:bg-zinc-700 text-slate-500 dark:text-zinc-400 border border-slate-200 dark:border-zinc-600/50'
                              }`}>
                                {op.isActive ? 'Aktif' : 'Nonaktif'}
                              </span>
                            </div>
                            <div className="text-xs text-slate-500 dark:text-zinc-400 flex flex-wrap items-center gap-x-3 gap-y-1">
                              <span className="flex items-center gap-1">
                                <Phone className="w-3.5 h-3.5 text-slate-400" />
                                <span>{op.phone}</span>
                              </span>
                              <span className="text-slate-300">•</span>
                              <span className="flex items-center gap-1 text-[11px] font-mono text-zinc-600 dark:text-zinc-300 bg-zinc-100 dark:bg-zinc-700 px-1.5 py-0.5 rounded">
                                <Key className="w-3 h-3 text-zinc-500 dark:text-zinc-400" /> Pass: <strong className="font-semibold text-zinc-900 dark:text-zinc-100">{op.password || '123456'}</strong>
                              </span>
                            </div>
                          </div>

                          <div className="flex items-center gap-4">
                            <div className="text-right text-xs">
                              <div className="text-slate-800 dark:text-zinc-100 font-medium">{assignedJamaahs.length} Ditugaskan</div>
                              <div className="text-[11px] text-amber-600 mt-0.5">{pendingCount} Butuh Barcode</div>
                            </div>

                            <div className="flex items-center gap-1.5">
                              {/* Toggle active / inactive status */}
                              <button
                                onClick={() => toggleOperatorStatus(op.id)}
                                className={`px-2.5 py-1 text-[11px] font-semibold rounded border transition-colors ${
                                  op.isActive 
                                    ? 'bg-slate-50 dark:bg-zinc-800/50 border-slate-200 dark:border-zinc-600 text-slate-600 dark:text-zinc-300 hover:bg-slate-100 dark:bg-zinc-700' 
                                    : 'bg-emerald-50 border-emerald-100 text-emerald-700 hover:bg-emerald-100'
                                }`}
                              >
                                {op.isActive ? 'Set Nonaktif' : 'Set Aktif'}
                              </button>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* TAB 5: SETTINGS */}
          {currentTab === 'settings' && (
            <div className="max-w-4xl mx-auto grid grid-cols-1 md:grid-cols-4 gap-6 bg-white dark:bg-zinc-900 border border-slate-200 dark:border-zinc-800 rounded-3xl shadow-xl p-6">
              {/* Settings Sidebar */}
              <div className="md:col-span-1 space-y-1 border-b md:border-b-0 md:border-r border-slate-100 dark:border-zinc-800 pb-4 md:pb-0 md:pr-4">
                <div className="mb-4">
                  <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">Pengaturan</h3>
                  <p className="text-[10px] text-slate-500 dark:text-zinc-400 mt-0.5">Konfigurasi & integrasi sistem.</p>
                </div>
                {[
                  { id: 'umum', label: 'Umum & Ekspor', icon: <Settings className="w-4 h-4" /> },
                  { id: 'database', label: 'Database & Sync', icon: <Building2 className="w-4 h-4" /> },
                  { id: 'nusuk', label: 'Kuota & Jadwal', icon: <Clock className="w-4 h-4" /> },
                  { id: 'gemini', label: 'Gemini AI', icon: <Sparkles className="w-4 h-4" /> },
                  { id: 'kolom', label: 'Kolom Kustom', icon: <Grid className="w-4 h-4" /> },
                ].map(tab => (
                  <button
                    key={tab.id}
                    type="button"
                    onClick={() => setActiveSettingSubTab(tab.id)}
                    className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-semibold transition-all cursor-pointer ${
                      activeSettingSubTab === tab.id
                        ? 'bg-red-50 dark:bg-red-500/10 text-red-700 dark:text-red-400'
                        : 'text-slate-600 dark:text-zinc-400 hover:bg-slate-50 dark:hover:bg-zinc-800/40 hover:text-slate-900 dark:hover:text-zinc-200'
                    }`}
                  >
                    {tab.icon}
                    <span>{tab.label}</span>
                  </button>
                ))}
              </div>

              {/* Settings Content Area */}
              <div className="md:col-span-3 space-y-6">
                {/* SUB TAB 1: UMUM */}
                {activeSettingSubTab === 'umum' && (
                  <div className="space-y-4 animate-in fade-in duration-200">
                    <div>
                      <h4 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">Pengaturan Umum</h4>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 mt-0.5">Konfigurasi nama agen travel dan preferensi tampilan ekspor.</p>
                    </div>
                    <div className="space-y-4 divide-y divide-slate-100 dark:divide-zinc-800">
                      {/* Row 1: Travel Name */}
                      <div className="pt-2 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Nama Travel Agen</span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Nama travel yang akan tertera pada ekspor manifest.</p>
                        </div>
                        <input
                          type="text"
                          value={settingsTravelName}
                          onChange={(e) => setSettingsTravelName(e.target.value.toUpperCase())}
                          className="text-xs border border-slate-200 dark:border-zinc-700 rounded-lg p-2 w-full sm:w-64 bg-white dark:bg-zinc-800 text-slate-800 dark:text-zinc-100 outline-hidden focus:border-red-500"
                        />
                      </div>

                      {/* Row 3: Admin Password */}
                      <div className="pt-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Password Utama Admin</span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Password untuk mengunci dan membatasi akses Admin.</p>
                        </div>
                        <input
                          type="text"
                          value={adminPassword}
                          onChange={(e) => setAdminPassword(e.target.value)}
                          className="text-xs border border-slate-200 dark:border-zinc-700 rounded-lg p-2 w-full sm:w-64 bg-white dark:bg-zinc-800 text-slate-800 dark:text-zinc-100 outline-hidden focus:border-red-500 font-mono"
                        />
                      </div>

                      {/* Row 4: Tanggal Referensi Hari Ini */}
                      <div className="pt-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Tanggal Referensi Hari Ini</span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Tanggal acuan "Hari Ini" untuk menghitung prioritas masuk Madinah.</p>
                        </div>
                        <input
                          type="date"
                          value={settingsReferenceDate}
                          onChange={(e) => setSettingsReferenceDate(e.target.value)}
                          className="text-xs border border-slate-200 dark:border-zinc-700 rounded-lg p-2 w-full sm:w-64 bg-white dark:bg-zinc-800 text-slate-800 dark:text-zinc-100 outline-hidden focus:border-red-500 font-mono text-center font-semibold"
                        />
                      </div>

                      {/* Excel Columns Config */}
                      <div className="pt-4 space-y-3">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Kustomisasi Ekspor Excel</span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Pilih kolom data jemaah yang ingin disertakan ke dalam file Excel hasil ekspor.</p>
                        </div>
                        <div className="bg-slate-50 dark:bg-zinc-800/30 p-4 border border-slate-200/60 dark:border-zinc-700/60 rounded-xl text-xs">
                          <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                            {[
                              { key: 'name', label: 'Nama Lengkap' },
                              { key: 'passport', label: 'Nomor Paspor' },
                              { key: 'visa', label: 'Nomor Visa' },
                              { key: 'gender', label: 'Jenis Kelamin' },
                              { key: 'phone', label: 'No WhatsApp' },
                              { key: 'email', label: 'Email' },
                              { key: 'entryMadinah', label: 'Tanggal Masuk' },
                              { key: 'exitMadinah', label: 'Tanggal Keluar' },
                              { key: 'travel', label: 'Nama Travel' },
                              { key: 'password', label: 'Password Akses' },
                              { key: 'status', label: 'Status Booking' },
                              { key: 'operator', label: 'Operator' },
                              { key: 'notes', label: 'Catatan' },
                              ...customFields.map(cf => ({ key: cf.id, label: cf.label }))
                            ].map((item) => (
                              <label key={item.key} className="flex items-center gap-2 cursor-pointer font-medium text-slate-700 dark:text-zinc-300 hover:text-slate-900 dark:hover:text-zinc-100">
                                <input 
                                  type="checkbox" 
                                  checked={!!(settingsExportColumns as any)[item.key]}
                                  onChange={(e) => setSettingsExportColumns({ ...settingsExportColumns, [item.key]: e.target.checked })}
                                  className="rounded text-red-600 focus:ring-red-500 border-slate-300 dark:border-zinc-600 bg-white dark:bg-zinc-800"
                                />
                                <span className="truncate" title={item.label}>{item.label}</span>
                              </label>
                            ))}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {/* SUB TAB 2: DATABASE */}
                {activeSettingSubTab === 'database' && (
                  <div className="space-y-4 animate-in fade-in duration-200">
                    <div>
                      <h4 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">Database & Sinkronisasi Online</h4>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 mt-0.5">Atur integrasi Supabase Anda untuk mengaktifkan kolaborasi multi-user.</p>
                    </div>

                    <div className="space-y-4">
                      {/* Status indicator */}
                      <div className="flex items-center justify-between p-3 rounded-xl border border-slate-100 dark:border-zinc-800 bg-slate-50 dark:bg-zinc-800/30">
                        <span className="text-xs font-semibold text-slate-700 dark:text-zinc-300">Status Koneksi Supabase</span>
                        <span className={`inline-flex items-center px-2.5 py-1 rounded-full text-[10px] font-bold ${
                          isSupabaseConnected 
                            ? 'bg-emerald-50 dark:bg-emerald-500/10 text-emerald-700 dark:text-emerald-400 border border-emerald-200/50' 
                            : isSupabaseLoading 
                            ? 'bg-amber-50 dark:bg-amber-500/10 text-amber-700 dark:text-amber-400 border border-amber-200/50' 
                            : 'bg-rose-50 dark:bg-rose-500/10 text-rose-700 dark:text-rose-400 border border-rose-200/50'
                        }`}>
                          <span className={`h-1.5 w-1.5 rounded-full mr-1.5 ${
                            isSupabaseConnected ? 'bg-emerald-500' : isSupabaseLoading ? 'bg-amber-500 animate-pulse' : 'bg-rose-500'
                          }`}></span>
                          {isSupabaseConnected ? 'Terhubung' : isSupabaseLoading ? 'Menghubungkan...' : 'Terputus'}
                        </span>
                      </div>

                      {/* Connection Fields */}
                      <div className="bg-slate-50 dark:bg-zinc-800/30 p-4 border border-slate-200/60 dark:border-zinc-700/60 rounded-xl space-y-3 text-xs">
                        <div className="grid grid-cols-1 gap-3">
                          <div className="space-y-1">
                            <label className="text-[10px] font-semibold text-slate-500 dark:text-zinc-400">SUPABASE URL</label>
                            <input 
                              type="text"
                              value={settingsSupabaseUrl}
                              onChange={(e) => setSettingsSupabaseUrl(e.target.value.trim())}
                              placeholder="https://xxxx.supabase.co"
                              className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2 bg-white dark:bg-zinc-800 text-slate-800 dark:text-zinc-100 focus:border-red-500 outline-hidden font-mono"
                            />
                          </div>
                          <div className="space-y-1">
                            <label className="text-[10px] font-semibold text-slate-500 dark:text-zinc-400">SUPABASE ANON KEY</label>
                            <input 
                              type="password"
                              value={settingsSupabaseAnonKey}
                              onChange={(e) => setSettingsSupabaseAnonKey(e.target.value.trim())}
                              placeholder="eyJhbGciOi..."
                              className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2 bg-white dark:bg-zinc-800 text-slate-800 dark:text-zinc-100 focus:border-red-500 outline-hidden font-mono"
                            />
                          </div>
                        </div>
                        {supabaseError && (
                          <p className="text-[10px] text-rose-600 dark:text-rose-400 font-semibold bg-rose-50 dark:bg-rose-500/10 border border-rose-100 dark:border-rose-500/20 p-2 rounded">
                            ⚠️ Error: {supabaseError}
                          </p>
                        )}
                        <div className="flex gap-2 justify-end pt-1">
                          <button
                            type="button"
                            onClick={handleMigrateToSupabase}
                            disabled={!isSupabaseConnected || (jamaahs.length === 0 && operators.length === 0)}
                            className="px-3.5 py-1.5 rounded-lg border border-blue-200 text-blue-700 bg-white hover:bg-blue-50 dark:border-blue-500/30 dark:text-blue-400 dark:bg-zinc-800 disabled:opacity-50 disabled:cursor-not-allowed transition-colors text-xs font-semibold cursor-pointer"
                          >
                            Migrasikan Data Lokal ke Supabase
                          </button>
                        </div>
                      </div>

                      {/* Reset DB Row */}
                      <div className="pt-4 border-t border-slate-100 dark:border-zinc-800 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-red-700 dark:text-red-400">Pembersihan / Reset Database</span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Kosongkan data manifest jemaah, data operator, atau kembalikan semua pengaturan ke semula.</p>
                        </div>
                        <button
                          type="button"
                          onClick={openResetModal}
                          className="px-3.5 py-1.5 rounded-lg bg-red-50 hover:bg-red-100 dark:bg-red-500/10 dark:hover:bg-red-500/20 text-red-700 dark:text-red-400 text-xs font-semibold border border-red-200/60 dark:border-red-500/30 transition-colors shrink-0 cursor-pointer"
                        >
                          Reset Database / Sistem
                        </button>
                      </div>
                    </div>
                  </div>
                )}

                {/* SUB TAB 3: NUSUK & JADWAL */}
                {activeSettingSubTab === 'nusuk' && (
                  <div className="space-y-4 animate-in fade-in duration-200">
                    <div>
                      <h4 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">Kuota & Batasan Jadwal Nusuk</h4>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 mt-0.5">Atur limit booking harian dan slot jadwal masuk raudhah untuk verifikasi otomatis.</p>
                    </div>

                    <div className="space-y-4 divide-y divide-slate-100 dark:divide-zinc-800">
                      {/* Row 2: Nusuk Limit */}
                      <div className="pt-2 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Limit Booking Harian Nusuk</span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Pemberitahuan kuota harian per operator.</p>
                        </div>
                        <input
                          type="number"
                          value={settingsNusukLimit}
                          onChange={(e) => setSettingsNusukLimit(parseInt(e.target.value) || 0)}
                          className="text-xs border border-slate-200 dark:border-discord-onyx rounded-xl p-2 w-full sm:w-64 bg-white dark:bg-discord-indigo text-slate-800 dark:text-white outline-hidden focus:border-discord-blurple focus:ring-2 focus:ring-discord-blurple/20 transition-all"
                        />
                      </div>

                      {/* Row: QR Distribution Lead Time */}
                      <div className="pt-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100 flex items-center gap-1.5">
                            <Bell className="w-3.5 h-3.5 text-discord-cyan" />
                            Jeda Distribusi QR (jam sebelum slot)
                          </span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Reminder untuk masuk Nusuk, download &amp; bagikan QR ke jemaah. Default 2 jam sebelum jam slot Raudhah (waktu Madinah).</p>
                        </div>
                        <input
                          type="number"
                          min={0}
                          step={0.5}
                          value={settingsQrLeadHours}
                          onChange={(e) => setSettingsQrLeadHours(Math.max(0, parseFloat(e.target.value) || 0))}
                          className="text-xs border border-slate-200 dark:border-discord-onyx rounded-xl p-2 w-full sm:w-64 bg-white dark:bg-discord-indigo text-slate-800 dark:text-white outline-hidden focus:border-discord-blurple focus:ring-2 focus:ring-discord-blurple/20 transition-all"
                        />
                      </div>

                      {/* Row: Notifikasi & Suara Alarm */}
                      <div className="pt-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-t border-slate-100 dark:border-zinc-800">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100 flex items-center gap-1.5">
                            <Volume2 className="w-3.5 h-3.5 text-blue-500" />
                            Notifikasi &amp; Suara Alarm Raudhah
                          </span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Mainkan suara alarm dan tampilkan pemberitahuan browser saat jeda waktu distribusi tiba.</p>
                        </div>
                        <div className="flex flex-wrap items-center gap-3">
                          {/* Test Button */}
                          <button
                            type="button"
                            onClick={() => {
                              playNotificationSound();
                              if (typeof Notification !== 'undefined') {
                                Notification.requestPermission().then(permission => {
                                  if (permission === 'granted') {
                                    new Notification('Uji Coba Notifikasi Raudhah', {
                                      body: 'Suara alarm & pemberitahuan berhasil diaktifkan!',
                                    });
                                  } else {
                                    alert('Izin notifikasi diblokir oleh browser. Harap aktifkan di pengaturan browser Anda.');
                                  }
                                });
                              }
                            }}
                            className="px-3 py-1.5 rounded-lg border border-slate-200 dark:border-zinc-700 bg-slate-50 dark:bg-zinc-800 text-slate-700 dark:text-zinc-200 text-[10px] font-semibold hover:bg-slate-100 dark:hover:bg-zinc-700 transition-all cursor-pointer"
                          >
                            🔔 Uji Coba Suara &amp; Pop-up
                          </button>

                          {/* Toggle Switch */}
                          <label className="relative inline-flex items-center cursor-pointer select-none">
                            <input
                              type="checkbox"
                              checked={settingsEnableSound}
                              onChange={(e) => setSettingsEnableSound(e.target.checked)}
                              className="sr-only peer"
                            />
                            <div className="w-9 h-5 bg-slate-200 dark:bg-zinc-700 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-blue-600"></div>
                            <span className="ml-2 text-xs font-semibold text-slate-700 dark:text-zinc-300">
                              {settingsEnableSound ? 'Suara Aktif' : 'Suara Senyap'}
                            </span>
                          </label>
                        </div>
                      </div>

                      {/* Row: Default Slot Jadwal Raudhah */}
                      <div className="pt-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100 flex items-center gap-1.5">
                            <Clock className="w-3.5 h-3.5 text-discord-green" />
                            Slot Jadwal Masuk Raudhah (Default)
                          </span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Jadwal slot default untuk pengisian otomatis jemaah baru.</p>
                        </div>
                        <input
                          type="datetime-local"
                          value={settingsDefaultRaudhahSlot}
                          onChange={(e) => setSettingsDefaultRaudhahSlot(e.target.value)}
                          className="text-xs border border-slate-200 dark:border-discord-onyx rounded-xl p-2 w-full sm:w-64 bg-white dark:bg-discord-indigo text-slate-800 dark:text-white outline-hidden focus:border-discord-blurple focus:ring-2 focus:ring-discord-blurple/20 font-mono transition-all"
                        />
                      </div>

                      {/* Row: Password Default Global */}
                      <div className="pt-4 flex flex-col sm:flex-row sm:items-end justify-between gap-4">
                        <div className="space-y-1 flex-1">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Password Global Jemaah</span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Password default saat membuat akun jemaah massal.</p>
                          <div className="relative max-w-xs">
                            <Key className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-discord-magenta pointer-events-none" />
                            <input
                              type={showDefaultPassword ? 'text' : 'password'}
                              value={settingsDefaultPassword}
                              onChange={(e) => setSettingsDefaultPassword(e.target.value)}
                              placeholder="mis. Visa2424@"
                              className="w-full text-xs border border-slate-200 dark:border-discord-onyx rounded-xl py-2.5 pl-9 pr-10 bg-slate-50 dark:bg-discord-indigo/40 text-slate-800 dark:text-white outline-hidden focus:border-discord-blurple focus:ring-2 focus:ring-discord-blurple/20 font-mono transition-all"
                            />
                            <button
                              type="button"
                              onClick={() => setShowDefaultPassword(!showDefaultPassword)}
                              className="absolute right-2.5 top-1/2 -translate-y-1/2 p-1 text-slate-400 dark:text-zinc-500 hover:text-discord-blurple transition-colors cursor-pointer"
                            >
                              {showDefaultPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                            </button>
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={applyGlobalPassword}
                          className="px-3.5 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 dark:bg-discord-indigo dark:hover:bg-discord-blurple/10 text-slate-800 dark:text-white text-xs font-semibold border border-slate-200 dark:border-discord-blurple/30 transition-all shrink-0 cursor-pointer inline-flex items-center gap-1.5 shadow-xs"
                        >
                          <Key className="w-3.5 h-3.5 text-discord-magenta" />
                          <span>Set ke Semua Jemaah</span>
                        </button>
                      </div>

                      {/* Akun & Jam Masuk Default */}
                      <div className="pt-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Buat Akun Login Otomatis</span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Generate email login default @mailnesia.com (dari nama) &amp; samakan password ke semua jemaah.</p>
                        </div>
                        <button
                          type="button"
                          onClick={applyDefaultAccounts}
                          className="px-3.5 py-1.5 rounded-xl bg-discord-blurple hover:bg-discord-blurple/90 text-white text-xs font-semibold border border-transparent transition-all shrink-0 cursor-pointer inline-flex items-center gap-1.5 shadow-xs"
                        >
                          <Sparkles className="w-3.5 h-3.5 text-discord-green" />
                          <span>Buat Akun Login Otomatis</span>
                        </button>
                      </div>

                    </div>
                  </div>
                )}

                {/* SUB TAB 4: GEMINI AI */}
                {activeSettingSubTab === 'gemini' && (
                  <div className="space-y-4 animate-in fade-in duration-200">
                    <div>
                      <h4 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">Integrasi Google Gemini AI</h4>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 mt-0.5">Konfigurasi API Key untuk deteksi OCR otomatis gambar visa jemaah.</p>
                    </div>

                    <div className="space-y-4 divide-y divide-slate-100 dark:divide-zinc-800">
                      {/* Row: Gemini API Key */}
                      <div className="pt-2 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Google Gemini API Key</span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Kunci akses untuk mendeteksi data visa secara otomatis.</p>
                        </div>
                        <div className="relative w-full sm:w-64">
                          <input
                            type={showGeminiApiKey ? 'text' : 'password'}
                            value={settingsGeminiApiKey}
                            onChange={(e) => setSettingsGeminiApiKey(e.target.value)}
                            placeholder="Masukkan API Key Gemini..."
                            className="text-xs border border-slate-200 dark:border-discord-onyx rounded-xl p-2 pr-10 w-full outline-hidden bg-white dark:bg-discord-indigo text-slate-800 dark:text-white focus:border-discord-blurple focus:ring-2 focus:ring-discord-blurple/20 font-mono transition-all"
                          />
                          <button
                            type="button"
                            onClick={() => setShowGeminiApiKey(!showGeminiApiKey)}
                            className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-discord-blurple dark:text-zinc-300 p-1 cursor-pointer"
                          >
                            {showGeminiApiKey ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                          </button>
                        </div>
                      </div>

                      {/* Row: Gemini Model Selection */}
                      <div className="pt-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-0.5">
                          <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Google Gemini Model</span>
                          <p className="text-[11px] text-slate-500 dark:text-zinc-400">Pilih tipe model AI yang akan digunakan.</p>
                        </div>
                        <div className="w-full sm:w-64">
                          <select
                            value={settingsGeminiModel}
                            onChange={(e) => setSettingsGeminiModel(e.target.value)}
                            className="text-xs border border-discord-onyx rounded-xl p-2 w-full outline-hidden bg-white dark:bg-discord-indigo text-slate-800 dark:text-white focus:border-discord-blurple focus:ring-2 focus:ring-discord-blurple/20 font-medium cursor-pointer transition-all"
                          >
                            <option value="gemini-2.0-flash">Gemini 2.0 Flash (Sangat Cepat & Direkomendasikan)</option>
                            <option value="gemini-2.5-flash">Gemini 2.5 Flash (Lebih Baru)</option>
                            <option value="gemini-3.5-flash">Gemini 3.5 Flash (Terbaru & Cerdas)</option>
                            <option value="gemini-1.5-flash">Gemini 1.5 Flash (Legacy)</option>
                            <option value="gemini-1.5-pro">Gemini 1.5 Pro (Legacy)</option>
                          </select>
                        </div>
                      </div>



                      {/* Quota & Usage Panel */}
                      <div className="pt-4">
                        <div className="bg-slate-50 dark:bg-discord-indigo/40 p-4 border border-slate-200/60 dark:border-discord-onyx rounded-xl space-y-4 text-xs">
                          <div className="flex items-center gap-2 font-bold text-slate-800 dark:text-white">
                            <Gauge className="w-4 h-4 text-discord-blurple" />
                            <span>Status Kuota & Penggunaan API ({settingsGeminiModel})</span>
                          </div>

                          {/* Indikator rate-limit NYATA: muncul saat request terakhir benar-benar ditolak 429 */}
                          {isRateLimitedNow ? (
                            <div className="flex items-center justify-between gap-2 p-2.5 rounded-lg bg-red-50 dark:bg-red-900/30 border border-red-200 dark:border-red-700/50">
                              <span className="flex items-center gap-2 text-[11px] font-bold text-red-600 dark:text-red-400">
                                <span className="animate-pulse">🔴</span>
                                Model sedang KENA RATE LIMIT
                              </span>
                              <span className="text-[11px] font-semibold text-red-500 dark:text-red-300">
                                Tunggu ± {rateLimitSecondsLeft} detik
                              </span>
                            </div>
                          ) : (
                            <div className="flex items-center gap-2 p-2.5 rounded-lg bg-emerald-50 dark:bg-emerald-900/20 border border-emerald-200 dark:border-emerald-700/40">
                              <span className="text-[11px] font-semibold text-emerald-600 dark:text-emerald-400 flex items-center gap-2">
                                <span>🟢</span>
                                Model siap (tidak ada penolakan rate limit terbaru)
                              </span>
                            </div>
                          )}

                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                            {/* Minute Limit (RPM) */}
                            <div className="space-y-1.5 p-3 bg-white dark:bg-discord-onyx/50 border border-slate-200/50 dark:border-discord-onyx rounded-lg">
                              <div className="flex justify-between items-center text-[11px] font-semibold text-slate-500 dark:text-zinc-400">
                                <span className="flex items-center gap-1">
                                  <Clock className="w-3.5 h-3.5 text-discord-cyan" />
                                  Kecepatan (Requests Per Minute)
                                </span>
                                <span>{currentRPM} / {limitRPM} RPM</span>
                              </div>
                              <div className="h-2 w-full bg-slate-100 dark:bg-discord-black rounded-full overflow-hidden">
                                <div 
                                  className={`h-full transition-all duration-500 ${
                                    currentRPM >= limitRPM ? 'bg-red-500' : currentRPM >= limitRPM * 0.75 ? 'bg-discord-magenta' : 'bg-discord-green'
                                  }`}
                                  style={{ width: `${rpmPercent}%` }}
                                ></div>
                              </div>
                              <div className="flex justify-between items-center text-[10px] text-slate-400">
                                <span>Batas: {limitRPM} request / menit</span>
                                {currentRPM >= limitRPM && <span className="text-red-500 font-semibold animate-pulse">⚠️ Terlampaui (Cooldown)</span>}
                              </div>
                            </div>

                            {/* Daily Limit (RPD) */}
                            <div className="space-y-1.5 p-3 bg-white dark:bg-discord-onyx/50 border border-slate-200/50 dark:border-discord-onyx rounded-lg">
                              <div className="flex justify-between items-center text-[11px] font-semibold text-slate-500 dark:text-zinc-400">
                                <span className="flex items-center gap-1">
                                  <Activity className="w-3.5 h-3.5 text-discord-magenta" />
                                  Kuota Harian (Requests Per Day)
                                </span>
                                <span>{currentRPD} / {limitRPD} RPD</span>
                              </div>
                              <div className="h-2 w-full bg-slate-100 dark:bg-discord-black rounded-full overflow-hidden">
                                <div 
                                  className={`h-full transition-all duration-500 ${
                                    currentRPD >= limitRPD ? 'bg-red-500' : currentRPD >= limitRPD * 0.75 ? 'bg-discord-magenta' : 'bg-discord-blurple'
                                  }`}
                                  style={{ width: `${rpdPercent}%` }}
                                ></div>
                              </div>
                              <div className="flex justify-between items-center text-[10px] text-slate-400">
                                <span>Batas: {limitRPD} request / hari</span>
                                <span>Token: {limitTPM}</span>
                              </div>
                            </div>
                          </div>

                          {/* Live Warnings */}
                          {currentRPM >= limitRPM * 0.8 && (
                            <div className="p-2.5 bg-amber-50 dark:bg-amber-500/10 border border-amber-200/50 dark:border-amber-500/20 text-amber-800 dark:text-amber-300 rounded-lg text-[10px] leading-relaxed flex gap-1.5 items-start">
                              <span>⚠️</span>
                              <div>
                                <strong>Peringatan RPM Tinggi:</strong> Penggunaan menit ini sudah mendekati batas {limitRPM} RPM. Pengunggahan file multipage atau batch scan berikutnya mungkin akan mengalami penundaan (cooldown) otomatis dari sistem untuk mencegah error 429.
                              </div>
                            </div>
                          )}

                          {currentRPD >= limitRPD && (
                            <div className="p-2.5 bg-rose-50 dark:bg-rose-500/10 border border-rose-200/50 dark:border-rose-500/20 text-rose-800 dark:text-rose-300 rounded-lg text-[10px] leading-relaxed flex gap-1.5 items-start">
                              <span>🚫</span>
                              <div>
                                <strong>Kuota Harian Habis:</strong> Penggunaan hari ini telah mencapai batas {limitRPD} request. API Key Anda akan menolak permintaan pemindaian baru hingga hari esok (reset waktu lokal). Harap gunakan API Key berbayar (pay-as-you-go) untuk kuota tanpa batas.
                              </div>
                            </div>
                          )}
                        </div>
                      </div>

                      {/* Row: Gemini API Diagnostics */}
                      <div className="pt-4 space-y-3">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                          <div className="space-y-0.5">
                            <span className="text-xs font-semibold text-slate-800 dark:text-zinc-100">Diagnostik Koneksi & Status Kuota</span>
                            <p className="text-[11px] text-slate-500 dark:text-zinc-400">Uji keaktifan API Key, ketersediaan model, dan limit kuota gratis.</p>
                          </div>
                          <button
                            type="button"
                            disabled={isTestingApi || !settingsGeminiApiKey}
                            onClick={handleTestGeminiConnection}
                            className="px-3.5 py-1.5 rounded-xl bg-discord-blurple hover:bg-discord-blurple/90 text-white text-xs font-semibold disabled:bg-slate-100 disabled:text-slate-400 disabled:cursor-not-allowed transition-all shrink-0 cursor-pointer flex items-center gap-1.5 shadow-xs"
                          >
                            {isTestingApi ? (
                              <>
                                <Loader2 className="w-3.5 h-3.5 animate-spin text-zinc-500" />
                                <span>Menguji...</span>
                              </>
                            ) : (
                              <span>Uji Koneksi & Kuota</span>
                            )}
                          </button>
                        </div>

                        {apiTestResult && (
                          <div className="p-4 bg-slate-50 dark:bg-discord-indigo border border-slate-200 dark:border-discord-blurple/30 rounded-xl space-y-3 animate-in fade-in slide-in-from-top-2 duration-200 text-xs">
                            <div className="flex items-start gap-2.5">
                              <div className={`mt-0.5 p-1 rounded-full ${apiTestResult.status === 'success' ? 'bg-emerald-100 text-emerald-700 dark:bg-discord-green/10 dark:text-discord-green' : 'bg-red-100 text-red-700 dark:bg-red-950/30 dark:text-red-400'}`}>
                                <CheckCircle2 className="w-4 h-4" />
                              </div>
                              <div className="space-y-0.5 flex-1">
                                <span className="font-bold text-slate-800 dark:text-zinc-100">Status Server Google API</span>
                                <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">{apiTestResult.message}</p>
                              </div>
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                )}

                {/* SUB TAB 5: KOLOM KUSTOM */}
                {activeSettingSubTab === 'kolom' && (
                  <div className="space-y-4 animate-in fade-in duration-200">
                    <div>
                      <h4 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">Kolom Kustom Jemaah (Metadata)</h4>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 mt-0.5">Tambahkan kolom tambahan sendiri yang spesifik (misal: Nomor Kamar, Penerbangan, Ukuran Baju) yang akan otomatis muncul di form jemaah.</p>
                    </div>

                    <div className="bg-slate-50 dark:bg-zinc-800/30 p-4 border border-slate-200/60 dark:border-zinc-700/60 rounded-xl space-y-3 text-xs">
                      {/* Add Field Form */}
                      <div className="flex gap-2">
                        <input
                          type="text"
                          placeholder="Nama kolom kustom (contoh: Nomor Kamar)..."
                          value={newCustomFieldLabel}
                          onChange={(e) => setNewCustomFieldLabel(e.target.value)}
                          className="flex-1 text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2 bg-white dark:bg-zinc-800 text-slate-800 dark:text-zinc-100 focus:border-red-500 outline-hidden font-medium"
                        />
                        <button
                          type="button"
                          onClick={handleAddCustomField}
                          className="px-3.5 py-2 rounded-lg bg-zinc-900 text-white text-xs font-semibold hover:bg-zinc-800 transition-colors shrink-0 cursor-pointer"
                        >
                          + Tambah Kolom
                        </button>
                      </div>

                      {/* Active Fields List */}
                      {customFields.length === 0 ? (
                        <p className="text-[11px] text-slate-500 dark:text-zinc-400 text-center py-2 italic font-medium">Belum ada kolom kustom yang ditambahkan.</p>
                      ) : (
                        <div className="space-y-2 border-t dark:border-zinc-700 pt-2 animate-in fade-in duration-200">
                          <span className="font-bold text-slate-800 dark:text-zinc-200 block mb-1">Kolom Kustom Aktif Anda:</span>
                          <div className="grid grid-cols-1 gap-2">
                            {customFields.map((cf) => (
                              <div
                                key={cf.id}
                                className="flex items-center justify-between p-2 rounded-lg border border-slate-200 dark:border-zinc-700 bg-white dark:bg-zinc-800/50"
                              >
                                <span className="font-medium text-slate-700 dark:text-zinc-200">{cf.label}</span>
                                <button
                                  type="button"
                                  onClick={() => handleDeleteCustomField(cf.id)}
                                  className="text-[10px] text-red-600 hover:text-red-700 font-semibold px-1 cursor-pointer"
                                >
                                  Hapus
                                </button>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {/* SUB TAB 6: RESET DB */}
                {activeSettingSubTab === 'reset' && (
                  <div className="space-y-4 animate-in fade-in duration-200">
                    <div>
                      <h4 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">Pembersihan Database</h4>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 mt-0.5">Kelola penyimpanan lokal browser Anda.</p>
                    </div>

                    <div className="pt-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                      <div className="space-y-0.5">
                        <span className="text-xs font-semibold text-red-600">Pembersihan Database</span>
                        <p className="text-[11px] text-slate-500 dark:text-zinc-400">Pilih bagian penyimpanan lokal browser (IndexedDB &amp; LocalStorage) yang ingin disetel ulang, lalu konfirmasi.</p>
                      </div>
                      <button
                        onClick={openResetModal}
                        className="px-3.5 py-1.5 rounded-lg bg-red-50 dark:bg-red-500/10 hover:bg-red-100 dark:hover:bg-red-500/20 text-red-700 dark:text-red-400 text-xs font-semibold border border-red-200/50 dark:border-red-500/30 transition-colors shrink-0 cursor-pointer inline-flex items-center gap-1.5"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                        <span>Pilih &amp; Bersihkan...</span>
                      </button>
                    </div>
                  </div>
                )}

              </div>
            </div>
          )}

          {/* Database Cleanup Modal */}
          {showResetModal && (
            <div
              className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in duration-150"
              onClick={() => setShowResetModal(false)}
            >
              <div
                className="bg-white dark:bg-zinc-900 rounded-2xl shadow-2xl w-full max-w-md border border-slate-200 dark:border-zinc-700 overflow-hidden animate-in zoom-in-95 duration-200"
                onClick={(e) => e.stopPropagation()}
              >
                {/* Header */}
                <div className="flex items-start justify-between gap-3 p-5 border-b border-slate-100 dark:border-zinc-800">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-xl bg-red-50 dark:bg-red-500/10 border border-red-100 dark:border-red-500/20 flex items-center justify-center shrink-0">
                      <Trash2 className="w-5 h-5 text-red-600 dark:text-red-400" />
                    </div>
                    <div>
                      <h3 className="text-sm font-bold text-slate-800 dark:text-zinc-100">Pembersihan Database</h3>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400">Pilih data yang ingin disetel ulang ke default.</p>
                    </div>
                  </div>
                  <button
                    onClick={() => setShowResetModal(false)}
                    className="p-1 rounded-md text-slate-400 hover:text-slate-700 dark:hover:text-zinc-200 hover:bg-slate-100 dark:hover:bg-zinc-800 transition-colors cursor-pointer"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>

                {/* Scope options */}
                <div className="p-5 space-y-2.5">
                  {[
                    { key: 'jamaah' as const, title: 'Data Jamaah', desc: 'Seluruh manifes jamaah (IndexedDB) kembali ke data awal.' },
                    { key: 'operators' as const, title: 'Data Operator', desc: 'Daftar operator & status penugasan dihapus.' },
                    { key: 'settings' as const, title: 'Pengaturan Aplikasi', desc: 'Nama travel, kuota, tanggal acuan, slot, kolom ekspor & field kustom.' },
                    { key: 'credentials' as const, title: 'Kredensial & API', desc: 'Password admin (ke admin123) & API Key Gemini.' },
                  ].map(opt => {
                    const checked = resetScopes[opt.key];
                    return (
                      <label
                        key={opt.key}
                        className={`flex items-start gap-3 p-3 rounded-xl border cursor-pointer transition-colors ${
                          checked
                            ? 'bg-red-50/60 dark:bg-red-500/10 border-red-200 dark:border-red-500/30'
                            : 'bg-slate-50 dark:bg-zinc-800/50 border-slate-200 dark:border-zinc-700 hover:bg-slate-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={(e) => setResetScopes(prev => ({ ...prev, [opt.key]: e.target.checked }))}
                          className="mt-0.5 h-4 w-4 rounded border-slate-300 dark:border-zinc-600 text-red-600 focus:ring-red-500 cursor-pointer shrink-0"
                        />
                        <div className="min-w-0">
                          <div className="text-xs font-semibold text-slate-800 dark:text-zinc-100">{opt.title}</div>
                          <div className="text-[10.5px] text-slate-500 dark:text-zinc-400 leading-snug mt-0.5">{opt.desc}</div>
                        </div>
                      </label>
                    );
                  })}
                </div>

                {/* Confirmation */}
                <div className="px-5 pb-5 space-y-3">
                  <div className="flex items-start gap-2 text-[11px] text-amber-800 dark:text-amber-300 bg-amber-50 dark:bg-amber-500/10 border border-amber-100 dark:border-amber-500/20 px-3 py-2 rounded-lg">
                    <AlertTriangle className="w-3.5 h-3.5 shrink-0 mt-0.5" />
                    <span>Tindakan ini <strong>tidak dapat dibatalkan</strong>. Data yang dipilih akan hilang permanen dari browser ini.</span>
                  </div>
                  <label className="flex items-center gap-2 cursor-pointer select-none">
                    <input
                      type="checkbox"
                      checked={resetConfirmed}
                      onChange={(e) => setResetConfirmed(e.target.checked)}
                      className="h-4 w-4 rounded border-slate-300 dark:border-zinc-600 text-red-600 focus:ring-red-500 cursor-pointer"
                    />
                    <span className="text-xs font-medium text-slate-700 dark:text-zinc-200">Ya, saya yakin dan mengerti risikonya.</span>
                  </label>
                </div>

                {/* Actions */}
                <div className="flex items-center justify-end gap-2 px-5 py-4 border-t border-slate-100 dark:border-zinc-800 bg-slate-50 dark:bg-zinc-800/50">
                  <button
                    onClick={() => setShowResetModal(false)}
                    className="px-4 py-2 rounded-lg border border-slate-200 dark:border-zinc-600 text-xs font-medium text-slate-700 dark:text-zinc-200 bg-white dark:bg-zinc-800 hover:bg-slate-100 dark:hover:bg-zinc-700 transition-colors cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    onClick={handleClearDatabase}
                    disabled={!resetConfirmed || !Object.values(resetScopes).some(Boolean)}
                    className="px-4 py-2 rounded-lg text-xs font-bold text-white bg-red-600 hover:bg-red-700 transition-colors cursor-pointer inline-flex items-center gap-1.5 disabled:opacity-40 disabled:cursor-not-allowed"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Bersihkan Sekarang</span>
                  </button>
                </div>
              </div>
            </div>
          )}

          {currentTab === 'guide' && (
            <div className="max-w-4xl mx-auto space-y-6">
              
              {/* Header Card */}
              <div className="bg-gradient-to-br from-zinc-900 via-zinc-800 to-zinc-950 text-white rounded-2xl p-6 md:p-8 shadow-md relative overflow-hidden">
                <div className="relative z-10 space-y-2.5">
                  <span className="text-[9px] bg-blue-500 text-white font-bold tracking-widest uppercase px-2.5 py-0.5 rounded-full select-none">Pusat Bantuan</span>
                  <h2 className="text-xl md:text-2xl font-bold font-sans tracking-tight">Panduan Pengoperasian & Penjelasan Fitur</h2>
                  <p className="text-xs text-zinc-300 max-w-2xl leading-relaxed font-sans">
                    Pelajari pembagian tugas peran, alur operasional di lapangan, serta cara memanfaatkan fitur-fitur Raudhah Barcode Manager secara optimal.
                  </p>
                </div>
                {/* Decorative glow */}
                <div className="absolute right-0 bottom-0 w-64 h-64 bg-blue-500/10 rounded-full blur-3xl -mr-16 -mb-16"></div>
              </div>

              {/* Grid: Admin vs Operator Per-Section Guide */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                
                {/* PANDUAN UNTUK ADMIN */}
                <div className="bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-700 rounded-xl p-6 shadow-3xs space-y-4">
                  <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-sm flex items-center gap-2 border-b border-slate-100 pb-3">
                    <Building2 className="w-4 h-4 text-red-600" />
                    <span>Panduan Akses Peran: Kantor Pusat (Admin)</span>
                  </h3>
                  <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                    Admin memegang otoritas penuh manajemen data manifes jemaah, operator, dan konfigurasi umum sistem.
                  </p>

                  <div className="space-y-3 pt-1">
                    <div className="text-[11px] space-y-1">
                      <strong className="text-slate-700 dark:text-zinc-200 block">1. Menu Dashboard (Memantau Rombongan)</strong>
                      <p className="text-slate-500 dark:text-zinc-400">
                        *Fungsi:* Melihat jemaah terdekat yang butuh barcode, dibagi per-travel.  
                        *Cara Pakai:* Klik dropdown **Tugas Op** di dalam rombongan untuk membagikan tugas ke operator lapangan secara cepat.
                      </p>
                    </div>

                    <div className="text-[11px] space-y-1">
                      <strong className="text-slate-700 dark:text-zinc-200 block">2. Menu Data Jamaah (Manifes Utama)</strong>
                      <p className="text-slate-500 dark:text-zinc-400">
                        *Fungsi:* Mengelola manifes lengkap (edit, hapus, tambah manual).  
                        *Cara Pakai:* Gunakan tombol tambah manual di pojok kanan atas, atau centang kotak jemaah untuk menghapus massal.
                      </p>
                    </div>

                    <div className="text-[11px] space-y-1">
                      <strong className="text-slate-700 dark:text-zinc-200 block">3. Menu Import Excel (Input Massal)</strong>
                      <p className="text-slate-500 dark:text-zinc-400">
                        *Fungsi:* Mengimpor ribuan jemaah sekaligus secara asinkron ke database lokal.  
                        *Cara Pakai:* Seret file Excel, pratinjau status (Valid/Duplikat), tentukan aksi duplikat (Skip/Overwrite), lalu klik simpan.
                      </p>
                    </div>

                    <div className="text-[11px] space-y-1">
                      <strong className="text-slate-700 dark:text-zinc-200 block">4. Menu Operator (Manajemen Staf)</strong>
                      <p className="text-slate-500 dark:text-zinc-400">
                        *Fungsi:* Mengatur password dan status keaktifan operator lapangan.  
                        *Cara Pakai:* Buat akun dengan nama, telepon, dan password. Status keaktifan bisa dinonaktifkan sementara jika operator tidak bertugas.
                      </p>
                    </div>

                    <div className="text-[11px] space-y-1">
                      <strong className="text-slate-700 dark:text-zinc-200 block">5. Menu Settings (Konfigurasi)</strong>
                      <p className="text-slate-500 dark:text-zinc-400">
                        *Fungsi:* Penyesuaian global (batas harian, tanggal acuan prioritas, reset database).  
                        *Cara Pakai:* Masukkan tanggal hari ini/demo di tanggal acuan untuk menyortir urutan prioritas kedatangan jemaah di dashboard.
                      </p>
                    </div>
                  </div>
                </div>

                {/* PANDUAN UNTUK OPERATOR */}
                <div className="bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-700 rounded-xl p-6 shadow-3xs space-y-4">
                  <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-sm flex items-center gap-2 border-b border-slate-100 pb-3">
                    <Laptop className="w-4 h-4 text-red-600" />
                    <span>Panduan Akses Peran: Operator Lapangan</span>
                  </h3>
                  <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                    Operator berfokus di lapangan untuk memantau antrean jemaah tugasnya, booking Nusuk, dan mengunggah barcode permit.
                  </p>

                  <div className="space-y-3 pt-1">
                    <div className="text-[11px] space-y-1">
                      <strong className="text-slate-700 dark:text-zinc-200 block">1. Menu Dashboard (Prioritas Harian)</strong>
                      <p className="text-slate-500 dark:text-zinc-400">
                        *Fungsi:* Menyaring dan menampilkan jemaah prioritas yang ditugaskan kepada diri sendiri saja.  
                        *Cara Pakai:* Pantau rombongan dengan waktu kedatangan terdekat. Ketika slot booking Nusuk rilis, ganti status ke **Sedang War** untuk memberi sinyal ke tim.
                      </p>
                    </div>

                    <div className="text-[11px] space-y-1">
                      <strong className="text-slate-700 dark:text-zinc-200 block">2. Menu Data Jamaah (Daftar Tugas)</strong>
                      <p className="text-slate-500 dark:text-zinc-400">
                        *Fungsi:* Mengakses jemaah tugas aktif secara rinci.  
                        *Cara Pakai:* Klik tombol **Detail** di kanan jemaah untuk membuka drawer detail. Di drawer tersebut, Anda bisa membaca password Nusuk jemaah dan mengunggah file screenshot permit.
                      </p>
                    </div>

                    <div className="text-[11px] space-y-1">
                      <strong className="text-slate-700 dark:text-zinc-200 block">3. Menu Panduan (Bantuan Operasional)</strong>
                      <p className="text-slate-500 dark:text-zinc-400">
                        *Fungsi:* Pusat instruksi sinkronisasi data lapangan.  
                        *Cara Pakai:* Dibuka kapan saja ketika operator mengalami kesulitan teknis atau lupa aturan pengisian kolom Excel.
                      </p>
                    </div>

                    <div className="p-3 bg-blue-50 text-blue-800 border border-blue-100 rounded text-[10px] leading-relaxed">
                      <strong>💡 Info Penting:</strong> Operator tidak dapat melihat menu **Import Excel**, **Staf Operator**, atau **Settings** guna menjamin keamanan kredensial password admin dan integritas database pusat.
                    </div>
                  </div>
                </div>

              </div>

              {/* LIST FITUR UNGGULAN & CARA MENGGUNAKAN */}
              <div className="bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-700 rounded-xl p-6 shadow-3xs space-y-5">
                <div>
                  <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-base flex items-center gap-2">
                    <span className="w-1.5 h-4 bg-red-600 rounded-xs"></span>
                    <span>Daftar Fitur Utama, Fungsi & Cara Penggunaan</span>
                  </h3>
                  <p className="text-xs text-slate-500 dark:text-zinc-400 mt-0.5">Penjelasan ringkas mengenai teknologi dan fitur canggih yang terintegrasi di dalam aplikasi.</p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  
                  {/* Fitur 1 */}
                  <div className="p-4 bg-slate-50 dark:bg-zinc-800/50 rounded-lg border border-slate-100 space-y-1.5">
                    <div className="flex items-center gap-1.5 text-slate-800 dark:text-zinc-100 font-bold text-xs">
                      <span className="w-1.5 h-1.5 bg-blue-500 rounded-full"></span>
                      <span>Penyimpanan Lokal IndexedDB</span>
                    </div>
                    <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                      *Fungsi:* Menyimpan data manifest jemaah dan foto screenshot QR Code dalam jumlah besar (GigaBytes) di browser tanpa batas memori.  
                      *Cara Pakai:* Berjalan otomatis di latar belakang. Setiap kali Anda menambah jemaah atau mengunggah bukti QR, data tersimpan permanen di memori lokal komputer Anda secara offline.
                    </p>
                  </div>

                  {/* Fitur 2 */}
                  <div className="p-4 bg-slate-50 dark:bg-zinc-800/50 rounded-lg border border-slate-100 space-y-1.5">
                    <div className="flex items-center gap-1.5 text-slate-800 dark:text-zinc-100 font-bold text-xs">
                      <span className="w-1.5 h-1.5 bg-blue-500 rounded-full"></span>
                      <span>Kompresi Foto Bukti Otomatis</span>
                    </div>
                    <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                      *Fungsi:* Mengompresi resolusi dan kualitas foto screenshot yang diunggah hingga berukuran di bawah 100KB (Format JPEG) demi performa browser yang tetap ringan.  
                      *Cara Pakai:* Cukup unggah gambar screenshot melalui drawer detail jemaah. Sistem secara otomatis memperkecil ukuran file sebelum menyimpannya ke database browser.
                    </p>
                  </div>

                  {/* Fitur 3 */}
                  <div className="p-4 bg-slate-50 dark:bg-zinc-800/50 rounded-lg border border-slate-100 space-y-1.5">
                    <div className="flex items-center gap-1.5 text-slate-800 dark:text-zinc-100 font-bold text-xs">
                      <span className="w-1.5 h-1.5 bg-blue-500 rounded-full"></span>
                      <span>Filter Duplikat Fleksibel</span>
                    </div>
                    <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                      *Fungsi:* Mencegah data paspor atau visa ganda terdaftar di database manifes saat impor Excel.  
                      *Cara Pakai:* Sebelum mengklik tombol impor travel di halaman Import Excel, pilih tindakan duplikat pada dropdown menu: **Abaikan** (Skip) data baru atau **Perbarui** (Overwrite) data lama.
                    </p>
                  </div>

                  {/* Fitur 4 */}
                  <div className="p-4 bg-slate-50 dark:bg-zinc-800/50 rounded-lg border border-slate-100 space-y-1.5">
                    <div className="flex items-center gap-1.5 text-slate-800 dark:text-zinc-100 font-bold text-xs">
                      <span className="w-1.5 h-1.5 bg-blue-500 rounded-full"></span>
                      <span>Indikator Warning Berkedip</span>
                    </div>
                    <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                      *Fungsi:* Menarik perhatian admin/operator secara visual agar tidak ada berkas jemaah yang terlupa atau tidak ditugaskan.  
                      *Cara Pakai:* Sistem mendeteksi otomatis jemaah yang memiliki status **Belum QR** (badge merah berkedip) atau **Belum Ditugaskan** (pilihan operator berwarna merah lembut).
                    </p>
                  </div>

                </div>
              </div>

              {/* TIMELINE ALUR KERJA UTAMA */}
              <div className="bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-700 rounded-xl p-6 shadow-3xs space-y-6">
                <div>
                  <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-base flex items-center gap-2">
                    <span className="w-1.5 h-4 bg-red-600 rounded-xs"></span>
                    <span>Alur Kerja Operasional Lapangan (Workflow)</span>
                  </h3>
                  <p className="text-xs text-slate-500 dark:text-zinc-400 mt-0.5">Urutan 5 langkah koordinasi kerja antara admin pusat dan operator di hotel.</p>
                </div>

                <div className="relative border-l-2 border-blue-100 pl-6 ml-3 space-y-6 py-2">
                  
                  {/* Step 1 */}
                  <div className="relative">
                    <div className="absolute -left-[31px] top-0.5 w-4 h-4 rounded-full bg-red-600 border-2 border-white ring-4 ring-blue-50 flex items-center justify-center">
                      <span className="text-[8px] text-white font-bold">1</span>
                    </div>
                    <div className="space-y-1">
                      <span className="text-xs font-bold text-slate-800 dark:text-zinc-100">Akses Peran & Login Akun</span>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                        Pilih peran Anda di halaman masuk utama. Masukkan password admin (`admin123`) atau password operator yang telah dibuat oleh admin.
                      </p>
                    </div>
                  </div>

                  {/* Step 2 */}
                  <div className="relative">
                    <div className="absolute -left-[31px] top-0.5 w-4 h-4 rounded-full bg-red-600 border-2 border-white ring-4 ring-blue-50 flex items-center justify-center">
                      <span className="text-[8px] text-white font-bold">2</span>
                    </div>
                    <div className="space-y-1">
                      <span className="text-xs font-bold text-slate-800 dark:text-zinc-100">Impor Manifest Jemaah (Excel)</span>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                        Admin mengimpor data Excel rombongan travel di tab **Import Excel**. Sistem secara pintar membagi jemaah per-tabel berdasarkan grup travel-nya.
                      </p>
                    </div>
                  </div>

                  {/* Step 3 */}
                  <div className="relative">
                    <div className="absolute -left-[31px] top-0.5 w-4 h-4 rounded-full bg-red-600 border-2 border-white ring-4 ring-blue-50 flex items-center justify-center">
                      <span className="text-[8px] text-white font-bold">3</span>
                    </div>
                    <div className="space-y-1">
                      <span className="text-xs font-bold text-slate-800 dark:text-zinc-100">Penugasan Operator Lapangan</span>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                        Admin menugaskan operator yang bertanggung jawab mengurus booking slot Raudhah jemaah bersangkutan melalui dropdown **Tugas Op**.
                      </p>
                    </div>
                  </div>

                  {/* Step 4 */}
                  <div className="relative">
                    <div className="absolute -left-[31px] top-0.5 w-4 h-4 rounded-full bg-red-600 border-2 border-white ring-4 ring-blue-50 flex items-center justify-center">
                      <span className="text-[8px] text-white font-bold">4</span>
                    </div>
                    <div className="space-y-1">
                      <span className="text-xs font-bold text-slate-800 dark:text-zinc-100">Ubah Status ke "Sedang War"</span>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                        Saat operator mulai memproses booking slot Nusuk jemaah di handphone, status diubah menjadi **Sedang War** agar operator lain mengetahui proses sedang berjalan.
                      </p>
                    </div>
                  </div>

                  {/* Step 5 */}
                  <div className="relative">
                    <div className="absolute -left-[31px] top-0.5 w-4 h-4 rounded-full bg-red-600 border-2 border-white ring-4 ring-blue-50 flex items-center justify-center">
                      <span className="text-[8px] text-white font-bold">5</span>
                    </div>
                    <div className="space-y-1">
                      <span className="text-xs font-bold text-slate-800 dark:text-zinc-100">Unggah Bukti Screenshot Barcode Nusuk</span>
                      <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                        Operator mengambil screenshot permit Raudhah dari Nusuk, lalu mengunggahnya ke detail jemaah. Status otomatis berganti menjadi **QR Berhasil** (Sukses).
                      </p>
                    </div>
                  </div>

                </div>
              </div>

              {/* PEMETAAN KOLOM EXCEL */}
              <div className="bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-700 rounded-xl p-6 shadow-3xs space-y-4">
                <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-sm flex items-center gap-2">
                  <span className="w-1.5 h-4 bg-red-600 rounded-xs"></span>
                  <span>Variasi Kolom Impor Excel yang Didukung</span>
                </h3>
                <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                  Sistem parser secara otomatis mencocokkan kata kunci header file Excel Anda dengan kolom database berikut secara cerdas:
                </p>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-[11px]">
                  <div className="p-3 bg-slate-50 dark:bg-zinc-800/50 rounded-lg border border-slate-100 space-y-1">
                    <span className="font-bold text-slate-800">Identitas Jemaah</span>
                    <div className="text-slate-500 dark:text-zinc-400 space-y-0.5">
                      • **Nama Jamaah:** `nama, name, nama_jamaah, jamaah`<br />
                      • **Email:** `email, surel, mail, e_mail`<br />
                      • **Gender:** `gender, jenis_kelamin, sex, kelamin`
                    </div>
                  </div>

                  <div className="p-3 bg-slate-50 dark:bg-zinc-800/50 rounded-lg border border-slate-100 space-y-1">
                    <span className="font-bold text-slate-800">Dokumen & Rombongan</span>
                    <div className="text-slate-500 dark:text-zinc-400 space-y-0.5">
                      • **No Paspor:** `passport, paspor, passport_no, no_paspor`<br />
                      • **No Visa:** `visa, no_visa, nomor_visa`<br />
                      • **Nama Travel:** `travel, group, travel_name, rombongan`
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

        </div>
      </main>

      {/* DETAILED SIDE-OVER DRAWER (FOR JAMAAH DETAIL) */}
      {selectedJamaah && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-xs flex justify-end z-50">
          
          <div className="w-full max-w-lg bg-white dark:bg-zinc-800 h-full shadow-2xl flex flex-col p-6 animate-in slide-in-from-right duration-200">
            
            {/* Drawer Header */}
            <div className="flex items-center justify-between pb-4 border-b border-slate-100">
              <div>
                <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-base">Detail Informasi Jamaah</h3>
                <p className="text-xs text-slate-400 font-mono mt-0.5">ID: {selectedJamaah.id}</p>
              </div>
              <button
                onClick={() => setSelectedJamaah(null)}
                className="p-1 rounded-md hover:bg-slate-100 dark:bg-zinc-700 text-slate-400 hover:text-slate-700 dark:text-zinc-200 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Drawer Body (Scrollable) */}
            <div className="flex-1 overflow-y-auto py-5 space-y-6">
              
              {/* Primary Identity Section */}
              <div className="bg-slate-50 dark:bg-zinc-800/50 rounded-xl p-4 space-y-3 border border-slate-100">
                <div className="flex items-center justify-between">
                  <span className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">Identitas Utama</span>
                  <span className={`inline-flex items-center gap-1 border px-2 py-0.5 rounded-full text-[10px] font-medium ${getPriorityInfo(selectedJamaah.entryMadinah, settingsReferenceDate).badgeColor}`}>
                    Prioritas: {getPriorityInfo(selectedJamaah.entryMadinah, settingsReferenceDate).level}
                  </span>
                </div>
                <div className="space-y-1.5">
                  <div className="font-bold text-slate-800 dark:text-zinc-100 text-lg leading-tight">{selectedJamaah.name}</div>
                  <div className="text-xs text-slate-500 dark:text-zinc-400 flex flex-wrap items-center gap-x-2 gap-y-1">
                    <span>{selectedJamaah.gender}</span>
                    <span>•</span>
                    <span>HP: {selectedJamaah.phone}</span>
                    {selectedJamaah.email && (
                      <>
                        <span>•</span>
                        <span className="text-slate-600 dark:text-zinc-300 font-mono">{selectedJamaah.email}</span>
                      </>
                    )}
                    {selectedJamaah.travel && (
                      <>
                        <span>•</span>
                        <span className="bg-zinc-100 dark:bg-zinc-700 text-zinc-800 dark:text-zinc-200 px-2 py-0.5 rounded text-[10px] font-bold uppercase">{selectedJamaah.travel}</span>
                      </>
                    )}
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-4 pt-2 text-xs font-mono">
                  <div className="p-2 bg-white dark:bg-zinc-800 rounded border border-slate-200 dark:border-zinc-600/50">
                    <span className="text-[9px] text-slate-400 uppercase tracking-wider block font-sans">Nomor Paspor</span>
                    <strong className="text-slate-700 dark:text-zinc-200">{selectedJamaah.passport}</strong>
                  </div>
                  <div className="p-2 bg-white dark:bg-zinc-800 rounded border border-slate-200 dark:border-zinc-600/50">
                    <span className="text-[9px] text-slate-400 uppercase tracking-wider block font-sans">Nomor Visa</span>
                    <strong className="text-slate-700 dark:text-zinc-200">{selectedJamaah.visa}</strong>
                  </div>
                </div>
              </div>

              {/* Kredensial untuk Nusuk — tombol copy cepat */}
              <div className="rounded-xl border border-slate-100 dark:border-zinc-700 overflow-hidden">
                <div className="px-3 py-2 bg-slate-50 dark:bg-zinc-800/50 text-[10px] uppercase font-bold text-slate-400 tracking-wider border-b border-slate-100 dark:border-zinc-700">
                  Kredensial untuk Nusuk
                </div>
                <div className="divide-y divide-slate-100 dark:divide-zinc-700">
                  {([
                    { key: 'email', label: 'Email', value: selectedJamaah.email || '' },
                    { key: 'passport', label: 'No. Paspor', value: selectedJamaah.passport || '' },
                    { key: 'visa', label: 'No. Visa', value: selectedJamaah.visa || '' },
                    { key: 'password', label: 'Password', value: selectedJamaah.password || '' },
                  ] as { key: string; label: string; value: string }[]).map(item => (
                    <div key={item.key} className="flex items-center justify-between gap-2 px-3 py-2">
                      <div className="min-w-0">
                        <span className="block text-[9px] uppercase tracking-wider text-slate-400 font-sans">{item.label}</span>
                        <span className="block font-mono text-xs text-slate-700 dark:text-zinc-200 truncate">{item.value || '-'}</span>
                      </div>
                      <button
                        type="button"
                        onClick={() => copyToClipboard(item.value, item.key)}
                        disabled={!item.value}
                        title={`Salin ${item.label}`}
                        className={`shrink-0 flex items-center gap-1 px-2 py-1.5 rounded-lg text-[10px] font-semibold transition-all ${
                          copiedField === item.key
                            ? 'bg-emerald-100 text-emerald-700 dark:bg-emerald-500/15 dark:text-emerald-300'
                            : 'bg-slate-100 text-slate-600 hover:bg-blue-50 hover:text-blue-600 dark:bg-zinc-700 dark:text-zinc-300 dark:hover:bg-zinc-600 disabled:opacity-40 disabled:cursor-not-allowed'
                        }`}
                      >
                        {copiedField === item.key ? <><Check className="w-3 h-3" /> Tersalin</> : <><Copy className="w-3 h-3" /> Salin</>}
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              {/* Booking Dates */}
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <span className="text-xs font-semibold text-slate-600 dark:text-zinc-300 flex items-center gap-1">
                    <Calendar className="w-3.5 h-3.5 text-slate-400" />
                    Masuk Madinah
                  </span>
                  <div className="p-2.5 bg-slate-50 dark:bg-zinc-800/50 border border-slate-100 rounded-lg text-xs font-medium text-slate-800 dark:text-zinc-100 font-mono">
                    {selectedJamaah.entryMadinah}
                  </div>
                </div>
                <div className="space-y-1">
                  <span className="text-xs font-semibold text-slate-600 dark:text-zinc-300 flex items-center gap-1">
                    <Calendar className="w-3.5 h-3.5 text-slate-400" />
                    Keluar Madinah
                  </span>
                  <div className="p-2.5 bg-slate-50 dark:bg-zinc-800/50 border border-slate-100 rounded-lg text-xs font-medium text-slate-800 dark:text-zinc-100 font-mono">
                    {selectedJamaah.exitMadinah}
                  </div>
                </div>
              </div>

              {/* Raudhah Slot & Countdown Reminder */}
              {selectedJamaah.raudhahSlot && (
                <div className="bg-slate-50 dark:bg-zinc-800/50 dark:bg-zinc-800/50 border border-slate-100 dark:border-zinc-700 rounded-xl p-4 space-y-2">
                  {/* Slot Raudhah & waktu QR tersedia, dalam Waktu Madinah + WITA. */}
                  {(() => {
                    const slotInstant = getDistributionInstant(selectedJamaah.raudhahSlot, 0);
                    const distInstant = getDistributionInstant(selectedJamaah.raudhahSlot, settingsQrLeadHours);
                    if (!slotInstant || !distInstant) return null;
                    return (
                      <div className="space-y-2">
                        <div className="flex items-start gap-2 bg-emerald-50/70 dark:bg-emerald-900/15 border border-emerald-100 dark:border-emerald-900/40 rounded-lg px-2.5 py-2">
                          <Clock className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                          <div className="text-[11px] text-slate-700 dark:text-zinc-200 leading-snug">
                            <span className="font-bold text-emerald-800 dark:text-emerald-400">Slot Raudhah:</span> {formatFullInZone(slotInstant, TZ_MADINAH)} <span className="text-slate-400 dark:text-zinc-500">(Waktu Madinah)</span> / <span className="font-semibold">{formatTimeColon(slotInstant, TZ_WITA)} WITA</span>
                          </div>
                        </div>
                        <div className="flex items-start gap-2 bg-sky-50/70 dark:bg-sky-900/15 border border-sky-100 dark:border-sky-900/40 rounded-lg px-2.5 py-2">
                          <Download className="w-4 h-4 text-sky-600 shrink-0 mt-0.5" />
                          <div className="text-[11px] text-slate-700 dark:text-zinc-200 leading-snug">
                            <span className="font-bold text-sky-800 dark:text-sky-300">QR tersedia:</span> {formatFullInZone(distInstant, TZ_MADINAH)} <span className="text-slate-400 dark:text-zinc-500">(Waktu Madinah)</span> / <span className="font-semibold">{formatTimeColon(distInstant, TZ_WITA)} WITA</span>
                            <span className="text-slate-400 dark:text-zinc-500"> · −{settingsQrLeadHours} jam</span>
                          </div>
                        </div>
                      </div>
                    );
                  })()}

                  {(() => {
                    if (selectedJamaah.status === 'QR Berhasil') {
                      return (
                        <div className="text-[11px] text-emerald-700 bg-emerald-50/50 border border-emerald-100 p-2.5 rounded-lg font-semibold flex items-center gap-1.5">
                          <Check className="w-4 h-4 text-emerald-600" />
                          <span>Barcode Raudhah Sudah Siap & Terunggah</span>
                        </div>
                      );
                    }

                    // Countdown targets the QR DISTRIBUTION time (slot − lead), computed from the
                    // Madinah-pinned instant so it stays correct on a WITA (or any) machine.
                    const distInstant = getDistributionInstant(selectedJamaah.raudhahSlot, settingsQrLeadHours);
                    const slotInstant = getDistributionInstant(selectedJamaah.raudhahSlot, 0);
                    if (!distInstant || !slotInstant) return null;
                    const nowTime = now.getTime();
                    const minsToDist = (distInstant.getTime() - nowTime) / 60000;
                    const minsToSlot = (slotInstant.getTime() - nowTime) / 60000;

                    if (minsToSlot <= 0) {
                      return (
                        <div className="text-[11px] text-slate-500 dark:text-zinc-400 bg-slate-50 dark:bg-zinc-800/50 border border-slate-200 dark:border-zinc-600/50 p-2.5 rounded-lg italic">
                          Melewati Jadwal Raudhah ({selectedJamaah.raudhahSlot.replace('T', ' ')} Madinah)
                        </div>
                      );
                    }

                    // Distribution window already open (past distribution time, before slot) — act NOW.
                    if (minsToDist <= 0) {
                      return (
                        <div className="text-[11px] text-rose-700 bg-rose-50 border border-rose-150 p-2.5 rounded-lg font-bold animate-pulse flex items-center gap-1.5">
                          <AlertTriangle className="w-4 h-4 text-rose-600" />
                          <span>SEKARANG: Masuk Nusuk, download & distribusikan QR! Slot Raudhah {Math.ceil(minsToSlot)} menit lagi.</span>
                        </div>
                      );
                    }

                    if (minsToDist <= 60) {
                      return (
                        <div className="text-[11px] text-rose-700 bg-rose-50 border border-rose-150 p-2.5 rounded-lg font-bold animate-pulse flex items-center gap-1.5">
                          <AlertTriangle className="w-4 h-4 text-rose-600" />
                          <span>Penting: Distribusi QR dalam {Math.ceil(minsToDist)} menit!</span>
                        </div>
                      );
                    }

                    if (minsToDist <= 180) {
                      return (
                        <div className="text-[11px] text-amber-700 bg-amber-50 border border-amber-100 p-2.5 rounded-lg font-semibold flex items-center gap-1.5">
                          <Clock className="w-4 h-4 text-amber-600" />
                          <span>Perhatian: Distribusi QR dalam {Math.floor(minsToDist / 60)} jam {Math.round(minsToDist % 60)} menit</span>
                        </div>
                      );
                    }

                    // Normal countdown (> 3 hours remaining until distribution)
                    const hours = Math.floor(minsToDist / 60);
                    const days = Math.floor(hours / 24);
                    const remainingHours = hours % 24;

                    return (
                      <div className="text-[11px] text-slate-600 dark:text-zinc-300 bg-slate-50 dark:bg-zinc-800/50 border border-slate-200 p-2.5 rounded-lg font-medium flex items-center gap-1.5">
                        <Clock className="w-3.5 h-3.5 text-slate-400" />
                        <span>
                          Distribusi QR dalam: {days > 0 ? `${days} hari ` : ''}{remainingHours} jam {Math.round(minsToDist % 60)} menit
                        </span>
                      </div>
                    );
                  })()}
                </div>
              )}

              {/* Dynamic Custom Fields Section */}
              {customFields.length > 0 && (
                <div className="bg-slate-50 dark:bg-zinc-800/50 rounded-xl p-4 border border-slate-100 dark:border-zinc-700/60 space-y-2.5">
                  <span className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">Informasi Tambahan (Kustom)</span>
                  <div className="grid grid-cols-2 gap-4 text-xs">
                    {customFields.map(cf => {
                      const val = selectedJamaah.customValues?.[cf.id] || '-';
                      return (
                        <div key={cf.id} className="p-2.5 bg-white dark:bg-zinc-800 rounded border border-slate-200 dark:border-zinc-600">
                          <span className="text-[9px] text-slate-400 uppercase tracking-wider block font-sans">{cf.label}</span>
                          <strong className="text-slate-700 dark:text-zinc-200 block truncate" title={val}>{val}</strong>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* Status Update & Assign Operator */}
              <div className="space-y-4">
                
                {/* Drodown Status */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-700 dark:text-zinc-200">Status Proses Booking</label>
                  <select
                    value={selectedJamaah.status}
                    onChange={(e) => handleQuickStatusChange(selectedJamaah.id, e.target.value as JamaahStatus)}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-white dark:bg-zinc-700 text-slate-800 dark:text-zinc-100 shadow-3xs outline-hidden focus:border-red-500"
                  >
                    <option value="Ready">Ready (Menunggu antrean)</option>
                    <option value="Sedang War">Sedang War (Nusuk sedang diproses)</option>
                    <option value="QR Berhasil">QR Berhasil (Barcode sukses diterbitkan)</option>
                    <option value="Belum Berhasil">Belum Berhasil (Quota limit/eror Nusuk)</option>
                  </select>
                </div>

                {/* Operator Assignment */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-700 dark:text-zinc-200">Tugaskan ke Operator</label>
                  {activeOperatorId === null ? (
                    <select
                      value={selectedJamaah.operatorId || ''}
                      onChange={(e) => handleAssignOperatorInDetail(e.target.value || null)}
                      className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-white dark:bg-zinc-700 text-slate-800 dark:text-zinc-100 shadow-3xs outline-hidden focus:border-red-500"
                    >
                      <option value="">Belum Ditugaskan (Unassigned)</option>
                      {operators.filter(o => o.isActive).map(op => (
                        <option key={op.id} value={op.id}>{op.name}</option>
                      ))}
                    </select>
                  ) : (
                    <div className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700 text-slate-800 dark:text-zinc-100">
                      {selectedJamaah.operatorId ? operators.find(o => o.id === selectedJamaah.operatorId)?.name || '-' : 'Belum Ditugaskan'}
                    </div>
                  )}
                </div>

                {/* Notes Textarea */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-700 dark:text-zinc-200">Catatan Operator / Kronologi</label>
                  <textarea
                    rows={3}
                    value={selectedJamaah.notes}
                    onChange={(e) => handleUpdateNotesInDetail(e.target.value)}
                    placeholder="Masukkan detail percobaan booking, kendala teknis, jam slot Nusuk, dll..."
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-white dark:bg-zinc-700 text-slate-800 dark:text-zinc-100 shadow-3xs outline-hidden focus:border-red-500 resize-none"
                  />
                </div>

              </div>

              {/* QR Code Screenshot Attachment */}
              <div className="space-y-2">
                <label className="text-xs font-semibold text-slate-700 dark:text-zinc-200 flex items-center justify-between">
                  <span>Dokumen Bukti QR Code (Screenshot Nusuk)</span>
                  {selectedJamaah.qrUploadedAt && (
                    <span className="text-[10px] text-red-600 font-mono">
                      Diunggah: {new Date(selectedJamaah.qrUploadedAt).toLocaleTimeString()}
                    </span>
                  )}
                </label>

                {selectedJamaah.qrCodeUrl ? (
                  <div className="border border-slate-200 dark:border-zinc-600 rounded-xl p-4 flex flex-col items-center justify-center gap-3 bg-slate-50 dark:bg-zinc-800/50 relative group">
                    <img
                      src={selectedJamaah.qrCodeUrl}
                      alt="Raudhah QR Barcode Screenshot"
                      className="max-h-48 rounded shadow-sm border border-slate-200 dark:border-zinc-600 bg-white dark:bg-zinc-800"
                      referrerPolicy="no-referrer"
                    />
                    
                    {activeOperatorId === null && (
                      <button
                        onClick={() => {
                          setJamaahs(prev => prev.map(j => j.id === selectedJamaah.id ? { ...j, qrCodeUrl: null, qrUploadedAt: null } : j));
                          setSelectedJamaah(prev => prev ? { ...prev, qrCodeUrl: null, qrUploadedAt: null } : null);
                        }}
                        className="absolute top-2 right-2 p-1.5 rounded-full bg-red-100 hover:bg-red-200 text-red-600 transition-colors shadow-sm"
                        title="Hapus screenshot"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    )}

                    <p className="text-[10px] text-slate-500 dark:text-zinc-400">Bukti Screenshot Raudhah Tersimpan Aman</p>
                  </div>
                ) : (
                  <div className="border border-dashed border-slate-200 dark:border-zinc-600 rounded-xl p-6 flex flex-col items-center justify-center gap-3 bg-slate-50 dark:bg-zinc-800/50">
                    <Upload className="w-6 h-6 text-slate-300" />
                    <div className="text-center space-y-1">
                      <span className="text-xs font-medium text-slate-700 dark:text-zinc-200 block">Belum ada screenshot QR Code</span>
                      <p className="text-[10px] text-slate-400">Silakan unggah tangkapan layar dari aplikasi Nusuk</p>
                    </div>

                    <div className="flex flex-col sm:flex-row gap-2 w-full max-w-xs pt-1">
                      <label className="flex-1 py-1.5 rounded bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-600 hover:bg-slate-50 dark:bg-zinc-800/50 dark:hover:bg-zinc-700 text-[10px] font-semibold text-slate-700 dark:text-zinc-200 transition-colors cursor-pointer text-center">
                        Pilih File Gambar
                        <input
                          type="file"
                          accept="image/*"
                          onChange={handleUploadScreenshotInDetail}
                          className="hidden"
                        />
                      </label>
                    </div>
                  </div>
                )}
              </div>

            </div>

            {/* Drawer Footer */}
            <div className="pt-4 border-t border-slate-100 flex items-center justify-end gap-2 shrink-0">
              <button
                onClick={() => { const j = selectedJamaah; setSelectedJamaah(null); handleOpenEditModal(j); }}
                className="px-4 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 transition-colors flex items-center gap-1.5"
              >
                <Edit className="w-3.5 h-3.5" /> Edit Data
              </button>
              <button
                onClick={() => setSelectedJamaah(null)}
                className="px-4 py-2 rounded-lg bg-zinc-900 text-white text-xs font-semibold hover:bg-zinc-800 transition-colors"
              >
                Selesai & Simpan
              </button>
            </div>

          </div>

        </div>
      )}

      {/* MODAL: TAMBAH JAMAAH MANUAL */}
      {showAddModal && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-xs flex items-center justify-center z-50 p-4">
          <div className="bg-white dark:bg-zinc-800 rounded-xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-hidden flex flex-col animate-in zoom-in-95 duration-150">

            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50 dark:bg-zinc-800/50 shrink-0">
              <div>
                <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">Tambah Jamaah Baru (Manual)</h3>
                <p className="text-xs text-slate-500 dark:text-zinc-400 mt-0.5">Input manifest jamaah visa terbit secara individual.</p>
              </div>
              <button
                onClick={() => { setShowAddModal(false); setOcrError(null); }}
                className="p-1 rounded-md hover:bg-slate-200 text-slate-400 hover:text-slate-700 dark:text-zinc-200 transition-all"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleAddJamaah} className="flex-1 overflow-y-auto p-6 space-y-4">

              {/* SCAN DOKUMEN VISA VIA GEMINI AI */}
              <div className="bg-slate-50 dark:bg-zinc-800/50 dark:bg-zinc-800/50 border border-slate-200 dark:border-zinc-600/60 rounded-xl p-4 space-y-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-1.5 text-xs font-bold text-slate-800 dark:text-zinc-100">
                    <Sparkles className="w-4 h-4 text-red-600 animate-pulse" />
                    <span>Scan & Deteksi Visa Otomatis (Gemini AI)</span>
                  </div>
                  {!settingsGeminiApiKey && (
                    <span className="text-[10px] bg-amber-50 text-amber-800 border border-amber-200 px-2 py-0.5 rounded font-semibold">Gemini API Key Kosong</span>
                  )}
                </div>
                
                <p className="text-[11px] text-slate-500 dark:text-zinc-400 leading-relaxed">
                  Unggah foto/file dokumen visa umrah/haji jemaah. Sistem akan membaca nomor paspor, nomor visa, nama lengkap, gender, dan travel secara otomatis menggunakan kecerdasan buatan Google Gemini.
                </p>

                {ocrError && (
                  <div className="p-3 bg-red-50 dark:bg-red-950/20 border border-red-200 dark:border-red-900/50 rounded-lg space-y-2">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-1.5 text-[11px] font-bold text-red-800 dark:text-red-300">
                        <AlertTriangle className="w-3.5 h-3.5 text-red-600 dark:text-red-400" />
                        <span>Gagal Scan Otomatis</span>
                      </div>
                      <button
                        type="button"
                        onClick={() => {
                          navigator.clipboard.writeText(ocrError);
                          alert('Detail error berhasil disalin ke clipboard!');
                        }}
                        className="text-[10px] text-red-600 dark:text-red-400 hover:underline font-semibold cursor-pointer"
                      >
                        Salin Detail Error
                      </button>
                    </div>
                    <p className="text-[10px] font-mono text-red-700 dark:text-red-400/90 break-all leading-relaxed select-all">
                      {ocrError}
                    </p>
                  </div>
                )}

                {isScanningVisa ? (
                  <div className="flex items-center justify-center gap-2.5 py-4 border border-dashed border-emerald-300 dark:border-emerald-700/50 rounded-lg bg-emerald-50/20 dark:bg-emerald-900/20 text-xs font-semibold text-emerald-800 dark:text-emerald-300">
                    <Loader2 className="w-4 h-4 text-emerald-600 animate-spin" />
                    <span>{scanVisaStatus}</span>
                  </div>
                ) : (
                  <div className="relative">
                    <input
                      type="file"
                      accept="image/*"
                      disabled={!settingsGeminiApiKey}
                      onChange={handleScanVisaImage}
                      className="absolute inset-0 w-full h-full opacity-0 cursor-pointer disabled:cursor-not-allowed"
                      id="visa-ocr-upload"
                    />
                    <label
                      htmlFor="visa-ocr-upload"
                      className={`flex items-center justify-center gap-1.5 py-2.5 rounded-lg border border-dashed text-xs font-bold transition-all cursor-pointer ${
                        settingsGeminiApiKey 
                          ? 'border-blue-300 bg-blue-50 text-blue-700 hover:bg-blue-100/70' 
                          : 'border-slate-200 dark:border-zinc-600 bg-slate-100 dark:bg-zinc-700 text-slate-400'
                      }`}
                    >
                      <Upload className="w-3.5 h-3.5" />
                      <span>{settingsGeminiApiKey ? 'Pilih Gambar Visa / Scan Sekarang' : 'Gemini API Key Belum Diatur (Fitur Dinonaktifkan)'}</span>
                    </label>
                  </div>
                )}
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                
                {/* Nama */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Nama Jamaah (Sesuai Paspor) *</label>
                  <input
                    type="text"
                    required
                    placeholder="Contoh: Muhammad Ali"
                    value={newJamaah.name}
                    onChange={(e) => {
                      const name = e.target.value;
                      setNewJamaah(prev => {
                        // Auto-fill email from name, but stop once the user edits email manually.
                        const autoPrev = buildDefaultEmail(prev.name || '');
                        const shouldAuto = !prev.email || prev.email === autoPrev;
                        return { ...prev, name, email: shouldAuto ? buildDefaultEmail(name) : prev.email };
                      });
                    }}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500"
                  />
                </div>

                {/* HP */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Nomor HP / WhatsApp</label>
                  <input
                    type="text"
                    placeholder="Contoh: +62812345678"
                    value={newJamaah.phone}
                    onChange={(e) => setNewJamaah({ ...newJamaah, phone: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500"
                  />
                </div>

                {/* Email */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Alamat Email</label>
                  <input
                    type="email"
                    placeholder="Contoh: jamaah@email.com"
                    value={newJamaah.email}
                    onChange={(e) => setNewJamaah({ ...newJamaah, email: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500"
                  />
                </div>

                {/* Passport */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Nomor Paspor *</label>
                  <input
                    type="text"
                    required
                    placeholder="Contoh: A1234567"
                    value={newJamaah.passport}
                    onChange={(e) => setNewJamaah({ ...newJamaah, passport: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 font-mono uppercase"
                  />
                </div>

                {/* Visa */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Nomor Visa *</label>
                  <input
                    type="text"
                    required
                    placeholder="Contoh: V123456789"
                    value={newJamaah.visa}
                    onChange={(e) => setNewJamaah({ ...newJamaah, visa: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 font-mono"
                  />
                </div>

                {/* Gender */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Gender *</label>
                  <select
                    value={newJamaah.gender}
                    onChange={(e) => setNewJamaah({ ...newJamaah, gender: e.target.value as Gender | '' })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-white dark:bg-zinc-700 text-slate-800 dark:text-zinc-100 outline-hidden focus:border-red-500"
                  >
                    <option value="">— Pilih —</option>
                    <option value="Laki-laki">Laki-laki</option>
                    <option value="Perempuan">Perempuan</option>
                  </select>
                </div>

                {/* Operator Assignment */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Tugaskan Operator</label>
                  <select
                    value={newJamaah.operatorId}
                    onChange={(e) => setNewJamaah({ ...newJamaah, operatorId: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-white dark:bg-zinc-700 text-slate-800 dark:text-zinc-100 outline-hidden focus:border-red-500"
                  >
                    <option value="">Belum Ditugaskan</option>
                    {operators.filter(o => o.isActive).map(op => (
                      <option key={op.id} value={op.id}>{op.name}</option>
                    ))}
                  </select>
                </div>

                {/* Nama Travel */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Nama Travel / Rombongan *</label>
                  <input
                    type="text"
                    required
                    placeholder="Contoh: An-Nahl Umrah & Haji, Al-Fatih Tour, dll."
                    value={newJamaah.travel || ''}
                    onChange={(e) => setNewJamaah({ ...newJamaah, travel: e.target.value.toUpperCase() })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500"
                  />
                </div>

                {/* Password Jemaah */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Password Akses Jemaah</label>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      placeholder="Kosongkan untuk acak otomatis..."
                      value={newJamaah.password || ''}
                      onChange={(e) => setNewJamaah({ ...newJamaah, password: e.target.value })}
                      className="flex-1 text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 font-mono font-semibold"
                    />
                    <button
                      type="button"
                      onClick={() => setNewJamaah({
                        ...newJamaah,
                        password: Math.floor(100000 + Math.random() * 900000).toString()
                      })}
                      className="px-3 bg-slate-100 dark:bg-zinc-700 hover:bg-slate-200 text-slate-700 dark:text-zinc-200 text-xs font-semibold rounded-lg border border-slate-200 dark:border-zinc-600 transition-colors"
                    >
                      Acak
                    </button>
                  </div>
                </div>

                {/* Entry Date */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Tanggal Masuk Madinah *</label>
                  <input
                    type="date"
                    required
                    value={toDateValue(newJamaah.entryMadinah)}
                    onChange={(e) => setNewJamaah({ ...newJamaah, entryMadinah: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500"
                  />
                </div>
                {/* Exit Date */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Tanggal Keluar Madinah *</label>
                  <input
                    type="date"
                    required
                    value={newJamaah.exitMadinah}
                    onChange={(e) => setNewJamaah({ ...newJamaah, exitMadinah: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500"
                  />
                </div>

                {/* Raudhah Entry Slot */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300 flex items-center gap-1">
                    <Clock className="w-3.5 h-3.5 text-emerald-600" />
                    <span>Jadwal Slot Raudhah (Nusuk)</span>
                  </label>
                  <input
                    type="datetime-local"
                    value={newJamaah.raudhahSlot || ''}
                    min={newJamaah.entryMadinah ? `${toDateValue(newJamaah.entryMadinah)}T00:00` : undefined}
                    max={newJamaah.exitMadinah ? `${toDateValue(newJamaah.exitMadinah)}T23:59` : undefined}
                    onChange={(e) => setNewJamaah({ ...newJamaah, raudhahSlot: e.target.value || null })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 font-mono"
                  />
                  {isSlotOutsideStay(newJamaah.raudhahSlot, newJamaah.entryMadinah, newJamaah.exitMadinah) && (
                    <div className="text-[10px] text-amber-700 bg-amber-50 border border-amber-100 p-2 rounded-lg font-semibold flex items-center gap-1.5 mt-1 dark:bg-amber-950/20 dark:border-amber-900/40 dark:text-amber-400">
                      <AlertTriangle className="w-3.5 h-3.5 text-amber-600 shrink-0" />
                      <span>Slot di luar rentang Masuk–Keluar Madinah ({formatDateLabel(newJamaah.entryMadinah || '')} – {formatDateLabel(newJamaah.exitMadinah || '')})!</span>
                    </div>
                  )}
                  {(() => {
                    const dist = getDistributionInstant(newJamaah.raudhahSlot, settingsQrLeadHours);
                    if (!dist) return null;
                    return (
                      <div className="text-[10px] text-sky-800 bg-sky-50/70 border border-sky-100 p-2 rounded-lg flex items-center gap-1.5 mt-1 dark:bg-sky-900/15 dark:border-sky-900/40 dark:text-sky-300">
                        <Download className="w-3.5 h-3.5 text-sky-600 shrink-0" />
                        <span>Distribusi QR (−{settingsQrLeadHours} jam): <strong>{formatInZone(dist, TZ_WITA, true)} WITA</strong> ({formatInZone(dist, TZ_MADINAH, true)} Madinah)</span>
                      </div>
                    );
                  })()}
                </div>

              </div>

              {/* Dynamic Custom Fields */}
              {customFields.length > 0 && (
                <div className="bg-slate-50 dark:bg-zinc-800/40 p-4 border border-slate-200/60 dark:border-zinc-700/60 rounded-xl space-y-3 mb-4">
                  <span className="text-xs font-bold text-slate-800 dark:text-zinc-100 block border-b pb-1">Kolom Tambahan (Kustom)</span>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    {customFields.map((cf) => (
                      <div className="space-y-1" key={cf.id}>
                        <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">{cf.label}</label>
                        <input
                          type="text"
                          placeholder={`Ketik ${cf.label.toLowerCase()}...`}
                          value={newJamaah.customValues?.[cf.id] || ''}
                          onChange={(e) => {
                            const vals = newJamaah.customValues || {};
                            setNewJamaah({
                              ...newJamaah,
                              customValues: {
                                ...vals,
                                [cf.id]: e.target.value
                              }
                            });
                          }}
                          className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-white dark:bg-zinc-700/40 outline-hidden focus:border-red-500 text-slate-800 dark:text-zinc-100"
                        />
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Notes */}
              <div className="space-y-1">
                <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Catatan Khusus (Opsional)</label>
                <textarea
                  rows={2}
                  placeholder="Contoh: Kursi roda, lansia, dll..."
                  value={newJamaah.notes}
                  onChange={(e) => setNewJamaah({ ...newJamaah, notes: e.target.value })}
                  className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 resize-none"
                />
              </div>

              {/* Form Actions */}
              <div className="flex items-center justify-end gap-2 pt-4 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setShowAddModal(false)}
                  className="px-4 py-2 rounded-lg border border-slate-200 dark:border-zinc-600 text-xs font-medium text-slate-700 dark:text-zinc-200 bg-white hover:bg-slate-50 dark:bg-zinc-800/50"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 rounded-lg bg-zinc-900 hover:bg-zinc-800 text-white text-xs font-semibold shadow-xs"
                >
                  Simpan Jamaah
                </button>
              </div>

            </form>

          </div>
        </div>
      )}

      {/* MODAL: EDIT JAMAAH */}
      {showEditModal && editingJamaah && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-xs flex items-center justify-center z-50 p-4 animate-in fade-in duration-200">
          <div className="bg-white dark:bg-zinc-800 rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-hidden flex flex-col animate-in zoom-in-95 duration-200 border border-slate-100 dark:border-zinc-700">
            
            {/* Header */}
            <div className="px-6 py-4 border-b border-slate-100 bg-slate-50 dark:bg-zinc-800/50 flex items-center justify-between">
              <div>
                <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-base">Edit Data Jamaah</h3>
                <p className="text-xs text-slate-400 font-medium">Ubah data jemaah dan kelola informasinya</p>
              </div>
              <button
                type="button"
                onClick={() => {
                  setShowEditModal(false);
                  setEditingJamaah(null);
                }}
                className="p-1.5 rounded-lg hover:bg-slate-200 text-slate-400 hover:text-slate-700 dark:text-zinc-200 transition-all cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Scrollable Form Body */}
            <form onSubmit={handleUpdateJamaah} className="flex-1 overflow-y-auto p-6 space-y-4">
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                
                {/* Full Name */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Nama Lengkap Jamaah *</label>
                  <input
                    type="text"
                    required
                    placeholder="Contoh: Muhammad Ibrahim"
                    value={editingJamaah.name}
                    onChange={(e) => setEditingJamaah({ ...editingJamaah, name: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500"
                  />
                </div>

                {/* Gender */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Jenis Kelamin *</label>
                  <select
                    value={editingJamaah.gender}
                    onChange={(e) => setEditingJamaah({ ...editingJamaah, gender: e.target.value as Gender | '' })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 cursor-pointer"
                  >
                    <option value="">— Pilih —</option>
                    <option value="Laki-laki">♂️ Laki-laki</option>
                    <option value="Perempuan">♀️ Perempuan</option>
                  </select>
                </div>

                {/* Passport */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Nomor Paspor *</label>
                  <input
                    type="text"
                    required
                    placeholder="Contoh: A1234567"
                    value={editingJamaah.passport}
                    onChange={(e) => setEditingJamaah({ ...editingJamaah, passport: e.target.value.toUpperCase() })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 font-mono font-bold"
                  />
                </div>

                {/* Visa */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Nomor Visa *</label>
                  <input
                    type="text"
                    required
                    placeholder="Contoh: 1234567890"
                    value={editingJamaah.visa}
                    onChange={(e) => setEditingJamaah({ ...editingJamaah, visa: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 font-mono font-bold"
                  />
                </div>

                {/* Phone */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">No. WhatsApp / Telepon</label>
                  <input
                    type="tel"
                    placeholder="Contoh: 08123456789"
                    value={editingJamaah.phone}
                    onChange={(e) => setEditingJamaah({ ...editingJamaah, phone: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500"
                  />
                </div>

                {/* Email */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Alamat Email (Opsional)</label>
                  <input
                    type="email"
                    placeholder="Contoh: jemaah@travel.com"
                    value={editingJamaah.email || ''}
                    onChange={(e) => setEditingJamaah({ ...editingJamaah, email: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500"
                  />
                </div>

                {/* Travel Agent — hanya admin/kantor pusat yang boleh mengubah nama travel */}
                {activeOperatorId === null ? (
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Agen Travel / Rombongan *</label>
                    <input
                      type="text"
                      required
                      placeholder="Contoh: Raudhah Al-Haramain Travel"
                      value={editingJamaah.travel || ''}
                      onChange={(e) => setEditingJamaah({ ...editingJamaah, travel: e.target.value.toUpperCase() })}
                      className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500"
                    />
                  </div>
                ) : (
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Agen Travel / Rombongan</label>
                    <div className="w-full text-xs border border-slate-200 dark:border-zinc-700 rounded-lg p-2.5 bg-slate-100 dark:bg-zinc-800 text-slate-500 dark:text-zinc-400 flex items-center justify-between">
                      <span>{editingJamaah.travel || '-'}</span>
                      <span className="text-[10px] italic">Hanya admin yang dapat mengubah</span>
                    </div>
                  </div>
                )}

                {/* Password */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Password Akses Jemaah *</label>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      required
                      placeholder="Ketik password..."
                      value={editingJamaah.password || ''}
                      onChange={(e) => setEditingJamaah({ ...editingJamaah, password: e.target.value })}
                      className="flex-1 text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 font-mono font-semibold"
                    />
                    <button
                      type="button"
                      onClick={() => setEditingJamaah({
                        ...editingJamaah,
                        password: Math.floor(100000 + Math.random() * 900000).toString()
                      })}
                      className="px-3 bg-slate-100 dark:bg-zinc-700 hover:bg-slate-200 text-slate-700 dark:text-zinc-200 text-xs font-semibold rounded-lg border border-slate-200 dark:border-zinc-600 transition-colors"
                    >
                      Acak
                    </button>
                  </div>
                </div>

                {/* Entry Date */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Tanggal Masuk Madinah *</label>
                  <input
                    type="date"
                    required
                    value={toDateValue(editingJamaah.entryMadinah)}
                    onChange={(e) => setEditingJamaah({ ...editingJamaah, entryMadinah: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 font-mono"
                  />
                </div>

                {/* Exit Date */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Tanggal Keluar Madinah *</label>
                  <input
                    type="date"
                    required
                    value={editingJamaah.exitMadinah}
                    onChange={(e) => setEditingJamaah({ ...editingJamaah, exitMadinah: e.target.value })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 font-mono"
                  />
                </div>

                {/* Raudhah Entry Slot */}
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300 flex items-center gap-1">
                    <Clock className="w-3.5 h-3.5 text-emerald-600" />
                    <span>Jadwal Slot Raudhah (Nusuk)</span>
                  </label>
                  <input
                    type="datetime-local"
                    value={editingJamaah.raudhahSlot || ''}
                    min={editingJamaah.entryMadinah ? `${toDateValue(editingJamaah.entryMadinah)}T00:00` : undefined}
                    max={editingJamaah.exitMadinah ? `${toDateValue(editingJamaah.exitMadinah)}T23:59` : undefined}
                    onChange={(e) => setEditingJamaah({ ...editingJamaah, raudhahSlot: e.target.value || null })}
                    className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 font-mono"
                  />
                  {isSlotOutsideStay(editingJamaah.raudhahSlot, editingJamaah.entryMadinah, editingJamaah.exitMadinah) && (
                    <div className="text-[10px] text-amber-700 bg-amber-50 border border-amber-100 p-2 rounded-lg font-semibold flex items-center gap-1.5 mt-1 dark:bg-amber-950/20 dark:border-amber-900/40 dark:text-amber-400">
                      <AlertTriangle className="w-3.5 h-3.5 text-amber-600 shrink-0" />
                      <span>Slot di luar rentang Masuk–Keluar Madinah ({formatDateLabel(editingJamaah.entryMadinah || '')} – {formatDateLabel(editingJamaah.exitMadinah || '')})!</span>
                    </div>
                  )}
                  {(() => {
                    const dist = getDistributionInstant(editingJamaah.raudhahSlot, settingsQrLeadHours);
                    if (!dist) return null;
                    return (
                      <div className="text-[10px] text-sky-800 bg-sky-50/70 border border-sky-100 p-2 rounded-lg flex items-center gap-1.5 mt-1 dark:bg-sky-900/15 dark:border-sky-900/40 dark:text-sky-300">
                        <Download className="w-3.5 h-3.5 text-sky-600 shrink-0" />
                        <span>Distribusi QR (−{settingsQrLeadHours} jam): <strong>{formatInZone(dist, TZ_WITA, true)} WITA</strong> ({formatInZone(dist, TZ_MADINAH, true)} Madinah)</span>
                      </div>
                    );
                  })()}
                </div>

              </div>

              {/* Dynamic Custom Fields */}
              {customFields.length > 0 && (
                <div className="bg-slate-50 dark:bg-zinc-800/40 p-4 border border-slate-200/60 dark:border-zinc-700/60 rounded-xl space-y-3 mb-4">
                  <span className="text-xs font-bold text-slate-800 dark:text-zinc-100 block border-b pb-1">Kolom Tambahan (Kustom)</span>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    {customFields.map((cf) => (
                      <div className="space-y-1" key={cf.id}>
                        <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">{cf.label}</label>
                        <input
                          type="text"
                          placeholder={`Ketik ${cf.label.toLowerCase()}...`}
                          value={editingJamaah.customValues?.[cf.id] || ''}
                          onChange={(e) => {
                            const vals = editingJamaah.customValues || {};
                            setEditingJamaah({
                              ...editingJamaah,
                              customValues: {
                                ...vals,
                                [cf.id]: e.target.value
                              }
                            });
                          }}
                          className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-white dark:bg-zinc-700/40 outline-hidden focus:border-red-500 text-slate-800 dark:text-zinc-100"
                        />
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Notes */}
              <div className="space-y-1">
                <label className="text-xs font-semibold text-slate-600 dark:text-zinc-300">Catatan Khusus (Opsional)</label>
                <textarea
                  rows={2}
                  placeholder="Contoh: Kursi roda, lansia, dll..."
                  value={editingJamaah.notes || ''}
                  onChange={(e) => setEditingJamaah({ ...editingJamaah, notes: e.target.value })}
                  className="w-full text-xs border border-slate-200 dark:border-zinc-600 rounded-lg p-2.5 bg-slate-50 dark:bg-zinc-700/40 outline-hidden focus:border-red-500 resize-none"
                />
              </div>

              {/* Form Actions */}
              <div className="flex items-center justify-end gap-2 pt-4 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => {
                    setShowEditModal(false);
                    setEditingJamaah(null);
                  }}
                  className="px-4 py-2 rounded-lg border border-slate-200 dark:border-zinc-600 text-xs font-medium text-slate-700 dark:text-zinc-200 bg-white dark:bg-zinc-800 hover:bg-slate-50 dark:bg-zinc-800/50 dark:hover:bg-zinc-700 transition-colors cursor-pointer"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 rounded-lg bg-red-600 hover:bg-red-700 text-white text-xs font-semibold shadow-xs transition-colors cursor-pointer"
                >
                  Simpan Perubahan
                </button>
              </div>

            </form>

          </div>
        </div>
      )}

      {/* MODAL: DETEKSI & BERSIHKAN DUPLIKAT */}
      {showDuplicateModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center z-50 p-4 animate-in fade-in duration-200">
          <div className="bg-white dark:bg-zinc-900 rounded-2xl shadow-2xl w-full max-w-3xl max-h-[90vh] overflow-hidden flex flex-col animate-in zoom-in-95 duration-200 border border-slate-200 dark:border-zinc-800">
            {/* Header */}
            <div className="px-6 py-4 border-b border-slate-100 dark:border-zinc-800 bg-amber-50 dark:bg-amber-950/20 flex items-center justify-between">
              <div>
                <h3 className="font-bold text-amber-800 dark:text-amber-300 text-base flex items-center gap-2">
                  <AlertTriangle className="w-5 h-5" />
                  <span>Deteksi & Bersihkan Duplikat</span>
                </h3>
                <p className="text-xs text-amber-700/80 dark:text-amber-400/70 font-medium">
                  Jemaah dengan Nomor Paspor atau Visa yang sama. Pertahankan satu, hapus sisanya.
                </p>
              </div>
              <button type="button" onClick={() => setShowDuplicateModal(false)} className="p-1.5 rounded-lg hover:bg-amber-100 dark:hover:bg-amber-900/30 text-amber-500 transition-all">
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Body */}
            <div className="flex-1 overflow-y-auto p-6 space-y-4">
              {duplicateGroups.length === 0 ? (
                <div className="text-center py-12 space-y-3">
                  <div className="w-14 h-14 rounded-2xl bg-emerald-50 dark:bg-emerald-950/30 text-emerald-500 flex items-center justify-center mx-auto">
                    <CheckCircle2 className="w-7 h-7" />
                  </div>
                  <p className="text-sm font-semibold text-slate-700 dark:text-zinc-200">Tidak ada duplikat 🎉</p>
                  <p className="text-xs text-slate-500 dark:text-zinc-400">Semua data jemaah memiliki Paspor & Visa yang unik.</p>
                </div>
              ) : (
                <>
                  <p className="text-xs text-slate-500 dark:text-zinc-400">
                    Ditemukan <strong className="text-amber-600">{duplicateGroups.length} grup duplikat</strong> ({duplicateIds.size} data). Klik <strong>Hapus</strong> pada data yang ingin dibuang.
                  </p>
                  {duplicateGroups.map((group, gi) => (
                    <div key={gi} className="border border-amber-200 dark:border-amber-900/40 rounded-xl overflow-hidden">
                      <div className="px-3 py-2 bg-amber-50/60 dark:bg-amber-950/10 text-[11px] font-bold text-amber-700 dark:text-amber-400 border-b border-amber-100 dark:border-amber-900/30">
                        Grup #{gi + 1} — {group.length} data sama
                      </div>
                      <div className="divide-y divide-slate-100 dark:divide-zinc-800">
                        {group.map((j, ji) => (
                          <div key={j.id} className="flex items-center gap-3 p-3 hover:bg-slate-50 dark:hover:bg-zinc-800/40">
                            <div className="flex-1 min-w-0">
                              <div className="text-xs font-semibold text-slate-800 dark:text-zinc-100 truncate">
                                {j.name || <span className="text-red-500 italic">[Tanpa Nama]</span>}
                                {ji === 0 && <span className="ml-2 text-[9px] px-1.5 py-0.5 rounded-full bg-emerald-100 text-emerald-700 font-bold">disarankan disimpan</span>}
                              </div>
                              <div className="text-[11px] text-slate-500 dark:text-zinc-400 font-mono flex flex-wrap gap-x-3">
                                <span>Paspor: {j.passport || '-'}</span>
                                <span>Visa: {j.visa || '-'}</span>
                                <span className="truncate">Travel: {j.travel || '-'}</span>
                              </div>
                            </div>
                            <button
                              type="button"
                              onClick={() => handleDeleteJamaah(j.id)}
                              className="shrink-0 flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 dark:bg-red-950/30 dark:hover:bg-red-950/50 dark:text-red-400 text-[11px] font-semibold transition-colors"
                            >
                              <Trash2 className="w-3.5 h-3.5" /> Hapus
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </>
              )}
            </div>

            {/* Footer */}
            <div className="px-6 py-3 border-t border-slate-100 dark:border-zinc-800 flex justify-end">
              <button type="button" onClick={() => setShowDuplicateModal(false)} className="px-4 py-2 rounded-lg bg-zinc-900 text-white text-xs font-semibold hover:bg-zinc-800 transition-colors">
                Selesai
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: BATCH VISA SCAN (PDF teks digital saja) */}
      {showBatchScanModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center z-50 p-4 animate-in fade-in duration-200">
          <div className="bg-white dark:bg-zinc-900 rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col animate-in zoom-in-95 duration-200 border border-slate-200 dark:border-zinc-800">
            
            {/* Header */}
            <div className="px-6 py-4 border-b border-slate-100 dark:border-zinc-800 bg-slate-50 dark:bg-zinc-800/40 flex items-center justify-between">
              <div>
                <h3 className="font-bold text-slate-800 dark:text-zinc-100 text-base flex items-center gap-2">
                  <Sparkles className="w-5 h-5 text-blue-600 animate-pulse" />
                  <span>Scan Massal Visa Jemaah (Maks 50 File)</span>
                </h3>
                <p className="text-xs text-slate-400 dark:text-zinc-400 font-medium">Unggah kumpulan PDF visa (teks digital) untuk dideteksi otomatis secara massal.</p>
              </div>
              <button
                type="button"
                disabled={isBatchScanning}
                onClick={() => setShowBatchScanModal(false)}
                className="p-1.5 rounded-lg hover:bg-slate-200 dark:hover:bg-zinc-800 text-slate-400 hover:text-slate-700 dark:text-zinc-300 transition-all cursor-pointer disabled:opacity-40"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Body */}
            <div className="flex-1 overflow-y-auto p-6 space-y-5">
              
              {/* API Key Warning Alert if empty */}
              {!settingsGeminiApiKey && (
                <div className="p-4 bg-amber-50 dark:bg-amber-950/20 border border-amber-200 dark:border-amber-900/40 rounded-xl flex items-start gap-3">
                  <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                  <div className="text-xs space-y-1">
                    <strong className="text-amber-800 dark:text-amber-400 block">Gemini API Key Belum Dikonfigurasi!</strong>
                    <p className="text-amber-700/90 dark:text-amber-300/80 leading-relaxed">Sistem membutuhkan API Key untuk menjalankan pemindaian AI. Harap buka menu <strong>Settings</strong> di dashboard untuk menyetel API Key Anda terlebih dahulu.</p>
                  </div>
                </div>
              )}

              {/* Upload Dropzone (When not scanning and no results or errors yet) */}
              {!isBatchScanning && batchScanResults.length === 0 && batchScanErrors.length === 0 && (
                <div className="space-y-4">
                  <div className="border-2 border-dashed border-slate-200 dark:border-zinc-700 hover:border-blue-500 dark:hover:border-blue-400 rounded-2xl p-8 text-center bg-slate-50/50 dark:bg-zinc-800/20 transition-all relative group">
                    <input
                      type="file"
                      multiple
                      accept="application/pdf"
                      disabled={!settingsGeminiApiKey}
                      onChange={(e) => {
                        if (e.target.files && e.target.files.length > 0) {
                          const files = Array.from(e.target.files);
                          if (files.length > 50) {
                            alert('Batas maksimal scan massal adalah 50 file sekaligus!');
                            return;
                          }
                          setBatchScanFiles(files);
                        }
                      }}
                      className="absolute inset-0 w-full h-full opacity-0 cursor-pointer disabled:cursor-not-allowed"
                    />
                    <div className="space-y-3 pointer-events-none">
                      <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-blue-500 to-indigo-600 text-white flex items-center justify-center mx-auto group-hover:scale-105 shadow-lg shadow-blue-500/20 transition-transform duration-200">
                        <Upload className="w-6 h-6" />
                      </div>
                      <div className="space-y-1">
                        <p className="text-xs font-semibold text-slate-700 dark:text-zinc-200">
                          <span className="text-blue-600 dark:text-blue-400 hover:underline">Klik untuk memilih berkas</span> atau seret berkas PDF Visa ke sini
                        </p>
                        <p className="text-[10px] text-slate-400 dark:text-zinc-400">Hanya mendukung berkas PDF Visa asli dengan teks digital (Maks. 50 file, maks. 20MB per file)</p>
                      </div>
                    </div>
                  </div>

                  {/* Selected Files List */}
                  {batchScanFiles.length > 0 && (
                    <div className="space-y-2.5 animate-in fade-in duration-150">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-700 dark:text-zinc-200 flex items-center gap-1.5">
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500" />
                          {batchScanFiles.length} berkas siap dipindai
                          <span className="font-medium text-slate-400 dark:text-zinc-500">
                            · {(batchScanFiles.reduce((s, f) => s + f.size, 0) / 1024 / 1024).toFixed(2)} MB total
                          </span>
                        </span>
                        <button
                          type="button"
                          onClick={() => setBatchScanFiles([])}
                          className="text-[10px] text-red-600 hover:text-red-700 hover:underline font-semibold flex items-center gap-1"
                        >
                          <Trash2 className="w-3 h-3" /> Bersihkan
                        </button>
                      </div>
                      <div className="max-h-48 overflow-y-auto border border-slate-200 dark:border-zinc-800 rounded-xl p-2 space-y-1.5 bg-white dark:bg-zinc-800/30">
                        {batchScanFiles.map((file, idx) => (
                          <div key={idx} className="group/file flex items-center gap-2.5 text-[11px] p-2 bg-slate-50 dark:bg-zinc-800/70 rounded-lg border border-slate-100 dark:border-zinc-700 hover:border-blue-300 dark:hover:border-blue-700/50 transition-colors">
                            <div className="w-7 h-7 shrink-0 rounded-lg bg-red-50 dark:bg-red-950/30 text-red-500 dark:text-red-400 flex items-center justify-center">
                              <FileText className="w-3.5 h-3.5" />
                            </div>
                            <span className="font-medium text-slate-700 dark:text-zinc-200 truncate flex-1" title={file.name}>{file.name}</span>
                            <span className="text-slate-400 shrink-0 tabular-nums">{(file.size / 1024 / 1024).toFixed(2)} MB</span>
                            <button
                              type="button"
                              onClick={() => setBatchScanFiles(prev => prev.filter((_, i) => i !== idx))}
                              className="shrink-0 p-1 rounded-md text-slate-300 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-950/30 opacity-0 group-hover/file:opacity-100 transition-all"
                              title="Hapus berkas ini"
                            >
                              <X className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        ))}
                      </div>

                      <button
                        type="button"
                        onClick={() => processBatchVisaScan(batchScanFiles)}
                        className="w-full py-3 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white rounded-xl text-xs font-bold transition-all shadow-sm hover:shadow-md hover:shadow-blue-500/20 flex items-center justify-center gap-2 cursor-pointer active:scale-[0.99]">
                        <Sparkles className="w-4 h-4" />
                        <span>Mulai Pemindaian Otomatis ({batchScanFiles.length} File)</span>
                      </button>
                    </div>
                  )}
                </div>
              )}

              {/* Progress Panel (When scanning is active) */}
              {isBatchScanning && (
                <div className="p-8 border border-slate-200 dark:border-zinc-800 rounded-2xl bg-slate-50/50 dark:bg-zinc-800/10 text-center space-y-5">
                  <div className="relative w-16 h-16 mx-auto">
                    <div className="absolute inset-0 rounded-full border-4 border-slate-200 dark:border-zinc-700"></div>
                    <div className="absolute inset-0 rounded-full border-4 border-t-blue-600 border-r-blue-600 animate-spin"></div>
                  </div>

                  <div className="space-y-2">
                    <h4 className="font-bold text-slate-800 dark:text-zinc-100 text-sm">
                      Memindai Visa Jemaah ({batchScanProgress.current} dari {batchScanProgress.total})
                    </h4>
                    <p className="text-xs text-slate-500 dark:text-zinc-400 font-mono tracking-tight font-medium max-w-md mx-auto truncate font-semibold">
                      {batchScanProgress.status}
                    </p>
                  </div>

                  {/* Progress Bar */}
                  <div className="w-full max-w-md mx-auto bg-slate-200 dark:bg-zinc-700 h-2.5 rounded-full overflow-hidden">
                    <div
                      className="bg-gradient-to-r from-blue-600 to-indigo-600 h-full rounded-full transition-all duration-300"
                      style={{ width: `${(batchScanProgress.current / batchScanProgress.total) * 100}%` }}
                    />
                  </div>

                  {/* File Metadata Card */}
                  {batchScanFileMeta && (
                    <div className="max-w-md mx-auto p-4 rounded-xl bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-700/60 text-left space-y-2.5 shadow-xs animate-in fade-in duration-200">
                      <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 dark:text-zinc-500">Berkas Aktif Sedang Diproses</span>
                      <div className="space-y-1.5 text-xs">
                        <div className="flex justify-between">
                          <span className="text-slate-500 dark:text-zinc-400">Nama File:</span>
                          <span className="font-semibold text-slate-800 dark:text-zinc-200 truncate max-w-[200px]" title={batchScanFileMeta.name}>
                            {batchScanFileMeta.name}
                          </span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-slate-500 dark:text-zinc-400">Tipe Dokumen:</span>
                          <span className="font-semibold px-2 py-0.5 rounded-full text-[10px] bg-blue-50 text-blue-700 dark:bg-blue-950/30 dark:text-blue-400">
                            {batchScanFileMeta.type}
                          </span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-slate-500 dark:text-zinc-400">Ukuran File:</span>
                          <span className="font-semibold text-slate-800 dark:text-zinc-200">{batchScanFileMeta.size}</span>
                        </div>
                        {batchScanFileMeta.pages > 1 && (
                          <div className="flex justify-between">
                            <span className="text-slate-500 dark:text-zinc-400">Jumlah Halaman (Pax):</span>
                            <span className="font-bold text-slate-800 dark:text-zinc-200">{batchScanFileMeta.pages} Halaman (Multipage PDF)</span>
                          </div>
                        )}
                        {batchScanFileMeta.charCount > 0 && (
                          <div className="flex justify-between">
                            <span className="text-slate-500 dark:text-zinc-400">Teks Terekstrak:</span>
                            <span className="font-semibold text-slate-800 dark:text-zinc-200">{batchScanFileMeta.charCount.toLocaleString()} Karakter</span>
                          </div>
                        )}

                        {/* Active File Progress Bar */}
                        <div className="pt-2.5 border-t border-slate-100 dark:border-zinc-700/50 space-y-1.5">
                          <div className="flex justify-between text-[10px] font-semibold text-slate-700 dark:text-zinc-300">
                            <span>Kemajuan Pemrosesan Berkas</span>
                            <span className="font-mono text-blue-600 dark:text-blue-400 font-bold">{batchActiveFileProgress}%</span>
                          </div>
                          <div className="w-full bg-slate-100 dark:bg-zinc-700 h-2 rounded-full overflow-hidden">
                            <div 
                              className="bg-gradient-to-r from-blue-500 to-blue-600 h-full rounded-full transition-all duration-300"
                              style={{ width: `${batchActiveFileProgress}%` }}
                            />
                          </div>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* Step Tracker Card */}
                  <div className="max-w-md mx-auto p-4 rounded-xl bg-white dark:bg-zinc-800 border border-slate-200 dark:border-zinc-700/60 text-left space-y-3 shadow-xs">
                    <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 dark:text-zinc-500">Tahapan Proses Saat Ini</span>
                    
                    <div className="space-y-3 text-xs">
                      {/* Step 1: PDF Text Extraction */}
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold border ${
                            batchActiveFileProgress >= 55
                              ? 'bg-emerald-50 border-emerald-200 text-emerald-700 dark:bg-emerald-950/30 dark:border-emerald-900/40 dark:text-emerald-400'
                              : 'bg-blue-50 border-blue-200 text-blue-700 dark:bg-blue-950/30 dark:border-blue-900/40 dark:text-blue-400'
                          }`}>
                            {batchActiveFileProgress >= 55 ? '✓' : '1'}
                          </span>
                          <span className={`${batchActiveFileProgress >= 55 ? 'line-through text-slate-400 dark:text-zinc-500' : 'font-semibold text-slate-800 dark:text-zinc-200'}`}>
                            Mengekstrak Teks Digital dari PDF
                          </span>
                        </div>
                        {batchActiveFileProgress < 55 && (
                          <Loader2 className="w-3.5 h-3.5 text-blue-500 animate-spin" />
                        )}
                      </div>

                      {/* Step 2: Gemini AI Processing */}
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold border ${
                            batchActiveFileProgress >= 100
                              ? 'bg-emerald-50 border-emerald-200 text-emerald-700 dark:bg-emerald-950/30 dark:border-emerald-900/40 dark:text-emerald-400'
                              : batchActiveFileProgress >= 55
                              ? 'bg-blue-50 border-blue-200 text-blue-700 dark:bg-blue-950/30 dark:border-blue-900/40 dark:text-blue-400'
                              : 'bg-slate-50 border-slate-200 text-slate-400 dark:bg-zinc-800 dark:border-zinc-700'
                          }`}>
                            {batchActiveFileProgress >= 100 ? '✓' : '2'}
                          </span>
                          <span className={`${
                            batchActiveFileProgress >= 100
                              ? 'line-through text-slate-400 dark:text-zinc-500'
                              : batchActiveFileProgress >= 55
                              ? 'font-semibold text-slate-800 dark:text-zinc-200'
                              : 'text-slate-400 dark:text-zinc-500'
                          }`}>
                            Analisis &amp; Pemetaan Teks oleh Gemini AI
                          </span>
                        </div>
                        {batchActiveFileProgress >= 55 && batchActiveFileProgress < 100 && (
                          <Loader2 className="w-3.5 h-3.5 text-blue-500 animate-spin" />
                        )}
                      </div>

                      {/* Step 3: Verification & Duplicates Prevention */}
                      <div className="flex items-center gap-2">
                        <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold border ${
                          batchActiveFileProgress >= 100
                            ? 'bg-emerald-50 border-emerald-200 text-emerald-700 dark:bg-emerald-950/30 dark:border-emerald-900/40 dark:text-emerald-400'
                            : 'bg-slate-50 border-slate-200 text-slate-400 dark:bg-zinc-800 dark:border-zinc-700'
                        }`}>
                          {batchActiveFileProgress >= 100 ? '✓' : '3'}
                        </span>
                        <span className={`${batchActiveFileProgress >= 100 ? 'font-semibold text-slate-800 dark:text-zinc-200' : 'text-slate-400 dark:text-zinc-500'}`}>
                          Proteksi Duplikat &amp; Validasi Data
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* Visual Status Explanation Card */}
                  <div className="max-w-md mx-auto p-4 rounded-xl bg-blue-50/50 dark:bg-blue-950/10 border border-blue-100 dark:border-blue-900/30 text-left text-xs space-y-2">
                    <span className="font-bold text-blue-800 dark:text-blue-400 flex items-center gap-1">
                      <Sparkles className="w-3.5 h-3.5" />
                      Informasi Pemrosesan AI
                    </span>
                    <ul className="list-disc pl-4 space-y-1 text-slate-600 dark:text-zinc-400 text-[11px] leading-relaxed">
                      <li><strong>Gemini AI sedang bekerja:</strong> Membaca dokumen teks besar (seperti PDF multipage {batchScanFileMeta?.pages || ''} halaman) membutuhkan waktu sekitar 10-30 detik untuk diurai.</li>
                      <li><strong>Status Cooldown/Rate-Limit:</strong> Jika Anda melihat teks status berubah ke <em>Rate limit terlampaui</em>, program akan otomatis menjeda sementara (cooldown) lalu melanjutkan memproses tanpa merusak atau menghilangkan data yang sudah terisi.</li>
                      <li><strong>Tindakan Pengguna:</strong> Harap pertahankan halaman browser tetap terbuka dan jangan menutup atau merefresh dashboard selama pemindaian masih berjalan.</li>
                    </ul>
                  </div>

                  {/* Stopwatch Ticker */}
                  <div className="flex items-center justify-center gap-1.5 text-xs font-semibold text-slate-600 dark:text-zinc-400 pt-1">
                    <Clock className="w-3.5 h-3.5 text-blue-500 animate-pulse" />
                    <span>Durasi Berjalan: <span className="font-mono text-slate-800 dark:text-zinc-200 font-bold bg-slate-100 dark:bg-zinc-800 px-1.5 py-0.5 rounded-md">{batchScanTimeElapsed} detik</span></span>
                  </div>

                  {/* Controls (Pause/Cancel) */}
                  <div className="flex justify-center gap-3 pt-2">
                    <button
                      type="button"
                      onClick={() => {
                        pauseBatchRef.current = !pauseBatchRef.current;
                        setIsBatchScanPaused(pauseBatchRef.current);
                      }}
                      className={`px-4 py-2 text-xs font-semibold rounded-lg border transition-all ${
                        isBatchScanPaused
                          ? 'bg-emerald-50 border-emerald-200 text-emerald-700 dark:bg-emerald-950/20 dark:border-emerald-900/30'
                          : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50 dark:bg-zinc-800 dark:border-zinc-700 dark:text-zinc-200 dark:hover:bg-zinc-700'
                      }`}
                    >
                      {isBatchScanPaused ? '▶️ Lanjutkan' : '⏸️ Jeda'}
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        cancelBatchRef.current = true;
                      }}
                      className="px-4 py-2 bg-red-50 hover:bg-red-100 border border-red-200/50 text-red-700 dark:bg-red-950/20 dark:border-red-900/30 dark:text-red-400 text-xs font-semibold rounded-lg transition-all"
                    >
                      🚫 Batalkan
                    </button>
                  </div>
                </div>
              )}

              {/* Review Table & Scan Summary (When scanning completes or results/errors are present) */}
              {!isBatchScanning && (batchScanResults.length > 0 || batchScanErrors.length > 0) && (
                <div className="space-y-5 animate-in fade-in duration-200">
                  
                  {/* Results Metrics */}
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <div className="p-3 bg-emerald-50 dark:bg-emerald-950/20 border border-emerald-100 dark:border-emerald-900/30 rounded-xl text-center">
                      <span className="text-[10px] text-emerald-600 dark:text-emerald-400 block font-bold uppercase tracking-wider">Berkas Sukses</span>
                      <strong className="text-xl text-emerald-700 dark:text-emerald-300 font-extrabold">{batchScanSuccessFilesCount}</strong>
                    </div>
                    <div className="p-3 bg-rose-50 dark:bg-rose-950/20 border border-rose-100 dark:border-rose-900/30 rounded-xl text-center">
                      <span className="text-[10px] text-rose-600 dark:text-rose-400 block font-bold uppercase tracking-wider">Berkas Gagal</span>
                      <strong className="text-xl text-rose-700 dark:text-rose-300 font-extrabold">{batchScanFailedFilesCount}</strong>
                    </div>
                    <div className="p-3 bg-slate-50 dark:bg-zinc-800/40 border border-slate-200 dark:border-zinc-700/60 rounded-xl text-center">
                      <span className="text-[10px] text-slate-500 block font-bold uppercase tracking-wider">Total Berkas</span>
                      <strong className="text-xl text-slate-700 dark:text-zinc-200 font-extrabold">{batchScanTotalFilesCount}</strong>
                    </div>
                    <div className="p-3 bg-blue-50 dark:bg-blue-950/20 border border-blue-100 dark:border-blue-900/30 rounded-xl text-center col-span-2 md:col-span-1">
                      <span className="text-[10px] text-blue-600 dark:text-blue-400 block font-bold uppercase tracking-wider">Jemaah Terbaca</span>
                      <strong className="text-xl text-blue-700 dark:text-blue-300 font-extrabold">{batchScanResults.length}</strong>
                    </div>
                  </div>

                  {/* Errors details if any */}
                  {batchScanErrors.length > 0 && (
                    <div className="p-3 bg-rose-500/5 border border-rose-200 dark:border-rose-900/40 rounded-xl space-y-2">
                      <span className="text-xs font-bold text-rose-700 dark:text-rose-400 flex items-center gap-1">
                        <AlertTriangle className="w-4 h-4 shrink-0" />
                        <span>Detail File Gagal Scan ({batchScanErrors.length}):</span>
                      </span>
                      <div className="max-h-28 overflow-y-auto space-y-1.5 text-[10px] font-mono p-1">
                        {batchScanErrors.map((err, idx) => (
                          <div key={idx} className="flex justify-between border-b dark:border-zinc-800 pb-1 text-slate-600 dark:text-zinc-300">
                            <span className="font-semibold truncate max-w-[50%]">{err.fileName}</span>
                            <span className="text-rose-600 text-right">{err.error}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Results review table */}
                  <div className="space-y-2">
                    <span className="text-xs font-bold text-slate-800 dark:text-zinc-200 block">Tinjau Hasil Deteksi Visa:</span>
                    <div className="overflow-x-auto border border-slate-200 dark:border-zinc-800 rounded-xl bg-white dark:bg-zinc-800/30">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-slate-50 dark:bg-zinc-800 border-b border-slate-200 dark:border-zinc-800 text-slate-500 dark:text-zinc-400 font-bold">
                            <th className="py-2.5 px-3">Berkas</th>
                            <th className="py-2.5 px-3">Nama Lengkap</th>
                            <th className="py-2.5 px-3">No Paspor</th>
                            <th className="py-2.5 px-3">No Visa</th>
                            <th className="py-2.5 px-3">Gender</th>
                            <th className="py-2.5 px-3">Travel</th>
                            <th className="py-2.5 px-3">Masuk Madinah</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 dark:divide-zinc-800">
                          {batchScanResults.map((r, idx) => (
                            <tr key={idx} className="hover:bg-slate-50/50 dark:hover:bg-zinc-700/20">
                              <td className="py-2.5 px-3 font-mono text-[10px] text-slate-400 truncate max-w-[120px]" title={r.fileName}>{r.fileName}</td>
                              <td className="py-2.5 px-3 font-medium text-slate-800 dark:text-zinc-100">{r.name || <span className="text-red-500 italic">[Kosong]</span>}</td>
                              <td className="py-2.5 px-3 font-mono text-slate-700 dark:text-zinc-300">{r.passport || <span className="text-red-500 italic">[Kosong]</span>}</td>
                              <td className="py-2.5 px-3 font-mono text-slate-700 dark:text-zinc-300">{r.visa || <span className="text-red-500 italic">[Kosong]</span>}</td>
                              <td className="py-2.5 px-3 text-slate-600 dark:text-zinc-300">{r.gender}</td>
                              <td className="py-2.5 px-3 text-slate-600 dark:text-zinc-300 truncate max-w-[120px]" title={r.travel}>{r.travel}</td>
                              <td className="py-2.5 px-3 font-mono text-slate-500 dark:text-zinc-400">{r.entryMadinah || '-'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                </div>
              )}

            </div>

            {/* Modal Footer */}
            <div className="px-6 py-4 border-t border-slate-100 dark:border-zinc-800 bg-slate-50 dark:bg-zinc-900/50 flex items-center justify-between gap-3">
              <button
                type="button"
                disabled={isBatchScanning}
                onClick={() => {
                  setBatchScanFiles([]);
                  setBatchScanResults([]);
                  setBatchScanErrors([]);
                  setBatchScanSuccessFilesCount(0);
                  setBatchScanFailedFilesCount(0);
                  setBatchScanTotalFilesCount(0);
                  setIsBatchScanning(false);
                }}
                className="px-4 py-2 border border-slate-200 dark:border-zinc-700 rounded-lg text-xs font-semibold text-slate-700 dark:text-zinc-200 hover:bg-slate-100 dark:hover:bg-zinc-800 transition-colors disabled:opacity-40 cursor-pointer"
              >
                Reset / Ulangi
              </button>

              <div className="flex gap-2">
                <button
                  type="button"
                  disabled={isBatchScanning}
                  onClick={() => setShowBatchScanModal(false)}
                  className="px-4 py-2 border border-slate-200 dark:border-zinc-700 rounded-lg text-xs font-medium text-slate-700 dark:text-zinc-200 hover:bg-slate-100 dark:hover:bg-zinc-800 transition-colors disabled:opacity-40 cursor-pointer"
                >
                  Batal
                </button>
                <button
                  type="button"
                  disabled={isBatchScanning || batchScanResults.length === 0}
                  onClick={saveBatchScanResults}
                  className="px-4 py-2 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white rounded-lg text-xs font-bold transition-all shadow-xs disabled:opacity-40 disabled:pointer-events-none cursor-pointer"
                >
                  Simpan Semua Ke Database ({batchScanResults.length} Jemaah)
                </button>
              </div>
            </div>

          </div>
        </div>
      )}

    </div>
  );
}
