import { createClient, SupabaseClient } from '@supabase/supabase-js';

// Dynamic connection initialization
let supabaseInstance: SupabaseClient | null = null;
let lastUrl = '';
let lastKey = '';

export const getSupabaseConfig = () => {
  const envUrl = (import.meta as any).env?.VITE_SUPABASE_URL;
  const envKey = (import.meta as any).env?.VITE_SUPABASE_ANON_KEY;

  const localUrl = localStorage.getItem('raudhah_supabase_url');
  const localKey = localStorage.getItem('raudhah_supabase_anon_key');

  return {
    url: envUrl || localUrl || '',
    key: envKey || localKey || '',
    isFromEnv: !!(envUrl && envKey),
  };
};

export const initSupabaseClient = (customUrl?: string, customKey?: string): SupabaseClient | null => {
  const { url, key } = getSupabaseConfig();
  const finalUrl = customUrl || url;
  const finalKey = customKey || key;

  if (!finalUrl || !finalKey) {
    supabaseInstance = null;
    return null;
  }

  if (supabaseInstance && finalUrl === lastUrl && finalKey === lastKey) {
    return supabaseInstance;
  }

  try {
    lastUrl = finalUrl;
    lastKey = finalKey;
    supabaseInstance = createClient(finalUrl, finalKey, {
      auth: {
        persistSession: false
      }
    });
    return supabaseInstance;
  } catch (error) {
    console.error('Failed to initialize Supabase client:', error);
    supabaseInstance = null;
    return null;
  }
};

export const getSupabase = (): SupabaseClient | null => {
  if (supabaseInstance) return supabaseInstance;
  return initSupabaseClient();
};

// --- Mapping Utilities ---

export const mapJamaahToDb = (j: any) => ({
  id: j.id,
  name: j.name,
  passport: j.passport,
  visa: j.visa,
  gender: j.gender,
  phone: j.phone || '-',
  entry_madinah: j.entryMadinah,
  exit_madinah: j.exitMadinah,
  operator_id: j.operatorId,
  status: j.status,
  notes: j.notes || '',
  qr_code_url: j.qrCodeUrl,
  qr_uploaded_at: j.qrUploadedAt,
  created_at: j.createdAt,
  travel: j.travel || '',
  email: j.email || '',
  password: j.password || '',
  raudhah_slot: j.raudhahSlot,
  custom_values: j.customValues || {},
});

export const mapJamaahFromDb = (db: any) => ({
  id: db.id,
  name: db.name,
  passport: db.passport,
  visa: db.visa,
  gender: db.gender,
  phone: db.phone,
  entryMadinah: db.entry_madinah,
  exitMadinah: db.exit_madinah,
  operatorId: db.operator_id,
  status: db.status,
  notes: db.notes,
  qrCodeUrl: db.qr_code_url,
  qrUploadedAt: db.qr_uploaded_at,
  createdAt: db.created_at,
  travel: db.travel,
  email: db.email,
  password: db.password,
  raudhahSlot: db.raudhah_slot,
  customValues: db.custom_values || {},
});

export const mapOperatorToDb = (o: any) => ({
  id: o.id,
  name: o.name,
  phone: o.phone || '-',
  password: o.password || '123456',
  is_active: o.isActive,
});

export const mapOperatorFromDb = (db: any) => ({
  id: db.id,
  name: db.name,
  phone: db.phone,
  password: db.password,
  isActive: db.is_active,
});

export const mapCustomFieldToDb = (cf: any) => ({
  id: cf.id,
  label: cf.label,
});

export const mapCustomFieldFromDb = (db: any) => ({
  id: db.id,
  label: db.label,
});
