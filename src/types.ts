export type Gender = 'Laki-laki' | 'Perempuan';

export type JamaahStatus = 'Ready' | 'Sedang War' | 'QR Berhasil' | 'QR Terdistribusi' | 'Belum Berhasil' | 'Visa Tidak Tersedia';

export interface CustomField {
  id: string;
  label: string;
}

export interface Jamaah {
  id: string;
  name: string;
  passport: string;
  visa: string;
  gender: Gender | ''; // '' = belum ditentukan (mis. tidak ada di Excel saat import)
  phone: string;
  entryMadinah: string; // YYYY-MM-DD
  exitMadinah: string; // YYYY-MM-DD
  operatorId: string | null; // ID of the assigned operator
  status: JamaahStatus;
  notes: string;
  qrCodeUrl: string | null; // Base64 or ObjectURL of uploaded screenshot
  qrUploadedAt: string | null; // ISO Date String
  createdAt: string;
  travel?: string; // Travel agency / Rombongan name
  email?: string; // Email address of the jamaah
  password?: string; // Access password for the jamaah
  raudhahSlot?: string | null; // Booked Raudhah slot datetime string (YYYY-MM-DDTHH:MM)
  customValues?: Record<string, string>; // Dynamic custom fields values mapping: fieldId -> value
}

export interface Operator {
  id: string;
  name: string;
  phone: string;
  password?: string; // Password for operator access
  isActive: boolean;
}

export type PriorityLevel = 'Tinggi' | 'Sedang' | 'Rendah' | 'Belum Ada';

export interface PriorityInfo {
  level: PriorityLevel;
  daysRemaining: number;
  badgeColor: string;
  dotColor: string;
}
