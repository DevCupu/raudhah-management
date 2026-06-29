// Logika murni untuk scan visa: ekstraksi teks PDF, render halaman PDF ke gambar,
// dan pemanggilan Gemini API. Dipisah dari App.tsx agar mudah dirawat & diuji.
// Tidak bergantung pada state React — semua input lewat parameter.
import * as pdfjsLib from 'pdfjs-dist';


// Prompt untuk jalur TEKS (PDF teks digital). Teks hasil ekstraksi disisipkan.
export const buildVisaTextPrompt = (pdfText: string): string =>
  `Berikut adalah hasil ekstraksi teks dari dokumen visa jemaah. Ekstrak data jemaah dari teks ini ke dalam format JSON Array. Setiap jemaah HARUS memiliki nomor paspor yang valid (diawali huruf diikuti angka, contoh: A1234567). JANGAN sertakan entri yang tidak memiliki nomor paspor valid. Gunakan kunci berikut:\n- name (Nama lengkap jemaah)\n- passport (Nomor Paspor, WAJIB diisi)\n- visa (Nomor Visa)\n- gender (Isi dengan \"Laki-laki\" atau \"Perempuan\" saja. Tebak secara cerdas dari nama jemaah jika jenis kelamin tidak tertulis eksplisit)\n- entryMadinah (Tanggal kedatangan/masuk Madinah dalam format YYYY-MM-DD. Jika tidak ditemukan, biarkan string kosong \"\")\n- exitMadinah (Tanggal keberangkatan/keluar Madinah dalam format YYYY-MM-DD. Jika tidak ditemukan, biarkan string kosong \"\")\n- travel (Nama travel/agen umrah yang tertera pada visa, cari di kolom 'External Agent' atau 'Umrah Operator'. Jika tertulis dalam bahasa/huruf Arab, terjemahkan atau transliterasikan ke huruf Latin Indonesia/Inggris secara cerdas, contoh: 'مجموعة بي TI رحمة الدولية' menjadi 'PT Rahma Internasional')\n\nTEKS VISA:\n${pdfText}\n\nHarap kembalikan HANYA string JSON Array mentah, misalnya: [ { \"name\": \"...\", \"passport\": \"...\" }, ... ]. Jangan gunakan blok format markdown atau teks penjelasan lainnya.`;

// Ekstrak lapisan teks digital dari PDF (kosong jika PDF hasil scan/gambar).
export const extractTextFromPdf = async (
  arrayBuffer: ArrayBuffer,
  onProgress?: (page: number, total: number) => void
): Promise<string> => {
  try {
    // Pakai SALINAN buffer (.slice) agar pdf.js tidak "melepas" (detach) buffer milik
    // pemanggil — buffer bisa dipakai lagi setelah fungsi ini selesai tanpa error.
    // @ts-ignore
    const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer.slice(0)) });
    const pdf = await loadingTask.promise;
    let fullText = '';
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map((item: any) => item.str).join(' ');
      fullText += `--- HALAMAN ${i} ---\n${pageText}\n\n`;
      if (onProgress) onProgress(i, pdf.numPages);
    }
    return fullText.trim();
  } catch (e) {
    console.error('Pdf text extraction failed:', e);
    return '';
  }
};


export interface VisaRequestOptions {
  apiKey: string;
  model: string;
  parts: any[];
  maxOut: number;
  thinkingConfig: Record<string, unknown>;
}

// Error scan dengan info rate-limit agar pemanggil bisa atur cooldown.
export interface VisaScanError extends Error {
  isRateLimit?: boolean;
  retryDelaySec?: number;
}

// Kirim sekumpulan "parts" ke Gemini, kembalikan array item JSON mentah.
// Melempar Error (dengan flag isRateLimit utk 429) agar penanganan retry seragam.
export const requestVisaExtraction = async ({
  apiKey,
  model,
  parts,
  maxOut,
  thinkingConfig,
}: VisaRequestOptions): Promise<any[]> => {
  const resp = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts }],
        generationConfig: {
          responseMimeType: 'application/json',
          maxOutputTokens: maxOut,
          temperature: 0,
          ...thinkingConfig,
        },
      }),
    }
  );

  if (!resp.ok) {
    let errorMsg = `API error status ${resp.status}`;
    let retryDelaySec: number | undefined;
    try {
      const errData = await resp.json();
      if (errData?.error?.message) errorMsg = errData.error.message;
      // Google menyertakan saran waktu tunggu di RetryInfo (mis. "47s").
      const retryInfo = errData?.error?.details?.find((d: any) => (d['@type'] || '').includes('RetryInfo'));
      const retryStr = retryInfo?.retryDelay;
      if (retryStr) {
        const secs = parseInt(String(retryStr).replace(/[^0-9]/g, ''), 10);
        if (!isNaN(secs)) retryDelaySec = secs;
      }
    } catch (e) {}
    const e: VisaScanError = new Error(errorMsg);
    if (resp.status === 429 || errorMsg.toLowerCase().includes('quota') || errorMsg.toLowerCase().includes('limit')) {
      e.isRateLimit = true;
      e.retryDelaySec = retryDelaySec;
    }
    throw e;
  }

  const resData = await resp.json();
  const candidate = resData.candidates?.[0];
  const finishReason = candidate?.finishReason;
  const text = candidate?.content?.parts?.[0]?.text;
  if (!text) {
    if (finishReason === 'MAX_TOKENS') throw new Error('Output AI terpotong (MAX_TOKENS) sebelum data keluar. Coba pisah PDF jadi lebih sedikit jemaah per berkas.');
    if (finishReason === 'SAFETY' || finishReason === 'RECITATION') throw new Error(`Diblokir oleh filter AI (${finishReason}).`);
    throw new Error('Data tidak dapat dibaca oleh AI (blank/safety filters).');
  }
  if (finishReason === 'MAX_TOKENS') throw new Error('Output AI terpotong (MAX_TOKENS) karena jumlah jemaah terlalu banyak dalam 1 berkas. Pisah PDF menjadi beberapa bagian (mis. 20-25 jemaah per file) lalu ulangi.');

  let cleanJson = text.trim();
  const jsonMatch = cleanJson.match(/\[[\s\S]*\]/) || cleanJson.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    console.warn('AI did not find any structured visa JSON on this page/chunk. Content:', cleanJson);
    return [];
  }
  cleanJson = jsonMatch[0];
  let parsed;
  try {
    parsed = JSON.parse(cleanJson);
  } catch (parseErr) {
    throw new Error('Respons AI bukan JSON valid (kemungkinan terpotong). Coba kurangi jumlah jemaah per berkas.');
  }
  return Array.isArray(parsed) ? parsed : [parsed];
};
