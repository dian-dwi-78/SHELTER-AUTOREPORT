import { NextRequest, NextResponse } from 'next/server';
import type { AiNarrativeResponse } from '@/types/report';

const FALLBACK_MODELS = [
  'gemini-3.5-flash-lite',
  'gemini-3.1-flash-lite',
  'gemini-flash-lite-latest',
  'gemini-3.6-flash',
  'gemini-3.5-flash',
];

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

export async function POST(req: NextRequest) {
  try {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return NextResponse.json(
        { error: 'GEMINI_API_KEY belum diatur di .env.local' },
        { status: 500 }
      );
    }

    const payload = await req.json();

    const systemInstruction = [
      'Anda adalah Spesialis Pelaporan Training Korporat untuk PT Shelter Indonesia Gemilang.',
      'Tugas Anda adalah menyusun narasi laporan giat training yang rapi, ringkas, objektif, dan berstandar industri berdasarkan teks mentah slide PPT serta ringkasan data evaluasi Excel.',
      'ATURAN WAJIB:',
      '1. Gunakan bahasa Indonesia formal, baku, dan profesional tingkat manajemen/korporat.',
      '2. DILARANG KERAS menggunakan emoji, ikon, simbol dekoratif, atau bahasa santai.',
      '3. DILARANG menulis teks dengan HURUF KAPITAL SEMUA (ALL CAPS). Gunakan Proper Case / Title Case untuk judul atau topik, dan Sentence Case untuk kalimat/uraian (kecuali singkatan baku seperti SOP, PHK, PT, WIB, KTP, K3, APD).',
      '4. Perbaiki salah ketik (typo) yang berasal dari slide PPT mentah.',
      '5. Jangan mengarang angka, nama, atau fakta di luar data yang diberikan.',
      '6. Untuk setiap submateri, tentukan secara cerdas apakah isinya lebih tepat disajikan sebagai "paragraf" atau "poin".',
      '7. Kembalikan keluaran HANYA dalam format JSON sesuai skema.',
    ].join('\n');

    const userPrompt =
      'Berikut adalah data mentah hasil ekstraksi file sumber:\n' +
      JSON.stringify(payload, null, 2) +
      '\n\nSusun JSON sesuai skema yang ditentukan.';

    const requestBody = {
      system_instruction: { parts: [{ text: systemInstruction }] },
      contents: [{ role: 'user', parts: [{ text: userPrompt }] }],
      generationConfig: {
        temperature: 0.2,
        responseMimeType: 'application/json',
        responseSchema: {
          type: 'OBJECT',
          properties: {
            namaTraining: { type: 'STRING' },
            targetPeserta: { type: 'STRING' },
            perusahaanSite: { type: 'STRING' },
            ringkasanFokus: { type: 'STRING' },
            tujuan: { type: 'ARRAY', items: { type: 'STRING' } },
            submateri: {
              type: 'ARRAY',
              items: {
                type: 'OBJECT',
                properties: {
                  topik: { type: 'STRING' },
                  format: { type: 'STRING', enum: ['paragraf', 'poin'] },
                  uraianParagraf: { type: 'STRING' },
                  uraianPoin: { type: 'ARRAY', items: { type: 'STRING' } },
                },
                required: ['topik', 'format', 'uraianParagraf', 'uraianPoin'],
              },
            },
            analisis: {
              type: 'ARRAY',
              items: {
                type: 'OBJECT',
                properties: {
                  judul: { type: 'STRING' },
                  isi: { type: 'STRING' },
                },
                required: ['judul', 'isi'],
              },
            },
            kesimpulan: { type: 'STRING' },
          },
          required: [
            'namaTraining',
            'targetPeserta',
            'perusahaanSite',
            'ringkasanFokus',
            'tujuan',
            'submateri',
            'analisis',
            'kesimpulan',
          ],
        },
      },
    };

    let lastError = '';
    let textOutput = '';

    for (const modelName of FALLBACK_MODELS) {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${encodeURIComponent(apiKey)}`;

      const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestBody),
      });

      const responseText = await res.text();

      if (res.ok) {
        const parsed = JSON.parse(responseText);
        textOutput = parsed?.candidates?.[0]?.content?.parts?.[0]?.text || '';
        if (textOutput) break;
      }

      lastError = `${modelName} (${res.status}): ${responseText}`;

      if ([503, 429, 404].includes(res.status)) {
        await sleep(1000);
        continue;
      }

      return NextResponse.json(
        { error: `Gemini API Error - ${lastError}` },
        { status: res.status }
      );
    }

    if (!textOutput) {
      return NextResponse.json(
        { error: `Semua model sedang sibuk/tidak tersedia. Error terakhir: ${lastError}` },
        { status: 503 }
      );
    }

    const result: AiNarrativeResponse = JSON.parse(textOutput);
    return NextResponse.json(result);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'Terjadi kesalahan internal.';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}