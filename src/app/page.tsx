'use client';

import React, { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { signIn, signOut, useSession } from 'next-auth/react';
import UploadModal from '@/components/UploadModal';
import HistoryModal from '@/components/HistoryModal';
import type {
  AiNarrativeResponse,
  PhotoItem,
  PptSlide,
  RawFileAttachment,
  ReportMeta,
  SendEmailPayload,
} from '@/types/report';
import {
  buildCleanReportHtml,
  buildEmailInlineHtml,
  compressImage,
  dateKey,
  fileToBase64,
  fmtDateFull,
  fmtDateMed,
  fmtDateShort,
  fmtNum,
  parseExcelDate,
  stripEmoji,
  toProperCase,
} from '@/lib/reportUtils';

type ExcelRow = Record<string, unknown>;

type EditableSubmateri = {
  topik: string;
  format: 'poin' | 'paragraf';
  uraianParagraf: string;
  uraianPoin: string[];
};

type EditableAnalisis = {
  judul: string;
  isi: string;
};

const EDIT_INPUT_CLS =
  'w-full rounded border border-dashed border-blue-400 bg-blue-50/50 px-2 py-1 text-slate-800 focus:border-blue-600 focus:bg-white focus:outline-none';

export default function TrainingReportPage() {
  const { data: session, status: sessionStatus } = useSession();
  const isLoggedIn = Boolean(session?.user?.email);
  const reportRef = useRef<HTMLDivElement>(null);

  const [meta, setMeta] = useState<ReportMeta>({
    entitas: 'PT Shelter Indonesia Gemilang',
    site: '',
    namaTraining: '',
    targetPeserta: '',
    metode: 'Offline',
    pemateri: '',
    waktuSesi: '',
    ttd: 'Tim Training',
    tujuan: '',
  });

  const [absensiRows, setAbsensiRows] = useState<ExcelRow[]>([]);
  const [testRows, setTestRows] = useState<ExcelRow[]>([]);
  const [pptSlides, setPptSlides] = useState<PptSlide[]>([]);
  const [photos, setPhotos] = useState<PhotoItem[]>([]);
  const [rawExcelFiles, setRawExcelFiles] = useState<RawFileAttachment[]>([]);
  const [rawPptFile, setRawPptFile] = useState<RawFileAttachment | null>(null);
  const [aiData, setAiData] = useState<AiNarrativeResponse | null>(null);

  // ============================================================================
  // INLINE EDIT MODE & OVERRIDE STATES
  // ============================================================================
  const [isEditMode, setIsEditMode] = useState(false);
  const [customGreeting, setCustomGreeting] = useState<string | null>(null);
  const [customIntro, setCustomIntro] = useState<string | null>(null);
  const [customSummary, setCustomSummary] = useState<string | null>(null);
  const [customGoals, setCustomGoals] = useState<string[] | null>(null);
  const [customSubmateri, setCustomSubmateri] = useState<EditableSubmateri[] | null>(null);
  const [customAnalisis, setCustomAnalisis] = useState<EditableAnalisis[] | null>(null);
  const [customKesimpulan, setCustomKesimpulan] = useState<string | null>(null);
  const [customClosing, setCustomClosing] = useState<string | null>(null);
  const [customAttendanceNames, setCustomAttendanceNames] = useState<Record<string, string>>({});

  // Process & Status indicators
  const [status, setStatus] = useState<{ msg: string; type: '' | 'ok' | 'err' }>({
    msg: 'Siap. Klik "Upload File" untuk mulai menyusun laporan.',
    type: '',
  });
  const [isReadingFile, setIsReadingFile] = useState(false);
  const [isAiLoading, setIsAiLoading] = useState(false);
  const [isSendingEmail, setIsSendingEmail] = useState(false);

  // Modal & Dropdown states
  const [uploadModalOpen, setUploadModalOpen] = useState(false);
  const [uploadModalError, setUploadModalError] = useState<string | null>(null);
  const [historyModalOpen, setHistoryModalOpen] = useState(false);
  const [downloadMenuOpen, setDownloadMenuOpen] = useState(false);
  const [metaModalOpen, setMetaModalOpen] = useState(false);
  const [draftMeta, setDraftMeta] = useState<ReportMeta>(meta);

  const [emailModalOpen, setEmailModalOpen] = useState(false);
  const [emailForm, setEmailForm] = useState({
    to: '',
    cc: '',
    bcc: '',
    senderName: 'PT Shelter Indonesia Gemilang',
    replyTo: '',
    subject: '',
    body: '',
    attachHtml: false,
    attachExcel: false,
    attachPpt: false,
  });

  // ============================================================================
  // 1. PARSE EXCEL, PPTX, FOTO
  // ============================================================================
  async function handleExcelUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    setIsReadingFile(true);
    setUploadModalError(null);
    setStatus({ msg: `Membaca ${files.length} file Excel...`, type: '' });

    try {
      const newAbsensi: ExcelRow[] = [];
      const newTests: ExcelRow[] = [];
      const newRaw: RawFileAttachment[] = [];

      for (const file of files) {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });

        wb.SheetNames.forEach((sheetName) => {
          const rows = XLSX.utils.sheet_to_json<ExcelRow>(wb.Sheets[sheetName], { defval: null });
          if (!rows.length) return;
          const cols = Object.keys(rows[0]).map((c) => c.trim().toLowerCase());

          if (cols.includes('jenis tes') && cols.some((c) => c.includes('skor'))) {
            newTests.push(...rows);
          } else if (cols.includes('judul meeting') || cols.includes('nama') || cols.includes('penempatan')) {
            newAbsensi.push(...rows);
          }
        });

        const base64 = await fileToBase64(file);
        newRaw.push({
          name: file.name,
          mimeType: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          base64,
        });
      }

      setAbsensiRows((prev) => [...prev, ...newAbsensi]);
      setTestRows((prev) => [...prev, ...newTests]);
      setRawExcelFiles((prev) => [...prev, ...newRaw]);
      setStatus({ msg: `Berhasil memuat ${files.length} file Excel.`, type: 'ok' });
    } catch (err: unknown) {
      const msg = `Gagal membaca Excel: ${err instanceof Error ? err.message : err}`;
      setStatus({ msg, type: 'err' });
      setUploadModalError(msg);
    } finally {
      setIsReadingFile(false);
      e.target.value = '';
    }
  }

  async function handlePptUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsReadingFile(true);
    setUploadModalError(null);
    setStatus({ msg: `Mengekstrak materi dari ${file.name}...`, type: '' });

    try {
      const zip = await JSZip.loadAsync(file);
      const slideFiles = Object.keys(zip.files)
        .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort((a, b) => {
          const na = parseInt(a.match(/slide(\d+)\.xml/)?.[1] || '0', 10);
          const nb = parseInt(b.match(/slide(\d+)\.xml/)?.[1] || '0', 10);
          return na - nb;
        });

      const parser = new DOMParser();
      const slides: PptSlide[] = [];

      for (const sf of slideFiles) {
        const xmlText = await zip.files[sf].async('string');
        const doc = parser.parseFromString(xmlText, 'application/xml');
        const paragraphs = Array.from(doc.getElementsByTagName('a:p'));
        const lines = paragraphs
          .map((p) =>
            Array.from(p.getElementsByTagName('a:t'))
              .map((t) => t.textContent)
              .join('')
              .trim()
          )
          .filter(Boolean);

        if (lines.length) {
          slides.push({ title: lines[0], body: lines.slice(1) });
        }
      }

      setPptSlides(slides);
      setAiData(null);
      if (slides.length) {
        setMeta((prev) => ({
          ...prev,
          namaTraining: prev.namaTraining || toProperCase(slides[0].title),
          targetPeserta: prev.targetPeserta || (slides[0].body.length ? toProperCase(slides[0].body.join(' · ')) : ''),
        }));
      }

      const base64 = await fileToBase64(file);
      setRawPptFile({
        name: file.name,
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        base64,
      });

      setStatus({
        msg: `Materi PPT berhasil diekstrak (${slides.length} slide). Klik "Simpan" untuk merapikan dengan AI.`,
        type: 'ok',
      });
    } catch (err: unknown) {
      const msg = `Gagal membaca PPTX: ${err instanceof Error ? err.message : err}`;
      setStatus({ msg, type: 'err' });
      setUploadModalError(msg);
    } finally {
      setIsReadingFile(false);
      e.target.value = '';
    }
  }

  async function handlePhotoUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    setIsReadingFile(true);
    setUploadModalError(null);
    setStatus({ msg: `Mengompresi ${files.length} foto dokumentasi...`, type: '' });

    try {
      const added: PhotoItem[] = [];
      for (let i = 0; i < files.length; i++) {
        const src = await compressImage(files[i], 1000, 0.75);
        added.push({ src, caption: `Dokumentasi ${photos.length + i + 1}` });
      }
      setPhotos((prev) => [...prev, ...added]);
      setStatus({ msg: `Berhasil menambahkan ${files.length} foto (terkompresi).`, type: 'ok' });
    } catch (err: unknown) {
      const msg = `Gagal memproses foto: ${err instanceof Error ? err.message : err}`;
      setStatus({ msg, type: 'err' });
      setUploadModalError(msg);
    } finally {
      setIsReadingFile(false);
      e.target.value = '';
    }
  }

  // ============================================================================
  // 2. KALKULASI DATA LAPORAN (CLIENT-SIDE PRESISI)
  // ============================================================================
  const computed = useMemo(() => {
    const judul =
      meta.namaTraining ||
      (absensiRows[0] && toProperCase(absensiRows[0]['Judul Meeting'])) ||
      (pptSlides[0] && toProperCase(pptSlides[0].title)) ||
      '-';

    const sites = Array.from(
      new Set(
        absensiRows
          .map((r) => r['Penempatan'])
          .filter(Boolean)
          .map((s) => String(s).trim())
      )
    );
    const site = meta.site || (sites.length ? sites.slice(0, 3).join(', ') : '-');
    const entitas = meta.entitas || 'PT Shelter Indonesia Gemilang';

    const byNik = new Map<
      string,
      { nik: string; nama: string; pre: number | null; post: number | null; preNull: boolean; date: string }
    >();
    const byDateScore: Record<string, { pre: number[]; post: number[] }> = {};
    let jumlahSoal = 20;

    testRows.forEach((r) => {
      const nik = String(r['NIK'] || r['Nama Karyawan'] || '').trim();
      if (!nik) return;
      const jenis = String(r['Jenis Tes'] || '').toLowerCase();
      const skor = Number(r['Skor (0-100)']);
      const dk = dateKey(parseExcelDate(r['Waktu']));

      if (r['Benar'] && String(r['Benar']).includes('/')) {
        const q = parseInt(String(r['Benar']).split('/')[1], 10);
        if (!isNaN(q)) jumlahSoal = q;
      }

      let hasNull = false;
      if (r['Detail Jawaban']) {
        try {
          const arr = JSON.parse(String(r['Detail Jawaban']));
          if (Array.isArray(arr) && arr.some((x) => x === null)) hasNull = true;
        } catch {}
      }

      if (!byNik.has(nik)) {
        byNik.set(nik, { nik, nama: String(r['Nama Karyawan'] || '').trim(), pre: null, post: null, preNull: false, date: dk });
      }
      const item = byNik.get(nik)!;
      if (jenis.includes('pre')) {
        item.pre = isNaN(skor) ? null : skor;
        item.preNull = hasNull;
        if (dk) item.date = dk;
      } else if (jenis.includes('post')) {
        item.post = isNaN(skor) ? null : skor;
        if (!item.date && dk) item.date = dk;
      }

      if (dk && !isNaN(skor)) {
        if (!byDateScore[dk]) byDateScore[dk] = { pre: [], post: [] };
        if (jenis.includes('pre')) byDateScore[dk].pre.push(skor);
        else if (jenis.includes('post')) byDateScore[dk].post.push(skor);
      }
    });

    if (!testRows.length && absensiRows.length) {
      absensiRows.forEach((r, idx) => {
        const pv = r['Nilai Post-test'];
        if (pv !== null && pv !== '' && !isNaN(Number(pv))) {
          const dk = dateKey(parseExcelDate(r['Tanggal Meeting'] || r['Timestamp']));
          byNik.set(`abs-${idx}`, {
            nik: String(r['NIK KTP'] || idx),
            nama: String(r['Nama'] || ''),
            pre: null,
            post: Number(pv),
            preNull: false,
            date: dk,
          });
          if (dk) {
            if (!byDateScore[dk]) byDateScore[dk] = { pre: [], post: [] };
            byDateScore[dk].post.push(Number(pv));
          }
        }
      });
    }

    const evalList = Array.from(byNik.values());
    const totalEval = evalList.length;
    const preScores = evalList.map((x) => x.pre).filter((x): x is number => x !== null);
    const postScores = evalList.map((x) => x.post).filter((x): x is number => x !== null);

    const avgPre = preScores.length ? preScores.reduce((a, b) => a + b, 0) / preScores.length : null;
    const avgPost = postScores.length ? postScores.reduce((a, b) => a + b, 0) / postScores.length : null;
    const delta = avgPre !== null && avgPost !== null ? avgPost - avgPre : null;

    let naik = 0, tetap = 0, turun = 0, postMin80 = 0, post100 = 0, postBawah70 = 0, preTidakLengkap = 0;
    evalList.forEach((p) => {
      if (p.pre !== null && p.post !== null) {
        if (p.post > p.pre) naik++;
        else if (p.post < p.pre) turun++;
        else tetap++;
      }
      if (p.post !== null) {
        if (p.post >= 80) postMin80++;
        if (p.post === 100) post100++;
        if (p.post < 70) postBawah70++;
      }
      if (p.pre !== null && p.preNull) preTidakLengkap++;
    });

    const sesiMap: Record<string, string[]> = {};
    if (absensiRows.length) {
      absensiRows.forEach((r) => {
        const dk = dateKey(parseExcelDate(r['Tanggal Meeting'] || r['Timestamp'])) || 'Tanpa Tanggal';
        if (!sesiMap[dk]) sesiMap[dk] = [];
        sesiMap[dk].push(String(r['Nama'] || r['NIK KTP'] || '-').trim());
      });
    } else if (evalList.length) {
      evalList.forEach((p) => {
        const dk = p.date || 'Tanpa Tanggal';
        if (!sesiMap[dk]) sesiMap[dk] = [];
        sesiMap[dk].push(p.nama || p.nik);
      });
    }

    const sesiKeys = Object.keys(sesiMap).sort();
    const validDates = sesiKeys.filter((k) => k !== 'Tanpa Tanggal');
    const tanggalRingkas = validDates.length ? validDates.map(fmtDateMed).join(' & ') : '-';
    const hariTanggal = validDates.length ? validDates.map(fmtDateFull).join(' dan ') : '-';

    const contentSlides = pptSlides.slice(1).filter((s) => !/^terima kasih$/i.test(stripEmoji(s.title)));

    return {
      judul, site, entitas, totalEval, preScores, postScores, avgPre, avgPost, delta,
      naik, tetap, turun, postMin80, post100, postBawah70, preTidakLengkap,
      jumlahSoal, byDateScore, sesiMap, sesiKeys, tanggalRingkas, hariTanggal, contentSlides,
    };
  }, [meta, absensiRows, testRows, pptSlides]);

  // ============================================================================
  // 2B. NILAI EFEKTIF (AUTO-SYNC DEFAULT VS CUSTOM EDIT USER)
  // ============================================================================
  const defaultGreeting = `Yth. Bapak/Ibu Manajemen${computed.site !== '-' ? ` ${computed.site}` : ''},`;
  const effectiveGreeting = customGreeting !== null ? customGreeting : defaultGreeting;

  const defaultIntroPlain = `Bersama laporan ini kami sampaikan dokumentasi pelaksanaan ${computed.judul} pada tanggal ${computed.tanggalRingkas}, beserta ringkasan materi, evaluasi pemahaman peserta, serta daftar hadir per sesi.`;
  const effectiveIntro = customIntro !== null ? customIntro : defaultIntroPlain;

  const defaultSummary = aiData?.ringkasanFokus
    ? stripEmoji(aiData.ringkasanFokus)
    : meta.tujuan
    ? stripEmoji(meta.tujuan)
    : computed.judul !== '-'
    ? `Program berfokus pada materi ${computed.judul} dalam pekerjaan operasional beserta penguatan standar kerja personel.`
    : 'Belum ada ringkasan materi. Unggah file PPT materi melalui tombol Upload File atau isi melalui Edit Informasi.';
  const effectiveSummary = customSummary !== null ? customSummary : defaultSummary;

  const defaultGoals: string[] = useMemo(() => {
    if (aiData?.tujuan?.length) return aiData.tujuan.map((t) => stripEmoji(t));
    if (computed.contentSlides.length > 0) {
      return computed.contentSlides.slice(0, 4).map((s) => `Memahami dan menerapkan materi ${toProperCase(s.title)}.`);
    }
    return [];
  }, [aiData, computed.contentSlides]);
  const effectiveGoals = customGoals !== null ? customGoals : defaultGoals;

  const defaultSubmateri: EditableSubmateri[] = useMemo(() => {
    if (aiData?.submateri?.length) {
      return aiData.submateri.map((m) => ({
        topik: toProperCase(m.topik),
        format: m.format === 'poin' ? 'poin' : 'paragraf',
        uraianParagraf: stripEmoji(m.uraianParagraf || ''),
        uraianPoin: (m.uraianPoin || []).map((p) => stripEmoji(p)),
      }));
    }
    if (computed.contentSlides.length > 0) {
      return computed.contentSlides.map((s) => {
        const cleanBody = s.body.map(stripEmoji).filter(Boolean).slice(0, 5).map(toProperCase);
        return {
          topik: toProperCase(s.title),
          format: cleanBody.length > 0 ? 'poin' : 'paragraf',
          uraianParagraf: cleanBody.length > 0 ? '' : '-',
          uraianPoin: cleanBody,
        };
      });
    }
    return [];
  }, [aiData, computed.contentSlides]);
  const effectiveSubmateri = customSubmateri !== null ? customSubmateri : defaultSubmateri;

  const defaultAnalisis: EditableAnalisis[] = useMemo(() => {
    if (aiData?.analisis?.length) {
      return aiData.analisis.map((a) => ({
        judul: stripEmoji(a.judul),
        isi: stripEmoji(a.isi),
      }));
    }
    if (computed.totalEval > 0) {
      return [
        {
          judul: 'Hasil akhir evaluasi peserta.',
          isi: `Sebanyak ${computed.postMin80} dari ${computed.totalEval} peserta evaluasi (${Math.round(
            (computed.postMin80 / computed.totalEval) * 100
          )}%) memperoleh nilai post-test minimal 80 dan ${computed.post100} peserta mencapai nilai 100.`,
        },
        {
          judul: 'Tindak lanjut.',
          isi: `Direkomendasikan pembinaan kepada ${computed.postBawah70} peserta dengan skor post-test di bawah 70${
            computed.preScores.length ? ` serta peninjauan untuk ${computed.turun} peserta yang mengalami penurunan nilai` : ''
          }.`,
        },
      ];
    }
    return [];
  }, [aiData, computed]);
  const effectiveAnalisis = customAnalisis !== null ? customAnalisis : defaultAnalisis;

  const defaultKesimpulan = aiData?.kesimpulan
    ? stripEmoji(aiData.kesimpulan)
    : `Training ${computed.judul} telah terdokumentasi dalam ${computed.sesiKeys.length || 1} sesi pada ${computed.tanggalRingkas}.${
        computed.totalEval
          ? ` Hasil evaluasi menunjukkan ${computed.postMin80} dari ${computed.totalEval} peserta mencapai nilai akhir minimal 80.`
          : ''
      }`;
  const effectiveKesimpulan = customKesimpulan !== null ? customKesimpulan : defaultKesimpulan;

  const defaultClosing =
    'Demikian laporan ini disampaikan sebagai dokumentasi kegiatan training dan bahan tindak lanjut pembinaan personel.';
  const effectiveClosing = customClosing !== null ? customClosing : defaultClosing;

  // Helper update list edit
  function updateGoalAt(idx: number, val: string) {
    const next = [...effectiveGoals];
    next[idx] = val;
    setCustomGoals(next);
  }
  function removeGoalAt(idx: number) {
    setCustomGoals(effectiveGoals.filter((_, i) => i !== idx));
  }
  function addGoal() {
    setCustomGoals([...effectiveGoals, '']);
  }

  function updateSubmateriAt(idx: number, patch: Partial<EditableSubmateri>) {
    const next = effectiveSubmateri.map((item, i) => (i === idx ? { ...item, ...patch } : item));
    setCustomSubmateri(next);
  }
  function removeSubmateriAt(idx: number) {
    setCustomSubmateri(effectiveSubmateri.filter((_, i) => i !== idx));
  }
  function addSubmateri() {
    setCustomSubmateri([
      ...effectiveSubmateri,
      { topik: 'Topik Baru', format: 'paragraf', uraianParagraf: '', uraianPoin: [] },
    ]);
  }

  function updateAnalisisAt(idx: number, patch: Partial<EditableAnalisis>) {
    const next = effectiveAnalisis.map((item, i) => (i === idx ? { ...item, ...patch } : item));
    setCustomAnalisis(next);
  }
  function removeAnalisisAt(idx: number) {
    setCustomAnalisis(effectiveAnalisis.filter((_, i) => i !== idx));
  }
  function addAnalisis() {
    setCustomAnalisis([...effectiveAnalisis, { judul: 'Catatan Tambahan.', isi: '' }]);
  }

  function getAttendanceName(dk: string, idx: number, fallback: string) {
    const key = `${dk}::${idx}`;
    return key in customAttendanceNames ? customAttendanceNames[key] : fallback;
  }
  function setAttendanceName(dk: string, idx: number, val: string) {
    const key = `${dk}::${idx}`;
    setCustomAttendanceNames((prev) => ({ ...prev, [key]: val }));
  }

  // ============================================================================
  // 3. CALL API GEMINI (DIPICU TOMBOL SIMPAN DI MODAL UPLOAD) & EMAIL
  // ============================================================================
  async function handleSaveUploadAndRefine() {
    setUploadModalError(null);

    if (!pptSlides.length && !absensiRows.length && !testRows.length) {
      if (photos.length > 0) {
        setUploadModalOpen(false);
        setStatus({ msg: `${photos.length} foto dokumentasi tersimpan ke laporan.`, type: 'ok' });
        return;
      }
      setUploadModalError('Unggah minimal satu file sebelum menekan Simpan.');
      return;
    }

    setIsAiLoading(true);
    setStatus({ msg: 'Menghubungi Gemini AI untuk merapikan materi & analisis...', type: '' });

    try {
      const res = await fetch('/api/ai-refine', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          metaEksisting: {
            namaTraining: computed.judul,
            perusahaanSite: computed.site,
            tanggal: computed.hariTanggal,
            jumlahSesi: computed.sesiKeys.length ? `${computed.sesiKeys.length} sesi` : '-',
          },
          ringkasanEvaluasi: {
            totalPesertaEvaluasi: computed.totalEval,
            rataRataPreTest: fmtNum(computed.avgPre),
            rataRataPostTest: fmtNum(computed.avgPost),
            kenaikanPoin: computed.delta !== null ? fmtNum(computed.delta) : '-',
            meningkat: computed.naik,
            tetap: computed.tetap,
            menurun: computed.turun,
            postMin80: computed.postMin80,
            postBawah70: computed.postBawah70,
          },
          slidePptMentah: pptSlides,
        }),
      });

      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Gagal memproses AI');

      setAiData(data);
      // Reset override narasi agar hasil AI terbaru langsung tampil
      setCustomSummary(null);
      setCustomGoals(null);
      setCustomSubmateri(null);
      setCustomAnalisis(null);
      setCustomKesimpulan(null);

      setMeta((prev) => ({
        ...prev,
        namaTraining: data.namaTraining ? toProperCase(data.namaTraining) : prev.namaTraining,
        targetPeserta: data.targetPeserta && data.targetPeserta !== '-' ? toProperCase(data.targetPeserta) : prev.targetPeserta,
        site: data.perusahaanSite && data.perusahaanSite !== '-' ? stripEmoji(data.perusahaanSite) : prev.site,
        tujuan: data.ringkasanFokus ? stripEmoji(data.ringkasanFokus) : prev.tujuan,
      }));

      setStatus({ msg: 'Laporan berhasil dirapikan dengan Gemini AI.', type: 'ok' });
      setUploadModalOpen(false);
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : 'Error tidak diketahui';
      setStatus({ msg: `Gagal memproses AI: ${message}`, type: 'err' });
      setUploadModalError(`Gagal memproses AI: ${message}`);
    } finally {
      setIsAiLoading(false);
    }
  }

  function handleDownloadHtml() {
    if (!reportRef.current || isEditMode) return;
    const html = buildCleanReportHtml(reportRef.current, computed.judul);
    const nama = (computed.judul || 'Training').replace(/[^a-zA-Z0-9]+/g, '_');
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Laporan_Training_${nama}.html`;
    a.click();
    URL.revokeObjectURL(a.href);
    setDownloadMenuOpen(false);
    setStatus({ msg: `File Laporan_Training_${nama}.html berhasil diunduh.`, type: 'ok' });
  }

  function openEmailDialog() {
    if (isEditMode) return;
    setEmailForm((prev) => ({
      ...prev,
      senderName: computed.entitas,
      subject: `Laporan Giat Training ${computed.judul} – ${computed.site}`,
      body: `${effectiveGreeting}\n\nBersama email ini kami sampaikan Laporan Giat Training "${computed.judul}" yang dilaksanakan pada ${computed.hariTanggal}.\n\nHormat kami,\n${meta.ttd || 'Tim Training'}\n${computed.entitas}`,
    }));
    setEmailModalOpen(true);
  }

  async function handleSendEmail() {
    if (!session?.user?.email) {
      alert('Silakan login dengan akun Google terlebih dahulu sebelum mengirim email.');
      return;
    }
    if (!emailForm.to.trim() || !emailForm.subject.trim()) {
      alert('Mohon lengkapi alamat email tujuan (To) dan Subjek.');
      return;
    }
    if (!reportRef.current) return;

    setIsSendingEmail(true);
    setStatus({ msg: `Mengirim email laporan ke ${emailForm.to}...`, type: '' });

    try {
      const extraAttachments: RawFileAttachment[] = [];
      if (emailForm.attachExcel) extraAttachments.push(...rawExcelFiles);
      if (emailForm.attachPpt && rawPptFile) extraAttachments.push(rawPptFile);

      const cleanNama = (computed.judul || 'Training').replace(/[^a-zA-Z0-9]+/g, '_');
      const fullCleanHtml = buildCleanReportHtml(reportRef.current, computed.judul);

      const payload: SendEmailPayload & { trainingName: string; siteName: string } = {
        to: emailForm.to,
        cc: emailForm.cc,
        bcc: emailForm.bcc,
        senderName: emailForm.senderName,
        replyTo: emailForm.replyTo,
        subject: emailForm.subject,
        body: emailForm.body,
        emailHtmlBody: buildEmailInlineHtml(reportRef.current, photos),
        inlinePhotos: photos.map((p, idx) => ({
          cid: `foto_${idx}`,
          base64: String(p.src || '').split(',')[1] || '',
        })),
        attachHtml: emailForm.attachHtml,
        htmlFileName: `Laporan_Training_${cleanNama}.html`,
        htmlContent: fullCleanHtml,
        extraAttachments,
        trainingName: computed.judul,
        siteName: computed.site,
      };

      const res = await fetch('/api/send-email', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Gagal mengirim email');

      setEmailModalOpen(false);
      setStatus({ msg: data.message, type: 'ok' });
      alert(data.message);
    } catch (err: unknown) {
      const errMsg = err instanceof Error ? err.message : String(err);
      setStatus({ msg: `Gagal mengirim email: ${errMsg}`, type: 'err' });
      alert(`Gagal mengirim email: ${errMsg}`);
    } finally {
      setIsSendingEmail(false);
    }
  }

  const isAnyProcessRunning = isReadingFile || isAiLoading || isSendingEmail;

  // ============================================================================
  // 4. RENDER UI
  // ============================================================================
  return (
    <div className="min-h-screen bg-slate-100 text-slate-800">
      {/* HEADER & TOOLBAR (TIDAK IKUT TERCETAK) */}
      <header className="sticky top-0 z-30 bg-[#001247] text-white shadow-md print:hidden">
        <div className="mx-auto max-w-5xl px-4 py-3.5">
          {/* Baris Atas: Judul + Akun / Login */}
          <div className="flex flex-wrap items-center justify-between gap-3">
            <span className="text-base font-extrabold tracking-wide text-white">
              GIAT TRAINING REPORT GENERATOR
            </span>

            {sessionStatus === 'loading' ? (
              <span className="rounded-lg bg-white/15 px-4 py-1.5 text-xs font-semibold text-blue-100">
                Memeriksa sesi...
              </span>
            ) : isLoggedIn ? (
              <div className="flex items-center gap-2">
                <span className="rounded-lg bg-white px-3.5 py-1.5 text-xs font-bold text-[#001247] shadow-sm">
                  {session?.user?.email}
                </span>
                <button
                  type="button"
                  onClick={() => signOut()}
                  className="rounded-lg border border-white/25 bg-white/10 px-3 py-1.5 text-xs font-semibold text-white transition-colors hover:border-rose-600 hover:bg-rose-600"
                  title="Keluar dari akun Google"
                >
                  Logout
                </button>
              </div>
            ) : (
              <button
                type="button"
                onClick={() => signIn('google')}
                className="rounded-lg bg-white px-5 py-2 text-xs font-extrabold tracking-wide text-[#001247] shadow-sm transition-colors hover:bg-blue-50"
              >
                LOGIN GOOGLE
              </button>
            )}
          </div>

          {/* Baris Bawah: Tombol Aksi (HANYA TAMPIL SETELAH LOGIN) */}
          {isLoggedIn && (
            <div className="mt-3 flex flex-wrap items-center gap-2 border-t border-white/10 pt-3">
              <button
                type="button"
                onClick={() => {
                  setUploadModalError(null);
                  setUploadModalOpen(true);
                }}
                className="rounded-lg bg-sky-600 px-3.5 py-1.5 text-xs font-bold text-white shadow-sm transition-colors hover:bg-sky-500"
              >
                Upload File
              </button>

              <button
                type="button"
                onClick={() => {
                  setDraftMeta({
                    ...meta,
                    site: computed.site === '-' ? '' : computed.site,
                    namaTraining: computed.judul === '-' ? '' : computed.judul,
                  });
                  setMetaModalOpen(true);
                }}
                className="rounded-lg bg-sky-600 px-3.5 py-1.5 text-xs font-bold text-white shadow-sm transition-colors hover:bg-sky-500"
              >
                Edit Informasi
              </button>

              {/* TOMBOL MODE EDIT LANGSUNG */}
              <button
                type="button"
                onClick={() => {
                  setDownloadMenuOpen(false);
                  setIsEditMode((prev) => {
                    const next = !prev;
                    setStatus({
                      msg: next
                        ? 'Mode Edit aktif. Silakan koreksi teks di kotak bertanda garis putus-putus, lalu klik "Selesai Edit".'
                        : 'Perubahan teks tersimpan. Laporan siap diunduh atau dikirim.',
                      type: next ? '' : 'ok',
                    });
                    return next;
                  });
                }}
                className={`rounded-lg px-3.5 py-1.5 text-xs font-bold text-white shadow-sm transition-colors ${
                  isEditMode
                    ? 'bg-amber-500 ring-2 ring-amber-300 hover:bg-amber-400'
                    : 'bg-indigo-600 hover:bg-indigo-500'
                }`}
              >
                {isEditMode ? '✓ Selesai Edit' : 'Mode Edit'}
              </button>

              {/* Dropdown Unduh PDF / HTML */}
              <div className="relative">
                <button
                  type="button"
                  disabled={isEditMode}
                  title={isEditMode ? 'Klik Selesai Edit terlebih dahulu' : ''}
                  onClick={() => setDownloadMenuOpen((prev) => !prev)}
                  className="inline-flex items-center gap-1.5 rounded-lg bg-sky-600 px-3.5 py-1.5 text-xs font-bold text-white shadow-sm transition-colors hover:bg-sky-500 disabled:cursor-not-allowed disabled:opacity-45"
                >
                  Unduh PDF/HTML
                  <span className="text-[10px]">▾</span>
                </button>
                {downloadMenuOpen && !isEditMode && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setDownloadMenuOpen(false)} />
                    <div className="absolute left-0 z-20 mt-1.5 w-44 overflow-hidden rounded-lg border border-slate-200 bg-white py-1 text-xs text-slate-700 shadow-lg">
                      <button
                        type="button"
                        onClick={() => {
                          setDownloadMenuOpen(false);
                          window.print();
                        }}
                        className="block w-full px-3.5 py-2 text-left font-semibold hover:bg-slate-100"
                      >
                        Cetak / Simpan PDF
                      </button>
                      <button
                        type="button"
                        onClick={handleDownloadHtml}
                        className="block w-full px-3.5 py-2 text-left font-semibold hover:bg-slate-100"
                      >
                        Unduh File HTML (.html)
                      </button>
                    </div>
                  </>
                )}
              </div>

              <button
                type="button"
                disabled={isEditMode}
                title={isEditMode ? 'Klik Selesai Edit terlebih dahulu' : ''}
                onClick={openEmailDialog}
                className="rounded-lg bg-emerald-600 px-3.5 py-1.5 text-xs font-bold text-white shadow-sm transition-colors hover:bg-emerald-500 disabled:cursor-not-allowed disabled:opacity-45"
              >
                Kirim Email
              </button>

              <button
                type="button"
                onClick={() => setHistoryModalOpen(true)}
                className="ml-auto rounded-lg bg-sky-600 px-4 py-1.5 text-xs font-bold text-white shadow-sm transition-colors hover:bg-sky-500"
              >
                Riwayat
              </button>
            </div>
          )}
        </div>

        {/* STATUS BAR DETAIL PROSES */}
        <div
          className={`border-t border-white/10 px-4 py-1.5 ${
            isEditMode ? 'bg-amber-600 text-white' : 'bg-[#000b2b]'
          }`}
        >
          <div className="mx-auto flex max-w-5xl items-center justify-between gap-2 text-[11px] font-medium">
            <div className="flex items-center gap-2">
              {!isLoggedIn ? (
                <span className="text-amber-300">
                  ● Silakan klik LOGIN GOOGLE di kanan atas untuk mulai menggunakan aplikasi.
                </span>
              ) : isEditMode ? (
                <span className="font-semibold text-white">
                  ✎ MODE EDIT AKTIF — Edit langsung teks pada kotak bergaris putus-putus. Klik "✓ Selesai Edit" sebelum mengunduh atau mengirim email.
                </span>
              ) : (
                <>
                  {isAnyProcessRunning ? (
                    <span className="inline-block h-3 w-3 shrink-0 animate-spin rounded-full border-2 border-sky-400 border-t-transparent" />
                  ) : (
                    <span
                      className={`inline-block h-2 w-2 shrink-0 rounded-full ${
                        status.type === 'ok'
                          ? 'bg-emerald-400'
                          : status.type === 'err'
                          ? 'bg-rose-400'
                          : 'bg-sky-400'
                      }`}
                    />
                  )}
                  <span
                    className={
                      status.type === 'ok'
                        ? 'text-emerald-300'
                        : status.type === 'err'
                        ? 'text-rose-300'
                        : 'text-slate-200'
                    }
                  >
                    {status.msg}
                  </span>
                </>
              )}
            </div>
          </div>
        </div>
      </header>

      {/* CONTAINER LAPORAN UTAMA */}
      <div className="wrapper p-4 md:p-8" ref={reportRef}>
        <main className="container mx-auto max-w-4xl overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
          <div className="topbar h-2 bg-gradient-to-r from-blue-900 via-blue-700 to-blue-500" />
          <div className="inner p-6 md:p-10">
            <p className="eyebrow text-xs font-bold uppercase tracking-widest text-blue-700">{computed.entitas}</p>
            <h1 className="mt-1 text-2xl font-extrabold text-slate-900">
              LAPORAN GIAT TRAINING
              <br />
              {computed.judul.toUpperCase()}
            </h1>
            <p className="site mt-1 text-sm font-semibold text-slate-500">{computed.site}</p>

            {/* SAPAAN & PARAGRAF PEMBUKA */}
            {isEditMode ? (
              <div className="mt-5 space-y-2 rounded-lg border border-blue-200 bg-blue-50/30 p-3">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] font-bold uppercase tracking-wide text-blue-700">
                    Bagian Pembuka
                  </span>
                  {(customGreeting !== null || customIntro !== null) && (
                    <button
                      type="button"
                      onClick={() => {
                        setCustomGreeting(null);
                        setCustomIntro(null);
                      }}
                      className="rounded bg-slate-200 px-2 py-0.5 text-[11px] font-semibold text-slate-700 hover:bg-slate-300"
                    >
                      Reset ke Default
                    </button>
                  )}
                </div>
                <input
                  type="text"
                  value={effectiveGreeting}
                  onChange={(e) => setCustomGreeting(e.target.value)}
                  className={`${EDIT_INPUT_CLS} text-sm font-bold`}
                  placeholder="Yth. Bapak/Ibu Manajemen..."
                />
                <textarea
                  rows={3}
                  value={effectiveIntro}
                  onChange={(e) => setCustomIntro(e.target.value)}
                  className={`${EDIT_INPUT_CLS} text-sm leading-relaxed`}
                  placeholder="Paragraf pengantar laporan..."
                />
              </div>
            ) : (
              <>
                <p className="greeting mt-5 text-sm font-bold">{effectiveGreeting}</p>
                {customIntro !== null ? (
                  <p className="mt-2 whitespace-pre-line text-sm text-slate-700">{customIntro}</p>
                ) : (
                  <p className="mt-2 text-sm text-slate-700">
                    Bersama laporan ini kami sampaikan dokumentasi pelaksanaan <b>{computed.judul}</b> pada tanggal{' '}
                    <b>{computed.tanggalRingkas}</b>, beserta ringkasan materi, evaluasi pemahaman peserta, serta daftar hadir per sesi.
                  </p>
                )}
              </>
            )}

            {/* METRICS */}
            <table className="metrics -mx-2.5 mt-6 w-full border-separate border-spacing-2.5" role="presentation">
              <tbody>
                <tr>
                  <td className="metric blue w-1/3 rounded-lg border border-slate-200 bg-slate-50 p-3.5">
                    <div className="label text-[11px] font-bold uppercase text-slate-500">Peserta Evaluasi</div>
                    <span className="num text-2xl font-extrabold text-blue-700">{computed.totalEval || '-'}</span>{' '}
                    <span className="unit text-xs font-semibold text-slate-500">orang</span>
                  </td>
                  <td className="metric green w-1/3 rounded-lg border border-slate-200 bg-slate-50 p-3.5">
                    <div className="label text-[11px] font-bold uppercase text-slate-500">Rata-rata Post-Test</div>
                    <span className="num text-2xl font-extrabold text-emerald-600">{fmtNum(computed.avgPost)}</span>{' '}
                    <span className="unit text-xs font-semibold text-slate-500">/ 100</span>
                  </td>
                  <td className="metric orange w-1/3 rounded-lg border border-slate-200 bg-slate-50 p-3.5">
                    <div className="label text-[11px] font-bold uppercase text-slate-500">Kenaikan Pre ke Post-Test</div>
                    <span className="num text-2xl font-extrabold text-amber-600">
                      {computed.delta !== null ? `${computed.delta >= 0 ? '+' : ''}${fmtNum(computed.delta)}` : '-'}
                    </span>{' '}
                    <span className="unit text-xs font-semibold text-slate-500">poin</span>
                  </td>
                </tr>
              </tbody>
            </table>

            {/* 1. INFORMASI PELAKSANAAN */}
            <h2 className="mb-3 mt-7 border-b-2 border-slate-200 pb-1.5 text-base font-bold text-blue-900">
              1. Informasi Pelaksanaan
            </h2>
            <table className="info w-full border-collapse text-sm">
              <tbody>
                {[
                  ['Perusahaan / Site', computed.site],
                  ['Entitas Pengelola', computed.entitas],
                  ['Hari & Tanggal', computed.hariTanggal],
                  ['Nama Training', computed.judul],
                  ['Target Peserta', meta.targetPeserta || '-'],
                  ['Metode', meta.metode || 'Offline'],
                  ['Jumlah Sesi', computed.sesiKeys.length ? `${computed.sesiKeys.length} sesi` : '-'],
                  ['Pelaksanaan Sesi', meta.waktuSesi || computed.tanggalRingkas],
                  [
                    'Evaluasi',
                    testRows.length
                      ? `Pre-test dan post-test, masing-masing ${computed.jumlahSoal} soal`
                      : computed.postScores.length
                      ? 'Post-test'
                      : '-',
                  ],
                  ['Pemateri', meta.pemateri || '-'],
                ].map(([label, val]) => (
                  <tr key={label}>
                    <th className="w-1/3 border border-slate-200 bg-slate-50 px-3 py-2 text-left font-semibold text-slate-700">
                      {label}
                    </th>
                    <td className="border border-slate-200 px-3 py-2 text-slate-800">{val}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            {/* 2. TUJUAN & FOKUS TRAINING */}
            <h2 className="mb-3 mt-7 border-b-2 border-slate-200 pb-1.5 text-base font-bold text-blue-900">
              2. Tujuan & Fokus Training
            </h2>

            {isEditMode ? (
              <div className="mb-4 space-y-3">
                <div className="summary-strip rounded-md border-l-4 border-blue-700 bg-blue-50/70 px-3 py-2.5">
                  <div className="mb-1 flex items-center justify-between">
                    <span className="text-[11px] font-bold text-blue-800">Ringkasan Fokus</span>
                    {customSummary !== null && (
                      <button
                        type="button"
                        onClick={() => setCustomSummary(null)}
                        className="rounded bg-white px-2 py-0.5 text-[11px] font-semibold text-slate-600 hover:bg-slate-100"
                      >
                        Reset
                      </button>
                    )}
                  </div>
                  <textarea
                    rows={3}
                    value={effectiveSummary}
                    onChange={(e) => setCustomSummary(e.target.value)}
                    className={`${EDIT_INPUT_CLS} text-sm`}
                  />
                </div>

                {/* Edit Poin Tujuan */}
                <div className="rounded-md border border-slate-200 bg-slate-50/60 p-3">
                  <div className="mb-2 flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700">Poin Tujuan Training</span>
                    <button
                      type="button"
                      onClick={addGoal}
                      className="rounded bg-blue-600 px-2 py-0.5 text-[11px] font-semibold text-white hover:bg-blue-500"
                    >
                      + Tambah Tujuan
                    </button>
                  </div>
                  <div className="space-y-1.5">
                    {effectiveGoals.map((g, i) => (
                      <div key={i} className="flex items-center gap-1.5">
                        <input
                          type="text"
                          value={g}
                          onChange={(e) => updateGoalAt(i, e.target.value)}
                          className={`${EDIT_INPUT_CLS} text-xs`}
                        />
                        <button
                          type="button"
                          onClick={() => removeGoalAt(i)}
                          className="rounded bg-rose-100 px-2 py-1 text-xs font-bold text-rose-700 hover:bg-rose-200"
                          title="Hapus poin"
                        >
                          ✕
                        </button>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Edit Tabel Submateri */}
                <div className="rounded-md border border-slate-200 bg-slate-50/60 p-3">
                  <div className="mb-2 flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700">Tabel Submateri</span>
                    <button
                      type="button"
                      onClick={addSubmateri}
                      className="rounded bg-blue-600 px-2 py-0.5 text-[11px] font-semibold text-white hover:bg-blue-500"
                    >
                      + Tambah Submateri
                    </button>
                  </div>
                  <div className="space-y-2.5">
                    {effectiveSubmateri.map((m, i) => (
                      <div key={i} className="rounded border border-slate-200 bg-white p-2.5 text-xs">
                        <div className="mb-1.5 flex flex-wrap items-center gap-2">
                          <input
                            type="text"
                            value={m.topik}
                            onChange={(e) => updateSubmateriAt(i, { topik: e.target.value })}
                            placeholder="Nama Topik / Submateri"
                            className={`${EDIT_INPUT_CLS} flex-1 font-semibold text-blue-900`}
                          />
                          <select
                            value={m.format}
                            onChange={(e) => {
                              const nextFmt = e.target.value === 'poin' ? 'poin' : 'paragraf';
                              updateSubmateriAt(i, {
                                format: nextFmt,
                                uraianPoin:
                                  nextFmt === 'poin' && m.uraianPoin.length === 0 && m.uraianParagraf
                                    ? [m.uraianParagraf]
                                    : m.uraianPoin,
                                uraianParagraf:
                                  nextFmt === 'paragraf' && !m.uraianParagraf && m.uraianPoin.length > 0
                                    ? m.uraianPoin.join(' ')
                                    : m.uraianParagraf,
                              });
                            }}
                            className="rounded border border-slate-300 bg-white px-2 py-1 text-[11px] font-semibold"
                          >
                            <option value="poin">Format Poin (1 baris = 1 poin)</option>
                            <option value="paragraf">Format Paragraf</option>
                          </select>
                          <button
                            type="button"
                            onClick={() => removeSubmateriAt(i)}
                            className="rounded bg-rose-100 px-2 py-1 font-bold text-rose-700 hover:bg-rose-200"
                          >
                            ✕
                          </button>
                        </div>
                        {m.format === 'poin' ? (
                          <textarea
                            rows={Math.max(2, m.uraianPoin.length)}
                            value={m.uraianPoin.join('\n')}
                            onChange={(e) =>
                              updateSubmateriAt(i, { uraianPoin: e.target.value.split('\n') })
                            }
                            placeholder="Tulis tiap poin di baris baru..."
                            className={`${EDIT_INPUT_CLS} text-xs`}
                          />
                        ) : (
                          <textarea
                            rows={2}
                            value={m.uraianParagraf}
                            onChange={(e) => updateSubmateriAt(i, { uraianParagraf: e.target.value })}
                            placeholder="Uraian paragraf submateri..."
                            className={`${EDIT_INPUT_CLS} text-xs`}
                          />
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            ) : (
              <>
                <div className="summary-strip mb-3 whitespace-pre-line rounded-md border-l-4 border-blue-700 bg-blue-50/70 px-4 py-3 text-sm text-slate-800">
                  {effectiveSummary}
                </div>

                {effectiveGoals.filter((g) => g.trim()).length > 0 && (
                  <ul className="goals mb-4 list-disc space-y-1 pl-5 text-sm text-slate-700">
                    {effectiveGoals
                      .filter((g) => g.trim())
                      .map((t, i) => (
                        <li key={i}>{t}</li>
                      ))}
                  </ul>
                )}

                {effectiveSubmateri.length > 0 && (
                  <table className="material w-full border-collapse text-sm">
                    <tbody>
                      {effectiveSubmateri.map((m, i) => {
                        const cleanPoin = m.uraianPoin.filter((p) => p.trim());
                        return (
                          <tr key={i}>
                            <th className="w-1/3 border border-slate-200 bg-slate-50 px-3 py-2 text-left font-semibold text-blue-900">
                              {m.topik}
                            </th>
                            <td className="border border-slate-200 px-3 py-2 text-slate-700">
                              {m.format === 'poin' && cleanPoin.length > 0 ? (
                                <ul className="list-disc space-y-0.5 pl-5">
                                  {cleanPoin.map((p, pi) => (
                                    <li key={pi}>{p}</li>
                                  ))}
                                </ul>
                              ) : (
                                <span className="whitespace-pre-line">{m.uraianParagraf || '-'}</span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}
              </>
            )}

            {/* 3. HASIL EVALUASI */}
            <h2 className="mb-3 mt-7 border-b-2 border-slate-200 pb-1.5 text-base font-bold text-blue-900">
              3. Hasil Evaluasi Training
            </h2>
            <table className="info mb-4 w-full border-collapse text-sm">
              <thead>
                <tr>
                  <th className="border border-slate-200 bg-slate-50 px-3 py-2 text-left">Indikator</th>
                  <th className="w-1/4 border border-slate-200 bg-slate-50 px-3 py-2 text-left">Sebelum</th>
                  <th className="w-1/4 border border-slate-200 bg-slate-50 px-3 py-2 text-left">Sesudah</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td className="border border-slate-200 px-3 py-2">Rata-rata nilai</td>
                  <td className="border border-slate-200 px-3 py-2">{fmtNum(computed.avgPre)}</td>
                  <td className="border border-slate-200 px-3 py-2 font-bold">{fmtNum(computed.avgPost)}</td>
                </tr>
              </tbody>
            </table>

            <div className="chartwrap mb-3 rounded-lg border border-slate-200 bg-slate-50/70 p-4">
              <div className="chart-title mb-2 text-xs font-bold text-blue-900">Perbandingan Rata-rata Nilai per Tanggal</div>
              {Object.keys(computed.byDateScore).length === 0 ? (
                <div className="text-xs text-slate-500">Belum ada data evaluasi pre-test/post-test.</div>
              ) : (
                Object.keys(computed.byDateScore)
                  .sort()
                  .map((dk) => {
                    const g = computed.byDateScore[dk];
                    const lbl = fmtDateShort(dk);
                    const avgP = g.pre.length ? g.pre.reduce((a, b) => a + b, 0) / g.pre.length : null;
                    const avgPo = g.post.length ? g.post.reduce((a, b) => a + b, 0) / g.post.length : null;
                    return (
                      <div key={dk} className="mb-2 space-y-1.5">
                        {avgP !== null && (
                          <div className="barline flex items-center gap-2 text-xs">
                            <span className="w-24 font-semibold text-slate-600">{lbl} · Pre</span>
                            <div className="bar-track h-3 flex-1 overflow-hidden rounded-full bg-slate-200">
                              <span className="bar-fill pre block h-full rounded-full bg-slate-400" style={{ width: `${avgP.toFixed(1)}%` }} />
                            </div>
                            <strong className="w-10 text-right">{fmtNum(avgP)}</strong>
                          </div>
                        )}
                        {avgPo !== null && (
                          <div className="barline flex items-center gap-2 text-xs">
                            <span className="w-24 font-semibold text-slate-600">{lbl} · Post</span>
                            <div className="bar-track h-3 flex-1 overflow-hidden rounded-full bg-slate-200">
                              <span className="bar-fill alt block h-full rounded-full bg-blue-600" style={{ width: `${avgPo.toFixed(1)}%` }} />
                            </div>
                            <strong className="w-10 text-right">{fmtNum(avgPo)}</strong>
                          </div>
                        )}
                      </div>
                    );
                  })
              )}
              <div className="legend mt-2 text-[11px] text-slate-500">Skala nilai 0–100. Per tanggal menggabungkan sesi pada tanggal tersebut.</div>
            </div>

            <table className="kpi-small -mx-2 w-full border-separate border-spacing-2" role="presentation">
              <tbody>
                <tr>
                  <td className="w-1/3 rounded-lg border border-slate-200 bg-slate-50 p-2.5 text-center text-xs text-slate-500">
                    <b className="block text-base text-slate-900">{computed.preScores.length ? `${computed.naik} orang` : '-'}</b>
                    Nilai meningkat
                  </td>
                  <td className="w-1/3 rounded-lg border border-slate-200 bg-slate-50 p-2.5 text-center text-xs text-slate-500">
                    <b className="block text-base text-slate-900">{computed.preScores.length ? `${computed.tetap} orang` : '-'}</b>
                    Nilai tetap
                  </td>
                  <td className="w-1/3 rounded-lg border border-slate-200 bg-slate-50 p-2.5 text-center text-xs text-slate-500">
                    <b className="block text-base text-slate-900">{computed.preScores.length ? `${computed.turun} orang` : '-'}</b>
                    Nilai menurun
                  </td>
                </tr>
              </tbody>
            </table>

            {/* 4. ANALISIS & TINDAK LANJUT */}
            <div className="mb-3 mt-7 flex items-center justify-between border-b-2 border-slate-200 pb-1.5">
              <h2 className="text-base font-bold text-blue-900">4. Analisis & Tindak Lanjut</h2>
              {isEditMode && (
                <button
                  type="button"
                  onClick={addAnalisis}
                  className="rounded bg-blue-600 px-2.5 py-0.5 text-[11px] font-semibold text-white hover:bg-blue-500"
                >
                  + Tambah Analisis
                </button>
              )}
            </div>

            <div className="space-y-2">
              {isEditMode ? (
                effectiveAnalisis.length === 0 ? (
                  <div className="rounded-md border border-dashed border-slate-300 p-3 text-xs text-slate-500">
                    Belum ada poin analisis. Klik "+ Tambah Analisis" di kanan atas.
                  </div>
                ) : (
                  effectiveAnalisis.map((a, i) => (
                    <div
                      key={i}
                      className="space-y-1.5 rounded-md border border-slate-200 border-l-4 border-l-blue-900 bg-slate-50 p-3 text-sm"
                    >
                      <div className="flex items-center gap-2">
                        <input
                          type="text"
                          value={a.judul}
                          onChange={(e) => updateAnalisisAt(i, { judul: e.target.value })}
                          placeholder="Judul poin analisis..."
                          className={`${EDIT_INPUT_CLS} text-xs font-bold`}
                        />
                        <button
                          type="button"
                          onClick={() => removeAnalisisAt(i)}
                          className="rounded bg-rose-100 px-2 py-1 text-xs font-bold text-rose-700 hover:bg-rose-200"
                        >
                          ✕
                        </button>
                      </div>
                      <textarea
                        rows={2}
                        value={a.isi}
                        onChange={(e) => updateAnalisisAt(i, { isi: e.target.value })}
                        placeholder="Isi analisis / tindak lanjut..."
                        className={`${EDIT_INPUT_CLS} text-xs`}
                      />
                    </div>
                  ))
                )
              ) : effectiveAnalisis.length > 0 ? (
                effectiveAnalisis.map((a, i) => (
                  <div key={i} className="insight rounded-md border border-slate-200 border-l-4 border-l-blue-900 bg-slate-50 p-3 text-sm">
                    <strong>{a.judul}</strong> {a.isi}
                  </div>
                ))
              ) : (
                <div className="insight rounded-md border border-slate-200 bg-slate-50 p-3 text-sm text-slate-500">
                  Belum ada data evaluasi untuk dianalisis.
                </div>
              )}
            </div>

            {/* 5. DOKUMENTASI TRAINING */}
            <h2 className="mb-3 mt-7 border-b-2 border-slate-200 pb-1.5 text-base font-bold text-blue-900">
              5. Dokumentasi Training
            </h2>
            <div className="mb-3 text-xs text-slate-500">
              <strong className="text-slate-800">{photos.length} Foto Dokumentasi Kegiatan</strong> · {computed.tanggalRingkas}
            </div>
            <div id="photoGrid" className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              {photos.length === 0 ? (
                <div className="col-span-full rounded-md border border-dashed border-slate-300 p-4 text-center text-xs text-slate-500">
                  Belum ada foto dokumentasi terlampir.
                </div>
              ) : (
                photos.map((photo, idx) => (
                  <div key={idx} className="rounded-lg border border-slate-200 bg-white p-2">
                    <img src={photo.src} alt={photo.caption} className="mb-2 h-48 w-full rounded object-cover" />
                    <div className="text-xs font-bold text-slate-800">{photo.caption}</div>
                    {isLoggedIn && (
                      <div className="no-export mt-2 flex gap-1.5 print:hidden">
                        <input
                          type="text"
                          value={photo.caption}
                          onChange={(e) =>
                            setPhotos((prev) => prev.map((p, i) => (i === idx ? { ...p, caption: e.target.value } : p)))
                          }
                          className="flex-1 rounded border border-slate-300 px-2 py-1 text-xs"
                          placeholder="Edit caption..."
                        />
                        <button
                          type="button"
                          onClick={() => setPhotos((prev) => prev.filter((_, i) => i !== idx))}
                          className="rounded bg-rose-100 px-2 py-1 text-xs font-semibold text-rose-700 hover:bg-rose-200"
                        >
                          Hapus
                        </button>
                      </div>
                    )}
                  </div>
                ))
              )}
            </div>

            {/* 6. DAFTAR HADIR */}
            <h2 className="mb-3 mt-7 border-b-2 border-slate-200 pb-1.5 text-base font-bold text-blue-900">
              6. Daftar Hadir Training
            </h2>
            {computed.sesiKeys.length === 0 ? (
              <div className="text-xs text-slate-500">Belum ada data daftar hadir. Unggah file Excel absensi melalui tombol Upload File.</div>
            ) : (
              computed.sesiKeys.map((dk, idx) => {
                const list = computed.sesiMap[dk];
                const half = Math.ceil(list.length / 2);
                const tglLbl = dk !== 'Tanpa Tanggal' ? fmtDateMed(dk) : '';
                return (
                  <div key={dk} className="mb-4">
                    <div className="session-heading mb-1.5 flex items-center justify-between">
                      <h3 className="text-sm font-bold text-blue-900">
                        Sesi {idx + 1} <span className="font-normal text-slate-500">{tglLbl}</span>
                      </h3>
                      <span className="session-count rounded-full bg-sky-100 px-2.5 py-0.5 text-xs font-bold text-sky-800">
                        {list.length} entri
                      </span>
                    </div>
                    <table className="attendance w-full border-collapse text-xs">
                      <thead>
                        <tr className="bg-blue-900 text-white">
                          <th className="no w-[7%] border border-slate-300 px-2 py-1.5 text-center">No.</th>
                          <th className="w-[43%] border border-slate-300 px-2 py-1.5 text-left">Nama Peserta</th>
                          <th className="no w-[7%] border border-slate-300 px-2 py-1.5 text-center">No.</th>
                          <th className="w-[43%] border border-slate-300 px-2 py-1.5 text-left">Nama Peserta</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Array.from({ length: half }).map((_, i) => {
                          const rIdx = i + half;
                          const hasR = rIdx < list.length;
                          const leftName = getAttendanceName(dk, i, list[i]);
                          const rightName = hasR ? getAttendanceName(dk, rIdx, list[rIdx]) : '';
                          return (
                            <tr key={i}>
                              <td className="no border border-slate-200 px-2 py-1 text-center">{i + 1}</td>
                              <td className="border border-slate-200 px-2 py-1">
                                {isEditMode ? (
                                  <input
                                    type="text"
                                    value={leftName}
                                    onChange={(e) => setAttendanceName(dk, i, e.target.value)}
                                    className={`${EDIT_INPUT_CLS} py-0.5 text-xs`}
                                  />
                                ) : (
                                  leftName
                                )}
                              </td>
                              <td className="no border border-slate-200 px-2 py-1 text-center">{hasR ? rIdx + 1 : ''}</td>
                              <td className="border border-slate-200 px-2 py-1">
                                {hasR ? (
                                  isEditMode ? (
                                    <input
                                      type="text"
                                      value={rightName}
                                      onChange={(e) => setAttendanceName(dk, rIdx, e.target.value)}
                                      className={`${EDIT_INPUT_CLS} py-0.5 text-xs`}
                                    />
                                  ) : (
                                    rightName
                                  )
                                ) : (
                                  ''
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                );
              })
            )}

            {/* 7. KESIMPULAN */}
            <h2 className="mb-3 mt-7 border-b-2 border-slate-200 pb-1.5 text-base font-bold text-blue-900">7. Kesimpulan</h2>
            {isEditMode ? (
              <div className="space-y-2.5 rounded-lg border border-blue-200 bg-blue-50/30 p-3">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] font-bold uppercase tracking-wide text-blue-700">
                    Kesimpulan & Penutup
                  </span>
                  {(customKesimpulan !== null || customClosing !== null) && (
                    <button
                      type="button"
                      onClick={() => {
                        setCustomKesimpulan(null);
                        setCustomClosing(null);
                      }}
                      className="rounded bg-slate-200 px-2 py-0.5 text-[11px] font-semibold text-slate-700 hover:bg-slate-300"
                    >
                      Reset ke Default
                    </button>
                  )}
                </div>
                <textarea
                  rows={3}
                  value={effectiveKesimpulan}
                  onChange={(e) => setCustomKesimpulan(e.target.value)}
                  className={`${EDIT_INPUT_CLS} text-sm`}
                  placeholder="Paragraf kesimpulan..."
                />
                <textarea
                  rows={2}
                  value={effectiveClosing}
                  onChange={(e) => setCustomClosing(e.target.value)}
                  className={`${EDIT_INPUT_CLS} text-sm`}
                  placeholder="Kalimat penutup..."
                />
                <div className="pt-2 text-sm">
                  <div className="mb-1 text-slate-700">Hormat kami,</div>
                  <div className="grid max-w-sm grid-cols-1 gap-1.5">
                    <input
                      type="text"
                      value={meta.ttd}
                      onChange={(e) => setMeta((prev) => ({ ...prev, ttd: e.target.value }))}
                      placeholder="Nama Penandatangan (misal: Tim Training)"
                      className={`${EDIT_INPUT_CLS} text-sm font-bold`}
                    />
                    <input
                      type="text"
                      value={meta.entitas}
                      onChange={(e) => setMeta((prev) => ({ ...prev, entitas: e.target.value }))}
                      placeholder="Entitas Pengelola"
                      className={`${EDIT_INPUT_CLS} text-sm font-bold`}
                    />
                  </div>
                </div>
              </div>
            ) : (
              <>
                <p className="whitespace-pre-line text-sm text-slate-700">{effectiveKesimpulan}</p>
                <p className="mt-2 whitespace-pre-line text-sm text-slate-700">{effectiveClosing}</p>
                <p className="sig mt-5 text-sm">
                  Hormat kami,
                  <br />
                  <b>{meta.ttd || 'Tim Training'}</b>
                  <br />
                  <b>{computed.entitas}</b>
                </p>
              </>
            )}
          </div>
          <div className="footer border-t border-slate-200 bg-slate-50 px-6 py-3 text-center text-xs text-slate-500">
            Laporan {computed.judul} • {computed.tanggalRingkas} • {computed.site}
          </div>
        </main>
      </div>

      {/* MODAL UPLOAD 3-IN-1 */}
      <UploadModal
        open={uploadModalOpen}
        onClose={() => setUploadModalOpen(false)}
        onExcelUpload={handleExcelUpload}
        onPptUpload={handlePptUpload}
        onPhotoUpload={handlePhotoUpload}
        onClearExcel={() => {
          setAbsensiRows([]);
          setTestRows([]);
          setRawExcelFiles([]);
          setCustomAttendanceNames({});
          setStatus({ msg: 'Data Excel direset.', type: '' });
        }}
        onClearPpt={() => {
          setPptSlides([]);
          setRawPptFile(null);
          setAiData(null);
          setCustomSummary(null);
          setCustomGoals(null);
          setCustomSubmateri(null);
          setCustomAnalisis(null);
          setCustomKesimpulan(null);
          setStatus({ msg: 'Data Materi PPT direset.', type: '' });
        }}
        onClearPhotos={() => {
          setPhotos([]);
          setStatus({ msg: 'Foto dokumentasi direset.', type: '' });
        }}
        rawExcelFiles={rawExcelFiles}
        absensiCount={absensiRows.length}
        testCount={testRows.length}
        rawPptFile={rawPptFile}
        pptSlideCount={pptSlides.length}
        photosCount={photos.length}
        isReadingFile={isReadingFile}
        isAiLoading={isAiLoading}
        status={status}
        modalError={uploadModalError}
        onSaveAndRefine={handleSaveUploadAndRefine}
      />

      {/* MODAL RIWAYAT */}
      <HistoryModal open={historyModalOpen} onClose={() => setHistoryModalOpen(false)} />

      {/* MODAL EDIT INFO */}
      {metaModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 print:hidden">
          <div className="w-full max-w-xl rounded-xl bg-white p-5 shadow-xl">
            <div className="mb-4 flex items-center justify-between">
              <h3 className="text-base font-bold text-slate-900">Edit Informasi Pelaksanaan</h3>
              <button type="button" onClick={() => setMetaModalOpen(false)} className="text-slate-400 hover:text-slate-600">
                ✕
              </button>
            </div>
            <div className="grid grid-cols-1 gap-3 text-xs sm:grid-cols-2">
              {(
                [
                  ['Entitas Pengelola', 'entitas'],
                  ['Perusahaan / Site', 'site'],
                  ['Nama Training', 'namaTraining'],
                  ['Target Peserta', 'targetPeserta'],
                  ['Metode', 'metode'],
                  ['Pemateri', 'pemateri'],
                  ['Pelaksanaan Sesi (Jam)', 'waktuSesi'],
                  ['Nama Penandatangan', 'ttd'],
                ] as [string, keyof ReportMeta][]
              ).map(([label, key]) => (
                <div key={key}>
                  <label className="mb-1 block font-semibold text-slate-700">{label}</label>
                  <input
                    type="text"
                    value={draftMeta[key]}
                    onChange={(e) => setDraftMeta({ ...draftMeta, [key]: e.target.value })}
                    className="w-full rounded-md border border-slate-300 px-2.5 py-1.5"
                  />
                </div>
              ))}
              <div className="sm:col-span-2">
                <label className="mb-1 block font-semibold text-slate-700">Ringkasan Tujuan & Fokus Training</label>
                <textarea
                  rows={3}
                  value={draftMeta.tujuan}
                  onChange={(e) => setDraftMeta({ ...draftMeta, tujuan: e.target.value })}
                  className="w-full rounded-md border border-slate-300 px-2.5 py-1.5"
                />
              </div>
            </div>
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setMetaModalOpen(false)}
                className="rounded-md border border-slate-300 px-3 py-1.5 text-xs font-semibold"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={() => {
                  setMeta(draftMeta);
                  setMetaModalOpen(false);
                  setStatus({ msg: 'Informasi pelaksanaan diperbarui.', type: 'ok' });
                }}
                className="rounded-md bg-blue-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-blue-500"
              >
                Simpan Perubahan
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL KIRIM EMAIL */}
      {emailModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 print:hidden">
          <div className="max-h-[90vh] w-full max-w-xl overflow-y-auto rounded-xl bg-white p-5 shadow-xl">
            <div className="mb-4 flex items-center justify-between">
              <h3 className="text-base font-bold text-slate-900">Kirim Laporan via Email</h3>
              <button type="button" onClick={() => setEmailModalOpen(false)} className="text-slate-400 hover:text-slate-600">
                ✕
              </button>
            </div>

            <div className="mb-3 rounded-md bg-slate-100 px-3 py-2 text-xs text-slate-700">
              Pengirim aktif: <b>{session?.user?.email}</b>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="mb-1 block font-semibold">Kepada (To) * (pisahkan koma jika lebih dari 1)</label>
                <input
                  type="text"
                  value={emailForm.to}
                  onChange={(e) => setEmailForm({ ...emailForm, to: e.target.value })}
                  placeholder="manajemen1@perusahaan.com, manajemen2@perusahaan.com"
                  className="w-full rounded-md border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                <div>
                  <label className="mb-1 block font-semibold">CC (opsional, pisahkan koma)</label>
                  <input
                    type="text"
                    value={emailForm.cc}
                    onChange={(e) => setEmailForm({ ...emailForm, cc: e.target.value })}
                    placeholder="cc1@perusahaan.com, cc2@perusahaan.com"
                    className="w-full rounded-md border border-slate-300 px-2.5 py-1.5"
                  />
                </div>
                <div>
                  <label className="mb-1 block font-semibold">BCC (opsional, pisahkan koma)</label>
                  <input
                    type="text"
                    value={emailForm.bcc}
                    onChange={(e) => setEmailForm({ ...emailForm, bcc: e.target.value })}
                    placeholder="bcc1@perusahaan.com"
                    className="w-full rounded-md border border-slate-300 px-2.5 py-1.5"
                  />
                </div>
              </div>
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                <div>
                  <label className="mb-1 block font-semibold">Nama Pengirim</label>
                  <input
                    type="text"
                    value={emailForm.senderName}
                    onChange={(e) => setEmailForm({ ...emailForm, senderName: e.target.value })}
                    className="w-full rounded-md border border-slate-300 px-2.5 py-1.5"
                  />
                </div>
                <div>
                  <label className="mb-1 block font-semibold">Reply-To (opsional)</label>
                  <input
                    type="email"
                    value={emailForm.replyTo}
                    onChange={(e) => setEmailForm({ ...emailForm, replyTo: e.target.value })}
                    className="w-full rounded-md border border-slate-300 px-2.5 py-1.5"
                  />
                </div>
              </div>
              <div>
                <label className="mb-1 block font-semibold">Subjek Email *</label>
                <input
                  type="text"
                  value={emailForm.subject}
                  onChange={(e) => setEmailForm({ ...emailForm, subject: e.target.value })}
                  className="w-full rounded-md border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">Pesan Pengantar (Plain-Text Fallback)</label>
                <textarea
                  rows={4}
                  value={emailForm.body}
                  onChange={(e) => setEmailForm({ ...emailForm, body: e.target.value })}
                  className="w-full rounded-md border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div className="space-y-1.5 pt-1">
                <label className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    checked={emailForm.attachHtml}
                    onChange={(e) => setEmailForm({ ...emailForm, attachHtml: e.target.checked })}
                  />
                  Lampirkan juga file Laporan HTML (.html)
                </label>
                <label className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    checked={emailForm.attachExcel}
                    onChange={(e) => setEmailForm({ ...emailForm, attachExcel: e.target.checked })}
                  />
                  File Excel sumber ({rawExcelFiles.length} file)
                </label>
                <label className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    checked={emailForm.attachPpt}
                    onChange={(e) => setEmailForm({ ...emailForm, attachPpt: e.target.checked })}
                  />
                  File Materi PPT ({rawPptFile ? rawPptFile.name : 'belum diunggah'})
                </label>
              </div>
            </div>
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setEmailModalOpen(false)}
                className="rounded-md border border-slate-300 px-3 py-1.5 text-xs font-semibold"
              >
                Batal
              </button>
              <button
                type="button"
                disabled={isSendingEmail || !session?.user?.email}
                onClick={handleSendEmail}
                className="rounded-md bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-emerald-500 disabled:opacity-50"
              >
                {isSendingEmail ? 'Mengirim...' : 'Kirim Sekarang'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}