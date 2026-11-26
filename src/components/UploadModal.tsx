'use client';

import React from 'react';
import type { RawFileAttachment } from '@/types/report';

interface UploadModalProps {
  open: boolean;
  onClose: () => void;
  onExcelUpload: (e: React.ChangeEvent<HTMLInputElement>) => Promise<void>;
  onPptUpload: (e: React.ChangeEvent<HTMLInputElement>) => Promise<void>;
  onPhotoUpload: (e: React.ChangeEvent<HTMLInputElement>) => Promise<void>;
  onClearExcel: () => void;
  onClearPpt: () => void;
  onClearPhotos: () => void;
  rawExcelFiles: RawFileAttachment[];
  absensiCount: number;
  testCount: number;
  rawPptFile: RawFileAttachment | null;
  pptSlideCount: number;
  photosCount: number;
  isReadingFile: boolean;
  isAiLoading: boolean;
  status: { msg: string; type: '' | 'ok' | 'err' };
  modalError: string | null;
  onSaveAndRefine: () => Promise<void>;
}

export default function UploadModal({
  open,
  onClose,
  onExcelUpload,
  onPptUpload,
  onPhotoUpload,
  onClearExcel,
  onClearPpt,
  onClearPhotos,
  rawExcelFiles,
  absensiCount,
  testCount,
  rawPptFile,
  pptSlideCount,
  photosCount,
  isReadingFile,
  isAiLoading,
  status,
  modalError,
  onSaveAndRefine,
}: UploadModalProps) {
  if (!open) return null;

  const isBusy = isReadingFile || isAiLoading;
  const hasAnyFile =
    rawExcelFiles.length > 0 || Boolean(rawPptFile) || photosCount > 0;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 backdrop-blur-[1px] print:hidden">
      <div className="flex max-h-[90vh] w-full max-w-2xl flex-col overflow-hidden rounded-xl bg-white shadow-2xl">
        {/* Header Modal */}
        <div className="flex items-center justify-between border-b border-slate-200 bg-[#001247] px-5 py-3.5 text-white">
          <div>
            <h3 className="text-sm font-bold tracking-wide">Upload File Laporan Training</h3>
            <p className="text-[11px] text-blue-200">
              Unggah file Absen/Test, Materi PPT, dan Foto Dokumentasi lalu klik Simpan.
            </p>
          </div>
          <button
            type="button"
            disabled={isBusy}
            onClick={onClose}
            className="rounded-md px-2 py-1 text-slate-300 hover:bg-white/10 hover:text-white disabled:opacity-40"
          >
            ✕
          </button>
        </div>

        {/* Body Modal */}
        <div className="space-y-4 overflow-y-auto p-5 text-xs">
          {/* 1. ABSENSI & EVALUASI */}
          <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-4">
            <div className="mb-2 flex items-center justify-between gap-2">
              <div>
                <h4 className="font-bold text-slate-800">1. File Absensi & Evaluasi (.xlsx / .xls)</h4>
                <p className="text-[11px] text-slate-500">
                  Bisa pilih lebih dari 1 file sekaligus (Absensi & Pre/Post Test).
                </p>
              </div>
              {rawExcelFiles.length > 0 && (
                <button
                  type="button"
                  disabled={isBusy}
                  onClick={onClearExcel}
                  className="rounded bg-rose-100 px-2 py-1 text-[11px] font-semibold text-rose-700 hover:bg-rose-200 disabled:opacity-50"
                >
                  Reset
                </button>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-3">
              <label
                className={`inline-flex cursor-pointer items-center gap-1.5 rounded-md bg-blue-600 px-3 py-1.5 font-semibold text-white hover:bg-blue-500 ${
                  isBusy ? 'pointer-events-none opacity-50' : ''
                }`}
              >
                <span>+ Pilih File Excel</span>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  multiple
                  disabled={isBusy}
                  onChange={onExcelUpload}
                  className="hidden"
                />
              </label>

              {rawExcelFiles.length > 0 ? (
                <span className="font-semibold text-emerald-700">
                  ✓ {rawExcelFiles.length} file terbaca ({absensiCount} baris absen, {testCount} baris tes)
                </span>
              ) : (
                <span className="text-slate-400">Belum ada file Excel diunggah</span>
              )}
            </div>

            {rawExcelFiles.length > 0 && (
              <div className="mt-2 flex flex-wrap gap-1.5">
                {rawExcelFiles.map((f, idx) => (
                  <span
                    key={`${f.name}-${idx}`}
                    className="rounded border border-slate-200 bg-white px-2 py-0.5 text-[11px] text-slate-700"
                  >
                    {f.name}
                  </span>
                ))}
              </div>
            )}
          </div>

          {/* 2. MATERI PPT */}
          <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-4">
            <div className="mb-2 flex items-center justify-between gap-2">
              <div>
                <h4 className="font-bold text-slate-800">2. File Materi Training (.pptx)</h4>
                <p className="text-[11px] text-slate-500">
                  Teks slide akan diekstrak otomatis untuk dirapikan oleh Gemini AI.
                </p>
              </div>
              {rawPptFile && (
                <button
                  type="button"
                  disabled={isBusy}
                  onClick={onClearPpt}
                  className="rounded bg-rose-100 px-2 py-1 text-[11px] font-semibold text-rose-700 hover:bg-rose-200 disabled:opacity-50"
                >
                  Reset
                </button>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-3">
              <label
                className={`inline-flex cursor-pointer items-center gap-1.5 rounded-md bg-blue-600 px-3 py-1.5 font-semibold text-white hover:bg-blue-500 ${
                  isBusy ? 'pointer-events-none opacity-50' : ''
                }`}
              >
                <span>+ Pilih File PPTX</span>
                <input
                  type="file"
                  accept=".pptx"
                  disabled={isBusy}
                  onChange={onPptUpload}
                  className="hidden"
                />
              </label>

              {rawPptFile ? (
                <span className="font-semibold text-emerald-700">
                  ✓ {rawPptFile.name} ({pptSlideCount} slide terbaca)
                </span>
              ) : (
                <span className="text-slate-400">Belum ada file PPTX diunggah</span>
              )}
            </div>
          </div>

          {/* 3. DOKUMENTASI FOTO */}
          <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-4">
            <div className="mb-2 flex items-center justify-between gap-2">
              <div>
                <h4 className="font-bold text-slate-800">3. Foto Dokumentasi Kegiatan (Multi)</h4>
                <p className="text-[11px] text-slate-500">
                  Foto otomatis dikompresi agar ringan saat dikirim via email.
                </p>
              </div>
              {photosCount > 0 && (
                <button
                  type="button"
                  disabled={isBusy}
                  onClick={onClearPhotos}
                  className="rounded bg-rose-100 px-2 py-1 text-[11px] font-semibold text-rose-700 hover:bg-rose-200 disabled:opacity-50"
                >
                  Reset
                </button>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-3">
              <label
                className={`inline-flex cursor-pointer items-center gap-1.5 rounded-md bg-blue-600 px-3 py-1.5 font-semibold text-white hover:bg-blue-500 ${
                  isBusy ? 'pointer-events-none opacity-50' : ''
                }`}
              >
                <span>+ Pilih Foto</span>
                <input
                  type="file"
                  accept="image/*"
                  multiple
                  disabled={isBusy}
                  onChange={onPhotoUpload}
                  className="hidden"
                />
              </label>

              {photosCount > 0 ? (
                <span className="font-semibold text-emerald-700">
                  ✓ {photosCount} foto siap dilampirkan
                </span>
              ) : (
                <span className="text-slate-400">Belum ada foto dokumentasi</span>
              )}
            </div>
          </div>

          {/* LIVE STATUS PROSES DI DALAM MODAL */}
          {(isBusy || modalError || status.type === 'ok') && (
            <div
              className={`flex items-center gap-2 rounded-lg border px-3 py-2.5 text-xs font-medium ${
                modalError
                  ? 'border-rose-200 bg-rose-50 text-rose-700'
                  : isBusy
                  ? 'border-blue-200 bg-blue-50 text-blue-800'
                  : 'border-emerald-200 bg-emerald-50 text-emerald-800'
              }`}
            >
              {isBusy && (
                <span className="inline-block h-3.5 w-3.5 shrink-0 animate-spin rounded-full border-2 border-blue-600 border-t-transparent" />
              )}
              <span>{modalError || status.msg}</span>
            </div>
          )}
        </div>

        {/* Footer Modal */}
        <div className="flex items-center justify-between border-t border-slate-200 bg-slate-50 px-5 py-3">
          <span className="text-[11px] text-slate-500">
            Klik <b>Simpan</b> untuk memproses & merapikan laporan dengan AI.
          </span>
          <div className="flex gap-2">
            <button
              type="button"
              disabled={isBusy}
              onClick={onClose}
              className="rounded-md border border-slate-300 bg-white px-3.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-100 disabled:opacity-50"
            >
              Tutup
            </button>
            <button
              type="button"
              disabled={isBusy || !hasAnyFile}
              onClick={onSaveAndRefine}
              className="inline-flex items-center gap-2 rounded-md bg-emerald-600 px-4 py-1.5 text-xs font-semibold text-white hover:bg-emerald-500 disabled:opacity-50"
            >
              {isAiLoading ? (
                <>
                  <span className="inline-block h-3 w-3 animate-spin rounded-full border-2 border-white border-t-transparent" />
                  Memproses AI...
                </>
              ) : (
                'Simpan'
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}