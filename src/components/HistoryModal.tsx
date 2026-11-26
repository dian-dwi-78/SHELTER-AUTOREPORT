'use client';

import React, { useCallback, useEffect, useState } from 'react';

interface HistoryItem {
  id: number;
  sentAt: string;
  senderEmail: string;
  toRecipients: string;
  ccRecipients: string | null;
  subject: string;
  trainingName: string | null;
  siteName: string | null;
  attachments: string | null;
  status: string;
  errorMessage: string | null;
}

interface HistoryDetail extends HistoryItem {
  htmlBody: string;
}

interface HistoryModalProps {
  open: boolean;
  onClose: () => void;
}

function formatDateTime(iso: string) {
  const d = new Date(iso);
  const tanggal = d.toLocaleDateString('id-ID', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  });
  const jam = d.toLocaleTimeString('id-ID', {
    hour: '2-digit',
    minute: '2-digit',
  });
  return { tanggal, jam };
}

function parseAttachments(raw: string | null): string[] {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export default function HistoryModal({ open, onClose }: HistoryModalProps) {
  const [items, setItems] = useState<HistoryItem[]>([]);
  const [senders, setSenders] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Filter states
  const [q, setQ] = useState('');
  const [sender, setSender] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');

  // Detail view states
  const [selectedDetail, setSelectedDetail] = useState<HistoryDetail | null>(null);
  const [loadingDetail, setLoadingDetail] = useState(false);

  const fetchHistory = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const params = new URLSearchParams();
      if (q.trim()) params.set('q', q.trim());
      if (sender) params.set('sender', sender);
      if (statusFilter) params.set('status', statusFilter);
      if (dateFrom) params.set('dateFrom', dateFrom);
      if (dateTo) params.set('dateTo', dateTo);

      const res = await fetch(`/api/history?${params.toString()}`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Gagal memuat riwayat.');

      setItems(data.items || []);
      setSenders(data.senders || []);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : 'Gagal memuat riwayat.');
    } finally {
      setLoading(false);
    }
  }, [q, sender, statusFilter, dateFrom, dateTo]);

  useEffect(() => {
    if (open) {
      setSelectedDetail(null);
      fetchHistory();
    }
  }, [open, fetchHistory]);

  async function handleOpenDetail(id: number) {
    setLoadingDetail(true);
    setError(null);
    try {
      const res = await fetch(`/api/history/${id}`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Gagal memuat detail.');
      setSelectedDetail(data.detail);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : 'Gagal memuat detail.');
    } finally {
      setLoadingDetail(false);
    }
  }

  function handleResetFilters() {
    setQ('');
    setSender('');
    setStatusFilter('');
    setDateFrom('');
    setDateTo('');
  }

  if (!open) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-3 backdrop-blur-[1px] print:hidden sm:p-5">
      <div className="flex h-[90vh] w-full max-w-6xl flex-col overflow-hidden rounded-xl bg-white shadow-2xl">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-200 bg-[#001247] px-5 py-3.5 text-white">
          <div className="flex items-center gap-3">
            {selectedDetail && (
              <button
                type="button"
                onClick={() => setSelectedDetail(null)}
                className="rounded-md bg-white/15 px-2.5 py-1 text-xs font-semibold hover:bg-white/25"
              >
                ← Kembali
              </button>
            )}
            <div>
              <h3 className="text-sm font-bold tracking-wide">
                {selectedDetail
                  ? `Detail Pengiriman #${selectedDetail.id}`
                  : 'Riwayat Pengiriman Laporan (Tim)'}
              </h3>
              <p className="text-[11px] text-blue-200">
                {selectedDetail
                  ? selectedDetail.subject
                  : 'Semua riwayat pengiriman email dari seluruh pengguna tercatat di sini.'}
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="rounded-md px-2 py-1 text-slate-300 hover:bg-white/10 hover:text-white"
          >
            ✕
          </button>
        </div>

        {/* Content */}
        {selectedDetail ? (
          <div className="flex-1 space-y-4 overflow-y-auto p-5 text-xs">
            {/* Metadata Pengiriman */}
            <div className="grid grid-cols-1 gap-3 rounded-lg border border-slate-200 bg-slate-50 p-4 sm:grid-cols-2">
              <div>
                <span className="block text-[11px] font-semibold text-slate-500">Waktu Pengiriman</span>
                <span className="font-bold text-slate-800">
                  {formatDateTime(selectedDetail.sentAt).tanggal} · {formatDateTime(selectedDetail.sentAt).jam} WIB
                </span>
              </div>
              <div>
                <span className="block text-[11px] font-semibold text-slate-500">Status</span>
                <span
                  className={`inline-block rounded-full px-2.5 py-0.5 text-[11px] font-bold ${
                    selectedDetail.status === 'BERHASIL'
                      ? 'bg-emerald-100 text-emerald-700'
                      : 'bg-rose-100 text-rose-700'
                  }`}
                >
                  {selectedDetail.status}
                </span>
              </div>
              <div>
                <span className="block text-[11px] font-semibold text-slate-500">Pengirim (From)</span>
                <span className="font-semibold text-slate-800">{selectedDetail.senderEmail}</span>
              </div>
              <div>
                <span className="block text-[11px] font-semibold text-slate-500">Kepada (To)</span>
                <span className="break-all text-slate-800">{selectedDetail.toRecipients}</span>
              </div>
              <div>
                <span className="block text-[11px] font-semibold text-slate-500">CC</span>
                <span className="break-all text-slate-800">{selectedDetail.ccRecipients || '-'}</span>
              </div>
              <div>
                <span className="block text-[11px] font-semibold text-slate-500">Lampiran</span>
                <div className="mt-0.5 flex flex-wrap gap-1">
                  {parseAttachments(selectedDetail.attachments).length > 0 ? (
                    parseAttachments(selectedDetail.attachments).map((name, i) => (
                      <span
                        key={i}
                        className="rounded border border-slate-200 bg-white px-2 py-0.5 text-[11px] text-slate-700"
                      >
                        {name}
                      </span>
                    ))
                  ) : (
                    <span className="text-slate-500">Tidak ada lampiran tambahan</span>
                  )}
                </div>
              </div>
              {(selectedDetail.trainingName || selectedDetail.siteName) && (
                <div className="sm:col-span-2">
                  <span className="block text-[11px] font-semibold text-slate-500">Training & Site</span>
                  <span className="text-slate-800">
                    {selectedDetail.trainingName || '-'} · {selectedDetail.siteName || '-'}
                  </span>
                </div>
              )}
              {selectedDetail.errorMessage && (
                <div className="rounded border border-rose-200 bg-rose-50 p-2.5 text-rose-700 sm:col-span-2">
                  <b>Pesan Error:</b> {selectedDetail.errorMessage}
                </div>
              )}
            </div>

            {/* Preview HTML Email */}
            <div>
              <h4 className="mb-2 text-xs font-bold uppercase tracking-wide text-slate-600">
                Preview Laporan yang Dikirimkan
              </h4>
              {selectedDetail.htmlBody ? (
                <iframe
                  title={`Preview Email #${selectedDetail.id}`}
                  srcDoc={selectedDetail.htmlBody}
                  sandbox="allow-same-origin"
                  className="h-[520px] w-full rounded-lg border border-slate-300 bg-white"
                />
              ) : (
                <div className="rounded-lg border border-dashed border-slate-300 p-6 text-center text-slate-500">
                  Konten preview HTML tidak tersedia untuk riwayat ini.
                </div>
              )}
            </div>
          </div>
        ) : (
          <div className="flex flex-1 flex-col overflow-hidden">
            {/* Bar Pencarian & Filter */}
            <div className="border-b border-slate-200 bg-slate-50 px-5 py-3 text-xs">
              <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-6">
                <input
                  type="text"
                  value={q}
                  onChange={(e) => setQ(e.target.value)}
                  placeholder="Cari subjek, tujuan, training, site..."
                  className="rounded-md border border-slate-300 bg-white px-2.5 py-1.5 lg:col-span-2"
                />
                <select
                  value={sender}
                  onChange={(e) => setSender(e.target.value)}
                  className="rounded-md border border-slate-300 bg-white px-2.5 py-1.5"
                >
                  <option value="">Semua Pengirim</option>
                  {senders.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
                <select
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                  className="rounded-md border border-slate-300 bg-white px-2.5 py-1.5"
                >
                  <option value="">Semua Status</option>
                  <option value="BERHASIL">BERHASIL</option>
                  <option value="GAGAL">GAGAL</option>
                </select>
                <input
                  type="date"
                  value={dateFrom}
                  onChange={(e) => setDateFrom(e.target.value)}
                  title="Dari Tanggal"
                  className="rounded-md border border-slate-300 bg-white px-2 py-1.5"
                />
                <input
                  type="date"
                  value={dateTo}
                  onChange={(e) => setDateTo(e.target.value)}
                  title="Sampai Tanggal"
                  className="rounded-md border border-slate-300 bg-white px-2 py-1.5"
                />
              </div>

              <div className="mt-2 flex items-center justify-between">
                <span className="text-[11px] text-slate-500">
                  Menampilkan <b>{items.length}</b> riwayat
                </span>
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={handleResetFilters}
                    className="rounded border border-slate-300 bg-white px-2.5 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100"
                  >
                    Reset Filter
                  </button>
                  <button
                    type="button"
                    onClick={fetchHistory}
                    className="rounded bg-blue-600 px-2.5 py-1 text-[11px] font-semibold text-white hover:bg-blue-500"
                  >
                    Refresh
                  </button>
                </div>
              </div>
            </div>

            {/* Tabel Riwayat */}
            <div className="flex-1 overflow-auto p-5">
              {error && (
                <div className="mb-3 rounded-md border border-rose-200 bg-rose-50 px-3 py-2 text-xs text-rose-700">
                  {error}
                </div>
              )}

              {loading ? (
                <div className="flex items-center justify-center gap-2 py-16 text-xs text-slate-500">
                  <span className="inline-block h-4 w-4 animate-spin rounded-full border-2 border-blue-600 border-t-transparent" />
                  Memuat riwayat pengiriman...
                </div>
              ) : items.length === 0 ? (
                <div className="rounded-lg border border-dashed border-slate-300 py-14 text-center text-xs text-slate-500">
                  Belum ada riwayat pengiriman yang sesuai dengan filter.
                </div>
              ) : (
                <table className="w-full border-collapse text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-100 text-slate-700">
                      <th className="px-3 py-2.5 font-bold">Tanggal</th>
                      <th className="px-3 py-2.5 font-bold">Jam</th>
                      <th className="px-3 py-2.5 font-bold">Pengirim</th>
                      <th className="px-3 py-2.5 font-bold">Kepada</th>
                      <th className="px-3 py-2.5 font-bold">Subjek</th>
                      <th className="px-3 py-2.5 font-bold">Status</th>
                      <th className="px-3 py-2.5 text-center font-bold">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {items.map((item) => {
                      const { tanggal, jam } = formatDateTime(item.sentAt);
                      return (
                        <tr key={item.id} className="hover:bg-slate-50">
                          <td className="whitespace-nowrap px-3 py-2.5 text-slate-700">{tanggal}</td>
                          <td className="whitespace-nowrap px-3 py-2.5 text-slate-600">{jam}</td>
                          <td className="max-w-[180px] truncate px-3 py-2.5 font-medium text-slate-800" title={item.senderEmail}>
                            {item.senderEmail}
                          </td>
                          <td className="max-w-[200px] truncate px-3 py-2.5 text-slate-700" title={item.toRecipients}>
                            {item.toRecipients}
                            {item.ccRecipients && (
                              <span className="ml-1 rounded bg-slate-200 px-1 py-0.5 text-[10px] font-semibold text-slate-600">
                                +CC
                              </span>
                            )}
                          </td>
                          <td className="max-w-[260px] truncate px-3 py-2.5 text-slate-800" title={item.subject}>
                            {item.subject}
                          </td>
                          <td className="whitespace-nowrap px-3 py-2.5">
                            <span
                              className={`inline-block rounded-full px-2 py-0.5 text-[10px] font-bold ${
                                item.status === 'BERHASIL'
                                  ? 'bg-emerald-100 text-emerald-700'
                                  : 'bg-rose-100 text-rose-700'
                              }`}
                            >
                              {item.status}
                            </span>
                          </td>
                          <td className="whitespace-nowrap px-3 py-2.5 text-center">
                            <button
                              type="button"
                              disabled={loadingDetail}
                              onClick={() => handleOpenDetail(item.id)}
                              className="rounded-md bg-blue-600 px-2.5 py-1 text-[11px] font-semibold text-white hover:bg-blue-500 disabled:opacity-50"
                            >
                              Lihat Detail
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}