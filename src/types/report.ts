export interface ReportMeta {
  entitas: string;
  site: string;
  namaTraining: string;
  targetPeserta: string;
  metode: string;
  pemateri: string;
  waktuSesi: string;
  ttd: string;
  tujuan: string;
}

export interface PptSlide {
  title: string;
  body: string[];
}

export interface PhotoItem {
  src: string; // Data URL base64 terkompresi
  caption: string;
}

export interface RawFileAttachment {
  name: string;
  mimeType: string;
  base64: string;
}

export interface AiSubmateri {
  topik: string;
  format: 'paragraf' | 'poin';
  uraianParagraf: string;
  uraianPoin: string[];
}

export interface AiAnalisis {
  judul: string;
  isi: string;
}

export interface AiNarrativeResponse {
  namaTraining: string;
  targetPeserta: string;
  perusahaanSite: string;
  ringkasanFokus: string;
  tujuan: string[];
  submateri: AiSubmateri[];
  analisis: AiAnalisis[];
  kesimpulan: string;
}

export interface SendEmailPayload {
  to: string;
  cc?: string;
  bcc?: string;
  senderName?: string;
  replyTo?: string;
  subject: string;
  body: string;
  emailHtmlBody: string;
  inlinePhotos: { cid: string; base64: string }[];
  attachHtml: boolean;
  htmlFileName?: string;
  htmlContent?: string;
  extraAttachments?: RawFileAttachment[];
}