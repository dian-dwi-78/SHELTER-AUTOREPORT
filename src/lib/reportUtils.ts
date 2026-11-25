import * as XLSX from 'xlsx';
import type { PhotoItem } from '@/types/report';

export const HARI_ID = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
export const BULAN_ID = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember',
];
export const BULAN_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

export function escapeHtml(str: unknown): string {
  const map: Record<string, string> = {
    '&': '\x26amp;',
    '<': '\x26lt;',
    '>': '\x26gt;',
    '"': '\x26quot;',
    "'": '\x26#39;',
  };
  return String(str == null ? '' : str).replace(/[&<>"']/g, (m) => map[m]);
}

export function stripEmoji(str: unknown): string {
  return String(str == null ? '' : str)
    .replace(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE00}-\u{FE0F}\u{200D}]/gu, '')
    .replace(/\s+/g, ' ')
    .trim();
}

export function toProperCase(str: unknown): string {
  const clean = stripEmoji(str);
  if (!clean) return '';
  const acronyms = ['PT', 'SOP', 'PHK', 'WIB', 'WITA', 'WIT', 'KTP', 'NIK', 'SPV', 'LOD', 'HRD', 'K3', 'APD', 'HO'];
  const letters = clean.replace(/[^a-zA-Z]/g, '');
  const uppers = clean.replace(/[^A-Z]/g, '');
  if (letters.length > 0 && uppers.length / letters.length < 0.7) {
    return clean;
  }
  return clean.toLowerCase().replace(/\b([a-z]+)\b/g, (word) => {
    const up = word.toUpperCase();
    if (acronyms.includes(up)) return up;
    return word.charAt(0).toUpperCase() + word.slice(1);
  });
}

export function fmtNum(val: number | null | undefined, dec = 1): string {
  if (val === null || val === undefined || isNaN(val)) return '-';
  return Number(val).toFixed(dec).replace('.', ',');
}

export function parseExcelDate(val: unknown): Date | null {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0);
  }
  const parsed = new Date(String(val).replace(' ', 'T'));
  return isNaN(parsed.getTime()) ? null : parsed;
}

export function dateKey(d: Date | null): string {
  if (!d) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

export function fmtDateFull(key: string): string {
  const [y, m, d] = key.split('-').map(Number);
  const dt = new Date(y, m - 1, d);
  return `${HARI_ID[dt.getDay()]}, ${d} ${BULAN_ID[m - 1]} ${y}`;
}

export function fmtDateShort(key: string): string {
  const [, m, d] = key.split('-').map(Number);
  return `${d} ${BULAN_SHORT[m - 1]}`;
}

export function fmtDateMed(key: string): string {
  const [y, m, d] = key.split('-').map(Number);
  return `${d} ${BULAN_ID[m - 1]} ${y}`;
}

export function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const res = String(e.target?.result || '');
      resolve(res.split(',')[1] || '');
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

export function compressImage(file: File, maxWidth = 1000, quality = 0.75): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (ev) => {
      const img = new Image();
      img.onload = () => {
        let { width: w, height: h } = img;
        if (w > maxWidth) {
          h = Math.round((h * maxWidth) / w);
          w = maxWidth;
        }
        const canvas = document.createElement('canvas');
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext('2d');
        ctx?.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.onerror = reject;
      img.src = String(ev.target?.result || '');
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// CSS inti untuk HTML ekspor & body email Gmail
export const REPORT_CORE_CSS = `
  body{margin:0;padding:0;background:#f4f7fb;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;color:#1e293b;line-height:1.55;}
  .wrapper{padding:24px 16px;}
  .container{max-width:860px;margin:0 auto;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #dbe5f0;box-shadow:0 6px 24px rgba(15,23,42,0.06);}
  .topbar{height:8px;background:linear-gradient(90deg,#0f3c78,#1d63b8,#3b82f6);}
  .inner{padding:32px 36px;}
  .eyebrow{font-size:11px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;color:#1d63b8;margin:0 0 6px;}
  h1{font-size:24px;line-height:1.25;color:#0f172a;margin:0 0 6px;font-weight:800;}
  h2{font-size:16px;color:#0f3c78;margin:28px 0 12px;padding-bottom:6px;border-bottom:2px solid #e2e8f0;}
  .site{font-size:14px;color:#64748b;margin:0 0 20px;font-weight:600;}
  .greeting{font-weight:700;margin:0 0 8px;}
  .metrics{width:100%;border-collapse:separate;border-spacing:10px;margin:18px -10px;}
  .metric{width:33.33%;border:1px solid #dbe5f0;border-radius:10px;padding:14px;background:#f8fafc;}
  .metric .label{font-size:11px;font-weight:700;text-transform:uppercase;color:#64748b;margin-bottom:4px;}
  .metric .num{font-size:24px;font-weight:800;margin-right:4px;}
  .metric.blue .num{color:#1d63b8;} .metric.green .num{color:#059669;} .metric.orange .num{color:#d97706;}
  .metric .unit{font-size:12px;color:#64748b;font-weight:600;}
  table.info,table.material,table.attendance{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:14px;}
  table.info th,table.info td,table.material th,table.material td,table.attendance th,table.attendance td{border:1px solid #dbe5f0;padding:8px 12px;text-align:left;vertical-align:top;}
  table.info th,table.material th{background:#f8fafc;color:#1e3a5f;font-weight:700;width:30%;}
  table.attendance th{background:#0f3c78;color:#ffffff;font-weight:700;}
  table.attendance td.no,table.attendance th.no{text-align:center;width:7%;}
  .summary-strip{background:#f0f7ff;border-left:4px solid #1d63b8;padding:12px 16px;border-radius:6px;font-size:13.5px;margin-bottom:12px;}
  ul.goals{margin:0 0 14px;padding-left:20px;font-size:13.5px;}
  ul.goals li{margin:4px 0;}
  .chartwrap{border:1px solid #dbe5f0;border-radius:10px;padding:14px 16px;background:#fafcff;margin-bottom:12px;}
  .chart-title{font-size:13px;font-weight:700;color:#0f3c78;margin-bottom:10px;}
  .barline{display:flex;align-items:center;gap:10px;margin:6px 0;font-size:12px;}
  .barline span:first-child{width:90px;color:#475569;font-weight:600;}
  .bar-track{flex:1;height:12px;background:#e2e8f0;border-radius:99px;overflow:hidden;}
  .bar-fill{display:block;height:100%;border-radius:99px;}
  .bar-fill.pre{background:#94a3b8;} .bar-fill.alt{background:#1d63b8;}
  .legend{font-size:11px;color:#64748b;margin-top:8px;}
  .kpi-small{width:100%;border-collapse:separate;border-spacing:8px;margin:8px -8px;}
  .kpi-small td{width:33.33%;border:1px solid #dbe5f0;border-radius:8px;padding:10px;text-align:center;font-size:12px;color:#64748b;background:#f8fafc;}
  .kpi-small td b{display:block;font-size:16px;color:#0f172a;margin-bottom:2px;}
  .insight{background:#f8fafc;border:1px solid #dbe5f0;border-left:4px solid #0f3c78;border-radius:6px;padding:10px 14px;font-size:13px;margin-bottom:8px;}
  .session-heading{display:flex;justify-content:space-between;align-items:center;margin:14px 0 6px;}
  .session-heading h3{margin:0;font-size:14px;color:#0f3c78;}
  .session-count{font-size:12px;font-weight:700;background:#e0f2fe;color:#0369a1;padding:2px 10px;border-radius:99px;}
  .sig{margin-top:20px;font-size:13.5px;}
  .footer{background:#f8fafc;border-top:1px solid #e2e8f0;padding:14px 36px;font-size:12px;color:#64748b;text-align:center;}
  p, .summary-strip, .insight, ul.goals li, table.material td { text-align: justify; }
  `;

export function buildCleanReportHtml(wrapperEl: HTMLElement, titleText: string): string {
  const clone = wrapperEl.cloneNode(true) as HTMLElement;
  clone.querySelectorAll('.no-export').forEach((el) => el.remove());

  return `<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Laporan Training ${escapeHtml(titleText)}</title>
<style>${REPORT_CORE_CSS}</style>
</head>
<body>
${clone.outerHTML}
</body>
</html>`;
}

export function buildEmailInlineHtml(wrapperEl: HTMLElement, photos: PhotoItem[]): string {
  const clone = wrapperEl.cloneNode(true) as HTMLElement;
  clone.querySelectorAll('.no-export').forEach((el) => el.remove());

  const photoGrid = clone.querySelector('#photoGrid');
  if (photoGrid && photos.length) {
    let tbl = '<table role="presentation" style="width:100%;border-collapse:separate;border-spacing:10px;">';
    for (let i = 0; i < photos.length; i += 2) {
      tbl += '<tr>';
      for (let j = i; j < i + 2; j++) {
        if (photos[j]) {
          tbl += `<td style="width:50%;vertical-align:top;border:1px solid #dbe5f0;border-radius:10px;padding:8px;background:#fff;">
            <img src="cid:foto_${j}" alt="${escapeHtml(photos[j].caption)}" style="width:100%;height:auto;border-radius:6px;display:block;margin-bottom:6px;"/>
            <div style="font-size:12px;font-weight:700;color:#253b58;">${escapeHtml(photos[j].caption)}</div>
          </td>`;
        } else {
          tbl += '<td style="width:50%;"></td>';
        }
      }
      tbl += '</tr>';
    }
    tbl += '</table>';
    photoGrid.outerHTML = tbl;
  }

  return `<!DOCTYPE html><html lang="id"><head><meta charset="utf-8"/><style>${REPORT_CORE_CSS}</style></head><body style="margin:0;padding:0;background:#f4f7fb;">${clone.outerHTML}</body></html>`;
}