import { NextRequest, NextResponse } from 'next/server';
import MailComposer from 'nodemailer/lib/mail-composer';
import { auth } from '@/auth';
import type { SendEmailPayload } from '@/types/report';

function normalizeEmails(raw?: string): string {
  if (!raw) return '';
  return raw
    .split(/[,;\n]+/)
    .map((e) => e.trim())
    .filter(Boolean)
    .join(', ');
}

export async function POST(req: NextRequest) {
  try {
    const session = (await auth()) as {
      user?: { email?: string | null; name?: string | null };
      accessToken?: string;
    } | null;

    if (!session?.accessToken || !session?.user?.email) {
      return NextResponse.json(
        { error: 'Silakan login dengan akun Google terlebih dahulu sebelum mengirim email.' },
        { status: 401 }
      );
    }

    const payload: SendEmailPayload = await req.json();

    const toList = normalizeEmails(payload.to);
    const ccList = normalizeEmails(payload.cc);
    const bccList = normalizeEmails(payload.bcc);

    if (!toList) {
      return NextResponse.json({ error: 'Alamat email tujuan (To) wajib diisi.' }, { status: 400 });
    }
    if (!payload.subject?.trim()) {
      return NextResponse.json({ error: 'Subjek email wajib diisi.' }, { status: 400 });
    }

    const attachments: Array<{
      filename: string;
      content: Buffer | string;
      contentType: string;
      cid?: string;
    }> = [];

    if (Array.isArray(payload.inlinePhotos)) {
      payload.inlinePhotos.forEach((item) => {
        if (item.cid && item.base64) {
          attachments.push({
            filename: `${item.cid}.jpg`,
            content: Buffer.from(item.base64, 'base64'),
            contentType: 'image/jpeg',
            cid: item.cid,
          });
        }
      });
    }

    if (payload.attachHtml && payload.htmlContent) {
      attachments.push({
        filename: payload.htmlFileName || 'Laporan_Training.html',
        content: payload.htmlContent,
        contentType: 'text/html',
      });
    }

    if (Array.isArray(payload.extraAttachments)) {
      payload.extraAttachments.forEach((item) => {
        if (item.base64 && item.name) {
          attachments.push({
            filename: item.name,
            content: Buffer.from(item.base64, 'base64'),
            contentType: item.mimeType || 'application/octet-stream',
          });
        }
      });
    }

    const senderName = payload.senderName?.trim() || session.user.name || 'PT Shelter Indonesia Gemilang';

    const mail = new MailComposer({
      from: `"${senderName}" <${session.user.email}>`,
      to: toList,
      cc: ccList || undefined,
      bcc: bccList || undefined,
      replyTo: payload.replyTo?.trim() || undefined,
      subject: payload.subject.trim(),
      text: payload.body || 'Laporan Giat Training terlampir.',
      html: payload.emailHtmlBody || undefined,
      attachments,
    });

    const rawMessage: Buffer = await new Promise((resolve, reject) => {
      mail.compile().build((err, message) => {
        if (err) reject(err);
        else resolve(message);
      });
    });

    const encodedMessage = rawMessage
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');

    const gmailRes = await fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${session.accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ raw: encodedMessage }),
    });

    if (!gmailRes.ok) {
      const errText = await gmailRes.text();
      return NextResponse.json(
        { error: `Gmail API Error (${gmailRes.status}): ${errText}` },
        { status: gmailRes.status }
      );
    }

    return NextResponse.json({
      success: true,
      message: `Email berhasil dikirim dari ${session.user.email} ke ${toList}.`,
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'Gagal mengirim email.';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}