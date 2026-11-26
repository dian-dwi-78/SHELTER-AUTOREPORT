import { NextRequest, NextResponse } from 'next/server';
import MailComposer from 'nodemailer/lib/mail-composer';
import { auth } from '@/auth';
import { prisma } from '@/lib/prisma';
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
  let senderEmail = '';
  let toList = '';
  let ccList = '';
  let subjectStr = '(Tanpa Subjek)';
  let htmlForPreview = '';
  let trainingName: string | null = null;
  let siteName: string | null = null;
  const attachmentNames: string[] = [];

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

    senderEmail = session.user.email;

    const payload = (await req.json()) as SendEmailPayload & {
      trainingName?: string;
      siteName?: string;
    };

    toList = normalizeEmails(payload.to);
    ccList = normalizeEmails(payload.cc);
    const bccList = normalizeEmails(payload.bcc);
    subjectStr = payload.subject?.trim() || '';
    htmlForPreview = payload.htmlContent || payload.emailHtmlBody || '';
    trainingName = payload.trainingName?.trim() || null;
    siteName = payload.siteName?.trim() || null;

    if (!toList) {
      return NextResponse.json({ error: 'Alamat email tujuan (To) wajib diisi.' }, { status: 400 });
    }
    if (!subjectStr) {
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
      const htmlName = payload.htmlFileName || 'Laporan_Training.html';
      attachments.push({
        filename: htmlName,
        content: payload.htmlContent,
        contentType: 'text/html',
      });
      attachmentNames.push(htmlName);
    }

    if (Array.isArray(payload.extraAttachments)) {
      payload.extraAttachments.forEach((item) => {
        if (item.base64 && item.name) {
          attachments.push({
            filename: item.name,
            content: Buffer.from(item.base64, 'base64'),
            contentType: item.mimeType || 'application/octet-stream',
          });
          attachmentNames.push(item.name);
        }
      });
    }

    const senderName =
      payload.senderName?.trim() || session.user.name || 'PT Shelter Indonesia Gemilang';

    const mail = new MailComposer({
      from: `"${senderName}" <${senderEmail}>`,
      to: toList,
      cc: ccList || undefined,
      bcc: bccList || undefined,
      replyTo: payload.replyTo?.trim() || undefined,
      subject: subjectStr,
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
      const errMsg = `Gmail API Error (${gmailRes.status}): ${errText}`;

      await prisma.emailHistory
        .create({
          data: {
            senderEmail,
            toRecipients: toList,
            ccRecipients: ccList || null,
            subject: subjectStr,
            trainingName,
            siteName,
            htmlBody: htmlForPreview,
            attachments: JSON.stringify(attachmentNames),
            status: 'GAGAL',
            errorMessage: errMsg,
          },
        })
        .catch((dbErr) => console.error('Gagal simpan log error ke DB:', dbErr));

      return NextResponse.json({ error: errMsg }, { status: gmailRes.status });
    }

    const savedHistory = await prisma.emailHistory.create({
      data: {
        senderEmail,
        toRecipients: toList,
        ccRecipients: ccList || null,
        subject: subjectStr,
        trainingName,
        siteName,
        htmlBody: htmlForPreview,
        attachments: JSON.stringify(attachmentNames),
        status: 'BERHASIL',
      },
    });

    return NextResponse.json({
      success: true,
      historyId: savedHistory.id,
      message: `Email berhasil dikirim dari ${senderEmail} ke ${toList}.`,
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'Gagal mengirim email.';

    if (senderEmail && toList) {
      await prisma.emailHistory
        .create({
          data: {
            senderEmail,
            toRecipients: toList,
            ccRecipients: ccList || null,
            subject: subjectStr || '(Tanpa Subjek)',
            trainingName,
            siteName,
            htmlBody: htmlForPreview,
            attachments: JSON.stringify(attachmentNames),
            status: 'GAGAL',
            errorMessage: message,
          },
        })
        .catch((dbErr) => console.error('Gagal simpan log error ke DB:', dbErr));
    }

    return NextResponse.json({ error: message }, { status: 500 });
  }
}