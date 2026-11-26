import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { prisma } from "@/lib/prisma";

export async function GET(req: NextRequest) {
  const session = await auth();
  if (!session?.user?.email) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  try {
    const { searchParams } = new URL(req.url);
    const q = searchParams.get("q")?.trim() || "";
    const sender = searchParams.get("sender")?.trim() || "";
    const status = searchParams.get("status")?.trim() || "";
    const dateFrom = searchParams.get("dateFrom")?.trim() || "";
    const dateTo = searchParams.get("dateTo")?.trim() || "";

    const where: any = {};

    if (q) {
      where.OR = [
        { subject: { contains: q } },
        { toRecipients: { contains: q } },
        { ccRecipients: { contains: q } },
        { trainingName: { contains: q } },
        { siteName: { contains: q } },
        { senderEmail: { contains: q } },
      ];
    }

    if (sender) {
      where.senderEmail = sender;
    }

    if (status && (status === "BERHASIL" || status === "GAGAL")) {
      where.status = status;
    }

    if (dateFrom || dateTo) {
      where.sentAt = {};
      if (dateFrom) {
        where.sentAt.gte = new Date(`${dateFrom}T00:00:00.000Z`);
      }
      if (dateTo) {
        where.sentAt.lte = new Date(`${dateTo}T23:59:59.999Z`);
      }
    }

    const [items, uniqueSenders] = await Promise.all([
      prisma.emailHistory.findMany({
        where,
        orderBy: { sentAt: "desc" },
        select: {
          id: true,
          sentAt: true,
          senderEmail: true,
          toRecipients: true,
          ccRecipients: true,
          subject: true,
          trainingName: true,
          siteName: true,
          attachments: true,
          status: true,
          errorMessage: true,
        },
        take: 200,
      }),
      prisma.emailHistory.findMany({
        distinct: ["senderEmail"],
        select: { senderEmail: true },
        orderBy: { senderEmail: "asc" },
      }),
    ]);

    return NextResponse.json({
      items,
      senders: uniqueSenders.map((s) => s.senderEmail),
    });
  } catch (error: any) {
    console.error("Error fetching history:", error);
    return NextResponse.json(
      { error: "Gagal memuat riwayat pengiriman." },
      { status: 500 }
    );
  }
}