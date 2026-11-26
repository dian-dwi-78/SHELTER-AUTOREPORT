import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/auth";
import { prisma } from "@/lib/prisma";

export async function GET(
  _req: NextRequest,
  context: { params: Promise<{ id: string }> }
) {
  const session = await auth();
  if (!session?.user?.email) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  try {
    const { id } = await context.params;
    const numericId = Number(id);

    if (Number.isNaN(numericId)) {
      return NextResponse.json({ error: "ID tidak valid" }, { status: 400 });
    }

    const detail = await prisma.emailHistory.findUnique({
      where: { id: numericId },
    });

    if (!detail) {
      return NextResponse.json(
        { error: "Riwayat tidak ditemukan" },
        { status: 404 }
      );
    }

    return NextResponse.json({ detail });
  } catch (error: any) {
    console.error("Error fetching history detail:", error);
    return NextResponse.json(
      { error: "Gagal memuat detail riwayat." },
      { status: 500 }
    );
  }
}