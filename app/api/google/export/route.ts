import { NextRequest, NextResponse } from "next/server";
import { googleAccess, saveGoogleCookie } from "@/lib/google-auth";

export async function GET(req: NextRequest) {
  try {
    const id = req.nextUrl.searchParams.get("id") || "";
    if (!/^[A-Za-z0-9_-]{10,}$/.test(id))
      return NextResponse.json({ error: "시트 번호가 올바르지 않습니다." }, { status: 400 });
    const auth = await googleAccess(req);
    const google = await fetch(
      `https://www.googleapis.com/drive/v3/files/${id}/export?mimeType=${encodeURIComponent("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}`,
      { headers: { Authorization: `Bearer ${auth.accessToken}` } },
    );
    if (!google.ok) throw new Error("선택한 Google Sheets를 가져오지 못했습니다.");
    const response = new NextResponse(await google.arrayBuffer(), {
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "cache-control": "no-store",
      },
    });
    if (auth.cookie) saveGoogleCookie(response, auth.cookie);
    return response;
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Google 가져오기 오류" },
      { status: 401 },
    );
  }
}
