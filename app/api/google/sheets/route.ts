import { NextRequest, NextResponse } from "next/server";
import { googleAccess, saveGoogleCookie } from "@/lib/google-auth";

export async function GET(req: NextRequest) {
  try {
    const auth = await googleAccess(req);
    const url = new URL("https://www.googleapis.com/drive/v3/files");
    url.search = new URLSearchParams({
      q: "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
      fields: "files(id,name,modifiedTime)",
      orderBy: "modifiedTime desc",
      pageSize: "100",
    }).toString();
    const google = await fetch(url, {
      headers: { Authorization: `Bearer ${auth.accessToken}` },
    });
    if (!google.ok) throw new Error("Google Sheets 목록을 불러오지 못했습니다.");
    const data = await google.json();
    const response = NextResponse.json({ files: data.files || [] });
    if (auth.cookie) saveGoogleCookie(response, auth.cookie);
    return response;
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Google 연결 오류" },
      { status: 401 },
    );
  }
}
