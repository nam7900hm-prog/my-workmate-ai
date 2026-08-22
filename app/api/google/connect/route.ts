import { NextRequest, NextResponse } from "next/server";

export async function GET(req: NextRequest) {
  if (!process.env.GOOGLE_CLIENT_ID || !process.env.GOOGLE_CLIENT_SECRET)
    return NextResponse.json(
      { error: "Google Sheets 연결 설정이 아직 완료되지 않았습니다." },
      { status: 503 },
    );
  const state = crypto.randomUUID();
  const redirectUri =
    process.env.GOOGLE_REDIRECT_URI ||
    `${req.nextUrl.origin}/api/google/callback`;
  const url = new URL("https://accounts.google.com/o/oauth2/v2/auth");
  url.search = new URLSearchParams({
    client_id: process.env.GOOGLE_CLIENT_ID,
    redirect_uri: redirectUri,
    response_type: "code",
    scope: "https://www.googleapis.com/auth/drive.readonly",
    access_type: "offline",
    prompt: "consent",
    state,
  }).toString();
  const response = NextResponse.redirect(url);
  response.cookies.set("mw_google_state", state, {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    path: "/",
    maxAge: 600,
  });
  return response;
}
