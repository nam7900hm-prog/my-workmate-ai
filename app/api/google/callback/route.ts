import { NextRequest, NextResponse } from "next/server";
import { saveGoogleCookie, sealGoogleToken } from "@/lib/google-auth";

export async function GET(req: NextRequest) {
  const code = req.nextUrl.searchParams.get("code");
  const state = req.nextUrl.searchParams.get("state");
  if (!code || !state || state !== req.cookies.get("mw_google_state")?.value)
    return NextResponse.redirect(`${req.nextUrl.origin}/?google=failed`);
  const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: process.env.GOOGLE_CLIENT_ID || "",
      client_secret: process.env.GOOGLE_CLIENT_SECRET || "",
      redirect_uri: `${req.nextUrl.origin}/api/google/callback`,
      grant_type: "authorization_code",
    }),
  });
  if (!tokenResponse.ok)
    return NextResponse.redirect(`${req.nextUrl.origin}/?google=failed`);
  const token = await tokenResponse.json();
  const response = NextResponse.redirect(`${req.nextUrl.origin}/?google=connected`);
  saveGoogleCookie(
    response,
    await sealGoogleToken({
      accessToken: token.access_token,
      refreshToken: token.refresh_token,
      expiresAt: Date.now() + Number(token.expires_in || 3600) * 1000,
    }),
  );
  response.cookies.delete("mw_google_state");
  return response;
}
