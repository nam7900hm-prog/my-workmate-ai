import { NextRequest, NextResponse } from "next/server";

export const GOOGLE_COOKIE = "mw_google_token";

type GoogleToken = {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
};

function secret() {
  const value = process.env.GOOGLE_TOKEN_SECRET || "";
  if (value.length < 32) throw new Error("GOOGLE_TOKEN_SECRET 설정이 필요합니다.");
  return value;
}

async function key() {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(secret()),
  );
  return crypto.subtle.importKey("raw", digest, "AES-GCM", false, [
    "encrypt",
    "decrypt",
  ]);
}

function encode(bytes: Uint8Array) {
  return Buffer.from(bytes).toString("base64url");
}

function decode(value: string) {
  return new Uint8Array(Buffer.from(value, "base64url"));
}

export async function sealGoogleToken(token: GoogleToken) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const encrypted = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv },
    await key(),
    new TextEncoder().encode(JSON.stringify(token)),
  );
  return `${encode(iv)}.${encode(new Uint8Array(encrypted))}`;
}

async function openGoogleToken(value: string): Promise<GoogleToken> {
  const [iv, encrypted] = value.split(".");
  if (!iv || !encrypted) throw new Error("Google 연결 정보가 올바르지 않습니다.");
  const plain = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: decode(iv) },
    await key(),
    decode(encrypted),
  );
  return JSON.parse(new TextDecoder().decode(plain));
}

export async function googleAccess(req: NextRequest) {
  const saved = req.cookies.get(GOOGLE_COOKIE)?.value;
  if (!saved) throw new Error("Google 계정을 먼저 연결해 주세요.");
  const token = await openGoogleToken(saved);
  if (token.expiresAt > Date.now() + 60_000)
    return { accessToken: token.accessToken };
  if (!token.refreshToken) throw new Error("Google 계정을 다시 연결해 주세요.");
  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.GOOGLE_CLIENT_ID || "",
      client_secret: process.env.GOOGLE_CLIENT_SECRET || "",
      refresh_token: token.refreshToken,
      grant_type: "refresh_token",
    }),
  });
  if (!response.ok) throw new Error("Google 연결을 갱신하지 못했습니다.");
  const fresh = await response.json();
  const updated: GoogleToken = {
    accessToken: fresh.access_token,
    refreshToken: token.refreshToken,
    expiresAt: Date.now() + Number(fresh.expires_in || 3600) * 1000,
  };
  return { accessToken: updated.accessToken, cookie: await sealGoogleToken(updated) };
}

export function saveGoogleCookie(response: NextResponse, value: string) {
  response.cookies.set(GOOGLE_COOKIE, value, {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    path: "/",
    maxAge: 60 * 60 * 24 * 30,
  });
}
