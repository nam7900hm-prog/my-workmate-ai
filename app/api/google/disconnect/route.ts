import { NextResponse } from "next/server";
import { GOOGLE_COOKIE } from "@/lib/google-auth";

export async function POST() {
  const response = NextResponse.json({ disconnected: true });
  response.cookies.delete(GOOGLE_COOKIE);
  return response;
}
