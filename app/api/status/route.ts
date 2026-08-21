import { NextResponse } from "next/server";

export async function GET() {
  return NextResponse.json({ ready: Boolean(process.env.OPENAI_API_KEY) });
}
