import { NextRequest, NextResponse } from "next/server";
import { openAIUrl } from "@/lib/openai-config";
export const runtime = "nodejs";

export async function POST(req: NextRequest) {
  if (!process.env.OPENAI_API_KEY)
    return NextResponse.json(
      { error: "인공지능 연결이 필요합니다." },
      { status: 503 },
    );
  const form = await req.formData();
  if (form.get("consent") !== "true")
    return NextResponse.json(
      { error: "사용자 확인 전에는 미디어를 전송하지 않습니다." },
      { status: 400 },
    );
  const file = form.get("file");
  if (!(file instanceof File))
    return NextResponse.json({ error: "파일이 없습니다." }, { status: 400 });
  if (file.size > 20 * 1024 * 1024)
    return NextResponse.json(
      { error: "미디어 파일은 20MB 이하만 처리할 수 있습니다." },
      { status: 413 },
    );
  if (file.type.startsWith("audio/")) {
    const outgoing = new FormData();
    outgoing.append("file", file, file.name);
    outgoing.append(
      "model",
      process.env.OPENAI_TRANSCRIBE_MODEL || "gpt-4o-mini-transcribe",
    );
    const response = await fetch(openAIUrl("audio/transcriptions"), {
      method: "POST",
      headers: { Authorization: `Bearer ${process.env.OPENAI_API_KEY}` },
      body: outgoing,
    });
    if (!response.ok)
      return NextResponse.json(
        { error: "음성 내용을 글로 바꾸지 못했습니다." },
        { status: 502 },
      );
    const data = await response.json();
    return NextResponse.json({ kind: "audio", text: data.text || "" });
  }
  if (file.type.startsWith("image/")) {
    const bytes = Buffer.from(await file.arrayBuffer()).toString("base64");
    const response = await fetch(openAIUrl("responses"), {
      method: "POST",
      headers: {
        Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL || "gpt-5-mini",
        store: false,
        input: [
          {
            role: "user",
            content: [
              {
                type: "input_text",
                text: "이 이미지의 글자와 표 구조를 빠짐없이 추출해 한국어 텍스트로 반환해 주세요. 보이지 않는 내용은 만들지 마세요.",
              },
              {
                type: "input_image",
                image_url: `data:${file.type};base64,${bytes}`,
              },
            ],
          },
        ],
      }),
    });
    if (!response.ok)
      return NextResponse.json(
        { error: "사진 내용을 분석하지 못했습니다." },
        { status: 502 },
      );
    const data = await response.json();
    return NextResponse.json({ kind: "image", text: data.output_text || "" });
  }
  return NextResponse.json(
    { error: "지원하지 않는 미디어 형식입니다." },
    { status: 415 },
  );
}
