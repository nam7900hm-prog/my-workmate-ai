import { NextRequest, NextResponse } from "next/server";
import { openAIUrl } from "@/lib/openai-config";
import { extractOpenAIText } from "@/lib/openai-response";
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
    outgoing.append(
      "prompt",
      "한국어 회의 녹음입니다. 날짜 시간 안건 참석자 결정사항 담당자 기한을 정확히 받아쓰고 들리지 않는 내용은 추측하지 마세요.",
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
    const text = extractOpenAIText(data);
    if (!text.trim())
      return NextResponse.json(
        { error: "사진 분석 결과가 비어 있습니다." },
        { status: 502 },
      );
    return NextResponse.json({ kind: "image", text });
  }
  if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
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
                text: "이 PDF의 보이는 글자와 표를 페이지 순서대로 추출하세요. 학생 활동지라면 학생 이름과 해당 학생이 실제로 쓴 내용만 분리하세요. 알아볼 수 없는 글자는 추측하지 말고 [판독 불가]로 표시하세요. JSON만 반환하세요. 형식은 {\"text\":\"전체 추출문\",\"students\":[{\"name\":\"학생 이름\",\"text\":\"해당 학생의 실제 글\"}],\"warnings\":[\"확인할 내용\"]}입니다.",
              },
              {
                type: "input_file",
                filename: file.name,
                file_data: `data:application/pdf;base64,${bytes}`,
              },
            ],
          },
        ],
        text: { format: { type: "json_object" } },
      }),
    });
    if (!response.ok)
      return NextResponse.json(
        { error: "PDF의 스캔 글자를 분석하지 못했습니다." },
        { status: 502 },
      );
    const raw = extractOpenAIText(await response.json());
    let parsed: Record<string, unknown>;
    try {
      parsed = JSON.parse(raw);
    } catch {
      return NextResponse.json(
        { error: "PDF 분석 결과의 형식이 올바르지 않습니다." },
        { status: 502 },
      );
    }
    const students = Array.isArray(parsed.students)
      ? parsed.students
          .filter(
            (item): item is { name: string; text: string } =>
              Boolean(item) &&
              typeof item === "object" &&
              typeof (item as Record<string, unknown>).name === "string" &&
              typeof (item as Record<string, unknown>).text === "string",
          )
          .map((item) => ({ name: item.name.trim(), text: item.text.trim() }))
          .filter((item) => item.name && item.text)
          .slice(0, 500)
      : [];
    const text = typeof parsed.text === "string" ? parsed.text.trim() : "";
    if (!text && !students.length)
      return NextResponse.json(
        { error: "PDF에서 확인할 수 있는 글자를 찾지 못했습니다." },
        { status: 422 },
      );
    return NextResponse.json({
      kind: "pdf",
      text,
      students,
      warnings: Array.isArray(parsed.warnings)
        ? parsed.warnings.filter((item): item is string => typeof item === "string")
        : [],
    });
  }
  return NextResponse.json(
    { error: "지원하지 않는 미디어 형식입니다." },
    { status: 415 },
  );
}
