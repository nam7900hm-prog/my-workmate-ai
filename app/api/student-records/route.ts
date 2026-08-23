import { NextRequest, NextResponse } from "next/server";
import { openAIError, parseOpenAIJson } from "@/lib/openai-response";
import { normalizeStudentSentence, studentDraftProblems } from "@/lib/student-text";
import { openAIUrl } from "@/lib/openai-config";
import {
  studentRecordPrompt,
  type StudentRecordKind,
} from "@/lib/student-record-guidelines";

export async function POST(req: NextRequest) {
  const body = await req.json();
  if (!body.consent)
    return NextResponse.json(
      { error: "사용자 확인 전에는 학생자료를 전송하지 않습니다." },
      { status: 400 },
    );
  if (!process.env.OPENAI_API_KEY)
    return NextResponse.json(
      { error: "AI 설정 후 학생별 세 가지 문안을 만들 수 있습니다." },
      { status: 503 },
    );
  const students = Array.isArray(body.students)
    ? body.students.slice(0, 200)
    : [];
  if (!students.length)
    return NextResponse.json(
      { error: "이름과 소감문이 들어 있는 학생 자료를 먼저 선택해 주세요." },
      { status: 400 },
    );
  const avoidPhrases = Array.isArray(body.avoidPhrases)
    ? body.avoidPhrases
        .filter((value: unknown) => typeof value === "string")
        .slice(-80)
    : [];
  const recordType = ["subject", "behavior", "activity", "general"].includes(body.recordType)
    ? (body.recordType as StudentRecordKind)
    : "general";
  const requestedLength = String(body.request || "").match(/(\d{2,3})\s*자/);
  const targetLength = requestedLength
    ? Math.max(200, Math.min(500, Number(requestedLength[1])))
    : null;
  const prompt = studentRecordPrompt({
    kind: recordType,
    request: String(body.request || ""),
    students,
    avoidDrafts: avoidPhrases,
    targetLength,
  });
  const response = await fetch(openAIUrl("responses"), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: process.env.OPENAI_MODEL || "gpt-5-mini",
      input: prompt,
      text: { format: { type: "json_object" } },
      store: false,
    }),
  });
  if (!response.ok)
    return NextResponse.json(
      {
        error: await openAIError(response, "학생별 초안 생성에 실패했습니다."),
      },
      { status: 502 },
    );
  const data = await response.json();
  let parsed: { students?: any[] };
  try {
    parsed = parseOpenAIJson(data, "학생별 초안");
  } catch (error) {
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "학생별 초안을 읽지 못했습니다.",
      },
      { status: 502 },
    );
  }
  const drafts = (parsed.students || [])
    .slice(0, students.length)
    .map((x: any) => ({
      ...x,
      factDraft: normalizeStudentSentence(x.factDraft),
      inferredDraft: normalizeStudentSentence(x.inferredDraft),
      recommendedDraft: normalizeStudentSentence(x.recommendedDraft),
      inferredParts: Array.isArray(x.inferredParts)
        ? x.inferredParts.map((value: unknown) => String(value))
        : [],
      recommendedInferredParts: Array.isArray(x.recommendedInferredParts)
        ? x.recommendedInferredParts.map((value: unknown) => String(value))
        : [],
      finalText: "",
      selected: undefined,
      reviewed: false,
    }));
  const invalid = drafts.flatMap((draft: any) =>
    [draft.factDraft, draft.inferredDraft, draft.recommendedDraft].flatMap(
      (value, index) =>
        studentDraftProblems(value).map(
          (problem) => `${draft.name} ${index + 1}안 ${problem}`,
        ),
    ),
  );
  if (invalid.length)
    return NextResponse.json(
      { error: `작성 기준을 통과하지 못했습니다. ${invalid.slice(0, 3).join(" ")}` },
      { status: 502 },
    );
  return NextResponse.json({ drafts });
}
