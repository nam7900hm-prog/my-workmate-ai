import { NextRequest, NextResponse } from "next/server";
import { openAIError, parseOpenAIJson } from "@/lib/openai-response";
import { normalizeStudentSentence } from "@/lib/student-text";
import { openAIUrl } from "@/lib/openai-config";

const RULES = `학생 생활기록부 초안 작성 규칙:
1. 입력된 학생 글에 없는 성과와 행동과 결과는 사실처럼 만들지 않는다.
2. 쉼표 문자를 절대 사용하지 않는다. 모든 완결 문장은 마침표로 끝낸다.
3. 영어 고유명사와 기술 용어는 가능한 범위에서 자연스러운 한글 의미로 풀어 쓴다.
4. 학생별 동기와 선택과 행동과 성취와 변화와 후속 관심 중 자료에서 확인되는 요소만 사용한다.
5. 우수함과 탁월함과 성실함 같은 추상적 칭찬을 근거 없이 사용하지 않는다.
6. 학생 수준을 벗어난 전문 지식이나 연구를 만들어내지 않는다.
7. 사실 중심 초안은 학생 글에서 직접 확인되는 내용만 교사 서술형으로 바꾼다.
8. 관찰 가능성 초안은 학생 글에서 자연스럽게 이어지는 행동만 조심스럽게 제안한다. 유추한 구절은 inferredParts에 그대로 적는다.
9. 두 초안은 시작 방식과 문장 구조와 마무리 의미가 달라야 한다.
10. 학생 사이의 문장과 의미와 종결 표현이 반복되지 않도록 한다.`;

export async function POST(req: NextRequest) {
  const body = await req.json();
  if (!body.consent)
    return NextResponse.json(
      { error: "사용자 확인 전에는 학생자료를 전송하지 않습니다." },
      { status: 400 },
    );
  if (!process.env.OPENAI_API_KEY)
    return NextResponse.json(
      { error: "AI 설정 후 학생별 두 가지 초안을 만들 수 있습니다." },
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
  const prompt = `당신은 한국 고등학교 교사를 돕는 생활기록부 초안 작성 비서다. ${RULES}
앞 묶음에서 이미 사용한 다음 표현과 문장 구조를 반복하지 않는다: ${JSON.stringify(avoidPhrases).slice(0, 30000)}
JSON만 반환한다. 형식은 {"students":[{"name":"학생명","source":"원문","factDraft":"사실 중심 초안","inferredDraft":"관찰 가능성 포함 초안","inferredParts":["교사가 확인할 유추 구절"]}]}이다. 학생은 입력 순서와 이름을 그대로 유지한다. 입력=${JSON.stringify(students).slice(0, 90000)}`;
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
      inferredParts: Array.isArray(x.inferredParts)
        ? x.inferredParts.map((value: unknown) => String(value))
        : [],
      finalText: normalizeStudentSentence(x.factDraft),
      selected: "fact",
      reviewed: false,
    }));
  return NextResponse.json({ drafts });
}
