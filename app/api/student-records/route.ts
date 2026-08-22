import { NextRequest, NextResponse } from "next/server";
import { openAIError, parseOpenAIJson } from "@/lib/openai-response";
import { normalizeStudentSentence } from "@/lib/student-text";
import { openAIUrl } from "@/lib/openai-config";

const COMMON_RULES = `학생 생활기록부 초안 작성 규칙:
1. 입력된 학생 글에 없는 성과와 행동과 결과는 사실처럼 만들지 않는다.
2. 쉼표 문자를 절대 사용하지 않는다. 모든 완결 문장은 마침표로 끝낸다.
3. 영어 고유명사와 기술 용어는 가능한 범위에서 자연스러운 한글 의미로 풀어 쓴다.
4. 학생별 동기와 선택과 행동과 성취와 변화와 후속 관심 중 자료에서 확인되는 요소만 사용한다.
5. 우수함과 탁월함과 성실함 같은 추상적 칭찬을 근거 없이 사용하지 않는다.
6. 학생 수준을 벗어난 전문 지식이나 연구를 만들어내지 않는다.
7. 학생 소감문은 보조 근거이며 교사가 기록한 관찰 내용과 실제 수업 활동과 결과물을 우선한다.
8. 1안은 교사가 확인할 수 있는 관찰과 활동 근거 중심으로 작성한다.
9. 2안은 확인된 근거에서 이어지는 변화와 발전 가능성을 강조한다. 유추한 구절은 inferredParts에 그대로 적는다.
10. 3안은 1안과 2안을 단순 연결하지 말고 관찰 근거와 성장 내용을 한 사람의 기록처럼 자연스럽게 다시 작성한다. 유추한 구절은 recommendedInferredParts에 적는다.
11. 세 초안은 시작 방식과 문장 구조와 마무리 의미가 달라야 한다.
12. 학생 사이의 문장과 의미와 종결 표현이 반복되지 않도록 한다.
13. 각 초안은 한글 기준 반드시 200자 이상 500자 이하이며 1500바이트를 넘지 않게 작성한다.
14. 입력이 짧아도 확인되는 주제와 표현과 관심을 바탕으로 태도와 발전 가능성을 긍정적으로 연결해 200자 이상 작성한다. 다만 입력에 없는 활동과 성과를 실제 사실처럼 만들지 않는다.
15. 조사와 문장 연결을 다시 읽어 어색한 표현을 고치고 같은 종결어미를 기계적으로 반복하지 않는다.`;

const TYPE_RULES = {
  subject:
    "교과 세부능력 및 특기사항이다. 교과 성취기준과 수업 참여와 탐구 발표 수행평가 문제 해결 과정과 자기주도적 변화와 성장을 중심으로 쓴다. 생활 전반의 성격 평가는 쓰지 않는다.",
  behavior:
    "행동특성 및 종합의견이다. 학년 동안 교사가 지속적으로 관찰한 책임감 협력 배려 의사소통 역할 수행 생활 태도와 행동 변화를 종합한다. 한 번의 수업 성취만으로 성격 전체를 단정하지 않는다.",
  general:
    "일반 생활기록부 초안이다. 사용자의 요청과 입력 자료에서 확인되는 기록 영역을 따르되 교사 관찰 근거를 중심으로 작성한다.",
};

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
  const recordType = ["subject", "behavior", "general"].includes(body.recordType)
    ? (body.recordType as keyof typeof TYPE_RULES)
    : "general";
  const requestedLength = String(body.request || "").match(/(\d{2,3})\s*자/);
  const targetLength = requestedLength
    ? Math.max(200, Math.min(500, Number(requestedLength[1])))
    : null;
  const lengthGuide = targetLength
    ? `사용자가 요청한 약 ${targetLength}자를 우선한다. 공백 포함 ${Math.floor(targetLength * 0.9)}자 이상 ${Math.ceil(targetLength * 1.1)}자 이하를 목표로 하되 근거 없는 내용을 보태 글자 수를 채우지 않는다.`
    : "사용자가 별도 길이를 정하지 않았으므로 각 초안은 200자 이상 500자 이하로 작성한다.";
  const prompt = `당신은 한국 학교 교사를 돕는 생활기록부 초안 작성 비서다. 작성 종류는 ${recordType}이다. ${TYPE_RULES[recordType]} ${COMMON_RULES}
${lengthGuide}
사용자 요청=${String(body.request || "").slice(0, 4000)}
앞 묶음에서 이미 사용한 다음 표현과 문장 구조를 반복하지 않는다: ${JSON.stringify(avoidPhrases).slice(0, 30000)}
JSON만 반환한다. 형식은 {"students":[{"name":"학생명","source":"원문","factDraft":"1안 관찰 중심","inferredDraft":"2안 성장 중심","recommendedDraft":"3안 AI 추천","inferredParts":["2안에서 교사가 확인할 유추 구절"],"recommendedInferredParts":["3안에서 교사가 확인할 유추 구절"]}]}이다. 학생은 입력 순서와 이름을 그대로 유지한다. 입력=${JSON.stringify(students).slice(0, 90000)}`;
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
      finalText: normalizeStudentSentence(x.recommendedDraft),
      selected: "recommended",
      reviewed: false,
    }));
  const invalidLength = drafts.find((draft: any) => {
    const values = [draft.factDraft, draft.inferredDraft, draft.recommendedDraft];
    return values.some((value) => {
      const characters = Array.from(String(value).trim()).length;
      const bytes = new TextEncoder().encode(String(value).trim()).length;
      return characters < 200 || characters > 500 || bytes > 1500;
    });
  });
  if (invalidLength)
    return NextResponse.json(
      {
        error:
          "학생별 세 초안은 각각 200자 이상 500자 이하로 다시 작성해야 합니다.",
      },
      { status: 502 },
    );
  return NextResponse.json({ drafts });
}
