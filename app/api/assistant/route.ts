import { NextRequest, NextResponse } from "next/server";
import { openAIError, parseOpenAIJson } from "@/lib/openai-response";
import { normalizePlan } from "@/lib/openai-schema";
import { openAIUrl } from "@/lib/openai-config";
export async function POST(req: NextRequest) {
  const body = await req.json();
  if (!body.consent)
    return NextResponse.json(
      { error: "사용자 확인 전에는 자료를 전송하지 않습니다." },
      { status: 400 },
    );
  if (!process.env.OPENAI_API_KEY)
    return NextResponse.json(
      { error: "Vercel 설정에서 인공지능 연결이 필요합니다." },
      { status: 503 },
    );
  const prompt = `당신은 한국어 범용 개인 업무비서다. 파일 형식에 치우치지 말고 모든 선택 자료의 실제 추출 내용과 사용자 요청을 함께 검토한다. 원본을 변경하지 않는다. 자료에 없는 사실은 만들지 않는다. 필요한 질문과 실행계획과 검증방법을 제시한다. 특정 전문 규칙은 사용자가 해당 작업을 요청했을 때만 적용한다. Excel 작업이면 검색과 추출과 정렬과 필터와 중복검사와 함수 적용과 교체 범위 및 결과 검증을 구체적으로 단계에 적는다. JSON만 반환한다. 키: understanding 문자열 materials 문자열배열 template 문자열 steps 문자열배열 resultFormat 문자열 questions 문자열배열 canExecute 불리언 limitation 문자열. 입력=${JSON.stringify(body).slice(0, 200000)}`;
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
      { error: await openAIError(response, "인공지능 요청에 실패했습니다.") },
      { status: 502 },
    );
  const data = await response.json();
  try {
    return NextResponse.json({
      configured: true,
      plan: normalizePlan(parseOpenAIJson(data, "작업계획")),
    });
  } catch (error) {
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "작업계획을 읽지 못했습니다.",
      },
      { status: 502 },
    );
  }
}
