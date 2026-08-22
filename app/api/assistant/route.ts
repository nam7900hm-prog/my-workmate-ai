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
  const prompt = `당신은 한국어 범용 개인 업무비서다. 파일 형식에 치우치지 말고 모든 선택 자료의 실제 추출 내용과 사용자 요청을 함께 검토한다. 원본을 변경하지 않는다. 자료에 없는 사실은 만들지 않는다. 필요한 질문과 실행계획과 검증방법을 제시한다. 특정 전문 규칙은 사용자가 해당 작업을 요청했을 때만 적용한다. Excel 작업이면 검색과 추출과 정렬과 필터와 중복검사와 함수 적용과 교체 범위 및 결과 검증을 구체적으로 단계에 적는다. 시간표를 반별 교과연계표로 바꾸는 요청에서는 반을 가로 열로 배치하고 시간별로 과목과 교사명과 수업내용을 왼쪽 항목 열에 한 번만 적는 구조를 계획한다. 교과연계표의 수업내용 값은 교사가 직접 입력하므로 모든 반에서 빈칸으로 두도록 계획한다. 시간표 수업 교체 요청에서 대상 학년과 반이 입력되지 않았다면 다른 질문보다 먼저 학년과 반을 질문한다. 학년과 반과 과목만 확인되고 교체할 현재 시간이 정해지지 않았다면 해당 과목의 모든 수업과 전체 요일을 비교해 후보를 찾는다. 출장 등으로 수업할 수 없는 요일과 교시가 입력되었다면 그 수업을 고정 교체 대상으로 삼고 사용자가 옮길 요일이나 상대 교사를 정하게 하지 않는다. 같은 학급의 다른 과목 수업 전체를 조사한다. 담당 교사가 후보 시간에 비어 있고 상대 교사가 고정 교체 시간에 비어 있는 경우만 남긴다. 학급 충돌과 양쪽 교사의 연속 수업과 공강 변화를 함께 검사해 후보를 좋은 순서대로 최대 3개 제안한다. 승인 전에는 변경하지 않는다. JSON만 반환한다. 키: understanding 문자열 materials 문자열배열 template 문자열 steps 문자열배열 resultFormat 문자열 questions 문자열배열 canExecute 불리언 limitation 문자열. 입력=${JSON.stringify(body).slice(0, 200000)}`;
  const meetingGuide = ` 회의록 요청이면 녹취에서 확인되는 회의명 일시 장소 참석자 안건별 논의 결정사항 후속 할 일 담당자 기한을 구분한다. 녹취에 없는 참석자 담당자 기한은 만들지 않고 추가 질문에 넣는다. 발언자가 불명확하면 임의 이름을 붙이지 말고 발언자 확인 필요로 표시한다.`;
  const response = await fetch(openAIUrl("responses"), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: process.env.OPENAI_MODEL || "gpt-5-mini",
      input: prompt + meetingGuide,
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
