import { NextRequest, NextResponse } from "next/server";
import { openAIError, parseOpenAIJson } from "@/lib/openai-response";
import { normalizeWorkResult } from "@/lib/openai-schema";
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
      { error: "인공지능 연결이 필요합니다." },
      { status: 503 },
    );
  const prompt = `사용자가 승인한 작업계획을 실제 결과로 작성한다. 모든 선택 자료를 공평하게 사용한다. 자료에 없는 사실을 만들지 않는다. 자료가 부족하면 warnings에 기록한다. Excel 검색과 추출과 교체 요청은 원래 행과 열의 의미를 유지하고 범위 밖 값은 바꾸지 않는다. 표가 적합하면 kind를 table로 하고 columns와 rows를 채운다. 문서가 적합하면 kind를 text로 하고 text를 채운다. Excel 원본 셀을 실제로 수정해야 할 때만 excelActions를 만든다. 허용 명령은 replace(sheet range find replace) set(sheet cell value) formula(sheet cell formula) highlight(sheet range value color) conditional(sheet range formula color)뿐이다. conditional은 사용자가 조건부서식을 요청한 경우에만 쓴다. sheet와 cell과 range는 입력 자료에서 확인한 실제 이름과 주소만 쓴다. 전체 치환이 아니라 요청 범위를 명시한다. 행 수와 합계와 중복과 누락과 교체 전후를 실제로 검증하여 validation에 기록한다. JSON만 반환한다. 형식은 {"kind":"text 또는 table","title":"결과 제목","text":"문서 본문","columns":["열"],"rows":[["값"]],"validation":["검증"],"warnings":["경고"],"excelActions":[{"type":"replace","sheet":"시트명","range":"A1:D20","find":"이전값","replace":"새값"}]}이다. 입력=${JSON.stringify(body).slice(0, 200000)}`;
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
      { error: await openAIError(response, "결과 생성에 실패했습니다.") },
      { status: 502 },
    );
  const data = await response.json();
  try {
    return NextResponse.json({
      result: normalizeWorkResult(parseOpenAIJson(data, "작업 결과")),
    });
  } catch (error) {
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "작업 결과를 읽지 못했습니다.",
      },
      { status: 502 },
    );
  }
}
