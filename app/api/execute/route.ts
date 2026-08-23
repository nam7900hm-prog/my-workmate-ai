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
  const prompt = `사용자가 승인한 작업계획을 실제 결과로 작성한다. 모든 선택 자료를 공평하게 사용한다. 자료에 없는 사실을 만들지 않는다. 자료가 부족하면 warnings에 기록한다. Excel 검색과 추출과 교체 요청은 원래 행과 열의 의미를 유지하고 범위 밖 값은 바꾸지 않는다. 표가 적합하면 kind를 table로 하고 columns와 rows를 채운다. 문서가 적합하면 kind를 text로 하고 text를 채운다. 시간표를 반별 교과연계표로 만드는 경우 columns는 시간 항목 2학년 1반 2학년 2반 2학년 3반 순서처럼 구성한다. 각 시간은 과목 교사명 수업내용의 3개 행을 사용한다. 과목 교사명 수업내용은 왼쪽 항목 열에 각각 한 번만 적고 반 이름은 가로 열로 배치한다. 같은 항목명을 각 반 셀 안에서 반복하지 않는다. 수업내용 행의 모든 반 값은 교사가 직접 입력할 수 있도록 반드시 빈 문자열로 둔다. 원본이나 AI 추정으로 수업내용을 채우지 않으며 확인 필요 같은 안내문도 넣지 않는다. 시간표 수업 교체는 학년과 반을 먼저 확정한 경우에만 실행한다. 선택 학급의 두 수업을 기준으로 양쪽 교사의 빈 시간과 학급 충돌을 검사하고 승인받은 두 수업만 서로 교환한다. 교사 이름 전체 치환이나 다른 시간표 셀 변경은 금지한다. Excel 원본 셀을 실제로 수정해야 할 때만 excelActions를 만든다. 허용 명령은 replace(sheet range find replace) set(sheet cell value) formula(sheet cell formula) highlight(sheet range value color) conditional(sheet range formula color) dataValidation(sheet range values 또는 sourceRange prompt)이다. conditional은 사용자가 조건부서식을 요청한 경우에만 쓴다. 교사 이름을 고르는 목록처럼 사용자가 드롭다운을 요청하면 dataValidation을 사용하고 입력 자료에서 확인한 이름 목록을 values에 넣는다. sheet와 cell과 range는 입력 자료에서 확인한 실제 이름과 주소만 쓴다. 전체 치환이 아니라 요청 범위를 명시한다. 행 수와 합계와 중복과 누락과 교체 전후를 실제로 검증하여 validation에 기록한다. JSON만 반환한다. 형식은 {"kind":"text 또는 table","title":"결과 제목","text":"문서 본문","columns":["열"],"rows":[["값"]],"validation":["검증"],"warnings":["경고"],"excelActions":[{"type":"replace","sheet":"시트명","range":"A1:D20","find":"이전값","replace":"새값"}]}이다. 입력=${JSON.stringify(body).slice(0, 200000)}`;
  const excelActionGuide = ` 추가 Excel 명령: 정렬은 {"type":"sort","sheet":"시트명","range":"A1:D20","column":2,"order":"asc","hasHeader":true} 필터는 {"type":"filter","sheet":"시트명","range":"A1:D20"} 중복제거는 {"type":"removeDuplicates","sheet":"시트명","range":"A1:D20","columns":[1,2],"hasHeader":true} 행열전환은 {"type":"transpose","sheet":"원본시트","range":"A1:D20","targetSheet":"변환결과","targetCell":"A1"} 행추가는 {"type":"insertRows","sheet":"시트명","startRow":3,"count":2} 열추가는 {"type":"insertColumns","sheet":"시트명","startColumn":4,"count":1} 병합은 {"type":"merge","sheet":"시트명","range":"A1:D1"} 표시형식은 {"type":"format","sheet":"시트명","range":"B2:B20","numberFormat":"#,##0원","bold":false,"horizontal":"right","wrapText":false,"fillColor":"FFF2CC","fontColor":"000000"} 인쇄설정은 {"type":"pageSetup","sheet":"시트명","orientation":"landscape","paperSize":"A4","fitToPage":true,"repeatRows":"1:1"} 일반 요약표는 {"type":"pivotSummary","sheet":"원본시트","range":"A1:D20","rowColumn":1,"valueColumn":4,"operation":"sum","targetSheet":"항목별 요약"} 형식만 사용한다. pivotSummary는 네이티브 피벗표가 아니라 값이 고정된 일반 요약표이다. 사용자가 네이티브 피벗표나 차트를 요구하면 자동 생성했다고 표시하지 말고 warnings에 현재 자동 생성 불가라고 기록하며 차트에 쓸 요약 데이터만 표로 만든다. 사용자가 요구한 작업에 필요한 명령만 만들고 원본 범위 밖은 변경하지 않는다.`;
  const meetingGuide = ` 회의록 요청이면 kind를 text로 하고 제목 다음에 [회의 개요] [참석자] [안건별 논의] [결정사항] [후속 할 일] [확인 필요] 순서로 구분한다. 후속 할 일은 할 일과 담당자와 기한과 상태를 한 줄씩 함께 적는다. 녹취에 없는 이름 날짜 결정은 만들지 않고 warnings에 기록한다. 발언자 구분이 확실하지 않으면 임의로 배정하지 않는다.`;
  const insuranceGuide = ` 보험계약서 분석 요청이면 반드시 kind를 table로 하고 선택한 모든 계약서를 공평하게 비교한다. columns는 보험사 상품명 보험종류 계약상태 보장명 보장조건 제외조건 가입금액 또는 한도 지급기준 원문근거 순서로 구성한다. 사용자가 질병이나 부상을 입력하면 관련 보장 행을 찾을 수 있도록 질병명과 부상명과 보장조건을 생략하지 않는다. 계약서에 없는 보험금 액수나 지급 가능성을 만들지 않는다. 지급 여부는 확정 표현 대신 계약서상 청구 검토 가능 확인 필요 보장 제외 중 하나로 적고 실제 지급액은 보험사 심사와 약관 확인이 필요하다고 warnings에 표시한다.`;
  const admissionGuide = ` 대학 입시요강 분석 요청이면 반드시 kind를 table로 하고 선택한 모든 대학 PDF를 공평하게 비교한다. columns는 대학 모집단위 전형명 전형유형 지역인재 여부 지원자격 학생부 반영 면접 여부 수능최저 제출서류 일정 특이사항 원문근거 순서로 구성한다. 면접 없는 곳 학생부전형 지역인재 등 사용자가 다시 찾을 조건을 각 셀에 명확히 적고 문서에 없는 값은 추정하지 말고 확인 필요로 둔다. 연도와 모집시기와 페이지 또는 항목 근거를 유지한다.`;
  const response = await fetch(openAIUrl("responses"), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: process.env.OPENAI_MODEL || "gpt-5-mini",
      input: prompt + excelActionGuide + meetingGuide + insuranceGuide + admissionGuide,
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
