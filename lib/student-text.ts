export function normalizeStudentSentence(value: unknown) {
  const text = String(value || "")
    .replace(/\bSupabase\b/gi, "자료를 저장하고 관리하는 장치")
    .replace(/[^\p{L}\p{N}\s.()]/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
  return text && !text.endsWith(".") ? `${text}.` : text;
}

const RECORD_ENDINGS =
  /(함|됨|보임|나타남|드러남|돋보임|가짐|생김|봄|있음|없음|짐|줌|냄|시킴|확인됨|이어짐|넓어짐|깊어짐|발견함|정리해 냄)$/;

export function studentDraftProblems(value: unknown) {
  const text = String(value || "").trim();
  const protectedText = text.replace(
    /\((\d{4})\.(\d{2})\.(\d{2})\.\)/g,
    "($1년$2월$3일)",
  );
  const sentences = protectedText
    .split(".")
    .map((sentence) => sentence.trim())
    .filter(Boolean);
  const problems: string[] = [];
  const characters = Array.from(text).length;
  const bytes = new TextEncoder().encode(text).length;
  if (characters < 200 || characters > 500 || bytes > 1500)
    problems.push("각 안은 200자 이상 500자 이하이며 1500바이트 이하여야 합니다.");
  if (!text.endsWith(".")) problems.push("마지막 문장에 마침표가 필요합니다.");
  const withoutDates = text.replace(/\(\d{4}\.\d{2}\.\d{2}\.\)/g, "");
  if (/[^\p{L}\p{N}\s.]/u.test(withoutDates))
    problems.push("날짜 괄호와 마침표 이외의 문장부호를 사용할 수 없습니다.");
  if (!sentences.length || sentences.some((sentence) => !RECORD_ENDINGS.test(sentence)))
    problems.push("중간 문장을 포함한 모든 문장을 명사형 음슴체 기록 문체로 마무리해야 합니다.");
  const endings = sentences.map((sentence) => sentence.match(RECORD_ENDINGS)?.[1] || "");
  for (let index = 2; index < endings.length; index++) {
    if (endings[index] && endings[index] === endings[index - 1] && endings[index] === endings[index - 2]) {
      problems.push("같은 종결 표현을 세 문장 이상 연속하여 사용할 수 없습니다.");
      break;
    }
  }
  if (/(역량이 우수함|능력이 탁월함|적극적으로 참여함)(\.|$)/.test(text))
    problems.push("근거 없는 추상적 칭찬을 구체적인 관찰 장면으로 바꿔야 합니다.");
  if (/구체화하는 계기가 됨/.test(text))
    problems.push("무엇을 구체화했는지 쉬운 말로 풀어 써야 합니다.");
  return [...new Set(problems)];
}

export function isStudentNominalStyle(value: unknown) {
  return !studentDraftProblems(value).some((problem) =>
    problem.includes("명사형 기록 문체"),
  );
}
