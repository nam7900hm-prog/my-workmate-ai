export function normalizeStudentSentence(value: unknown) {
  const text = String(value || "")
    .replace(/\bSupabase\b/gi, "자료를 저장하고 관리하는 장치")
    .replaceAll(",", " ")
    .replaceAll("，", " ")
    .replace(/\s+/g, " ")
    .trim();
  return text && !/[.!?]$/.test(text) ? `${text}.` : text;
}
