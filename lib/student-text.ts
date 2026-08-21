export function normalizeStudentSentence(value: unknown) {
  const text = String(value || "")
    .replaceAll(",", " ")
    .replaceAll("，", " ")
    .replace(/\s+/g, " ")
    .trim();
  return text && !/[.!?]$/.test(text) ? `${text}.` : text;
}
