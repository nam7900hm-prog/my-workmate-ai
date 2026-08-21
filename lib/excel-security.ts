export function safeExcelFormula(value: string) {
  const formula = value.replace(/^=/, "").trim();
  return Boolean(
    formula &&
      formula.length <= 500 &&
      !/https?:|\[|\]|WEBSERVICE|HYPERLINK|DDE|_xll|\bCALL\s*\(|\bEXEC\s*\(/i.test(
        formula,
      ),
  );
}
