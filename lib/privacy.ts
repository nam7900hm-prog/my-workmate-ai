const patterns = [
  /\b01[016789][- ]?\d{3,4}[- ]?\d{4}\b/g,
  /\b\d{6}[- ]?[1-4]\d{6}\b/g,
  /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi,
  /\b\d{2,3}[- ]\d{3,4}[- ]\d{4}\b/g,
];
export function maskText(text: string, names: string[] = []) {
  let out = text;
  for (const p of patterns) out = out.replace(p, "[개인정보]");
  const map = new Map<string, string>();
  [...new Set(names.filter((x) => x.trim()).map((x) => x.trim()))].forEach(
    (name, i) => {
      const id = `학생${String(i + 1).padStart(3, "0")}`;
      map.set(id, name);
      out = out.replaceAll(name, id);
    },
  );
  return { text: out, map: Object.fromEntries(map) };
}
