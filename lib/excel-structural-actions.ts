import * as XLSX from "xlsx";
import type { WorkResult } from "./model";

type Action = Extract<
  NonNullable<WorkResult["excelActions"]>[number],
  { type: "sort" | "filter" | "removeDuplicates" | "transpose" }
>;

const RANGE = /^[A-Z]{1,3}[1-9][0-9]{0,6}:[A-Z]{1,3}[1-9][0-9]{0,6}$/i;
const CELL = /^[A-Z]{1,3}[1-9][0-9]{0,6}$/i;

export function applyStructuralExcelAction(book: any, action: Action) {
  const sheet = book.Sheets[action.sheet];
  if (!sheet || !RANGE.test(action.range)) return false;
  const bounds = XLSX.utils.decode_range(action.range);
  const height = bounds.e.r - bounds.s.r + 1;
  const width = bounds.e.c - bounds.s.c + 1;
  if (height * width > 100000) return false;
  if (action.type === "filter") {
    sheet["!autofilter"] = { ref: action.range.toUpperCase() };
    return true;
  }
  const sourceRows = Array.from({ length: height }, (_, row) =>
    Array.from({ length: width }, (_, column) => {
      const address = XLSX.utils.encode_cell({
        r: bounds.s.r + row,
        c: bounds.s.c + column,
      });
      return sheet[address] ? { ...sheet[address] } : undefined;
    }),
  );
  if (action.type === "transpose") {
    if (!CELL.test(action.targetCell) || !action.targetSheet.trim()) return false;
    const safeName = action.targetSheet.trim().slice(0, 31);
    let targetSheet = book.Sheets[safeName];
    if (!targetSheet) {
      targetSheet = XLSX.utils.aoa_to_sheet([]);
      XLSX.utils.book_append_sheet(book, targetSheet, safeName);
    }
    const start = XLSX.utils.decode_cell(action.targetCell);
    sourceRows.forEach((row, rowIndex) =>
      row.forEach((cell: any, columnIndex) => {
        const address = XLSX.utils.encode_cell({
          r: start.r + columnIndex,
          c: start.c + rowIndex,
        });
        if (!cell) return;
        const copied = { ...cell };
        if (copied.f) delete copied.f;
        targetSheet[address] = copied;
      }),
    );
    const targetRange = {
      s: start,
      e: { r: start.r + width - 1, c: start.c + height - 1 },
    };
    const existing = targetSheet["!ref"]
      ? XLSX.utils.decode_range(targetSheet["!ref"])
      : targetRange;
    existing.s.r = Math.min(existing.s.r, targetRange.s.r);
    existing.s.c = Math.min(existing.s.c, targetRange.s.c);
    existing.e.r = Math.max(existing.e.r, targetRange.e.r);
    existing.e.c = Math.max(existing.e.c, targetRange.e.c);
    targetSheet["!ref"] = XLSX.utils.encode_range(existing);
    return true;
  }
  if (sourceRows.some((row) => row.some((cell: any) => Boolean(cell?.f))))
    return false;
  const hasHeader = action.hasHeader !== false;
  const header = hasHeader ? sourceRows.shift() : undefined;
  let outputRows = sourceRows;
  if (action.type === "sort") {
    const key = action.column - 1;
    if (key < 0 || key >= width) return false;
    outputRows = [...sourceRows].sort((left: any[], right: any[]) => {
      const a = left[key]?.v ?? "";
      const b = right[key]?.v ?? "";
      const result =
        typeof a === "number" && typeof b === "number"
          ? a - b
          : String(a).localeCompare(String(b), "ko", { numeric: true });
      return action.order === "desc" ? -result : result;
    });
  } else {
    const keys = action.columns
      .map((column) => column - 1)
      .filter((column) => column >= 0 && column < width);
    if (!keys.length) return false;
    const seen = new Set<string>();
    outputRows = sourceRows.filter((row: any[]) => {
      const key = JSON.stringify(keys.map((column) => row[column]?.v));
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }
  const finalRows = header ? [header, ...outputRows] : outputRows;
  for (let row = 0; row < height; row++)
    for (let column = 0; column < width; column++) {
      const address = XLSX.utils.encode_cell({
        r: bounds.s.r + row,
        c: bounds.s.c + column,
      });
      const cell = finalRows[row]?.[column];
      if (cell) sheet[address] = { ...cell };
      else delete sheet[address];
    }
  return true;
}
