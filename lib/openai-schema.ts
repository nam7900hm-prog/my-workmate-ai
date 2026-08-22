import type { Plan, WorkResult } from "./model";

const strings = (value: unknown) =>
  Array.isArray(value)
    ? value.filter((item): item is string => typeof item === "string")
    : [];

export function normalizePlan(value: unknown): Plan {
  if (!value || typeof value !== "object")
    throw new Error("작업계획의 내용이 올바르지 않습니다.");
  const item = value as Record<string, unknown>;
  if (typeof item.understanding !== "string" || !item.understanding.trim())
    throw new Error("작업계획에서 이해한 요청이 빠졌습니다.");
  return {
    understanding: item.understanding,
    materials: strings(item.materials),
    template:
      typeof item.template === "string" ? item.template : "선택하지 않음",
    steps: strings(item.steps),
    resultFormat:
      typeof item.resultFormat === "string"
        ? item.resultFormat
        : "사용자 확인 필요",
    questions: strings(item.questions),
    canExecute: item.canExecute === true,
    limitation:
      typeof item.limitation === "string" ? item.limitation : undefined,
  };
}

export function normalizeWorkResult(value: unknown): WorkResult {
  if (!value || typeof value !== "object")
    throw new Error("작업 결과의 내용이 올바르지 않습니다.");
  const item = value as Record<string, unknown>;
  const kind = item.kind === "table" ? "table" : "text";
  if (typeof item.title !== "string" || !item.title.trim())
    throw new Error("작업 결과 제목이 빠졌습니다.");
  const columns = strings(item.columns);
  const rows = Array.isArray(item.rows)
    ? item.rows
        .filter(Array.isArray)
        .map((row) => row.map((cell: unknown) => String(cell ?? "")))
    : [];
  const excelActions: NonNullable<WorkResult["excelActions"]> = [];
  if (Array.isArray(item.excelActions))
    for (const value of item.excelActions) {
      if (!value || typeof value !== "object") continue;
      const action = value as Record<string, unknown>;
      if (typeof action.sheet !== "string" || !action.sheet.trim()) continue;
      const range = typeof action.range === "string" ? action.range : undefined;
      const color = typeof action.color === "string" ? action.color : undefined;
      if (
        action.type === "replace" &&
        typeof action.range === "string" &&
        typeof action.find === "string" &&
        typeof action.replace === "string"
      )
        excelActions.push({
          type: "replace" as const,
          sheet: action.sheet,
          range: action.range,
          find: action.find,
          replace: action.replace,
        });
      if (
        action.type === "set" &&
        typeof action.cell === "string" &&
        (typeof action.value === "string" || typeof action.value === "number")
      )
        excelActions.push({
          type: "set" as const,
          sheet: action.sheet,
          cell: action.cell,
          value: action.value,
        });
      if (
        action.type === "formula" &&
        typeof action.cell === "string" &&
        typeof action.formula === "string"
      )
        excelActions.push({
          type: "formula" as const,
          sheet: action.sheet,
          cell: action.cell,
          formula: action.formula,
        });
      if (
        action.type === "highlight" &&
        typeof action.range === "string" &&
        typeof action.value === "string"
      )
        excelActions.push({
          type: "highlight" as const,
          sheet: action.sheet,
          range: action.range,
          value: action.value,
          color,
        });
      if (
        action.type === "conditional" &&
        typeof action.range === "string" &&
        typeof action.formula === "string"
      )
        excelActions.push({
          type: "conditional" as const,
          sheet: action.sheet,
          range: action.range,
          formula: action.formula,
          color,
        });
      if (
        action.type === "dataValidation" &&
        typeof action.range === "string"
      ) {
        const values = strings(action.values).slice(0, 1000);
        const sourceRange =
          typeof action.sourceRange === "string"
            ? action.sourceRange
            : undefined;
        if (values.length || sourceRange)
          excelActions.push({
            type: "dataValidation" as const,
            sheet: action.sheet,
            range: action.range,
            values,
            sourceRange,
            prompt:
              typeof action.prompt === "string" ? action.prompt : undefined,
          });
      }
      if (
        action.type === "sort" &&
        typeof action.range === "string" &&
        typeof action.column === "number" &&
        Number.isInteger(action.column) &&
        action.column >= 1
      )
        excelActions.push({
          type: "sort" as const,
          sheet: action.sheet,
          range: action.range,
          column: action.column,
          order: action.order === "desc" ? "desc" : "asc",
          hasHeader: action.hasHeader !== false,
        });
      if (action.type === "filter" && typeof action.range === "string")
        excelActions.push({
          type: "filter" as const,
          sheet: action.sheet,
          range: action.range,
        });
      if (
        action.type === "removeDuplicates" &&
        typeof action.range === "string"
      ) {
        const columns = Array.isArray(action.columns)
          ? action.columns
              .filter(
                (column): column is number =>
                  typeof column === "number" &&
                  Number.isInteger(column) &&
                  column >= 1,
              )
              .slice(0, 100)
          : [];
        if (columns.length)
          excelActions.push({
            type: "removeDuplicates" as const,
            sheet: action.sheet,
            range: action.range,
            columns,
            hasHeader: action.hasHeader !== false,
          });
      }
      if (
        action.type === "transpose" &&
        typeof action.range === "string" &&
        typeof action.targetSheet === "string" &&
        typeof action.targetCell === "string"
      )
        excelActions.push({
          type: "transpose" as const,
          sheet: action.sheet,
          range: action.range,
          targetSheet: action.targetSheet,
          targetCell: action.targetCell,
        });
    }
  return {
    kind,
    title: item.title,
    text: typeof item.text === "string" ? item.text : "",
    columns,
    rows,
    validation: strings(item.validation),
    warnings: strings(item.warnings),
    excelActions,
  };
}
