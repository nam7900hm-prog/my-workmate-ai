export type StudentSource = { name: string; text: string };
export type StudentDraft = {
  name: string;
  source: string;
  factDraft: string;
  inferredDraft: string;
  recommendedDraft?: string;
  inferredParts: string[];
  recommendedInferredParts?: string[];
  selected?: "fact" | "inferred" | "recommended" | "merged";
  finalText: string;
  reviewed: boolean;
};
export type WorkResult = {
  kind: "text" | "table";
  title: string;
  text?: string;
  columns?: string[];
  rows?: string[][];
  validation: string[];
  warnings: string[];
  excelActions?: Array<
    | {
        type: "replace";
        sheet: string;
        range?: string;
        find: string;
        replace: string;
      }
    | {
        type: "set";
        sheet: string;
        cell: string;
        value: string | number;
      }
    | {
        type: "formula";
        sheet: string;
        cell: string;
        formula: string;
      }
    | {
        type: "highlight";
        sheet: string;
        range?: string;
        value: string;
        color?: string;
      }
    | {
        type: "conditional";
        sheet: string;
        range: string;
        formula: string;
        color?: string;
      }
    | {
        type: "dataValidation";
        sheet: string;
        range: string;
        values?: string[];
        sourceRange?: string;
        prompt?: string;
      }
    | {
        type: "sort";
        sheet: string;
        range: string;
        column: number;
        order: "asc" | "desc";
        hasHeader?: boolean;
      }
    | { type: "filter"; sheet: string; range: string }
    | {
        type: "removeDuplicates";
        sheet: string;
        range: string;
        columns: number[];
        hasHeader?: boolean;
      }
    | {
        type: "transpose";
        sheet: string;
        range: string;
        targetSheet: string;
        targetCell: string;
      }
  >;
};
export type FileAnalysis = {
  kind: string;
  summary: string;
  text: string;
  details: string[];
  sheets?: {
    name: string;
    range: string;
    rows: number;
    columns: number;
    formulaCount: number;
    mergedCount: number;
    headers: string[];
    tableAreas?: string[];
    titleCandidates?: string[];
    dateCount?: number;
    numberCount?: number;
    textCount?: number;
    styledCellCount?: number;
    repeatedValues?: string[];
  }[];
  relationships?: string[];
  personalNames?: string[];
  warnings: string[];
  analyzedAt: string;
};
export type FileItem = {
  id: string;
  name: string;
  type: string;
  size: number;
  addedAt: string;
  status: "analyzing" | "ready" | "partial" | "unsupported" | "error";
  data?: string;
  students?: StudentSource[];
  analysis?: FileAnalysis;
};
export type TemplateItem = {
  id: string;
  name: string;
  type: string;
  size?: number;
  addedAt: string;
  data?: string;
  analysis?: FileAnalysis;
};
export type Task = {
  id: string;
  title: string;
  step: number;
  fileIds: string[];
  templateId?: string;
  request: string;
  conversation: { role: "user" | "assistant"; text: string }[];
  plan?: Plan;
  result?: string;
  workResult?: WorkResult;
  outputFormat?: "auto" | "xlsx" | "docx" | "pdf" | "pptx" | "txt" | "csv";
  studentDrafts?: StudentDraft[];
  studentValidation?: string[];
  studentRecordType?: "subject" | "behavior" | "general";
  apiConsent?: { approvedAt: string; fileIds: string[]; masked: boolean };
  createdAt: string;
  updatedAt: string;
  status: "draft" | "completed";
};
export type Plan = {
  understanding: string;
  materials: string[];
  template: string;
  steps: string[];
  resultFormat: string;
  questions: string[];
  canExecute: boolean;
  limitation?: string;
};
export type Store = {
  files: FileItem[];
  templates: TemplateItem[];
  tasks: Task[];
  archive: Task[];
  trash: {
    files: FileItem[];
    templates: TemplateItem[];
    tasks: Task[];
  };
};
export const initialStore: Store = {
  files: [],
  templates: [],
  tasks: [],
  archive: [],
  trash: { files: [], templates: [], tasks: [] },
};
