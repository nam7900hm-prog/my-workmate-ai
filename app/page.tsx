"use client";
import { useEffect, useRef, useState } from "react";
import type {
  FileItem,
  Plan,
  Store,
  StudentDraft,
  StudentSource,
  Task,
  TemplateItem,
  WorkResult,
} from "@/lib/model";
import { initialStore } from "@/lib/model";
import {
  backup,
  load,
  restore,
  save,
  savedBackups,
  saveImportant,
  restoreSaved,
} from "@/lib/store";
import { analyzeFile } from "@/lib/analyze";
import { maskText } from "@/lib/privacy";
import { saveFileBlob, storedFile } from "@/lib/file-db";
import { safeExcelFormula } from "@/lib/excel-security";
type View =
  | "home"
  | "help"
  | "files"
  | "jobs"
  | "templates"
  | "recent"
  | "archive"
  | "settings"
  | "task";
const tabs: [View, string, string][] = [
  ["home", "⌂", "홈"],
  ["help", "?", "사용설명"],
  ["files", "▤", "내 자료"],
  ["jobs", "☷", "진행 중 작업"],
  ["templates", "▣", "내 양식"],
  ["recent", "◷", "최근 작업"],
  ["archive", "□", "보관함"],
  ["settings", "⚙", "설정"],
];
const steps = [
  ["선택", "자료·양식·이전 작업 고르기"],
  ["설명", "글쓰기 또는 마이크로 말하기"],
  ["AI 질문", "필요한 내용 함께 확인"],
  ["계획 확인", "변경 내용과 결과 미리보기"],
  ["실행·검증", "원본을 보호하며 결과 생성"],
  ["결과 확인", "다운로드하거나 다시 수정"],
];
function resultRows(result: WorkResult) {
  if (result.kind === "table")
    return [result.columns || [], ...(result.rows || [])].map((row) =>
      row.map(String),
    );
  return (result.text || "")
    .split(/\n+/)
    .filter(Boolean)
    .map((line) => [line]);
}

async function wordBlob(result: WorkResult) {
  const docx = await import("docx");
  const children: any[] = [
    new docx.Paragraph({
      text: result.title,
      heading: docx.HeadingLevel.TITLE,
    }),
  ];
  if (result.kind === "table") {
    children.push(
      new docx.Table({
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
        rows: resultRows(result).map(
          (row, rowIndex) =>
            new docx.TableRow({
              tableHeader: rowIndex === 0,
              children: row.map(
                (value) =>
                  new docx.TableCell({
                    shading:
                      rowIndex === 0
                        ? { fill: "DCEBFA", type: docx.ShadingType.CLEAR }
                        : undefined,
                    children: [
                      new docx.Paragraph({
                        children: [
                          new docx.TextRun({
                            text: value,
                            bold: rowIndex === 0,
                          }),
                        ],
                      }),
                    ],
                  }),
              ),
            }),
        ),
      }),
    );
  } else {
    children.push(
      ...(result.text || "")
        .split(/\n+/)
        .filter(Boolean)
        .map((text) => new docx.Paragraph({ text, spacing: { after: 160 } })),
    );
  }
  if (result.validation.length) {
    children.push(
      new docx.Paragraph({
        text: "검증 결과",
        heading: docx.HeadingLevel.HEADING_1,
      }),
      ...result.validation.map(
        (text) => new docx.Paragraph({ text: `✓ ${text}` }),
      ),
    );
  }
  return docx.Packer.toBlob(new docx.Document({ sections: [{ children }] }));
}

function xmlText(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

async function wordTemplateBlob(template: File, result: WorkResult) {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(await template.arrayBuffer());
  const documentPart = zip.file("word/document.xml");
  if (!documentPart)
    throw new Error("선택한 Word 양식의 본문을 읽지 못했습니다.");
  const documentXml = await documentPart.async("string");
  const paragraphs = [
    result.title,
    ...resultRows(result).map((row) => row.join(" | ")),
    ...(result.validation.length ? ["검증 결과", ...result.validation] : []),
  ];
  const inserted = paragraphs
    .filter(Boolean)
    .map(
      (text, index) =>
        `<w:p><w:pPr>${index === 0 ? '<w:pStyle w:val="Title"/>' : ""}</w:pPr><w:r><w:t xml:space="preserve">${xmlText(text)}</w:t></w:r></w:p>`,
    )
    .join("");
  const marker = documentXml.includes("<w:sectPr") ? "<w:sectPr" : "</w:body>";
  zip.file("word/document.xml", documentXml.replace(marker, inserted + marker));
  return zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function pdfBlob(result: WorkResult) {
  const { PDFDocument } = await import("pdf-lib");
  const pdf = await PDFDocument.create();
  const lines = [
    result.title,
    "",
    ...resultRows(result).map((row) => row.join("  |  ")),
    "",
    "검증 결과",
    ...result.validation.map((text) => `✓ ${text}`),
  ];
  const pages: string[][] = [];
  let current: string[] = [];
  for (const raw of lines) {
    const pieces: string[] = [];
    let line = "";
    for (const character of raw) {
      line += character;
      if (line.length >= 48) {
        pieces.push(line);
        line = "";
      }
    }
    pieces.push(line);
    for (const piece of pieces) {
      if (current.length >= 34) {
        pages.push(current);
        current = [];
      }
      current.push(piece);
    }
  }
  if (current.length) pages.push(current);
  for (const pageLines of pages) {
    const canvas = document.createElement("canvas");
    canvas.width = 1240;
    canvas.height = 1754;
    const context = canvas.getContext("2d")!;
    context.fillStyle = "white";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.fillStyle = "#102A43";
    context.font = "34px Arial, 'Malgun Gothic', sans-serif";
    pageLines.forEach((line, index) =>
      context.fillText(line, 90, 110 + index * 46),
    );
    const image = await new Promise<Blob>((resolve, reject) =>
      canvas.toBlob(
        (blob) =>
          blob ? resolve(blob) : reject(new Error("PDF 화면 생성 실패")),
        "image/png",
      ),
    );
    const embedded = await pdf.embedPng(await image.arrayBuffer());
    const page = pdf.addPage([595.28, 841.89]);
    page.drawImage(embedded, { x: 0, y: 0, width: 595.28, height: 841.89 });
  }
  const bytes = await pdf.save();
  const buffer = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
  return new Blob([buffer], { type: "application/pdf" });
}

async function powerPointBlob(result: WorkResult) {
  if (!(window as any).PptxGenJS) {
    await new Promise<void>((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "/vendor/pptxgen.bundle.js";
      script.onload = () => resolve();
      script.onerror = () =>
        reject(new Error("PowerPoint 생성기를 불러오지 못했습니다."));
      document.head.appendChild(script);
    });
  }
  const PptxGenJS = (window as any).PptxGenJS;
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "my workmate ai";
  pptx.subject = result.title;
  if (result.kind === "table") {
    const rows = resultRows(result);
    const header = rows[0] || [];
    for (let start = 1; start < rows.length || start === 1; start += 12) {
      const slide = pptx.addSlide();
      slide.addText(result.title, {
        x: 0.6,
        y: 0.35,
        w: 12,
        h: 0.55,
        fontSize: 24,
        bold: true,
        color: "102A43",
      });
      slide.addTable([header, ...rows.slice(start, start + 12)], {
        x: 0.6,
        y: 1.15,
        w: 12,
        h: 5.6,
        border: { type: "solid", color: "D8E2EC", pt: 1 },
        fill: "FFFFFF",
        color: "243B53",
        fontSize: 11,
        bold: false,
        rowH: 0.35,
        margin: 0.06,
      });
    }
  } else {
    const paragraphs = (result.text || "").split(/\n+/).filter(Boolean);
    for (let start = 0; start < Math.max(1, paragraphs.length); start += 10) {
      const slide = pptx.addSlide();
      slide.addText(result.title, {
        x: 0.7,
        y: 0.45,
        w: 11.8,
        h: 0.6,
        fontSize: 26,
        bold: true,
        color: "102A43",
      });
      slide.addText(paragraphs.slice(start, start + 10).join("\n\n"), {
        x: 0.8,
        y: 1.35,
        w: 11.6,
        h: 5.4,
        fontSize: 18,
        color: "334E68",
        breakLine: false,
        valign: "top",
      });
    }
  }
  return (await pptx.write({ outputType: "blob" })) as Blob;
}
function findDraftDuplicates(drafts: StudentDraft[]) {
  const warnings: string[] = [];
  const grams = (text: string) => {
    const value = text.replace(/[\s.,!?]/g, "");
    return new Set(
      Array.from({ length: Math.max(0, value.length - 1) }, (_, i) =>
        value.slice(i, i + 2),
      ),
    );
  };
  const similarity = (a: Set<string>, b: Set<string>) => {
    if (!a.size || !b.size) return 0;
    let common = 0;
    a.forEach((x) => b.has(x) && common++);
    return common / Math.max(a.size, b.size);
  };
  const prepared = drafts.map((draft) => ({
    name: draft.name,
    fact: grams(draft.factDraft),
    inferred: grams(draft.inferredDraft),
  }));
  for (let i = 0; i < prepared.length; i++) {
    for (let j = i + 1; j < prepared.length; j++) {
      if (
        similarity(prepared[i].fact, prepared[j].fact) >= 0.82 ||
        similarity(prepared[i].inferred, prepared[j].inferred) >= 0.82
      ) {
        warnings.push(
          `${prepared[i].name}과 ${prepared[j].name}의 문장 유사도가 높습니다.`,
        );
        if (warnings.length >= 30) return warnings;
      }
    }
  }
  return warnings;
}
export default function App() {
  const [mounted, setMounted] = useState(false),
    [view, setView] = useState<View>("home"),
    [store, setStore] = useState<Store>(initialStore),
    [taskId, setTaskId] = useState(""),
    [selected, setSelected] = useState<string[]>([]),
    [notice, setNotice] = useState(""),
    [listening, setListening] = useState(false),
    [restorePreview, setRestorePreview] = useState<Store | null>(null),
    [pendingApi, setPendingApi] = useState<"plan" | "students" | null>(null);
  const rec = useRef<any>(null),
    keep = useRef(false),
    voice = useRef(""),
    rawFiles = useRef(new Map<string, File>());
  const task = store.tasks.find((x) => x.id === taskId);
  useEffect(() => {
    setMounted(true);
    setStore(load());
    const sync = () => {
      const h = location.hash.slice(1) as View;
      setView(tabs.some((x) => x[0] === h) || h === "task" ? h : "home");
    };
    sync();
    addEventListener("hashchange", sync);
    return () => removeEventListener("hashchange", sync);
  }, []);
  useEffect(() => {
    if (mounted) save(store);
  }, [store, mounted]);
  const go = (v: View) => {
    if (location.hash !== "#" + v) history.pushState({}, "", "#" + v);
    setView(v);
  };
  const update = (fn: (s: Store) => Store) => setStore((s) => fn(s));
  function newTask(fileIds: string[] = [], templateId?: string) {
    const now = new Date().toISOString(),
      t: Task = {
        id: crypto.randomUUID(),
        title: "새 작업",
        // A task created from the home screen intentionally has no source yet.
        // It must open the writing/voice screen instead of getting stuck on
        // the optional source-selection step.
        step: 2,
        fileIds,
        templateId,
        request: "",
        outputFormat: "auto",
        conversation: [],
        createdAt: now,
        updatedAt: now,
        status: "draft",
      };
    update((s) => ({ ...s, tasks: [t, ...s.tasks] }));
    setTaskId(t.id);
    setSelected(fileIds);
    go("task");
  }
  function changeTask(p: Partial<Task>) {
    if (!task) return;
    update((s) => ({
      ...s,
      tasks: s.tasks.map((t) =>
        t.id === task.id
          ? { ...t, ...p, updatedAt: new Date().toISOString() }
          : t,
      ),
    }));
  }
  async function addFiles(list: FileList | null, asTemplate = false) {
    if (!list) return;
    const made: any[] = [];
    setNotice(`${list.length}개 자료의 실제 내용을 이 기기에서 읽고 있습니다.`);
    for (const f of Array.from(list)) {
      const id = crypto.randomUUID();
      rawFiles.current.set(id, f);
      await saveFileBlob(id, f);
      const ext = f.name.split(".").pop()?.toLowerCase() || "",
        supported = [
          "xlsx",
          "xls",
          "csv",
          "docx",
          "pdf",
          "pptx",
          "txt",
          "jpg",
          "jpeg",
          "png",
          "webp",
          "mp3",
          "m4a",
          "wav",
          "webm",
          "hwpx",
        ].includes(ext);
      try {
        const found = supported ? await analyzeFile(f) : undefined;
        made.push({
          id,
          name: f.name,
          type: ext.toUpperCase(),
          size: f.size,
          addedAt: new Date().toISOString(),
          status: found
            ? found.analysis.warnings.length
              ? "partial"
              : "ready"
            : "unsupported",
          analysis: found?.analysis,
          students: found?.students,
        });
      } catch (e: any) {
        made.push({
          id,
          name: f.name,
          type: ext.toUpperCase(),
          size: f.size,
          addedAt: new Date().toISOString(),
          status: "error",
          analysis: {
            kind: ext,
            summary: "내용 분석에 실패했습니다.",
            text: "",
            details: [],
            warnings: [e.message || "알 수 없는 오류"],
            analyzedAt: new Date().toISOString(),
          },
        });
      }
    }
    if (asTemplate)
      update((s) => ({
        ...s,
        templates: [
          ...made.map((x) => ({
            id: x.id,
            name: x.name,
            type: x.type,
            size: x.size,
            addedAt: x.addedAt,
            data: x.data,
            analysis: x.analysis,
          })),
          ...s.templates,
        ],
      }));
    else update((s) => ({ ...s, files: [...made, ...s.files] }));
    setNotice(`${made.length}개 자료 분석을 마쳤습니다.`);
  }
  function startVoice() {
    const SR =
      (window as any).SpeechRecognition ||
      (window as any).webkitSpeechRecognition;
    if (!SR)
      return setNotice(
        "이 브라우저는 연속 받아쓰기를 지원하지 않습니다. Chrome 또는 Edge를 사용하거나 음성파일을 추가해 주세요.",
      );
    keep.current = true;
    setListening(true);
    voice.current = task?.request || "";
    const begin = () => {
      if (!keep.current) return;
      const r = new SR();
      rec.current = r;
      r.lang = "ko-KR";
      r.continuous = true;
      r.interimResults = true;
      r.onresult = (e: any) => {
        let interim = "";
        for (let i = e.resultIndex; i < e.results.length; i++) {
          const x = e.results[i][0].transcript;
          if (e.results[i].isFinal)
            voice.current = (voice.current + " " + x).trim();
          else interim += x;
        }
        changeTask({ request: (voice.current + " " + interim).trim() });
      };
      r.onerror = (e: any) => {
        if (["not-allowed", "service-not-allowed"].includes(e.error)) {
          keep.current = false;
          setListening(false);
          setNotice("마이크 권한을 허용해 주세요.");
        }
      };
      r.onend = () => {
        rec.current = null;
        if (keep.current) setTimeout(begin, 350);
        else setListening(false);
      };
      try {
        r.start();
      } catch {
        setTimeout(begin, 500);
      }
    };
    begin();
  }
  function stopVoice() {
    keep.current = false;
    rec.current?.stop();
    rec.current = null;
    setListening(false);
    if (task) changeTask({ request: voice.current.trim() || task.request });
    setNotice("받아쓰기를 완료했습니다.");
  }
  function makePlan() {
    if (!task?.request.trim())
      return setNotice("원하는 일을 글이나 말로 설명해 주세요.");
    setPendingApi("plan");
  }
  async function runPlan() {
    if (!task) return;
    let chosen = store.files.filter((x) => task.fileIds.includes(x.id));
    const media = chosen.filter((x) =>
      ["image", "audio"].includes(x.analysis?.kind || ""),
    );
    if (media.length) {
      setNotice("확인한 사진과 음성 내용을 읽고 있습니다.");
      const refreshed = new Map<string, FileItem>();
      for (const item of media) {
        let file = rawFiles.current.get(item.id);
        if (!file) file = await storedFile(item.id, item.name);
        if (!file && item.data) {
          const blob = await (await fetch(item.data)).blob();
          file = new File([blob], item.name, { type: blob.type });
        }
        if (!file) {
          setPendingApi(null);
          return setNotice(
            `${item.name} 원본을 다시 선택해 주세요. 새로고침 뒤에는 큰 사진과 음성 원본을 다시 선택해야 합니다.`,
          );
        }
        const form = new FormData();
        form.append("consent", "true");
        form.append("file", file, file.name);
        const response = await fetch("/api/media", {
          method: "POST",
          body: form,
        });
        const data = await response.json();
        if (!response.ok) {
          setPendingApi(null);
          return setNotice(
            data.error || `${item.name} 내용을 읽지 못했습니다.`,
          );
        }
        refreshed.set(item.id, {
          ...item,
          status: "ready",
          analysis: {
            ...(item.analysis as NonNullable<FileItem["analysis"]>),
            summary:
              item.analysis?.kind === "audio"
                ? "음성을 글로 변환했습니다."
                : "사진의 글과 표를 읽었습니다.",
            text: data.text || "",
            details: [
              ...(item.analysis?.details || []),
              "OpenAI API 분석 완료",
            ],
            warnings: [],
            analyzedAt: new Date().toISOString(),
          },
        });
      }
      if (refreshed.size) {
        update((s) => ({
          ...s,
          files: s.files.map((x) => refreshed.get(x.id) || x),
        }));
        chosen = chosen.map((x) => refreshed.get(x.id) || x);
      }
    }
    const names = chosen.flatMap((x) => [
      ...(x.students || []).map((s) => s.name),
      ...(x.analysis?.personalNames || []),
    ]);
    const perFileLimit = Math.max(
      4000,
      Math.floor(160000 / Math.max(1, chosen.length)),
    );
    const files = chosen.map(({ name, type, size, analysis }) => ({
      name,
      type,
      size,
      summary: analysis?.summary,
      details: analysis?.details,
      text: maskText(analysis?.text || "", names).text.slice(0, perFileLimit),
      warnings: analysis?.warnings,
    }));
    const template = store.templates.find((x) => x.id === task.templateId);
    setPendingApi(null);
    setNotice("개인정보를 가린 자료로 작업계획을 만들고 있습니다.");
    try {
      const r = await fetch("/api/assistant", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          request: maskText(task.request, names).text,
          files,
          outputFormat: task.outputFormat || "auto",
          template: template
            ? {
                name: template.name,
                type: template.type,
                analysis: template.analysis,
              }
            : undefined,
          consent: true,
        }),
      });
      const d = await r.json();
      if (!r.ok) throw new Error(d.error);
      changeTask({
        apiConsent: {
          approvedAt: new Date().toISOString(),
          fileIds: task.fileIds,
          masked: true,
        },
        plan: d.plan,
        step: d.plan.questions?.length ? 3 : 4,
        conversation: [
          ...task.conversation,
          { role: "user", text: task.request },
          {
            role: "assistant",
            text: d.plan.questions?.join(" ") || "작업계획을 확인해 주세요.",
          },
        ],
      });
      setNotice("실제 자료 내용을 바탕으로 계획했습니다.");
    } catch (e: any) {
      setNotice(e.message);
    }
  }
  function makeStudentDrafts() {
    if (!task) return;
    const students = store.files
      .filter((x) => task.fileIds.includes(x.id))
      .flatMap((x) => x.students || [])
      .slice(0, 200);
    if (!students.length)
      return setNotice(
        "이름과 소감문 열이 있는 Excel 또는 CSV 자료가 필요합니다.",
      );
    setPendingApi("students");
  }
  async function runStudentDrafts() {
    if (!task) return;
    const students = store.files
      .filter((x) => task.fileIds.includes(x.id))
      .flatMap((x) => x.students || [])
      .slice(0, 200);
    const masked = students.map((s, i) => ({
      name: `학생${String(i + 1).padStart(3, "0")}`,
      text: maskText(
        s.text,
        students.map((x) => x.name),
      ).text,
    }));
    setPendingApi(null);
    setNotice(`${students.length}명의 두 가지 초안을 만들고 있습니다.`);
    try {
      const generated: StudentDraft[] = [];
      const batchSize = 20;
      for (let start = 0; start < masked.length; start += batchSize) {
        setNotice(
          `${students.length}명 중 ${start + 1}~${Math.min(start + batchSize, students.length)}명 초안을 만들고 있습니다.`,
        );
        let completed = false;
        let lastError = "학생별 초안 생성에 실패했습니다.";
        for (let attempt = 1; attempt <= 3 && !completed; attempt++) {
          const r = await fetch("/api/student-records", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              students: masked.slice(start, start + batchSize),
              avoidPhrases: generated
                .slice(-40)
                .flatMap((draft) => [draft.factDraft, draft.inferredDraft])
                .filter(Boolean),
              consent: true,
            }),
          });
          const d = await r.json();
          const expected = Math.min(batchSize, masked.length - start);
          if (r.ok && Array.isArray(d.drafts) && d.drafts.length === expected) {
            generated.push(...d.drafts);
            completed = true;
          } else {
            lastError =
              d.error ||
              `요청한 ${expected}명과 응답 인원이 달라 다시 확인합니다.`;
            if (attempt < 3)
              await new Promise((resolve) =>
                setTimeout(resolve, attempt * 800),
              );
          }
        }
        if (!completed) throw new Error(lastError);
      }
      const drafts = generated.map((x: StudentDraft, i: number) => ({
        ...x,
        name: students[i].name,
        source: students[i].text,
      }));
      changeTask({
        apiConsent: {
          approvedAt: new Date().toISOString(),
          fileIds: task.fileIds,
          masked: true,
        },
        studentDrafts: drafts,
        studentValidation: findDraftDuplicates(drafts),
        step: 6,
        result: `학생 ${drafts.length}명의 두 가지 초안을 생성했습니다.`,
      });
      setNotice("학생별 두 가지 초안을 만들었습니다.");
    } catch (e: any) {
      setNotice(e.message);
    }
  }
  async function executeTask() {
    if (!task?.plan || !task.apiConsent)
      return setNotice("작업계획 승인이 필요합니다.");
    const chosen = store.files.filter((x) => task.fileIds.includes(x.id));
    const names = chosen.flatMap((x) => [
      ...(x.students || []).map((s) => s.name),
      ...(x.analysis?.personalNames || []),
    ]);
    const perFileLimit = Math.max(
      4000,
      Math.floor(160000 / Math.max(1, chosen.length)),
    );
    const files = chosen.map((x) => ({
      name: x.name,
      summary: x.analysis?.summary,
      details: x.analysis?.details,
      text: maskText(x.analysis?.text || "", names).text.slice(0, perFileLimit),
    }));
    setNotice("승인한 계획에 따라 결과를 만들고 있습니다.");
    try {
      const response = await fetch("/api/execute", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          request: maskText(task.request, names).text,
          plan: task.plan,
          files,
          consent: true,
        }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error);
      changeTask({
        workResult: data.result,
        result: data.result.title,
        step: 6,
        status: "completed",
      });
      setNotice("결과 생성과 검증을 마쳤습니다.");
    } catch (e: any) {
      setNotice(e.message);
    }
  }
  function changeDraft(index: number, next: Partial<StudentDraft>) {
    if (!task?.studentDrafts) return;
    changeTask({
      studentDrafts: task.studentDrafts.map((d, i) =>
        i === index ? { ...d, ...next } : d,
      ),
    });
  }
  async function downloadStudentResult() {
    if (!task?.studentDrafts) return;
    const columns = ["학생 이름", "학생 원문", "선택", "교사 최종본"];
    const rows = task.studentDrafts.map((d) => [
      d.name,
      d.source,
      d.selected || "",
      d.finalText,
    ]);
    const result: WorkResult = {
      kind: "table",
      title: "학생별 생기부 교사수정본",
      columns,
      rows,
      validation: [
        `전체 ${task.studentDrafts.length}명 교사 확인 완료`,
        "쉼표 금지와 문장 마침표 검사 완료",
        ...(task.studentValidation?.length
          ? [`유사 문장 확인 필요 ${task.studentValidation.length}건`]
          : ["높은 유사도의 문장 쌍 없음"]),
      ],
      warnings: task.studentValidation || [],
    };
    const format =
      !task.outputFormat || task.outputFormat === "auto"
        ? "xlsx"
        : task.outputFormat;
    if (format === "docx")
      return download(await wordBlob(result), `${result.title}.docx`);
    if (format === "pdf")
      return download(await pdfBlob(result), `${result.title}.pdf`);
    if (format === "pptx")
      return download(await powerPointBlob(result), `${result.title}.pptx`);
    const csv =
      "\ufeff" +
      [columns, ...rows]
        .map((r) =>
          r.map((x) => `"${String(x).replaceAll('"', '""')}"`).join(","),
        )
        .join("\r\n");
    if (format === "csv" || format === "txt")
      return download(
        new Blob([csv], { type: "text/csv;charset=utf-8" }),
        `학생별_생기부_교사수정본.${format === "txt" ? "txt" : "csv"}`,
      );
    const XLSX = await import("xlsx");
    const sheet = XLSX.utils.aoa_to_sheet([columns, ...rows]);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, "교사수정본");
    XLSX.writeFile(book, "학생별_생기부_교사수정본.xlsx");
  }
  async function downloadWorkResult() {
    if (!task?.workResult) return;
    const result = task.workResult;
    const safeTitle = (result.title || "결과").replace(/[\\/:*?"<>|]/g, "_");
    const format =
      !task.outputFormat || task.outputFormat === "auto"
        ? result.kind === "table"
          ? "xlsx"
          : "docx"
        : task.outputFormat;
    if (format === "csv") {
      const csv =
        "\ufeff" +
        resultRows(result)
          .map((row) =>
            row.map((value) => `"${value.replaceAll('"', '""')}"`).join(","),
          )
          .join("\r\n");
      return download(
        new Blob([csv], { type: "text/csv;charset=utf-8" }),
        `${safeTitle}.csv`,
      );
    }
    if (format === "txt")
      return download(
        new Blob(
          [
            result.kind === "table"
              ? resultRows(result)
                  .map((row) => row.join("\t"))
                  .join("\r\n")
              : result.text || "",
          ],
          { type: "text/plain;charset=utf-8" },
        ),
        `${safeTitle}.txt`,
      );
    if (format === "docx") {
      const template = store.templates.find(
        (x) => x.id === task.templateId && x.type === "DOCX",
      );
      if (template) {
        const original = await storedFile(template.id, template.name);
        if (!original)
          throw new Error("선택한 Word 양식 원본을 다시 등록해 주세요.");
        return download(
          await wordTemplateBlob(original, result),
          `${safeTitle}.docx`,
        );
      }
      return download(await wordBlob(result), `${safeTitle}.docx`);
    }
    if (format === "pdf")
      return download(await pdfBlob(result), `${safeTitle}.pdf`);
    if (format === "pptx")
      return download(await powerPointBlob(result), `${safeTitle}.pptx`);
    if (format === "xlsx") {
      const XLSX = await import("xlsx");
      const matrix = resultRows(result);
      const columns = matrix[0] || ["내용"];
      const rows = matrix.slice(1);
      const template = store.templates.find((x) => x.id === task.templateId);
      const source = store.files.find(
        (x) =>
          task.fileIds.includes(x.id) &&
          ["XLSX", "XLS", "CSV"].includes(x.type),
      );
      const base = template || source;
      let book: any;
      let filled = false;
      if (base && ["XLSX", "XLS", "CSV"].includes(base.type)) {
        const original =
          (await storedFile(base.id, base.name)) ||
          (base.data
            ? await (async () => {
                const blob = await (await fetch(base.data as string)).blob();
                return new File([blob], base.name, { type: blob.type });
              })()
            : undefined);
        if (!original)
          throw new Error("선택한 Excel 원본을 다시 등록해 주세요.");
        const bytes = await original.arrayBuffer();
        book = XLSX.read(bytes, {
          type: "array",
          cellStyles: true,
          cellFormula: true,
        });
        let actionCount = 0;
        for (const action of (result.excelActions || []).slice(0, 500)) {
          const sheet = book.Sheets[action.sheet];
          if (!sheet) continue;
          if (action.type === "set" || action.type === "formula") {
            if (!/^[A-Z]{1,3}[1-9][0-9]{0,6}$/i.test(action.cell)) continue;
            const address = action.cell.toUpperCase();
            if (action.type === "formula") {
              const formula = action.formula.replace(/^=/, "").trim();
              if (!safeExcelFormula(formula)) continue;
              sheet[address] = { t: "n", f: formula };
            } else {
              sheet[address] = {
                t: typeof action.value === "number" ? "n" : "s",
                v: action.value,
              };
            }
            const existing = XLSX.utils.decode_range(sheet["!ref"] || address);
            const point = XLSX.utils.decode_cell(address);
            existing.s.r = Math.min(existing.s.r, point.r);
            existing.s.c = Math.min(existing.s.c, point.c);
            existing.e.r = Math.max(existing.e.r, point.r);
            existing.e.c = Math.max(existing.e.c, point.c);
            sheet["!ref"] = XLSX.utils.encode_range(existing);
            actionCount++;
            continue;
          }
          if (action.type === "conditional") {
            actionCount++;
            continue;
          }
          let target;
          try {
            target = XLSX.utils.decode_range(
              action.range || sheet["!ref"] || "A1:A1",
            );
          } catch {
            continue;
          }
          const total =
            (target.e.r - target.s.r + 1) * (target.e.c - target.s.c + 1);
          if (total > 100000) continue;
          for (let row = target.s.r; row <= target.e.r; row++) {
            for (let column = target.s.c; column <= target.e.c; column++) {
              const address = XLSX.utils.encode_cell({ r: row, c: column });
              const cell = sheet[address];
              if (!cell) continue;
              const currentValue = String(cell.v ?? "").trim();
              if (
                action.type === "replace" &&
                currentValue !== action.find.trim()
              )
                continue;
              if (
                action.type === "highlight" &&
                currentValue !== action.value.trim()
              )
                continue;
              if (action.type === "replace") {
                cell.v = action.replace;
                cell.t = "s";
                delete cell.f;
              } else if (action.type === "highlight") {
                const color = (action.color || "FFF2CC")
                  .replace(/^#/, "")
                  .toUpperCase();
                if (!/^[0-9A-F]{6}$/.test(color)) continue;
                cell.s = {
                  ...(cell.s || {}),
                  fill: { patternType: "solid", fgColor: { rgb: color } },
                };
              }
              actionCount++;
            }
          }
        }
        filled = actionCount > 0;
        for (const sheetName of book.SheetNames) {
          if (result.excelActions?.length) break;
          const sheet = book.Sheets[sheetName];
          const grid = XLSX.utils.sheet_to_json<any[]>(sheet, {
            header: 1,
            defval: "",
          });
          for (
            let rowIndex = 0;
            rowIndex < Math.min(grid.length, 40);
            rowIndex++
          ) {
            const normalized = (grid[rowIndex] || []).map((x: unknown) =>
              String(x).trim(),
            );
            const positions = columns.map((name) =>
              normalized.indexOf(String(name).trim()),
            );
            if (columns.length && positions.every((x) => x >= 0)) {
              rows.forEach((row, offset) =>
                row.forEach((value, colIndex) => {
                  const address = XLSX.utils.encode_cell({
                    r: rowIndex + 1 + offset,
                    c: positions[colIndex],
                  });
                  sheet[address] = {
                    t: typeof value === "number" ? "n" : "s",
                    v: value,
                  };
                }),
              );
              const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
              range.e.r = Math.max(range.e.r, rowIndex + rows.length);
              sheet["!ref"] = XLSX.utils.encode_range(range);
              filled = true;
              break;
            }
          }
          if (filled) break;
        }
      } else book = XLSX.utils.book_new();
      if (!filled) {
        const sheet = XLSX.utils.aoa_to_sheet([columns, ...rows]);
        let name = "AI 결과";
        let suffix = 2;
        while (book.SheetNames.includes(name)) name = `AI 결과 ${suffix++}`;
        XLSX.utils.book_append_sheet(book, sheet, name);
      }
      let output = XLSX.write(book, {
        type: "array",
        bookType: "xlsx",
        cellStyles: true,
      }) as ArrayBuffer;
      const styleActions = (result.excelActions || []).filter(
        (action) =>
          action.type === "highlight" || action.type === "conditional",
      );
      if (styleActions.length) {
        const ExcelJS = await import("exceljs");
        const styled = new ExcelJS.Workbook();
        await styled.xlsx.load(output);
        for (const action of styleActions) {
          const sheet = styled.getWorksheet(action.sheet);
          if (!sheet) continue;
          const color = (action.color || "FFF2CC")
            .replace(/^#/, "")
            .toUpperCase();
          if (!/^[0-9A-F]{6}$/.test(color)) continue;
          if (action.type === "conditional") {
            if (
              !/^[A-Z]{1,3}[1-9][0-9]{0,6}:[A-Z]{1,3}[1-9][0-9]{0,6}$/i.test(
                action.range,
              ) ||
              !action.formula.trim() ||
              !safeExcelFormula(action.formula)
            )
              continue;
            sheet.addConditionalFormatting({
              ref: action.range.toUpperCase(),
              rules: [
                {
                  type: "expression",
                  priority: 1,
                  formulae: [action.formula.replace(/^=/, "")],
                  style: {
                    fill: {
                      type: "pattern",
                      pattern: "solid",
                      bgColor: { argb: `FF${color}` },
                      fgColor: { argb: `FF${color}` },
                    },
                  },
                },
              ],
            });
            continue;
          }
          let bounds;
          try {
            bounds = XLSX.utils.decode_range(
              action.range ||
                XLSX.utils.encode_range({
                  s: { r: 0, c: 0 },
                  e: {
                    r: Math.max(0, sheet.rowCount - 1),
                    c: Math.max(0, sheet.columnCount - 1),
                  },
                }),
            );
          } catch {
            continue;
          }
          const total =
            (bounds.e.r - bounds.s.r + 1) * (bounds.e.c - bounds.s.c + 1);
          if (total > 100000) continue;
          for (let row = bounds.s.r + 1; row <= bounds.e.r + 1; row++)
            for (
              let column = bounds.s.c + 1;
              column <= bounds.e.c + 1;
              column++
            ) {
              const cell = sheet.getCell(row, column);
              if (String(cell.value ?? "").trim() !== action.value.trim())
                continue;
              cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: `FF${color}` },
              };
            }
        }
        const styledOutput = await styled.xlsx.writeBuffer();
        output = styledOutput.slice(0) as ArrayBuffer;
      }
      download(
        new Blob([output], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        `${safeTitle}.xlsx`,
      );
      setNotice(
        filled
          ? `${template ? "선택한 양식" : "원본 사본"}의 해당 열에 결과를 채웠습니다.`
          : `${template ? "양식" : "원본 사본"}을 보존하고 AI 결과 시트를 추가했습니다.`,
      );
    }
  }
  function answer(text: string) {
    if (!task || !text.trim()) return;
    changeTask({
      conversation: [
        ...task.conversation,
        { role: "user", text },
        {
          role: "assistant",
          text: "답변을 반영했습니다. 계획을 다시 확인해 주세요.",
        },
      ],
      step: 4,
      plan: task.plan ? { ...task.plan, questions: [] } : task.plan,
    });
  }
  async function exportBackup() {
    const b = await backup(store);
    download(
      b,
      `my-workmate-ai_백업_${new Date().toISOString().slice(0, 10)}.zip`,
    );
  }
  function download(b: Blob, n: string) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(b);
    a.download = n;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }
  if (!mounted) return <div className="loading">my workmate ai</div>;
  return (
    <div className="app">
      <aside>
        <div className="logo">
          <span>●</span>
          <b>my workmate ai</b>
        </div>
        {tabs.map(([v, icon, name]) => (
          <button
            key={v}
            className={view === v ? "active" : ""}
            onClick={() => go(v)}
          >
            <i>{icon}</i>
            {name}
            {v === "jobs" &&
              store.tasks.filter((x) => x.status === "draft").length > 0 && (
                <em>
                  {store.tasks.filter((x) => x.status === "draft").length}
                </em>
              )}
          </button>
        ))}
      </aside>
      <main>
        {view === "home" && (
          <Home
            jobs={store.tasks.filter((x) => x.status === "draft").length}
            onStart={() => newTask()}
            onJobs={() => go("jobs")}
          />
        )}{" "}
        {view === "help" && <Help />}
        {view === "files" && (
          <Files
            store={store}
            selected={selected}
            setSelected={setSelected}
            add={addFiles}
            start={() => newTask(selected)}
            remove={(id) => {
              setSelected((x) => x.filter((value) => value !== id));
              update((s) => {
                const item = s.files.find((x) => x.id === id);
                return item
                  ? {
                      ...s,
                      files: s.files.filter((x) => x.id !== id),
                      trash: { ...s.trash, files: [item, ...s.trash.files] },
                    }
                  : s;
              });
            }}
          />
        )}{" "}
        {view === "templates" && (
          <Templates
            store={store}
            add={(f) => addFiles(f, true)}
            start={(id) => newTask([], id)}
            remove={(id) =>
              update((s) => {
                const item = s.templates.find((x) => x.id === id);
                return item
                  ? {
                      ...s,
                      templates: s.templates.filter((x) => x.id !== id),
                      trash: {
                        ...s.trash,
                        templates: [item, ...s.trash.templates],
                      },
                    }
                  : s;
              })
            }
          />
        )}{" "}
        {view === "jobs" && (
          <Tasks
            title="진행 중 작업"
            items={store.tasks.filter((x) => x.status === "draft")}
            open={(t) => {
              setTaskId(t.id);
              go("task");
            }}
            remove={(id) =>
              update((s) => {
                const item = s.tasks.find((x) => x.id === id);
                return item
                  ? {
                      ...s,
                      tasks: s.tasks.filter((x) => x.id !== id),
                      trash: { ...s.trash, tasks: [item, ...s.trash.tasks] },
                    }
                  : s;
              })
            }
          />
        )}
        {view === "recent" && (
          <Tasks
            title="최근 작업"
            items={store.tasks}
            open={(t) => {
              setTaskId(t.id);
              go("task");
            }}
            duplicate={(source) => {
              const now = new Date().toISOString();
              const copied: Task = {
                id: crypto.randomUUID(),
                title: `${source.title} 다시 만들기`,
                step: 1,
                fileIds: [],
                templateId: source.templateId,
                request: source.request,
                outputFormat: source.outputFormat || "auto",
                conversation: [],
                createdAt: now,
                updatedAt: now,
                status: "draft",
              };
              update((s) => ({ ...s, tasks: [copied, ...s.tasks] }));
              setTaskId(copied.id);
              setSelected([]);
              go("task");
            }}
            remove={(id) =>
              update((s) => {
                const item = s.tasks.find((x) => x.id === id);
                return item
                  ? {
                      ...s,
                      tasks: s.tasks.filter((x) => x.id !== id),
                      trash: { ...s.trash, tasks: [item, ...s.trash.tasks] },
                    }
                  : s;
              })
            }
          />
        )}
        {view === "archive" && (
          <Archive
            store={store}
            restoreFile={(id) =>
              update((s) => {
                const item = s.trash.files.find((x) => x.id === id);
                return item
                  ? {
                      ...s,
                      files: [item, ...s.files],
                      trash: {
                        ...s.trash,
                        files: s.trash.files.filter((x) => x.id !== id),
                      },
                    }
                  : s;
              })
            }
            restoreTemplate={(id) =>
              update((s) => {
                const item = s.trash.templates.find((x) => x.id === id);
                return item
                  ? {
                      ...s,
                      templates: [item, ...s.templates],
                      trash: {
                        ...s.trash,
                        templates: s.trash.templates.filter((x) => x.id !== id),
                      },
                    }
                  : s;
              })
            }
            restoreTask={(id) =>
              update((s) => {
                const item = s.trash.tasks.find((x) => x.id === id);
                return item
                  ? {
                      ...s,
                      tasks: [item, ...s.tasks],
                      trash: {
                        ...s.trash,
                        tasks: s.trash.tasks.filter((x) => x.id !== id),
                      },
                    }
                  : s;
              })
            }
          />
        )}
        {view === "settings" && (
          <Settings
            store={store}
            backup={exportBackup}
            inspect={async (f) => {
              try {
                setRestorePreview(await restore(f));
              } catch (e: any) {
                setNotice(e.message);
              }
            }}
            preview={restorePreview}
            apply={() => {
              if (restorePreview) {
                setStore(restorePreview);
                setRestorePreview(null);
                setNotice("확인한 백업을 복원했습니다.");
              }
            }}
            restoreInternal={(id) => {
              const found = restoreSaved(id);
              if (found) setRestorePreview(found);
            }}
          />
        )}{" "}
        {view === "task" && task && (
          <TaskFlow
            task={task}
            store={store}
            selected={selected}
            setSelected={(ids) => {
              setSelected(ids);
              changeTask({ fileIds: ids });
            }}
            change={changeTask}
            plan={makePlan}
            answer={answer}
            listening={listening}
            startVoice={startVoice}
            stopVoice={stopVoice}
            generateStudents={makeStudentDrafts}
            changeDraft={changeDraft}
            downloadStudents={downloadStudentResult}
            execute={executeTask}
            downloadWork={downloadWorkResult}
            addFiles={addFiles}
          />
        )}
      </main>
      {view !== "task" && (
        <div className="history">
          <button onClick={() => history.back()}>← 이전</button>
          <button onClick={() => history.forward()}>다음 →</button>
        </div>
      )}
      {notice && (
        <div className="toast" onClick={() => setNotice("")}>
          {notice}
        </div>
      )}
      {pendingApi && (
        <div className="modalShade">
          <div className="consent">
            <h2>인공지능 분석 안내</h2>
            <p>이 작업은 OpenAI API를 사용합니다.</p>
            <p>
              이름과 개인정보는 자동으로 가린 뒤 작업에 필요한 내용만
              전송합니다.
            </p>
            {task &&
              store.files.some(
                (x) =>
                  task.fileIds.includes(x.id) &&
                  ["image", "audio"].includes(x.analysis?.kind || ""),
              ) && (
                <p className="mediaWarning">
                  사진과 음성은 원본 내용을 글로 바꾸기 위해 가리지 않은 파일이
                  전송됩니다.
                </p>
              )}
            <p>API 사용료가 발생할 수 있습니다.</p>
            <b>계속하시겠습니까?</b>
            <div>
              <button onClick={() => setPendingApi(null)}>취소</button>
              <button
                className="primary"
                onClick={pendingApi === "plan" ? runPlan : runStudentDrafts}
              >
                확인
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
function Home({
  jobs,
  onStart,
  onJobs,
}: {
  jobs: number;
  onStart: () => void;
  onJobs: () => void;
}) {
  return (
    <div className="home">
      <section className="guide">
        <img src="/assistant-character.png" />
        <div>
          <b>자료나 양식을 선택하고</b>
          <strong>만들고 싶은 결과를 글이나 말로 설명해 주세요.</strong>
          <span>제가 필요한 내용을 질문하고 작업계획을 먼저 보여드릴게요.</span>
        </div>
        <button onClick={onStart}>자료 없이 글·말로 바로 시작 →</button>
      </section>
      <div className="stepgrid">
        {steps.map((s, i) => (
          <div className="step" key={s[0]}>
            <em>{i + 1}</em>
            <b>{s[0]}</b>
            <span>{s[1]}</span>
            {i === 1 && <small>⌨ 글쓰기　🎙 마이크</small>}
          </div>
        ))}
      </div>
      <button className="jobsbar" onClick={onJobs}>
        <span>▤ 진행 중인 작업 {jobs}개</span>
        <b>확인하기 →</b>
      </button>
    </div>
  );
}
function Help() {
  return (
    <Section
      title="사용설명"
      sub="어디에서 시작해도 같은 1~6단계 작업으로 연결됩니다."
    >
      <div className="routes">
        <p>
          <b>내 자료에서</b> → 자료 선택 → 글·말로 설명 → 양식 선택(필요할 때) →
          계획 확인 → 결과
        </p>
        <p>
          <b>내 양식에서</b> → 양식 선택 → 자료·내용 추가 → 글·말로 설명 → 계획
          확인 → 결과
        </p>
        <p>
          <b>최근 작업에서</b> → 지난 작업 선택 → 새 자료 선택 → 수정 내용 설명
          → 계획 확인 → 결과
        </p>
        <p>
          <b>보관함에서</b> → 기존 결과 선택 → 수정·변환 설명 → 계획 확인 → 새
          결과
        </p>
        <p>
          <b>진행 중 작업에서</b> → 저장된 단계부터 계속
        </p>
      </div>
      <div className="note">
        <b>자료 없이 시작하면</b>
        <p>
          바로 설명 단계로 이동합니다. 필요한 자료나 양식이 있으면 다음 단계에서
          추가할 수 있습니다.
        </p>
      </div>
    </Section>
  );
}
function Files({
  store,
  selected,
  setSelected,
  add,
  start,
  remove,
}: {
  store: Store;
  selected: string[];
  setSelected: (x: string[]) => void;
  add: (f: FileList | null) => void;
  start: () => void;
  remove: (id: string) => void;
}) {
  return (
    <Section
      title={`내 자료 ${store.files.length}개`}
      sub="필요한 자료를 여러 개 선택한 뒤 작업을 시작하세요."
    >
      <label className="upload">
        ＋ 새 자료 추가
        <input type="file" multiple onChange={(e) => add(e.target.files)} />
      </label>
      <div className="filelist">
        {store.files.map((f) => (
          <label
            key={f.id}
            className={
              ["unsupported", "error"].includes(f.status) ? "unsupported" : ""
            }
          >
            <input
              type="checkbox"
              checked={selected.includes(f.id)}
              onChange={(e) =>
                setSelected(
                  e.target.checked
                    ? [...selected, f.id]
                    : selected.filter((x) => x !== f.id),
                )
              }
            />
            <div>
              <b>{f.name}</b>
              <span>
                {f.type} · {(f.size / 1024).toFixed(0)}KB ·{" "}
                {f.status === "ready"
                  ? "분석 완료"
                  : f.status === "partial"
                    ? "일부 분석"
                    : f.status === "error"
                      ? "분석 실패"
                      : "지원하지 않음"}
              </span>
              {f.analysis && <strong>{f.analysis.summary}</strong>}
              {f.analysis?.details.map((x) => (
                <small key={x}>{x}</small>
              ))}
              {f.analysis?.warnings.map((x) => (
                <small className="warn" key={x}>
                  {x}
                </small>
              ))}
            </div>
            <button
              className="trashButton"
              onClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
                remove(f.id);
              }}
            >
              삭제
            </button>
          </label>
        ))}
      </div>
      <div className="nextbar">
        <span>
          {selected.length}개 선택 / 전체 {store.files.length}개
        </span>
        <button disabled={!selected.length} onClick={start}>
          선택한 자료로 시작 →
        </button>
      </div>
    </Section>
  );
}
function Templates({
  store,
  add,
  start,
  remove,
}: {
  store: Store;
  add: (f: FileList | null) => void;
  start: (id: string) => void;
  remove: (id: string) => void;
}) {
  const [chosen, setChosen] = useState("");
  const selectedTemplate = store.templates.find((x) => x.id === chosen);
  return (
    <Section
      title="내 양식"
      sub="자주 쓰는 Word·Excel 양식을 등록하고 해당 양식에서 작업을 시작하세요."
    >
      <label className="upload">
        ＋ 양식 등록
        <input type="file" multiple onChange={(e) => add(e.target.files)} />
      </label>
      <div className="cards">
        {store.templates.map((t) => (
          <label className={chosen === t.id ? "templatePicked" : ""} key={t.id}>
            <input
              type="radio"
              name="template"
              checked={chosen === t.id}
              onChange={() => setChosen(t.id)}
            />
            <b>{t.name}</b>
            <span>
              {t.type} · {t.analysis?.summary || "구조 확인 필요"}
            </span>
            <button
              className="trashButton"
              onClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
                if (chosen === t.id) setChosen("");
                remove(t.id);
              }}
            >
              삭제
            </button>
          </label>
        ))}
      </div>
      {selectedTemplate?.analysis && (
        <div className="templatePreview">
          <b>선택한 양식 구조</b>
          {selectedTemplate.analysis.details.map((x) => (
            <p key={x}>{x}</p>
          ))}
          {selectedTemplate.analysis.warnings.map((x) => (
            <p className="warn" key={x}>
              {x}
            </p>
          ))}
        </div>
      )}
      {!!store.templates.length && (
        <div className="nextbar">
          <span>{chosen ? "양식 1개 선택됨" : "사용할 양식을 체크하세요"}</span>
          <button disabled={!chosen} onClick={() => start(chosen)}>
            다음 →
          </button>
        </div>
      )}
      {!store.templates.length && <Empty text="등록된 양식이 없습니다." />}
    </Section>
  );
}
function Tasks({
  title,
  items,
  open,
  duplicate,
  remove,
}: {
  title: string;
  items: Task[];
  open: (t: Task) => void;
  duplicate?: (t: Task) => void;
  remove?: (id: string) => void;
}) {
  return (
    <Section
      title={title}
      sub="저장된 작업을 선택하면 마지막 단계부터 계속합니다."
    >
      {items.map((t) => (
        <button className="taskrow" key={t.id} onClick={() => open(t)}>
          <div>
            <b>{t.title}</b>
            <span>
              {steps[Math.max(0, t.step - 1)][0]} ·{" "}
              {new Date(t.updatedAt).toLocaleString("ko-KR")}
            </span>
          </div>
          <strong>계속하기 →</strong>
          {duplicate && t.status === "completed" && (
            <span
              className="repeatButton"
              role="button"
              tabIndex={0}
              onClick={(event) => {
                event.stopPropagation();
                duplicate(t);
              }}
            >
              지난번처럼 새로 시작
            </span>
          )}
          {remove && (
            <span
              className="trashButton"
              role="button"
              tabIndex={0}
              onClick={(event) => {
                event.stopPropagation();
                remove(t.id);
              }}
            >
              삭제
            </span>
          )}
        </button>
      ))}
      {!items.length && <Empty text="저장된 작업이 없습니다." />}
    </Section>
  );
}
function Archive({
  store,
  restoreFile,
  restoreTemplate,
  restoreTask,
}: {
  store: Store;
  restoreFile: (id: string) => void;
  restoreTemplate: (id: string) => void;
  restoreTask: (id: string) => void;
}) {
  const empty =
    !store.trash.files.length &&
    !store.trash.templates.length &&
    !store.trash.tasks.length &&
    !store.archive.length;
  return (
    <Section
      title="보관함"
      sub="삭제한 자료와 양식을 확인한 뒤 원래 위치로 복원할 수 있습니다."
    >
      {store.trash.files.map((item) => (
        <div className="archiveRow" key={item.id}>
          <span>
            <b>자료</b> {item.name}
          </span>
          <button onClick={() => restoreFile(item.id)}>복원</button>
        </div>
      ))}
      {store.trash.templates.map((item) => (
        <div className="archiveRow" key={item.id}>
          <span>
            <b>양식</b> {item.name}
          </span>
          <button onClick={() => restoreTemplate(item.id)}>복원</button>
        </div>
      ))}
      {store.trash.tasks.map((item) => (
        <div className="archiveRow" key={item.id}>
          <span>
            <b>작업</b> {item.title}
          </span>
          <button onClick={() => restoreTask(item.id)}>복원</button>
        </div>
      ))}
      {store.archive.map((item) => (
        <div className="archiveRow" key={item.id}>
          <span>
            <b>완료 작업</b> {item.title}
          </span>
        </div>
      ))}
      {empty && <Empty text="보관된 항목이 없습니다." />}
    </Section>
  );
}
function Settings({
  store,
  backup,
  inspect,
  preview,
  apply,
  restoreInternal,
}: {
  store: Store;
  backup: () => void;
  inspect: (f: File) => void;
  preview: Store | null;
  apply: () => void;
  restoreInternal: (id: string) => void;
}) {
  const [refresh, setRefresh] = useState(0);
  const [aiReady, setAiReady] = useState<boolean | null>(null);
  useEffect(() => {
    let active = true;
    fetch("/api/status")
      .then((response) => response.json())
      .then((data) => active && setAiReady(Boolean(data.ready)))
      .catch(() => active && setAiReady(false));
    return () => {
      active = false;
    };
  }, []);
  const backups = savedBackups();
  return (
    <Section title="설정" sub="자동저장과 음성 및 백업을 관리합니다.">
      <div className="setting">
        <b>인공지능</b>
        <span>
          {aiReady === null
            ? "연결 상태를 확인하고 있습니다."
            : aiReady
              ? "글과 자료를 이해하고 결과를 만들 준비가 되었습니다."
              : "아직 연결 전입니다. 연결 전에도 자료 분석과 저장 및 백업은 사용할 수 있습니다."}
        </span>
        <i>{aiReady ? "사용 준비됨" : "연결 전"}</i>
      </div>
      <div className="setting">
        <b>자동저장</b>
        <span>작업 내용은 이 브라우저에 자동으로 저장됩니다.</span>
        <i>사용 중</i>
      </div>
      <div className="setting block">
        <b>백업</b>
        <span>자료 목록과 양식 및 작업 내용을 ZIP으로 보관합니다.</span>
        <div>
          <button onClick={backup}>백업하기</button>
          <button
            onClick={() => {
              saveImportant(
                store,
                `중요 백업 ${new Date().toLocaleString("ko-KR")}`,
              );
              setRefresh(refresh + 1);
            }}
          >
            중요 백업
          </button>
          <label>
            백업 가져오기
            <input
              type="file"
              accept=".zip"
              onChange={(e) =>
                e.target.files?.[0] && inspect(e.target.files[0])
              }
            />
          </label>
        </div>
        <p>
          최근 자동백업 {backups.auto.length}/4개 · 중요 백업{" "}
          {backups.important.length}개
        </p>
        {[...backups.auto, ...backups.important].map((x) => (
          <button key={x.id} onClick={() => restoreInternal(x.id)}>
            {x.name} · {new Date(x.createdAt).toLocaleString("ko-KR")}
          </button>
        ))}
      </div>
      {preview && (
        <div className="restore">
          <b>복원 미리보기</b>
          <p>
            자료 {preview.files.length}개 · 양식 {preview.templates.length}개 ·
            작업 {preview.tasks.length}개
          </p>
          {!!preview.files.length && (
            <p>
              자료:{" "}
              {preview.files
                .slice(0, 5)
                .map((x) => x.name)
                .join(" · ")}
            </p>
          )}
          {!!preview.templates.length && (
            <p>
              양식:{" "}
              {preview.templates
                .slice(0, 5)
                .map((x) => x.name)
                .join(" · ")}
            </p>
          )}
          {!!preview.tasks.length && (
            <p>
              작업:{" "}
              {preview.tasks
                .slice(0, 5)
                .map((x) => x.title)
                .join(" · ")}
            </p>
          )}
          <button onClick={apply}>확인하고 복원</button>
        </div>
      )}
      <div className="setting">
        <b>음성</b>
        <span>
          Chrome·Edge에서 사용자가 완료할 때까지 연속 받아쓰기를 시도합니다.
        </span>
      </div>
    </Section>
  );
}
function TaskFlow({
  task,
  store,
  selected,
  setSelected,
  change,
  plan,
  answer,
  listening,
  startVoice,
  stopVoice,
  generateStudents,
  changeDraft,
  downloadStudents,
  execute,
  downloadWork,
  addFiles,
}: {
  task: Task;
  store: Store;
  selected: string[];
  setSelected: (x: string[]) => void;
  change: (p: Partial<Task>) => void;
  plan: () => void;
  answer: (x: string) => void;
  listening: boolean;
  startVoice: () => void;
  stopVoice: () => void;
  generateStudents: () => void;
  changeDraft: (i: number, p: Partial<StudentDraft>) => void;
  downloadStudents: () => void;
  execute: () => void;
  downloadWork: () => void;
  addFiles: (list: FileList | null, asTemplate?: boolean) => void;
}) {
  const [q, setQ] = useState("");
  const isStudent = /생기부|생활기록부/.test(task.request);
  const count = store.files
    .filter((x) => task.fileIds.includes(x.id))
    .reduce((n, x) => n + (x.students?.length || 0), 0);
  const next = () => {
    if (task.step === 1) change({ step: 2 });
    else if (task.step === 2) plan();
    else if (task.step === 3) {
      if (!q.trim()) return;
      answer(q);
      setQ("");
    } else if (task.step === 4) change({ step: 5 });
    else if (task.step === 5) {
      if (isStudent) generateStudents();
      else execute();
    }
  };
  const nextDisabled =
    (task.step === 2 && !task.request.trim()) ||
    (task.step === 3 && !q.trim()) ||
    (task.step === 5 && (isStudent ? !count : !task.plan?.canExecute));
  return (
    <Section
      title={task.title}
      sub={`${task.step}/6 ${steps[task.step - 1][0]}`}
    >
      {task.step === 1 && (
        <>
          <h2>사용할 자료를 선택하세요</h2>
          <div className="filelist">
            {store.files.map((f) => (
              <label key={f.id}>
                <input
                  type="checkbox"
                  checked={selected.includes(f.id)}
                  onChange={(e) =>
                    setSelected(
                      e.target.checked
                        ? [...selected, f.id]
                        : selected.filter((x) => x !== f.id),
                    )
                  }
                />
                <b>{f.name}</b>
                {f.students && (
                  <span>학생 자료 {f.students.length}명 확인</span>
                )}
              </label>
            ))}
          </div>
        </>
      )}
      {task.step === 2 && (
        <>
          <h2>원하는 일을 설명해 주세요</h2>
          <div className="chosen">
            선택 자료 {task.fileIds.length}개 · 선택 양식{" "}
            {task.templateId ? "1개" : "없음"}
          </div>
          <textarea
            value={task.request}
            onChange={(e) =>
              change({
                request: e.target.value,
                title: e.target.value.slice(0, 28) || "새 작업",
              })
            }
            placeholder="예: 학생 소감문을 학생 이름별로 구분하고 서로 다른 생기부 문구 초안으로 정리해줘."
          />
          <div className="voicebar">
            {listening ? (
              <button className="stop" onClick={stopVoice}>
                ■ 받아쓰기 완료
              </button>
            ) : (
              <button onClick={startVoice}>🎙 계속 말하기</button>
            )}
            <label>
              음성파일 추가
              <input
                type="file"
                accept="audio/*"
                onChange={(e) => addFiles(e.target.files)}
              />
            </label>
            <span>
              {listening
                ? "듣는 중 · 잠시 멈춰도 다시 시작합니다."
                : "글쓰기와 말하기를 함께 사용할 수 있습니다."}
            </span>
          </div>
        </>
      )}
      {task.step === 3 && (
        <>
          <h2>AI 질문에 답해 주세요</h2>
          <div className="chat">
            {task.conversation.map((m, i) => (
              <p className={m.role} key={i}>
                {m.text}
              </p>
            ))}
          </div>
          <textarea
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="답변을 글로 적거나 마이크로 말하세요."
          />
        </>
      )}
      {task.step === 4 && (
        <>
          <PlanView p={task.plan} />
          <label className="formatChoice">
            <b>결과 파일 형식</b>
            <select
              value={task.outputFormat || "auto"}
              onChange={(event) =>
                change({
                  outputFormat: event.target.value as Task["outputFormat"],
                })
              }
            >
              <option value="auto">요청에 맞게 자동 선택</option>
              <option value="xlsx">Excel 통합문서</option>
              <option value="docx">Word 문서</option>
              <option value="pdf">PDF 문서</option>
              <option value="pptx">PowerPoint 발표자료</option>
              <option value="txt">텍스트 문서</option>
              <option value="csv">CSV 표</option>
            </select>
          </label>
        </>
      )}{" "}
      {task.step === 5 && (
        <div className="execute">
          <h2>실행·검증</h2>
          {isStudent ? (
            <div className="studentStart">
              <b>학생별 두 가지 생기부 초안</b>
              <p>
                선택한 자료에서 이름과 소감문이 확인된 학생은 {count}명입니다.
              </p>
              <ul>
                <li>1안은 학생 글에서 확인되는 사실만 사용합니다.</li>
                <li>
                  2안은 학생 글에서 연결 가능한 관찰 가능성을 포함하고 유추
                  부분을 표시합니다.
                </li>
                <li>
                  쉼표는 사용하지 않으며 모든 완결 문장을 마침표로 끝냅니다.
                </li>
              </ul>
            </div>
          ) : task.plan?.canExecute ? (
            <>
              <p>작업 사본에서 결과를 생성하고 누락과 오류를 검사합니다.</p>
            </>
          ) : (
            <div className="warning">
              <b>현재 요청 실행 보류</b>
              <p>
                {task.plan?.limitation || "현재 이 요청을 실행할 수 없습니다."}
              </p>
            </div>
          )}
        </div>
      )}
      {task.step === 6 &&
        (task.studentDrafts?.length ? (
          <StudentReview
            drafts={task.studentDrafts}
            validation={task.studentValidation || []}
            change={changeDraft}
            download={downloadStudents}
          />
        ) : (
          <div className="result">
            <h2>결과 확인</h2>
            <h3>{task.workResult?.title || task.result}</h3>
            {task.workResult?.text && <p>{task.workResult.text}</p>}
            {task.workResult?.kind === "table" && (
              <div className="tablePreview">
                <table>
                  <thead>
                    <tr>
                      {task.workResult.columns?.map((x) => (
                        <th key={x}>{x}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {task.workResult.rows?.slice(0, 20).map((row, i) => (
                      <tr key={i}>
                        {row.map((x, j) => (
                          <td key={j}>{x}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            {task.workResult?.validation.map((x) => (
              <p className="good" key={x}>
                ✓ {x}
              </p>
            ))}
            {task.workResult?.warnings.map((x) => (
              <p className="warn" key={x}>
                확인 필요 · {x}
              </p>
            ))}
            <button onClick={() => change({ step: 2, status: "draft" })}>
              수정 요청
            </button>
            <button onClick={downloadWork}>다운로드</button>
          </div>
        ))}
      <div className="workNav">
        <button
          disabled={task.step === 1}
          onClick={() => change({ step: Math.max(1, task.step - 1) })}
        >
          ← 이전
        </button>
        {task.step < 6 && (
          <button className="primary" disabled={nextDisabled} onClick={next}>
            다음 →
          </button>
        )}
      </div>
    </Section>
  );
}

function StudentReview({
  drafts,
  validation,
  change,
  download,
}: {
  drafts: StudentDraft[];
  validation: string[];
  change: (i: number, p: Partial<StudentDraft>) => void;
  download: () => void;
}) {
  const [i, setI] = useState(0),
    d = drafts[i];
  const reviewed = drafts.filter((x) => x.reviewed).length;
  const choose = (selected: "fact" | "inferred" | "merged") => {
    const finalText =
      selected === "fact"
        ? d.factDraft
        : selected === "inferred"
          ? d.inferredDraft
          : `${d.factDraft} ${d.inferredDraft}`;
    change(i, { selected, finalText });
  };
  return (
    <div className="studentReview">
      <div className="reviewTop">
        <div>
          <h2>학생별 두 가지 초안</h2>
          <span>
            {reviewed}/{drafts.length}명 확인 완료
          </span>
        </div>
        <select value={i} onChange={(e) => setI(Number(e.target.value))}>
          {drafts.map((x, n) => (
            <option key={n} value={n}>
              {n + 1}. {x.name}
              {x.reviewed ? " ✓" : ""}
            </option>
          ))}
        </select>
      </div>
      <div className={validation.length ? "warning" : "validationOk"}>
        <b>전체 학생 문장 중복 검사</b>
        {validation.length ? (
          validation.map((item) => <p key={item}>{item}</p>)
        ) : (
          <p>높은 유사도의 문장 쌍이 발견되지 않았습니다.</p>
        )}
      </div>
      <div className="sourceBox">
        <b>{d.name} 학생 원문</b>
        <p>{d.source}</p>
      </div>
      <div className="draftChoices">
        <article className={d.selected === "fact" ? "picked" : ""}>
          <h3>1안 사실 중심</h3>
          <p>{d.factDraft}</p>
          <button onClick={() => choose("fact")}>1안 선택</button>
        </article>
        <article className={d.selected === "inferred" ? "picked" : ""}>
          <h3>2안 관찰 가능성 포함</h3>
          <p>{d.inferredDraft}</p>
          {d.inferredParts?.length > 0 && (
            <div className="inference">
              <b>교사 확인이 필요한 유추</b>
              {d.inferredParts.map((x, n) => (
                <span key={n}>{x}</span>
              ))}
            </div>
          )}
          <button onClick={() => choose("inferred")}>2안 선택</button>
        </article>
      </div>
      <button onClick={() => choose("merged")}>두 초안 합쳐서 수정</button>
      <h3>교사 최종 수정본</h3>
      <textarea
        value={d.finalText}
        onChange={(e) =>
          change(i, { finalText: e.target.value, reviewed: false })
        }
      />
      <div className="validation">
        <span className={d.finalText.includes(",") ? "bad" : "good"}>
          {d.finalText.includes(",") ? "쉼표가 있습니다" : "쉼표 없음"}
        </span>
        <span className={d.finalText.trim().endsWith(".") ? "good" : "bad"}>
          {d.finalText.trim().endsWith(".") ? "마침표 확인" : "마침표 필요"}
        </span>
      </div>
      <div className="reviewNav">
        <button disabled={i === 0} onClick={() => setI(i - 1)}>
          ← 이전 학생
        </button>
        <button
          className="primary"
          disabled={
            d.finalText.includes(",") || !d.finalText.trim().endsWith(".")
          }
          onClick={() => {
            change(i, { reviewed: true });
            if (i < drafts.length - 1) setI(i + 1);
          }}
        >
          확인하고 다음 학생 →
        </button>
      </div>
      <div className="downloadBar">
        <b>
          전체 {drafts.length}명 중 {reviewed}명 확인
        </b>
        <button disabled={reviewed !== drafts.length} onClick={download}>
          교사 수정본 다운로드
        </button>
      </div>
    </div>
  );
}
function PlanView({ p }: { p?: Plan }) {
  if (!p) return <Empty text="작업계획이 없습니다." />;
  return (
    <div className="plan">
      <h2>작업계획을 확인하세요</h2>
      <dl>
        <dt>이해한 요청</dt>
        <dd>{p.understanding}</dd>
        <dt>사용 자료</dt>
        <dd>{p.materials.join(" · ") || "없음"}</dd>
        <dt>사용 양식</dt>
        <dd>{p.template}</dd>
        <dt>결과 형식</dt>
        <dd>{p.resultFormat}</dd>
        <dt>처리 단계</dt>
        <dd>{p.steps.join(" → ")}</dd>
      </dl>
      {p.limitation && <p className="warning">{p.limitation}</p>}
    </div>
  );
}
function Section({
  title,
  sub,
  children,
}: {
  title: string;
  sub: string;
  children: React.ReactNode;
}) {
  return (
    <section className="section">
      <h1>{title}</h1>
      <p className="sub">{sub}</p>
      {children}
    </section>
  );
}
function Empty({ text }: { text: string }) {
  return <div className="empty">{text}</div>;
}
