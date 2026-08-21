import JSZip from "jszip";
import * as XLSX from "xlsx";
import type { FileAnalysis, StudentSource } from "./model";

const clean = (s: string) =>
  s
    .replace(/\u0000/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
const xmlText = (s: string) =>
  clean(
    s
      .replace(/<[^>]+>/g, " ")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&"),
  );
const now = () => new Date().toISOString();

function findStudents(rows: Record<string, unknown>[]): StudentSource[] {
  if (!rows.length) return [];
  const keys = Object.keys(rows[0]);
  const nameKey = keys.find((k) => /(학생명|성명|이름)/.test(k));
  const textKey = keys.find(
    (k) =>
      /(소감문|소감|활동내용|활동 내용|내용|학생글|학생 글)/.test(k) &&
      k !== nameKey,
  );
  if (!nameKey || !textKey) return [];
  return rows
    .map((r) => ({
      name: String(r[nameKey] ?? "").trim(),
      text: String(r[textKey] ?? "").trim(),
    }))
    .filter((x) => x.name && x.text)
    .slice(0, 500);
}

function tableHeaderIndex(grid: any[][]) {
  let best = 0;
  let bestScore = -1;
  grid.slice(0, 20).forEach((row, index) => {
    const values = row
      .map((value) => String(value ?? "").trim())
      .filter(Boolean);
    const unique = new Set(values).size;
    const words = values.filter(
      (value) => !/^[-+]?\d+(?:\.\d+)?$/.test(value),
    ).length;
    const keywords = values.filter((value) =>
      /(이름|성명|번호|코드|날짜|부서|수량|금액|시간|과목|담당|내용|구분)/.test(
        value,
      ),
    ).length;
    const score = values.length + unique + words + keywords * 4;
    if (values.length >= 2 && score > bestScore) {
      best = index;
      bestScore = score;
    }
  });
  return best;
}

function objectRows(grid: any[][], headerRow: number) {
  const headers = (grid[headerRow] || []).map((value) =>
    String(value ?? "").trim(),
  );
  return grid
    .slice(headerRow + 1)
    .map((row) =>
      Object.fromEntries(
        headers
          .map((header, index) => [header, row[index]] as const)
          .filter(([header]) => Boolean(header)),
      ),
    );
}

export async function analyzeFile(
  file: File,
): Promise<{ analysis: FileAnalysis; students?: StudentSource[] }> {
  const ext = file.name.split(".").pop()?.toLowerCase() || "";
  if (["xlsx", "xls", "csv"].includes(ext)) {
    const wb = XLSX.read(await file.arrayBuffer(), {
      type: "array",
      cellStyles: true,
      cellFormula: true,
      cellDates: true,
    });
    const allStudents: StudentSource[] = [];
    const personalNames = new Set<string>();
    const allTexts: string[] = [];
    const headerOwners = new Map<string, string[]>();
    const formulaRelations = new Set<string>();
    const sheets = wb.SheetNames.map((name) => {
      const ws = wb.Sheets[name];
      const ref = ws["!ref"] || "A1:A1";
      const range = XLSX.utils.decode_range(ref);
      const grid = XLSX.utils.sheet_to_json<any[]>(ws, {
        header: 1,
        defval: "",
        raw: true,
      });
      const headerRow = tableHeaderIndex(grid);
      const headers = (grid[headerRow] || []).map(String).filter(Boolean);
      headers.forEach((header) => {
        const key = header.trim();
        if (key)
          headerOwners.set(key, [...(headerOwners.get(key) || []), name]);
      });
      let formulaCount = 0;
      let dateCount = 0;
      let numberCount = 0;
      let textCount = 0;
      let styledCellCount = 0;
      const frequency = new Map<string, number>();
      Object.keys(ws).forEach((k) => {
        if (k.startsWith("!")) return;
        const cell = ws[k] as any;
        if (cell?.f) {
          formulaCount++;
          for (const match of String(cell.f).matchAll(
            /(?:'([^']+)'|([\p{L}\p{N}_ ]+))!/gu,
          )) {
            const target = (match[1] || match[2] || "").trim();
            if (target && target !== name)
              formulaRelations.add(`${name} → ${target}`);
          }
        }
        if (cell?.s !== undefined || cell?.z) styledCellCount++;
        const value = cell?.v;
        if (value instanceof Date || cell?.t === "d") dateCount++;
        else if (typeof value === "number") numberCount++;
        else if (typeof value === "string" && value.trim()) {
          textCount++;
          const normalized = value.trim();
          if (normalized.length >= 2)
            frequency.set(normalized, (frequency.get(normalized) || 0) + 1);
        }
      });
      const titleCandidates = grid
        .slice(0, Math.max(1, headerRow + 1))
        .flatMap((row) => row.slice(0, 8))
        .map((x) => String(x).trim())
        .filter((x) => x.length >= 4)
        .slice(0, 8);
      const repeatedValues = [...frequency.entries()]
        .filter(([, count]) => count > 1)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8)
        .map(([value, count]) => `${value}(${count}회)`);
      const objects = objectRows(grid, headerRow);
      const nameHeaders = headers.filter((header) =>
        /^(학생명|성명|이름|직원명|교사명|담당자|작성자)$/.test(header.trim()),
      );
      for (const row of objects)
        for (const header of nameHeaders) {
          const value = String(row[header] ?? "").trim();
          if (value && value.length <= 30) personalNames.add(value);
        }
      allStudents.push(...findStudents(objects));
      allTexts.push(`[시트: ${name}]\n${XLSX.utils.sheet_to_csv(ws)}`);
      return {
        name,
        range: ref,
        rows: range.e.r - range.s.r + 1,
        columns: range.e.c - range.s.c + 1,
        formulaCount,
        mergedCount: ws["!merges"]?.length || 0,
        headers,
        tableAreas: [
          XLSX.utils.encode_range({
            s: { r: headerRow, c: range.s.c },
            e: range.e,
          }),
        ],
        titleCandidates,
        dateCount,
        numberCount,
        textCount,
        styledCellCount,
        repeatedValues,
      };
    });
    const sharedHeaders = [...headerOwners.entries()]
      .filter(([, owners]) => new Set(owners).size > 1)
      .map(
        ([header, owners]) => `${header}: ${[...new Set(owners)].join(" ↔ ")}`,
      );
    const relationships = [...formulaRelations, ...sharedHeaders].slice(0, 30);
    const text = clean(allTexts.join("\n\n").slice(0, 160000));
    return {
      analysis: {
        kind: "spreadsheet",
        summary: `시트 ${sheets.length}개와 사용 범위를 분석했습니다.`,
        text,
        details: [
          ...sheets.map(
            (s) =>
              `${s.name}: 표 영역 ${s.range} · ${s.rows}행 · ${s.columns}열 · 날짜 ${s.dateCount}개 · 숫자 ${s.numberCount}개 · 문자 ${s.textCount}개 · 수식 ${s.formulaCount}개 · 서식 ${s.styledCellCount}개 · 병합 ${s.mergedCount}개`,
          ),
          relationships.length
            ? `자료 관계: ${relationships.join(" · ")}`
            : "시트 사이의 직접 관계는 확인되지 않았습니다.",
        ],
        sheets,
        relationships,
        personalNames: [...personalNames].slice(0, 1000),
        warnings: [],
        analyzedAt: now(),
      },
      students: allStudents.length ? allStudents.slice(0, 500) : undefined,
    };
  }
  if (ext === "txt") {
    const text = clean(await file.text());
    return {
      analysis: {
        kind: "text",
        summary: `텍스트 ${text.length.toLocaleString()}자를 읽었습니다.`,
        text: text.slice(0, 120000),
        details: [`문단 ${text.split(/\n\s*\n/).filter(Boolean).length}개`],
        warnings: [],
        analyzedAt: now(),
      },
    };
  }
  if (ext === "docx") {
    const mammoth = await import("mammoth/mammoth.browser");
    const result = await mammoth.extractRawText({
      arrayBuffer: await file.arrayBuffer(),
    });
    const text = clean(result.value);
    return {
      analysis: {
        kind: "word",
        summary: `Word 본문 ${text.length.toLocaleString()}자를 읽었습니다.`,
        text: text.slice(0, 120000),
        details: [`변환 알림 ${result.messages.length}개`],
        warnings: result.messages.map((x) => x.message),
        analyzedAt: now(),
      },
    };
  }
  if (ext === "pdf") {
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    const pdf = await pdfjs.getDocument({
      data: new Uint8Array(await file.arrayBuffer()),
    }).promise;
    const pages: string[] = [];
    let extractedCharacters = 0;
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const pageText = content.items.map((x: any) => x.str || "").join(" ");
      extractedCharacters += pageText.replace(/\s/g, "").length;
      pages.push(`[${i}쪽] ${pageText}`);
    }
    const text = clean(pages.join("\n"));
    return {
      analysis: {
        kind: "pdf",
        summary: `PDF ${pdf.numPages}쪽의 글자를 읽었습니다.`,
        text: text.slice(0, 160000),
        details: [`전체 ${pdf.numPages}쪽`],
        warnings: extractedCharacters
          ? []
          : ["글자가 없는 스캔 PDF는 사진 문자 인식이 필요합니다."],
        analyzedAt: now(),
      },
    };
  }
  if (["pptx", "hwpx"].includes(ext)) {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const pattern =
      ext === "pptx"
        ? /^ppt\/slides\/slide\d+\.xml$/
        : /^Contents\/section\d+\.xml$/i;
    const names = Object.keys(zip.files)
      .filter((x) => pattern.test(x))
      .sort();
    const parts = [] as string[];
    for (const n of names) {
      const raw = await zip.file(n)?.async("string");
      if (raw) parts.push(xmlText(raw));
    }
    const text = clean(parts.join("\n"));
    return {
      analysis: {
        kind: ext,
        summary: `${ext.toUpperCase()} 문서의 본문 구역 ${names.length}개를 읽었습니다.`,
        text: text.slice(0, 120000),
        details: [`본문 구역 ${names.length}개`],
        warnings: [],
        analyzedAt: now(),
      },
    };
  }
  if (["jpg", "jpeg", "png", "webp"].includes(ext))
    return {
      analysis: {
        kind: "image",
        summary: "사진은 등록되었습니다.",
        text: "",
        details: [],
        warnings: ["글자와 표 분석에는 인공지능 연결이 필요합니다."],
        analyzedAt: now(),
      },
    };
  if (["mp3", "m4a", "wav", "webm"].includes(ext))
    return {
      analysis: {
        kind: "audio",
        summary: "음성파일은 등록되었습니다.",
        text: "",
        details: [],
        warnings: ["음성 내용을 글로 바꾸려면 인공지능 연결이 필요합니다."],
        analyzedAt: now(),
      },
    };
  return {
    analysis: {
      kind: "unsupported",
      summary: "현재 내용을 읽을 수 없는 형식입니다.",
      text: "",
      details: [],
      warnings: ["다른 형식으로 변환해 주세요."],
      analyzedAt: now(),
    },
  };
}
