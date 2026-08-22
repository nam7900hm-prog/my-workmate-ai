export function extractOpenAIText(data: unknown) {
  const response = data as {
    output_text?: unknown;
    output?: Array<{
      content?: Array<{ type?: string; text?: unknown; refusal?: unknown }>;
    }>;
  } | null;
  const direct = response?.output_text;
  const nested = response?.output
    ?.flatMap((item) => item.content || [])
    .filter((item) => item.type === "output_text" || typeof item.text === "string")
    .map((item) => item.text)
    .filter((item): item is string => typeof item === "string")
    .join("");
  const text = typeof direct === "string" && direct.trim() ? direct : nested;
  return typeof text === "string" ? text : "";
}

export function parseOpenAIJson<T>(data: unknown, label: string): T {
  const text = extractOpenAIText(data);
  if (typeof text !== "string" || !text.trim())
    throw new Error(`${label} 응답에 결과 내용이 없습니다.`);
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new Error(`${label} 응답 형식을 읽지 못했습니다.`);
  }
}

export async function openAIError(response: Response, fallback: string) {
  try {
    const data = await response.json();
    const message = data?.error?.message;
    return typeof message === "string" && message.trim() ? message : fallback;
  } catch {
    return fallback;
  }
}
