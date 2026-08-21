export function parseOpenAIJson<T>(data: unknown, label: string): T {
  const text =
    data && typeof data === "object" && "output_text" in data
      ? (data as { output_text?: unknown }).output_text
      : undefined;
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
