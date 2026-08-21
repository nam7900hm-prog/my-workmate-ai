const OFFICIAL = "https://api.openai.com/v1";

export function openAIUrl(path: string) {
  const base =
    process.env.OPENAI_TEST_MODE === "true" &&
    process.env.OPENAI_TEST_BASE_URL?.startsWith("http://127.0.0.1:")
      ? process.env.OPENAI_TEST_BASE_URL
      : OFFICIAL;
  return `${base.replace(/\/$/, "")}/${path.replace(/^\//, "")}`;
}
