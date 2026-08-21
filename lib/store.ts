import JSZip from "jszip";
import type { Store } from "./model";
import { initialStore } from "./model";
import { getFileBlob, saveFileBlob } from "./file-db";
const KEY = "my_workmate_ai_v4_store";
const AUTO = "my_workmate_ai_v4_auto_backups";
const IMPORTANT = "my_workmate_ai_v4_important_backups";
export type SavedBackup = {
  id: string;
  name: string;
  createdAt: string;
  store: Store;
};
const readList = (key: string): SavedBackup[] => {
  try {
    return JSON.parse(localStorage.getItem(key) || "[]");
  } catch {
    return [];
  }
};
export function load(): Store {
  try {
    const value = JSON.parse(localStorage.getItem(KEY) || "{}");
    return {
      ...initialStore,
      ...value,
      trash: { ...initialStore.trash, ...(value.trash || {}) },
    };
  } catch {
    return initialStore;
  }
}
export function save(s: Store) {
  localStorage.setItem(KEY, JSON.stringify(s));
  const list = readList(AUTO);
  const last = list[0] && new Date(list[0].createdAt).getTime();
  if (!last || Date.now() - last > 120000) {
    list.unshift({
      id: crypto.randomUUID(),
      name: "자동백업",
      createdAt: new Date().toISOString(),
      store: s,
    });
    localStorage.setItem(AUTO, JSON.stringify(list.slice(0, 4)));
  }
}
export function savedBackups() {
  return { auto: readList(AUTO), important: readList(IMPORTANT) };
}
export function saveImportant(s: Store, name: string) {
  const list = readList(IMPORTANT);
  list.unshift({
    id: crypto.randomUUID(),
    name: name.trim() || "중요 백업",
    createdAt: new Date().toISOString(),
    store: s,
  });
  localStorage.setItem(IMPORTANT, JSON.stringify(list));
}
export function restoreSaved(id: string) {
  return [...readList(AUTO), ...readList(IMPORTANT)].find((x) => x.id === id)
    ?.store;
}
export async function fileToData(file: File) {
  return new Promise<string>((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result));
    r.onerror = () => reject(r.error);
    r.readAsDataURL(file);
  });
}
export async function backup(s: Store) {
  const z = new JSZip();
  z.file(
    "manifest.json",
    JSON.stringify(
      {
        app: "my workmate ai",
        version: 4,
        createdAt: new Date().toISOString(),
      },
      null,
      2,
    ),
  );
  z.file("workspace.json", JSON.stringify(s));
  for (const item of [
    ...s.files,
    ...s.templates,
    ...s.trash.files,
    ...s.trash.templates,
  ]) {
    const blob = await getFileBlob(item.id);
    if (blob) z.file(`originals/${item.id}`, await blob.arrayBuffer());
  }
  return z.generateAsync({ type: "blob" });
}
export async function restore(file: File) {
  const z = await JSZip.loadAsync(await file.arrayBuffer());
  const manifestRaw = await z.file("manifest.json")?.async("string");
  const raw = await z.file("workspace.json")?.async("string");
  if (!raw || !manifestRaw) throw new Error("올바른 백업 ZIP이 아닙니다.");
  const manifest = JSON.parse(manifestRaw);
  if (manifest.app !== "my workmate ai")
    throw new Error("다른 앱의 백업입니다.");
  const value = JSON.parse(raw);
  const restored = {
    ...initialStore,
    ...value,
    trash: { ...initialStore.trash, ...(value.trash || {}) },
  } as Store;
  for (const item of [
    ...restored.files,
    ...restored.templates,
    ...restored.trash.files,
    ...restored.trash.templates,
  ]) {
    const original = z.file(`originals/${item.id}`);
    if (original) await saveFileBlob(item.id, await original.async("blob"));
  }
  return restored;
}
