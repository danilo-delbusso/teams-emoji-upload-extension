import { UploadState } from "./types";

// The popup's file list, persisted so closing the popup does not lose it
export interface DraftItem {
  id: string;
  name: string;
  type: string;
  shortcut: string;
  state?: UploadState;
  step: number;
  error?: string;
}

export interface Draft {
  items: DraftItem[];
  // Ids of the items sent to the background worker, in upload order
  inFlightIds: string[];
}

const DRAFT_KEY = "draft";
// File contents are stored once per file, so renaming only rewrites the small draft
const FILE_PREFIX = "draftFile:";

export function isDraftKey(key: string): boolean {
  return key === DRAFT_KEY || key.startsWith(FILE_PREFIX);
}

export async function saveDraft(draft: Draft): Promise<void> {
  await chrome.storage.local.set({ [DRAFT_KEY]: draft });
}

export async function saveDraftFile(id: string, base64: string): Promise<void> {
  await chrome.storage.local.set({ [FILE_PREFIX + id]: base64 });
}

export async function removeDraftFile(id: string): Promise<void> {
  await chrome.storage.local.remove(FILE_PREFIX + id);
}

export async function loadDraft(): Promise<{
  draft: Draft;
  files: Record<string, string>;
} | null> {
  const stored = await chrome.storage.local.get(null);
  const draft = stored[DRAFT_KEY] as Draft | undefined;
  if (!draft) {
    return null;
  }

  const files: Record<string, string> = {};
  for (const item of draft.items) {
    const base64 = stored[FILE_PREFIX + item.id];
    if (typeof base64 === "string") {
      files[item.id] = base64;
    }
  }
  return { draft, files };
}

export async function clearDraft(): Promise<void> {
  const stored = await chrome.storage.local.get(null);
  await chrome.storage.local.remove(Object.keys(stored).filter(isDraftKey));
}

// Clear captured tokens and upload state, but keep the user's file list
export async function clearAllExceptDraft(): Promise<void> {
  const stored = await chrome.storage.local.get(null);
  await chrome.storage.local.remove(
    Object.keys(stored).filter((key) => !isDraftKey(key)),
  );
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  // Chunked to stay under the argument limit of String.fromCharCode
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}
