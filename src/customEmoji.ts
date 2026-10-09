import { CustomEmojiResponseItem } from "./types";

export interface CustomEmoji {
  id: string;
  documentId: string;
  shortcut: string;
  keywords: string[];
  // Milliseconds since the epoch, when known
  createdOn?: number;
  creator?: string;
}

export function toCustomEmojis(
  items: CustomEmojiResponseItem[],
): CustomEmoji[] {
  return items
    .filter((item) => !item.isDeleted && item.id && item.documentId)
    .map((item) => ({
      id: item.id,
      documentId: item.documentId,
      shortcut: item.shortcuts?.[0] || item.id.split(";")[0],
      keywords: item.keywords || [],
      // Accept seconds as well as milliseconds
      createdOn:
        typeof item.createdOn === "number"
          ? item.createdOn < 1e12
            ? item.createdOn * 1000
            : item.createdOn
          : undefined,
      creator: item.creator,
    }))
    .sort(
      (a, b) =>
        (b.createdOn || 0) - (a.createdOn || 0) ||
        a.shortcut.localeCompare(b.shortcut),
    );
}

export function filterEmojis(
  emojis: CustomEmoji[],
  query: string,
  creator?: string,
): CustomEmoji[] {
  const needle = query.trim().toLowerCase().replace(/^:|:$/g, "");
  return emojis.filter(
    (emoji) =>
      (!creator || emoji.creator?.toLowerCase() === creator.toLowerCase()) &&
      (!needle ||
        emoji.shortcut.toLowerCase().includes(needle) ||
        emoji.keywords.some((k) => k.toLowerCase().includes(needle))),
  );
}

// The signed-in user's object id ("oid" claim) from a Teams JWT, read locally
export function userIdFromToken(token: string | null): string | null {
  try {
    const payload = token?.split(".")[1];
    if (!payload) return null;
    const base64 = payload.replace(/-/g, "+").replace(/_/g, "/");
    const padded = base64 + "=".repeat((4 - (base64.length % 4)) % 4);
    const claims = JSON.parse(atob(padded));
    return typeof claims.oid === "string" ? claims.oid : null;
  } catch {
    return null;
  }
}
