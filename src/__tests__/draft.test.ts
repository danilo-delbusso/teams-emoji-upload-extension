import {
  base64ToBytes,
  bytesToBase64,
  clearAllExceptDraft,
  clearDraft,
  Draft,
  isDraftKey,
  loadDraft,
  removeDraftFile,
  saveDraft,
  saveDraftFile,
} from "../draft";

describe("draft", () => {
  let store: Record<string, unknown>;

  beforeEach(() => {
    store = {};
    window.chrome = {
      storage: {
        local: {
          get: jest.fn(async () => ({ ...store })),
          set: jest.fn(async (items: Record<string, unknown>) => {
            Object.assign(store, items);
          }),
          remove: jest.fn(async (keys: string | string[]) => {
            for (const key of Array.isArray(keys) ? keys : [keys]) {
              delete store[key];
            }
          }),
        },
      },
    } as any;
  });

  const draft: Draft = {
    items: [
      { id: "1", name: "a.png", type: "image/png", shortcut: "a", step: 0 },
      { id: "2", name: "b.gif", type: "image/gif", shortcut: "b", step: 0 },
    ],
    inFlightIds: [],
  };

  it("should round-trip the draft and its files", async () => {
    await saveDraftFile("1", "AAAA");
    await saveDraftFile("2", "BBBB");
    await saveDraft(draft);

    expect(await loadDraft()).toEqual({
      draft,
      files: { "1": "AAAA", "2": "BBBB" },
    });
  });

  it("should return null when nothing is saved", async () => {
    expect(await loadDraft()).toBeNull();
  });

  it("should skip files that are missing from storage", async () => {
    await saveDraftFile("1", "AAAA");
    await saveDraft(draft);

    expect((await loadDraft())?.files).toEqual({ "1": "AAAA" });
  });

  it("should remove a single file", async () => {
    await saveDraftFile("1", "AAAA");
    await removeDraftFile("1");

    expect(store).toEqual({});
  });

  it("should clear only draft keys", async () => {
    store = { teams_ic3: { secret: "x" } };
    await saveDraftFile("1", "AAAA");
    await saveDraft(draft);

    await clearDraft();

    expect(store).toEqual({ teams_ic3: { secret: "x" } });
  });

  it("should clear everything except draft keys", async () => {
    store = { teams_ic3: { secret: "x" }, uploadProgress: [] };
    await saveDraftFile("1", "AAAA");
    await saveDraft(draft);

    await clearAllExceptDraft();

    expect(Object.keys(store).sort()).toEqual(["draft", "draftFile:1"]);
  });

  it("should recognise draft keys", () => {
    expect(isDraftKey("draft")).toBe(true);
    expect(isDraftKey("draftFile:123")).toBe(true);
    expect(isDraftKey("uploadProgress")).toBe(false);
  });

  it("should convert bytes to base64 and back", () => {
    // Larger than one 0x8000 chunk to cover the chunked conversion
    const bytes = new Uint8Array(70000).map((_, i) => i % 256);
    const base64 = bytesToBase64(bytes);

    expect(base64.slice(0, 8)).toBe(btoa("\x00\x01\x02\x03\x04\x05"));
    expect(base64ToBytes(base64)).toEqual(bytes);
  });
});
