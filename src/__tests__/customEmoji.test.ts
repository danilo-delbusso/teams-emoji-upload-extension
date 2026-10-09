import { filterEmojis, toCustomEmojis, userIdFromToken } from "../customEmoji";
import { CustomEmojiResponseItem } from "../types";

const item = (
  shortcut: string,
  extra: Partial<CustomEmojiResponseItem> = {},
): CustomEmojiResponseItem => ({
  id: `${shortcut};0-wuk-d1-${shortcut.padEnd(32, "0").slice(0, 32)}`,
  documentId: `0-wuk-d1-${shortcut.padEnd(32, "0").slice(0, 32)}`,
  shortcuts: [shortcut],
  ...extra,
});

describe("toCustomEmojis", () => {
  it("should drop deleted entries and sort newest first", () => {
    const emojis = toCustomEmojis([
      item("old", { createdOn: 1_600_000_000_000 }),
      item("gone", { isDeleted: true, createdOn: 1_800_000_000_000 }),
      item("new", { createdOn: 1_700_000_000_000 }),
    ]);

    expect(emojis.map((e) => e.shortcut)).toEqual(["new", "old"]);
  });

  it("should accept createdOn in seconds", () => {
    const [emoji] = toCustomEmojis([item("cat", { createdOn: 1_700_000_000 })]);
    expect(emoji.createdOn).toBe(1_700_000_000_000);
  });

  it("should fall back to the id for the shortcut", () => {
    const [emoji] = toCustomEmojis([item("cat", { shortcuts: [] })]);
    expect(emoji.shortcut).toBe("cat");
  });

  it("should sort entries without a date by name", () => {
    const emojis = toCustomEmojis([item("b"), item("a")]);
    expect(emojis.map((e) => e.shortcut)).toEqual(["a", "b"]);
  });
});

describe("filterEmojis", () => {
  const emojis = toCustomEmojis([
    item("party-parrot", { keywords: ["bird"], creator: "AAA" }),
    item("ship-it", { creator: "bbb" }),
  ]);

  it("should match shortcuts and keywords case-insensitively", () => {
    expect(filterEmojis(emojis, "PARROT").map((e) => e.shortcut)).toEqual([
      "party-parrot",
    ]);
    expect(filterEmojis(emojis, "bird").map((e) => e.shortcut)).toEqual([
      "party-parrot",
    ]);
  });

  it("should ignore surrounding colons", () => {
    expect(filterEmojis(emojis, ":ship-it:")).toHaveLength(1);
  });

  it("should filter by creator", () => {
    expect(filterEmojis(emojis, "", "aaa").map((e) => e.shortcut)).toEqual([
      "party-parrot",
    ]);
  });

  it("should return everything for an empty query", () => {
    expect(filterEmojis(emojis, "  ")).toHaveLength(2);
  });
});

describe("userIdFromToken", () => {
  const jwt = (claims: object) =>
    `header.${btoa(JSON.stringify(claims)).replace(/=+$/, "").replace(/\+/g, "-").replace(/\//g, "_")}.signature`;

  it("should read the oid claim", () => {
    expect(userIdFromToken(jwt({ oid: "user-1", name: "x" }))).toBe("user-1");
  });

  it("should return null for missing or malformed tokens", () => {
    expect(userIdFromToken(null)).toBeNull();
    expect(userIdFromToken("not-a-jwt")).toBeNull();
    expect(userIdFromToken("a.%%%.c")).toBeNull();
    expect(userIdFromToken(jwt({ sub: "x" }))).toBeNull();
  });
});
