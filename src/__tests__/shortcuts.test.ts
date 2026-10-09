import { defaultShortcut, validateShortcuts } from "../shortcuts";

describe("defaultShortcut", () => {
  it("should strip the file extension", () => {
    expect(defaultShortcut("partyparrot.gif")).toBe("partyparrot");
  });

  it("should only strip the last extension", () => {
    expect(defaultShortcut("party.parrot.png")).toBe("party.parrot");
  });

  it("should keep names without an extension", () => {
    expect(defaultShortcut("partyparrot")).toBe("partyparrot");
    expect(defaultShortcut(".hidden")).toBe(".hidden");
  });

  it("should replace whitespace and drop colons", () => {
    expect(defaultShortcut(" party  parrot:.png")).toBe("party-parrot");
  });
});

describe("validateShortcuts", () => {
  it("should accept valid unique names", () => {
    expect(validateShortcuts(["cat", "dog_2", "fox-3"])).toEqual([
      null,
      null,
      null,
    ]);
  });

  it("should reject empty names", () => {
    expect(validateShortcuts(["", "  "])).toEqual([
      "Name is required",
      "Name is required",
    ]);
  });

  it("should reject spaces and colons", () => {
    expect(validateShortcuts(["party parrot", "party:parrot"])).toEqual([
      "No spaces or colons",
      "No spaces or colons",
    ]);
  });

  it("should flag duplicates case-insensitively", () => {
    expect(validateShortcuts(["Cat", "cat", "dog"])).toEqual([
      "Duplicate name",
      "Duplicate name",
      null,
    ]);
  });
});
