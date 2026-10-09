// Emoji names are typed as :name: in Teams, so whitespace and colons are not allowed
const INVALID_CHARS = /[\s:]/;

export function defaultShortcut(fileName: string): string {
  const dot = fileName.lastIndexOf(".");
  const base = dot > 0 ? fileName.slice(0, dot) : fileName;
  return base.trim().replace(/\s+/g, "-").replace(/:/g, "");
}

export function validateShortcuts(shortcuts: string[]): (string | null)[] {
  const counts = new Map<string, number>();
  for (const shortcut of shortcuts) {
    const key = shortcut.trim().toLowerCase();
    counts.set(key, (counts.get(key) || 0) + 1);
  }

  return shortcuts.map((shortcut) => {
    const trimmed = shortcut.trim();
    if (!trimmed) {
      return "Name is required";
    }
    if (INVALID_CHARS.test(trimmed)) {
      return "No spaces or colons";
    }
    if ((counts.get(trimmed.toLowerCase()) || 0) > 1) {
      return "Duplicate name";
    }
    return null;
  });
}
