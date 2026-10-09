import TokenStore from "./tokenStore";
import { moonIcon, sunIcon, svgIcon } from "./icons";
import MsTeamsClient from "./msTeams";
import { userIdFromToken } from "./customEmoji";
import { createEmojiBrowser } from "./emojiBrowser";
import { defaultShortcut, validateShortcuts } from "./shortcuts";
import {
  base64ToBytes,
  bytesToBase64,
  clearAllExceptDraft,
  clearDraft,
  loadDraft,
  removeDraftFile,
  saveDraft,
  saveDraftFile,
} from "./draft";
import {
  ProcessResult,
  FileDetails,
  UploadProgressItem,
  UploadState,
  UPLOAD_STEPS,
} from "./types";

type StatusType = "ready" | "success" | "error" | "processing";

// Teams web is served from both domains depending on the tenant
const TEAMS_ORIGINS: [string, ...string[]] = [
  "https://teams.microsoft.com",
  "https://teams.cloud.microsoft",
];
const TEAMS_TAB_URLS = TEAMS_ORIGINS.map((origin) => `${origin}/*`);

// The side panel stays open across tabs, so look for Teams rather than assuming
// the active tab: prefer the active tab if it is Teams, else any Teams tab
async function findTeamsTab(): Promise<chrome.tabs.Tab | undefined> {
  const [active] = await chrome.tabs.query({
    active: true,
    currentWindow: true,
    url: TEAMS_TAB_URLS,
  });
  if (active) return active;
  const [other] = await chrome.tabs.query({
    currentWindow: true,
    url: TEAMS_TAB_URLS,
  });
  return other;
}

// Caches only: sign-in lives in cookies and local storage, which are kept
const TEAMS_CACHE_DATA: chrome.browsingData.DataTypeSet = {
  cache: true,
  cacheStorage: true,
  serviceWorkers: true,
  indexedDB: true,
};
// Everything, which signs you out (cookies are cleared for the whole domain)
const TEAMS_ALL_DATA: chrome.browsingData.DataTypeSet = {
  ...TEAMS_CACHE_DATA,
  cookies: true,
  localStorage: true,
};

interface QueueItem {
  id: string;
  file: File;
  base64: string;
  name: string;
  shortcut: string;
  previewUrl?: string;
  state?: UploadState;
  step: number;
  error?: string;
  row?: {
    input: HTMLInputElement;
    remove: HTMLButtonElement;
    stateLabel: HTMLElement;
    message: HTMLElement;
    progress: HTMLElement;
    bar: HTMLElement;
  };
}

let queue: QueueItem[] = [];
// Items in the order they were sent to the background worker
let inFlight: QueueItem[] = [];
let uploading = false;
// Set when the popup reopens while the background worker is still uploading
let restoredUpload = false;
let nextId = 0;

// Run storage writes in order, so a slow clear cannot undo a newer save
let storageWrites = Promise.resolve();
function queueStorageWrite(write: () => Promise<void>) {
  storageWrites = storageWrites
    .then(write)
    .catch((err) => console.error("Error saving file list:", err));
}

function persistDraft() {
  // Once everything is uploaded there is nothing worth restoring
  const keep = uploading || queue.some((item) => item.state !== "done");
  const draft = {
    items: queue.map((item) => ({
      id: item.id,
      name: item.name,
      type: item.file.type,
      shortcut: item.shortcut,
      state: item.state,
      step: item.step,
      error: item.error,
    })),
    inFlightIds: inFlight.map((item) => item.id),
  };
  queueStorageWrite(() => (keep ? saveDraft(draft) : clearDraft()));
}

document.addEventListener("DOMContentLoaded", async () => {
  try {
    // Check for existing processing state
    const state = await chrome.storage.local.get<{
      processingState?: { status: string; type: string };
      uploadProgress?: UploadProgressItem[];
    }>(["processingState", "uploadProgress"]);
    if (state.processingState) {
      updateStatus(
        state.processingState.status,
        state.processingState.type as StatusType,
      );
    }
    if (queue.length === 0) {
      await restoreDraft(state.uploadProgress);
    }
    updateControls();
  } catch (err) {
    console.error("Error restoring state:", err);
  }

  await captureTokens();
});

// Copy the Teams auth tokens from the Teams tab's localStorage
async function captureTokens() {
  try {
    const tab = await findTeamsTab();

    if (tab?.id) {
      const result = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: () => {
          const items: Record<string, string> = {};
          for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (
              key &&
              (key.includes("chatsvcagg") || key.includes("ic3.teams.office"))
            ) {
              items[key] = localStorage.getItem(key) || "";
            }
          }
          return items;
        },
      });

      if (result && result[0].result) {
        for (const [key, value] of Object.entries(result[0].result)) {
          try {
            const parsedValue = JSON.parse(value as string);
            await chrome.storage.local.set({ [key]: parsedValue });
            console.log("Stored token:", key);
          } catch (e) {
            console.error("Failed to parse token:", key, e);
          }
        }
        console.log("Teams tokens captured from page");
      }
    }
  } catch (err) {
    console.error("Error capturing tokens:", err);
  }
}

const statusIcons = {
  success: svgIcon(
    '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" /><polyline points="22 4 12 14.01 9 11.01" />',
  ),
  error: svgIcon(
    '<circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" />',
  ),
  processing: svgIcon('<path d="M21 12a9 9 0 1 1-6.219-8.56" />', "spin"),
  ready: svgIcon(
    '<circle cx="12" cy="12" r="10" /><line x1="12" y1="16" x2="12" y2="12" /><line x1="12" y1="8" x2="12.01" y2="8" />',
  ),
};

const removeIcon = svgIcon(
  '<line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" />',
  "icon",
);

function updateStatus(message: string, type: StatusType = "ready") {
  const statusDiv = document.getElementById("status")!;
  const statusText = document.getElementById("statusText")!;
  const statusIcon = document.getElementById("statusIcon")!;

  statusDiv.className = "status"; // Reset class
  statusDiv.classList.add(`status-${type}`);
  statusIcon.innerHTML = statusIcons[type];
  statusText.textContent = message;
}

function formatSize(bytes: number): string {
  if (bytes >= 1024 * 1024) {
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }
  return `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

function stateLabel(item: QueueItem): string {
  switch (item.state) {
    case "pending":
      return "Waiting";
    case "uploading":
      return ["Creating…", "Uploading image…", "Saving name…"][item.step - 1];
    case "done":
      return "Uploaded";
    case "error":
      return "Failed";
    default:
      return "";
  }
}

function renderList() {
  const list = document.getElementById("fileList")!;
  list.replaceChildren();

  for (const item of queue) {
    const li = document.createElement("li");
    li.className = "file-item";

    let preview: HTMLElement;
    if (item.previewUrl) {
      const img = document.createElement("img");
      img.src = item.previewUrl;
      img.alt = "";
      preview = img;
    } else {
      preview = document.createElement("div");
    }
    preview.classList.add("file-preview");

    const field = document.createElement("label");
    field.className = "file-name-field";
    const input = document.createElement("input");
    input.type = "text";
    input.value = item.shortcut;
    input.spellcheck = false;
    input.setAttribute("aria-label", `Emoji name for ${item.name}`);
    input.addEventListener("input", () => {
      item.shortcut = input.value;
      refreshValidation();
      persistDraft();
    });
    field.append(":", input, ":");

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "file-remove";
    remove.title = "Remove";
    remove.setAttribute("aria-label", `Remove ${item.name}`);
    remove.innerHTML = removeIcon;
    remove.addEventListener("click", () => removeItem(item));

    const meta = document.createElement("div");
    meta.className = "file-meta";
    const fileName = document.createElement("span");
    fileName.className = "file-original";
    fileName.textContent = `${item.name} · ${formatSize(item.file.size)}`;
    fileName.title = item.name;
    const state = document.createElement("span");
    meta.append(fileName, state);

    const message = document.createElement("div");
    message.className = "file-message";

    const progress = document.createElement("div");
    progress.className = "file-progress";
    const bar = document.createElement("div");
    bar.className = "file-progress-bar";
    progress.append(bar);

    li.append(preview, field, remove, meta, message, progress);
    list.append(li);

    item.row = { input, remove, stateLabel: state, message, progress, bar };
    updateRow(item);
  }

  refreshValidation();
  updateControls();
}

function updateRow(item: QueueItem) {
  if (!item.row) return;
  const { input, remove, stateLabel: label, progress, bar } = item.row;

  // Uploaded files keep their name; any row can leave the list between uploads
  input.disabled = uploading || item.state === "done";
  remove.hidden = uploading;
  remove.title =
    item.state === "done" ? "Remove from list (stays in Teams)" : "Remove";

  label.textContent = stateLabel(item);
  label.className = item.state ? `file-state-${item.state}` : "";

  progress.hidden = !item.state;
  bar.className = "file-progress-bar";
  if (item.state === "uploading") {
    bar.style.width = `${Math.max(5, ((item.step - 1) / UPLOAD_STEPS) * 100)}%`;
  } else if (item.state === "done" || item.state === "error") {
    bar.style.width = "100%";
    bar.classList.add(item.state);
  } else {
    bar.style.width = "0";
  }
}

function refreshValidation() {
  const errors = validateShortcuts(queue.map((item) => item.shortcut));
  queue.forEach((item, i) => {
    if (!item.row) return;
    const validationError = errors[i];
    item.row.input.classList.toggle("invalid", !!validationError);
    item.row.message.textContent = validationError || item.error || "";
  });
  updateControls();
}

function updateControls() {
  const button = document.getElementById("processButton") as HTMLButtonElement;
  const summary = document.getElementById("selectedFiles")!;
  const label = document.getElementById("fileInputLabel")!;
  const remaining = queue.filter((item) => item.state !== "done");

  button.disabled = uploading || remaining.length === 0;
  const clearDone = document.getElementById("clearDoneButton")!;
  const clearAll = document.getElementById("clearAllButton")!;
  clearDone.hidden = uploading || !queue.some((item) => item.state === "done");
  clearAll.hidden = uploading || queue.length === 0;
  label.classList.toggle("disabled", uploading);
  (document.getElementById("fileInput") as HTMLInputElement).disabled =
    uploading;

  if (queue.length === 0) {
    summary.textContent = "No files selected";
  } else if (remaining.length === 0) {
    summary.textContent = `${queue.length} file${queue.length === 1 ? "" : "s"}`;
  } else {
    summary.textContent = `${remaining.length} of ${queue.length} file${queue.length === 1 ? "" : "s"} to upload`;
  }
}

// Removes rows from the list only; uploaded emojis stay in Teams
function removeItems(items: QueueItem[]) {
  for (const item of items) {
    if (item.previewUrl) {
      URL.revokeObjectURL(item.previewUrl);
    }
    queueStorageWrite(() => removeDraftFile(item.id));
  }
  queue = queue.filter((item) => !items.includes(item));
  renderList();
  persistDraft();

  // An empty list starts fresh, so drop the last upload's result too
  if (queue.length === 0) {
    updateStatus("Ready to process");
    queueStorageWrite(() => chrome.storage.local.remove("processingState"));
  }
}

function removeItem(item: QueueItem) {
  removeItems([item]);
}

function resetUnfinished() {
  for (const item of queue) {
    if (item.state === "pending" || item.state === "uploading") {
      item.state = undefined;
      item.step = 0;
    }
  }
}

async function restoreDraft(progress?: UploadProgressItem[]) {
  const saved = await loadDraft();
  if (!saved) return;

  queue = saved.draft.items
    .filter((item) => saved.files[item.id])
    .map((item) => {
      const base64 = saved.files[item.id];
      const file = new File([base64ToBytes(base64)], item.name, {
        type: item.type,
      });
      return {
        id: item.id,
        file,
        base64,
        name: item.name,
        shortcut: item.shortcut,
        previewUrl: URL.createObjectURL(file),
        state: item.state,
        step: item.step,
        error: item.error,
      };
    });
  inFlight = saved.draft.inFlightIds
    .map((id) => queue.find((item) => item.id === id))
    .filter((item): item is QueueItem => !!item);

  let backgroundUploading = false;
  try {
    const status = await chrome.runtime.sendMessage({
      action: "getUploadStatus",
    });
    backgroundUploading = !!status?.uploading;
  } catch (err) {
    console.error("Error checking upload status:", err);
  }
  uploading = backgroundUploading;
  restoredUpload = backgroundUploading;

  renderList();
  // The background worker kept going while the popup was closed
  if (progress && progress.length === inFlight.length) {
    applyProgress(progress);
  }
  if (
    !backgroundUploading &&
    queue.some((item) => item.state === "pending" || item.state === "uploading")
  ) {
    resetUnfinished();
    queue.forEach(updateRow);
    updateStatus("The upload was interrupted. Upload again to retry.", "error");
  }
  refreshValidation();
  persistDraft();
}

function applyProgress(progress: UploadProgressItem[]) {
  progress.forEach((p, i) => {
    const item = inFlight[i];
    if (!item) return;
    item.state = p.state;
    item.step = p.step;
    item.error = p.error;
    updateRow(item);
  });
  refreshValidation();
  persistDraft();

  const finished = progress.filter(
    (p) => p.state === "done" || p.state === "error",
  ).length;
  if (finished < progress.length) {
    updateStatus(
      `Uploading… ${finished} of ${progress.length} done`,
      "processing",
    );
  }
}

async function addFiles(files: File[]) {
  const added: QueueItem[] = [];
  for (const file of files) {
    const item: QueueItem = {
      id: `${Date.now()}-${nextId++}`,
      file,
      base64: bytesToBase64(new Uint8Array(await file.arrayBuffer())),
      name: file.name,
      shortcut: defaultShortcut(file.name),
      previewUrl: URL.createObjectURL(file),
      step: 0,
    };
    added.push(item);
    queueStorageWrite(() => saveDraftFile(item.id, item.base64));
  }
  queue.push(...added);
  renderList();
  persistDraft();
}

function toFileDetails(item: QueueItem): FileDetails {
  return {
    name: item.file.name,
    size: item.file.size,
    type: item.file.type,
    base64: item.base64,
    shortcut: item.shortcut.trim(),
  };
}

async function handleFileProcessing() {
  if (uploading) return;

  const toUpload = queue.filter((item) => item.state !== "done");
  if (toUpload.length === 0) {
    updateStatus("Please select files first.", "error");
    return;
  }
  if (validateShortcuts(queue.map((item) => item.shortcut)).some((e) => e)) {
    updateStatus("Fix the highlighted emoji names first.", "error");
    return;
  }

  uploading = true;
  inFlight = toUpload;
  for (const item of toUpload) {
    item.state = "pending";
    item.step = 0;
    item.error = undefined;
  }
  queue.forEach(updateRow);
  refreshValidation();
  // Drop the previous upload's progress so a reopened popup cannot mix them up
  await chrome.storage.local.remove("uploadProgress");
  persistDraft();

  try {
    const files = toUpload.map(toFileDetails);
    const filesJson = JSON.stringify(files);

    updateStatus("Processing...", "processing");

    // Tokens may have been cleared by an earlier refresh in this popup session
    await captureTokens();
    const tokens = await TokenStore.collectTokensFromStorage();

    const result = await chrome.runtime.sendMessage({
      action: "processFiles",
      files: filesJson,
      tokens: tokens,
    });

    const { success, error, status } = result as ProcessResult;
    if (error) {
      updateStatus(error, "error");
    } else {
      updateStatus(
        success
          ? `${status}. If they don't show up in Teams, use Refresh Teams below.`
          : status || "",
        success ? "success" : "processing",
      );
    }
  } catch (error) {
    let errorMessage = "";

    if (error instanceof Error) {
      errorMessage = error.message;
    } else if (typeof error === "string") {
      try {
        const parsed = JSON.parse(error);
        errorMessage = JSON.stringify(parsed, null, 2);
      } catch {
        errorMessage = error;
      }
    } else {
      errorMessage = String(error);
    }

    updateStatus(errorMessage, "error");
  } finally {
    uploading = false;
    // Anything the background worker never reached goes back to the editable list
    resetUnfinished();
    queue.forEach(updateRow);
    refreshValidation();
    persistDraft();
    // New emojis exist now, so the browser should fetch the list again
    emojiBrowser.invalidate();
  }
}

// A client for the signed-in user, with tokens freshly read from the Teams tab
async function getTeamsClient() {
  await captureTokens();
  const tokens = await TokenStore.collectTokensFromStorage();
  if (!tokens.chatsvcagg || !tokens.ic3 || !tokens.permissionsId) {
    throw new Error(
      "Couldn't find your Teams sign-in. Open Teams in this window, sign in, then try again.",
    );
  }
  return {
    client: new MsTeamsClient(
      tokens.ic3,
      tokens.chatsvcagg,
      tokens.permissionsId,
    ),
    userId: userIdFromToken(tokens.chatsvcagg),
  };
}

const emojiBrowser = createEmojiBrowser({
  getClient: getTeamsClient,
  updateStatus,
});

type Theme = "light" | "dark";

// Follows the system setting until the user picks a theme
function currentTheme(): Theme {
  const chosen = document.documentElement.dataset.theme;
  if (chosen === "light" || chosen === "dark") return chosen;
  return window.matchMedia?.("(prefers-color-scheme: dark)").matches
    ? "dark"
    : "light";
}

function updateThemeToggle() {
  const toggle = document.getElementById("themeToggle");
  if (!toggle) return;
  const next = currentTheme() === "dark" ? "light" : "dark";
  toggle.innerHTML = next === "light" ? sunIcon : moonIcon;
  toggle.title = `Switch to ${next} theme`;
  toggle.setAttribute("aria-label", toggle.title);
}

document.getElementById("themeToggle")?.addEventListener("click", () => {
  const next: Theme = currentTheme() === "dark" ? "light" : "dark";
  document.documentElement.dataset.theme = next;
  updateThemeToggle();
  try {
    localStorage.setItem("theme", next);
  } catch {
    // Storage unavailable
  }
});
window
  .matchMedia?.("(prefers-color-scheme: dark)")
  .addEventListener?.("change", updateThemeToggle);
try {
  const saved = localStorage.getItem("theme");
  if (saved === "light" || saved === "dark") {
    document.documentElement.dataset.theme = saved;
  }
} catch {
  // Storage unavailable
}
updateThemeToggle();

type View = "upload" | "browse";

function showView(view: View) {
  for (const [name, tab, panel] of [
    ["upload", "uploadTab", "uploadView"],
    ["browse", "browseTab", "browseView"],
  ] as const) {
    const active = name === view;
    document.getElementById(tab)?.setAttribute("aria-selected", String(active));
    document.getElementById(panel)!.hidden = !active;
  }
  if (view === "browse") {
    emojiBrowser.show();
  }
  try {
    // Per-viewer convenience only, so losing it is harmless
    localStorage.setItem("view", view);
  } catch {
    // Storage unavailable
  }
}

document
  .getElementById("uploadTab")
  ?.addEventListener("click", () => showView("upload"));
document
  .getElementById("browseTab")
  ?.addEventListener("click", () => showView("browse"));
try {
  if (localStorage.getItem("view") === "browse") {
    showView("browse");
  }
} catch {
  // Storage unavailable
}

const fileInput = document.getElementById("fileInput") as HTMLInputElement;
fileInput?.addEventListener("change", () => {
  addFiles(Array.from(fileInput.files || []));
  // Allow choosing the same file again later
  fileInput.value = "";
});
document
  .getElementById("processButton")
  ?.addEventListener("click", () => handleFileProcessing());
document
  .getElementById("clearDoneButton")
  ?.addEventListener("click", () =>
    removeItems(queue.filter((item) => item.state === "done")),
  );
document
  .getElementById("clearAllButton")
  ?.addEventListener("click", () => removeItems(queue));

// Clear Teams site data so it fetches the emoji list again, then reload the tab
async function clearTeamsData(mode: "refresh" | "reset") {
  await chrome.browsingData.remove(
    { origins: TEAMS_ORIGINS },
    mode === "reset" ? TEAMS_ALL_DATA : TEAMS_CACHE_DATA,
  );

  await clearAllExceptDraft();

  const tab = await findTeamsTab();

  if (tab?.id) {
    await chrome.tabs.reload(tab.id, { bypassCache: true });
  }
}

async function refreshTeams(mode: "refresh" | "reset") {
  try {
    updateStatus(
      mode === "reset" ? "Resetting Teams..." : "Refreshing Teams...",
      "processing",
    );
    await clearTeamsData(mode);
    updateStatus(
      mode === "reset"
        ? "Teams was reset. Sign in again to see your emojis."
        : "Teams refreshed successfully!",
      "success",
    );
  } catch (error) {
    updateStatus(`Refresh failed: ${error}`, "error");
  }
}

document
  .getElementById("refreshButton")
  ?.addEventListener("click", () => refreshTeams("refresh"));
document
  .getElementById("resetButton")
  ?.addEventListener("click", () => refreshTeams("reset"));

chrome.runtime.onMessage.addListener(
  (message: {
    type: string;
    error: any;
    status: any;
    success: any;
    progress?: UploadProgressItem[];
  }) => {
    if (message.type === "uploadProgress" && message.progress) {
      applyProgress(message.progress);
    }

    if (message.type === "processUpdate") {
      const status = message.error || message.status;
      const type = message.success
        ? "success"
        : message.error
          ? "error"
          : "processing";

      // Update the UI
      updateStatus(status, type as "success" | "error" | "processing");

      // A restored session has no pending sendMessage to finish it, so unlock here
      if (restoredUpload) {
        restoredUpload = false;
        uploading = false;
        resetUnfinished();
        queue.forEach(updateRow);
        refreshValidation();
        persistDraft();
      }

      // Ensure we store the state in case it was sent directly from background.ts
      // Background should be storing the state, but this is a good fallback
      chrome.storage.local.set({
        processingState: {
          status: status,
          type: type,
        },
      });
    }
  },
);
