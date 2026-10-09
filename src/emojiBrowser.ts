import MsTeamsClient from "./msTeams";
import { CustomEmoji, filterEmojis, toCustomEmojis } from "./customEmoji";
import { runWithConcurrency } from "./pool";
import { trashIcon } from "./icons";
import { RateLimiter, RateLimitError } from "./rateLimit";

type StatusType = "ready" | "success" | "error" | "processing";

export interface EmojiBrowserDeps {
  getClient: () => Promise<{ client: MsTeamsClient; userId: string | null }>;
  updateStatus: (message: string, type?: StatusType) => void;
}

const DELETE_CONCURRENCY = 3;
// Teams allows 10 deletions per user per minute (it says so in its 429 replies)
const DELETE_LIMIT = 10;
const DELETE_WINDOW_MS = 60_000;
const MAX_DELETE_ATTEMPTS = 5;

type TileState =
  | { kind: "waiting" }
  | { kind: "deleting" }
  | { kind: "failed"; message: string };
const IMAGE_CONCURRENCY = 6;
// Deleting this many at once also requires typing "delete" to confirm
const TYPED_CONFIRMATION_THRESHOLD = 10;
// The confirmation dialog lists at most this many names
const CONFIRM_LIST_LIMIT = 30;

const byId = <T extends HTMLElement = HTMLElement>(id: string) =>
  document.getElementById(id) as T;

const errorMessage = (error: unknown) =>
  error instanceof Error ? error.message : String(error);

export function createEmojiBrowser(deps: EmojiBrowserDeps) {
  const grid = byId("emojiGrid");
  const search = byId<HTMLInputElement>("emojiSearch");
  const onlyMine = byId<HTMLInputElement>("onlyMine");
  const count = byId("emojiCount");
  const empty = byId("emojiEmpty");
  const reload = byId<HTMLButtonElement>("reloadEmojis");
  const selectShown = byId<HTMLButtonElement>("selectShown");
  const clearSelection = byId<HTMLButtonElement>("clearSelection");
  const deleteSelected = byId<HTMLButtonElement>("deleteSelected");

  let client: MsTeamsClient | null = null;
  let userId: string | null = null;
  let emojis: CustomEmoji[] = [];
  let shown: CustomEmoji[] = [];
  let loaded = false;
  let loading = false;
  let deleting = false;
  const selected = new Set<string>();
  // Per-emoji delete progress and failures, shown on the tiles
  const tileStates = new Map<string, TileState>();
  const tileElements = new Map<string, HTMLLIElement>();
  // Shared across batches so back-to-back deletes still respect the limit
  const deleteLimiter = new RateLimiter(DELETE_LIMIT, DELETE_WINDOW_MS);
  // Object URLs per documentId, kept across reloads
  const images = new Map<string, Promise<string | null>>();
  let observer: IntersectionObserver | null = null;

  // Limit parallel image downloads so the list stays responsive
  let activeImages = 0;
  const waitingImages: (() => void)[] = [];
  async function withImageSlot<T>(task: () => Promise<T>): Promise<T> {
    if (activeImages >= IMAGE_CONCURRENCY) {
      await new Promise<void>((resolve) => waitingImages.push(resolve));
    }
    activeImages++;
    try {
      return await task();
    } finally {
      activeImages--;
      waitingImages.shift()?.();
    }
  }

  function imageUrl(documentId: string): Promise<string | null> {
    let url = images.get(documentId);
    if (!url) {
      const current = client;
      url = current
        ? withImageSlot(() => current.fetchEmojiImage(documentId))
            .then((blob) => URL.createObjectURL(blob))
            .catch((error) => {
              console.error("Error loading emoji image:", error);
              // Forget the failure so a later render or reload retries it
              images.delete(documentId);
              return null;
            })
        : Promise.resolve(null);
      images.set(documentId, url);
    }
    return url;
  }

  function showImage(img: HTMLImageElement, documentId: string) {
    imageUrl(documentId).then((url) => {
      if (url) {
        img.src = url;
      } else {
        img.classList.add("failed");
      }
    });
  }

  function tile(emoji: CustomEmoji): HTMLLIElement {
    const li = document.createElement("li");
    li.className = "emoji-tile";
    li.classList.toggle("selected", selected.has(emoji.id));

    const card = document.createElement("label");
    card.className = "emoji-card";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.className = "emoji-check";
    checkbox.checked = selected.has(emoji.id);
    checkbox.setAttribute("aria-label", `Select :${emoji.shortcut}:`);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        selected.add(emoji.id);
      } else {
        selected.delete(emoji.id);
      }
      li.classList.toggle("selected", checkbox.checked);
      updateSelectionBar();
    });

    const img = document.createElement("img");
    img.className = "emoji-image";
    img.alt = "";
    img.dataset.documentId = emoji.documentId;

    const name = document.createElement("span");
    name.className = "emoji-name";
    name.textContent = `:${emoji.shortcut}:`;
    name.title = `:${emoji.shortcut}:`;

    const meta = document.createElement("span");
    meta.className = "emoji-meta";
    meta.textContent = [
      emoji.createdOn ? new Date(emoji.createdOn).toLocaleDateString() : "",
      userId && emoji.creator?.toLowerCase() === userId.toLowerCase()
        ? "by you"
        : "",
    ]
      .filter(Boolean)
      .join(" · ");

    card.append(checkbox, img, name, meta);

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "emoji-delete";
    remove.title = `Delete :${emoji.shortcut}:`;
    remove.setAttribute("aria-label", `Delete :${emoji.shortcut}:`);
    remove.innerHTML = trashIcon;
    remove.disabled = deleting;
    remove.addEventListener("click", () => deleteEmojis([emoji]));

    const status = document.createElement("span");
    status.className = "emoji-status";
    status.setAttribute("role", "status");

    li.append(card, remove, status);
    tileElements.set(emoji.id, li);
    applyTileState(li, emoji.id);

    if (observer) {
      observer.observe(img);
    } else {
      showImage(img, emoji.documentId);
    }
    return li;
  }

  function render() {
    const creator = onlyMine.checked ? userId || undefined : undefined;
    shown = filterEmojis(emojis, search.value, creator);

    count.textContent = loading
      ? "Loading…"
      : shown.length === emojis.length
        ? `${emojis.length} custom emoji${emojis.length === 1 ? "" : "s"}`
        : `${shown.length} of ${emojis.length} shown`;

    // Only download previews as tiles scroll into view
    observer?.disconnect();
    observer =
      typeof IntersectionObserver === "undefined"
        ? null
        : new IntersectionObserver(
            (entries) => {
              for (const entry of entries) {
                if (!entry.isIntersecting) continue;
                const img = entry.target as HTMLImageElement;
                observer?.unobserve(img);
                showImage(img, img.dataset.documentId!);
              }
            },
            { rootMargin: "200px" },
          );

    tileElements.clear();
    grid.replaceChildren(...shown.map(tile));

    empty.hidden = loading || shown.length > 0 || !loaded;
    empty.textContent =
      emojis.length === 0
        ? "Your organization has no custom emojis yet."
        : "No custom emojis match your search.";

    updateSelectionBar();
  }

  function updateSelectionBar() {
    const n = selected.size;
    deleteSelected.textContent = n
      ? `Delete selected (${n})`
      : "Delete selected";
    deleteSelected.disabled = deleting || n === 0;
    clearSelection.hidden = n === 0;
    selectShown.disabled = deleting || shown.length === 0;
    reload.disabled = deleting || loading;
    onlyMine.disabled = !userId;
    grid
      .querySelectorAll<HTMLButtonElement>(".emoji-delete")
      .forEach((button) => (button.disabled = deleting));
  }

  async function load() {
    if (loading || deleting) return;
    loading = true;
    render();
    try {
      ({ client, userId } = await deps.getClient());
      emojis = toCustomEmojis(await client.listCustomEmojis());
      loaded = true;
      // Keep the selection for emojis that still exist
      const ids = new Set(emojis.map((emoji) => emoji.id));
      [...selected]
        .filter((id) => !ids.has(id))
        .forEach((id) => selected.delete(id));
    } catch (error) {
      deps.updateStatus(errorMessage(error), "error");
    } finally {
      loading = false;
      render();
      if (!loaded) {
        count.textContent = "Couldn't load custom emojis";
      }
    }
  }

  function confirmDelete(list: CustomEmoji[]): Promise<boolean> {
    const dialog = byId<HTMLDialogElement>("confirmDialog");
    const title = byId("confirmTitle");
    const items = byId("confirmList");
    const typedRow = byId("confirmTypedRow");
    const typed = byId<HTMLInputElement>("confirmTyped");
    const cancel = byId<HTMLButtonElement>("confirmCancel");
    const confirm = byId<HTMLButtonElement>("confirmDelete");

    title.textContent =
      list.length === 1
        ? `Delete :${list[0].shortcut}:?`
        : `Delete ${list.length} custom emojis?`;

    items.replaceChildren(
      ...list.slice(0, CONFIRM_LIST_LIMIT).map((emoji) => {
        const li = document.createElement("li");
        const img = document.createElement("img");
        img.alt = "";
        showImage(img, emoji.documentId);
        li.append(img, `:${emoji.shortcut}:`);
        return li;
      }),
    );
    if (list.length > CONFIRM_LIST_LIMIT) {
      const more = document.createElement("li");
      more.className = "confirm-more";
      more.textContent = `and ${list.length - CONFIRM_LIST_LIMIT} more`;
      items.append(more);
    }

    // Large bulk deletes need the word typed out, not just a click
    const needsTyping = list.length >= TYPED_CONFIRMATION_THRESHOLD;
    typedRow.hidden = !needsTyping;
    typed.value = "";
    confirm.disabled = needsTyping;
    typed.oninput = () => {
      confirm.disabled = typed.value.trim().toLowerCase() !== "delete";
    };
    confirm.textContent =
      list.length === 1 ? "Delete emoji" : `Delete ${list.length} emojis`;

    dialog.returnValue = "";
    dialog.showModal();
    // Cancel is the default, so Enter or a stray click never deletes
    cancel.focus();

    return new Promise((resolve) => {
      dialog.addEventListener(
        "close",
        () => resolve(dialog.returnValue === "delete"),
        { once: true },
      );
    });
  }

  function setTileState(id: string, state?: TileState) {
    if (state) {
      tileStates.set(id, state);
    } else {
      tileStates.delete(id);
    }
    const li = tileElements.get(id);
    if (li) applyTileState(li, id);
  }

  function applyTileState(li: HTMLLIElement, id: string) {
    const state = tileStates.get(id);
    for (const name of ["waiting", "deleting", "failed"]) {
      li.classList.toggle(name, state?.kind === name);
    }
    const status = li.querySelector(".emoji-status")!;
    status.textContent =
      state?.kind === "waiting"
        ? "Waiting…"
        : state?.kind === "deleting"
          ? "Deleting…"
          : state?.kind === "failed"
            ? state.message
            : "";
    (status as HTMLElement).title = status.textContent || "";
  }

  async function deleteEmojis(list: CustomEmoji[]) {
    if (deleting || list.length === 0) return;
    if (!(await confirmDelete(list))) return;

    deleting = true;
    updateSelectionBar();
    let removedCount = 0;
    const failures = new Map<string, string[]>();
    const progress = () => `${removedCount} of ${list.length} deleted`;
    list.forEach((emoji) => setTileState(emoji.id, { kind: "waiting" }));
    deps.updateStatus(`Deleting… ${progress()}`, "processing");

    const fail = (emoji: CustomEmoji, message: string) => {
      setTileState(emoji.id, { kind: "failed", message });
      failures.set(message, [...(failures.get(message) || []), emoji.shortcut]);
    };

    try {
      const { client: current } = await deps.getClient();
      client = current;
      await runWithConcurrency(list.length, DELETE_CONCURRENCY, async (i) => {
        const emoji = list[i];
        for (let attempt = 1; ; attempt++) {
          // Stay under Teams' limit, showing a countdown while we wait
          await deleteLimiter.acquire((msLeft) =>
            deps.updateStatus(
              `Pausing for Teams' rate limit. Continuing in ${Math.ceil(msLeft / 1000)}s… (${progress()})`,
              "processing",
            ),
          );
          setTileState(emoji.id, { kind: "deleting" });
          try {
            await current.deleteCustomEmoji(emoji.id);
            removedCount++;
            setTileState(emoji.id);
            emojis = emojis.filter((other) => other.id !== emoji.id);
            selected.delete(emoji.id);
            tileElements.get(emoji.id)?.remove();
            deps.updateStatus(`Deleting… ${progress()}`, "processing");
            return;
          } catch (error) {
            if (
              error instanceof RateLimitError &&
              attempt < MAX_DELETE_ATTEMPTS
            ) {
              // Teams said slow down: wait as long as it asked, then retry
              deleteLimiter.backOff(error);
              setTileState(emoji.id, { kind: "waiting" });
              continue;
            }
            fail(emoji, errorMessage(error));
            return;
          }
        }
      });
    } catch (error) {
      // Could not even start, e.g. no Teams sign-in found
      list
        .filter((emoji) => tileStates.get(emoji.id)?.kind !== "failed")
        .forEach((emoji) => fail(emoji, errorMessage(error)));
    } finally {
      deleting = false;
      render();
    }

    if (failures.size > 0) {
      const failedCount = [...failures.values()].flat().length;
      const reasons = [...failures.entries()]
        .map(
          ([message, names]) =>
            `• ${message} (${names.map((n) => `:${n}:`).join(", ")})`,
        )
        .join("\n");
      deps.updateStatus(
        `Deleted ${removedCount} of ${list.length}. ${failedCount} failed and ${failedCount === 1 ? "is" : "are"} still selected, so you can retry:\n${reasons}`,
        "error",
      );
    } else {
      deps.updateStatus(
        `Deleted ${removedCount} emoji${removedCount === 1 ? "" : "s"}. Refresh Teams to update its emoji picker.`,
        "success",
      );
    }
  }

  search.addEventListener("input", render);
  onlyMine.addEventListener("change", render);
  reload.addEventListener("click", load);
  selectShown.addEventListener("click", () => {
    shown.forEach((emoji) => selected.add(emoji.id));
    render();
  });
  clearSelection.addEventListener("click", () => {
    selected.clear();
    render();
  });
  deleteSelected.addEventListener("click", () =>
    deleteEmojis(emojis.filter((emoji) => selected.has(emoji.id))),
  );

  return {
    // Load the list the first time the view is shown
    show() {
      if (!loaded) load();
    },
    // The list changed elsewhere (e.g. after an upload), so fetch it again
    invalidate() {
      loaded = false;
    },
  };
}
