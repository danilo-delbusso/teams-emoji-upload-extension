import MsTeamsClient from "./msTeams";
import { ProcessResult, FileDetails, UploadProgressItem } from "./types";

// Open the side panel from the toolbar icon (absent in unit tests)
chrome.sidePanel
  ?.setPanelBehavior({ openPanelOnActionClick: true })
  .catch((error) => console.error(error));

// Lets a reopened popup tell a running upload from one cut short by a worker restart
let uploadInProgress = false;

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === "getUploadStatus") {
    sendResponse({ uploading: uploadInProgress });
    return;
  }

  if (message.action === "processFiles") {
    const files = JSON.parse(message.files) as FileDetails[];
    // Set initial processing state
    chrome.storage.local.set({
      processingState: { status: "Processing...", type: "processing" },
    });
    uploadInProgress = true;
    handleFileProcessing(files, message.tokens)
      .then((result) => {
        uploadInProgress = false;
        sendResponse(result);
      })
      .catch((error) => {
        uploadInProgress = false;
        const errorMessage = formatErrorMessage(error);
        // Store error state
        chrome.storage.local.set({
          processingState: { status: errorMessage, type: "error" },
        });
        sendResponse({ success: false, error: errorMessage });
      });
    return true;
  }
});

// The popup may be closed, in which case nobody receives the message
function notifyPopup(message: object) {
  chrome.runtime.sendMessage(message)?.catch(() => {});
}

export function formatErrorMessage(error: any): string {
  // If it's a standard error object
  if (error instanceof Error) {
    return error.message;
  }

  // If it's a string that might be JSON
  if (typeof error === "string") {
    try {
      const parsed = JSON.parse(error);
      return JSON.stringify(parsed, null, 2);
    } catch {
      // Not JSON, return as is
      return error;
    }
  }

  // If it's already an object
  if (typeof error === "object") {
    try {
      return JSON.stringify(error, null, 2);
    } catch {
      return String(error);
    }
  }

  // Fallback
  return String(error);
}

export async function handleFileProcessing(
  files: FileDetails[],
  tokens: {
    chatsvcagg: string | null;
    ic3: string | null;
    permissionsId: string | null;
  },
): Promise<ProcessResult> {
  try {
    if (!tokens.chatsvcagg || !tokens.ic3 || !tokens.permissionsId) {
      const errorMsg = "Could not find required tokens";
      chrome.storage.local.set({
        processingState: { status: errorMsg, type: "error" },
      });
      throw new Error(errorMsg);
    }
    const teams = new MsTeamsClient(
      tokens.ic3,
      tokens.chatsvcagg,
      tokens.permissionsId,
    );
    const progress: UploadProgressItem[] = files.map((file) => ({
      name: file.name,
      shortcut: file.shortcut || file.name.split(".")[0],
      state: "pending",
      step: 0,
    }));
    const publishProgress = () => {
      // Persist so the popup can show progress again if it is reopened
      chrome.storage.local.set({ uploadProgress: progress });
      notifyPopup({ type: "uploadProgress", progress });
    };
    publishProgress();

    const result = await teams.uploadFiles(files, (update) => {
      const { index, ...rest } = update;
      progress[index] = { ...progress[index], error: undefined, ...rest };
      publishProgress();
    });

    // Store processing result state
    chrome.storage.local.set({
      processingState: {
        status: result.status || "",
        type: result.success ? "success" : "error",
      },
    });

    chrome.runtime.sendMessage({
      type: "processUpdate",
      success: result.success,
      status: result.status,
      error: result.error,
    });

    return result;
  } catch (error) {
    const errorMessage = formatErrorMessage(error);
    const errorResult = { success: false, error: errorMessage };

    // Store error state
    chrome.storage.local.set({
      processingState: { status: errorMessage, type: "error" },
    });

    chrome.runtime.sendMessage({
      type: "processUpdate",
      ...errorResult,
    });

    return errorResult;
  }
}
