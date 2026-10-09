import {
  TeamsUploadResponse,
  TeamsMetadataResponse,
  FileDetails,
  FileProgress,
  UPLOAD_STEPS,
  CustomEmojiListResponse,
  CustomEmojiResponseItem,
} from "./types";
import { runWithConcurrency } from "./pool";
import { rateLimitError } from "./rateLimit";

// Files uploaded in parallel; too many risks Teams rate limiting (429)
export const UPLOAD_CONCURRENCY = 10;

// Custom emoji metadata: POST creates, GET lists, DELETE /{id} removes
const CUSTOM_EMOJI_METADATA_URL =
  "https://teams.microsoft.com/api/csa/apac/api/v1/customemoji/metadata";

// Image renditions to try for previews, in order
const EMOJI_IMAGE_VIEWS = ["imgt2_anim", "imgpsh_fullsize_anim"];

// HTTP/2 responses have no status text, so fall back to a readable name
const STATUS_NAMES: Record<number, string> = {
  400: "Bad Request",
  401: "Unauthorized",
  403: "Forbidden",
  404: "Not Found",
  409: "Conflict",
  413: "Payload Too Large",
  415: "Unsupported Media Type",
  429: "Too Many Requests",
  500: "Internal Server Error",
  503: "Service Unavailable",
};

export async function describeFailure(response: Response): Promise<string> {
  const statusText = response.statusText || STATUS_NAMES[response.status];
  const parts = [[response.status, statusText].filter(Boolean).join(" ")];

  try {
    const body = (await response.text()).trim();
    if (body) {
      parts.push(body.length > 200 ? `${body.slice(0, 200)}…` : body);
    }
  } catch {
    // No readable body
  }

  return parts.filter(Boolean).join(" - ") || "Unknown error";
}

class MsTeamsClient {
  private _ic3Token: string;
  private _chatToken: string;
  private _permissionsId: string;

  constructor(ic3Token: string, chatToken: string, permissionsId: string) {
    this._ic3Token = ic3Token;
    this._chatToken = chatToken;
    this._permissionsId = permissionsId;
  }

  public async createObject(): Promise<{ id: string; msCv: string }> {
    const response = await fetch(
      "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/",
      {
        method: "POST",
        headers: {
          authorization: `Bearer ${this._ic3Token}`,
          "content-type": "application/json",
          accept: "application/json",
          "x-ms-client-version": "1415/25021400912",
          origin: "https://teams.microsoft.com",
          referer: "https://teams.microsoft.com/",
        },
        mode: "cors",
        body: JSON.stringify({
          type: "pish/image",
          permissions: {
            ["*:tid:" + this._permissionsId]: ["read"],
          },
          sharingMode: "Unknown",
        }),
      },
    );

    if (!response.ok) {
      throw new Error(
        `Failed to create object: ${await describeFailure(response)}`,
      );
    }

    const msCv = response.headers.get("ms-cv") || "";
    const data = (await response.json()) as TeamsUploadResponse;
    return { id: data.id, msCv };
  }

  public async uploadFileContent(
    documentId: string,
    msCv: string,
    file: FileDetails,
  ): Promise<void> {
    console.log("Uploading file content", file.name);
    var buf = Buffer.from(file.base64, "base64");
    const response = await fetch(
      `https://as-prod.asyncgw.teams.microsoft.com/v1/objects/${documentId}/content/imgpsh`,
      {
        method: "PUT",
        headers: {
          authorization: `Bearer ${this._ic3Token}`,
          "content-type": "application/octet-stream",
          "x-ms-client-version": "1415/25021400912",
          DNT: "1",
          origin: "https://teams.microsoft.com",
          referer: "https://teams.microsoft.com/",
          "x-ms-migration": "True",
          "x-ms-test-user": "False",
          "ms-cv": msCv,
        },
        body: buf,
      },
    );

    if (!response.ok) {
      throw new Error(
        `Failed to upload image: ${await describeFailure(response)}`,
      );
    }
  }

  public async sendMetadata(
    documentId: string,
    msCv: string,
    shortcuts: string[],
  ): Promise<void> {
    const response = await fetch(CUSTOM_EMOJI_METADATA_URL, {
      method: "POST",
      headers: {
        authorization: `Bearer ${this._chatToken}`,
        "content-type": "application/json",
        accept: "application/json",
        "x-ms-client-version": "1415/25021400912",
        "x-ms-client-type": "web",
        "ms-cv": msCv,
      },
      body: JSON.stringify({
        shortcuts: shortcuts,
        documentId: documentId,
      }),
    });

    if (response.status === 429) {
      throw await rateLimitError(response, "new emojis");
    }
    if (!response.ok) {
      const reason = await describeFailure(response);
      if (response.status === 409) {
        throw new Error(
          `An emoji named :${shortcuts.join(", ")}: probably already exists, so try another name (${reason})`,
        );
      }
      throw new Error(`Failed to create metadata: ${reason}`);
    }

    const metadataData = (await response.json()) as TeamsMetadataResponse;
    console.log("Metadata:", metadataData);
  }

  public async uploadFile(
    file: FileDetails,
    onStep?: (step: number) => void,
  ): Promise<void> {
    const shortcut = file.shortcut || `${file.name.split(".")[0]}`;
    onStep?.(1);
    const { id: documentId, msCv } = await this.createObject();
    console.log("Document ID:", documentId);
    onStep?.(2);
    await this.uploadFileContent(documentId, msCv, file);
    onStep?.(3);
    await this.sendMetadata(documentId, msCv, [shortcut]);
  }

  public async uploadFiles(
    files: FileDetails[],
    onProgress?: (progress: FileProgress) => void,
  ): Promise<{ success: boolean; error?: string; status?: string }> {
    if (files.length === 0) {
      return { success: false, error: "Please select files first" };
    }

    // Indexed by file so the summary keeps the original order
    const results: (string | null)[] = files.map(() => null);
    const uploadAt = async (index: number) => {
      const file = files[index];
      try {
        await this.uploadFile(file, (step) =>
          onProgress?.({ index, state: "uploading", step }),
        );
        onProgress?.({ index, state: "done", step: UPLOAD_STEPS });
      } catch (error) {
        console.error("Error uploading file:", file.name, error);
        const reason = error instanceof Error ? error.message : String(error);
        results[index] = `${file.shortcut || file.name}: ${reason}`;
        onProgress?.({ index, state: "error", step: 0, error: reason });
      }
    };

    await runWithConcurrency(files.length, UPLOAD_CONCURRENCY, uploadAt);

    const failures = results.filter((r): r is string => r !== null);
    if (failures.length > 0) {
      return {
        success: false,
        error: `Failed to upload ${failures.length} of ${files.length} emojis:\n${failures.join("\n")}`,
      };
    }

    return { success: true, status: "Your emojis have been uploaded" };
  }

  public async listCustomEmojis(): Promise<CustomEmojiResponseItem[]> {
    const response = await fetch(CUSTOM_EMOJI_METADATA_URL, {
      headers: {
        authorization: `Bearer ${this._chatToken}`,
        accept: "application/json",
        "x-ms-client-version": "1415/25021400912",
        "x-ms-client-type": "web",
      },
    });

    if (response.status === 429) {
      throw await rateLimitError(response, "requests");
    }
    if (!response.ok) {
      throw new Error(
        `Failed to load custom emojis: ${await describeFailure(response)}`,
      );
    }

    const data = (await response.json()) as CustomEmojiListResponse;
    return (data.categories || []).flatMap((category) => category.emoticons);
  }

  // `id` is the list entry's id, "<shortcut>;<documentId>"
  public async deleteCustomEmoji(id: string): Promise<void> {
    // Teams sends the separator unencoded, so keep ";" as is
    const path = encodeURIComponent(id).replace(/%3B/gi, ";");
    const response = await fetch(`${CUSTOM_EMOJI_METADATA_URL}/${path}`, {
      method: "DELETE",
      headers: {
        authorization: `Bearer ${this._chatToken}`,
        "x-ms-client-version": "1415/25021400912",
        "x-ms-client-type": "web",
      },
    });

    if (response.status === 429) {
      throw await rateLimitError(response, "deletions");
    }
    if (!response.ok) {
      const reason = await describeFailure(response);
      if (response.status === 403) {
        throw new Error(
          `Not allowed to delete this emoji, probably because someone else added it (${reason})`,
        );
      }
      throw new Error(`Failed to delete emoji: ${reason}`);
    }
  }

  public async fetchEmojiImage(documentId: string): Promise<Blob> {
    let lastError = "";
    for (const view of EMOJI_IMAGE_VIEWS) {
      const response = await fetch(
        `https://as-prod.asyncgw.teams.microsoft.com/v1/objects/${encodeURIComponent(documentId)}/views/${view}`,
        { headers: { authorization: `Bearer ${this._ic3Token}` } },
      );
      if (response.ok) {
        return response.blob();
      }
      lastError = await describeFailure(response);
    }
    throw new Error(`Failed to load emoji image: ${lastError}`);
  }
}

export default MsTeamsClient;
