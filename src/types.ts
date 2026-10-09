export interface ProcessResult {
  success: boolean;
  error?: string;
  status?: string;
}

export interface TeamsUploadResponse {
  id: string;
}

export interface TeamsMetadataResponse {
  success: boolean;
}

export interface FileDetails {
  name: string;
  size: number;
  type: string;
  base64: string;
  // Emoji name to register; defaults to the file name without its extension
  shortcut?: string;
}

export type UploadState = "pending" | "uploading" | "done" | "error";

// Each file goes through three requests: create object, upload image, save name
export const UPLOAD_STEPS = 3;

export interface FileProgress {
  index: number;
  state: UploadState;
  step: number;
  error?: string;
}

export interface UploadProgressItem {
  name: string;
  shortcut: string;
  state: UploadState;
  step: number;
  error?: string;
}

// One entry from GET /api/csa/{region}/api/v1/customemoji/metadata
export interface CustomEmojiResponseItem {
  id: string; // "<shortcut>;<documentId>", as used by the delete endpoint
  documentId: string;
  shortcuts: string[];
  description?: string;
  createdOn?: number;
  isDeleted?: boolean;
  keywords?: string[];
  creator?: string;
  etag?: string;
}

export interface CustomEmojiListResponse {
  continuationToken?: string;
  categories?: {
    id: string;
    title: string;
    emoticons: CustomEmojiResponseItem[];
  }[];
}
