// filepath: my-app/src/types/attachments.ts
// Email attachment types and enums for the attachment lifecycle

export const AttachmentStatus = {
  Selected: 'selected',
  Validating: 'validating',
  Uploading: 'uploading',
  Uploaded: 'uploaded',
  Failed: 'failed',
  Cancelled: 'cancelled',
  Removed: 'removed',
} as const;

export type AttachmentStatus = (typeof AttachmentStatus)[keyof typeof AttachmentStatus];

export interface EmailAttachment {
  /** Local UI ID (UUID) */
  id: string;
  fileName: string;
  contentType: string;
  /** Size in bytes */
  size: number;
  status: AttachmentStatus;
  /** Progress percentage 0-100 */
  progress?: number;
  /** Optional error message for Failed status */
  error?: string;
  /** Microsoft Graph attachment ID (after successful upload) */
  graphAttachmentId?: string;
  /** Microsoft Graph upload session URL (for chunked uploads) */
  uploadSessionUrl?: string;
  /** Chunked upload tracking */
  totalChunks?: number;
  currentChunk?: number;
}

export interface AttachmentValidationResult {
  valid: boolean;
  error?: string;
}

export interface AttachmentValidationConfig {
  /** Max single file size in bytes (default 150 MB) */
  maxFileSize: number;
  /** Max total attachments size in bytes (default 150 MB) */
  maxTotalSize: number;
  /** Allowed extensions (lowercase, no leading dot) */
  allowedExtensions: string[];
  /** Blocked extensions (lowercase, no leading dot) */
  blockedExtensions: string[];
  /** Max filename length */
  maxFilenameLength: number;
}

export const DEFAULT_VALIDATION_CONFIG: AttachmentValidationConfig = {
  maxFileSize: 150 * 1024 * 1024, // 150 MB
  maxTotalSize: 150 * 1024 * 1024, // 150 MB
  allowedExtensions: [
    // Documents
    'pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'csv',
    // Images
    'jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp',
    // Archives
    'zip', 'rar', '7z',
    // Media
    'mp4', 'mov', 'mp3', 'wav',
  ],
  blockedExtensions: [
    'exe', 'bat', 'cmd', 'com', 'dll', 'msi',
    'vbs', 'js', 'jse', 'ps1', 'scr', 'reg',
  ],
  maxFilenameLength: 255,
};

export const BLOCKED_FILENAME_CHARS = /[<>:"/\\|?*]/;
