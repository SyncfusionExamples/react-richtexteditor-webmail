// filepath: my-app/src/utils/attachmentValidation.ts
// Client-side validation rules for email attachments

import {
  AttachmentStatus,
  DEFAULT_VALIDATION_CONFIG,
  type AttachmentValidationConfig,
  type AttachmentValidationResult,
  type EmailAttachment,
  BLOCKED_FILENAME_CHARS,
} from '../types/attachments';

/**
 * Get the extension of a filename, lowercased, no leading dot.
 * Returns empty string if no extension.
 */
export function getFileExtension(fileName: string): string {
  const idx = fileName.lastIndexOf('.');
  if (idx <= 0 || idx === fileName.length - 1) return '';
  return fileName.substring(idx + 1).toLowerCase();
}

/**
 * Get the mime type guess from a filename. Falls back to
 * the provided default when the extension is unknown.
 */
export function guessContentType(fileName: string, fallback = 'application/octet-stream'): string {
  const ext = getFileExtension(fileName);
  const map: Record<string, string> = {
    pdf: 'application/pdf',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ppt: 'application/vnd.ms-powerpoint',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    txt: 'text/plain',
    csv: 'text/csv',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    png: 'image/png',
    gif: 'image/gif',
    bmp: 'image/bmp',
    webp: 'image/webp',
    zip: 'application/zip',
    rar: 'application/x-rar-compressed',
    '7z': 'application/x-7z-compressed',
    mp4: 'video/mp4',
    mov: 'video/quicktime',
    mp3: 'audio/mpeg',
    wav: 'audio/wav',
  };
  return map[ext] || fallback;
}

/**
 * Generate a fingerprint used to detect duplicates in the attachment list.
 * Combines name + size + lastModified (when available).
 */
export function getFileFingerprint(file: File): string {
  const lastModified = file.lastModified ?? 0;
  return `${file.name}::${file.size}::${lastModified}`;
}

interface ValidateFileOptions {
  file: File;
  existingAttachments: EmailAttachment[];
  config?: AttachmentValidationConfig;
}

/**
 * Validate a single file against the configured rules. Returns a result
 * describing whether the file is acceptable and a human-readable error
 * message when it is not.
 */
export function validateFile({
  file,
  existingAttachments,
  config = DEFAULT_VALIDATION_CONFIG,
}: ValidateFileOptions): AttachmentValidationResult {
  // Filename checks
  if (!file.name || file.name.trim().length === 0) {
    return { valid: false, error: 'Invalid file name.' };
  }
  if (file.name.length > config.maxFilenameLength) {
    return { valid: false, error: 'Invalid file name.' };
  }
  if (BLOCKED_FILENAME_CHARS.test(file.name)) {
    return { valid: false, error: 'Invalid file name.' };
  }

  // Extension checks
  const ext = getFileExtension(file.name);
  if (config.blockedExtensions.includes(ext)) {
    return { valid: false, error: 'This file type is not supported.' };
  }
  if (ext && !config.allowedExtensions.includes(ext)) {
    return { valid: false, error: 'This file type is not supported.' };
  }

  // Size checks
  if (file.size <= 0) {
    return { valid: false, error: 'Invalid file name.' };
  }
  if (file.size > config.maxFileSize) {
    return { valid: false, error: 'File exceeds maximum allowed size of 150 MB.' };
  }

  // Total size check
  const currentTotal = existingAttachments
    .filter((a) => a.status !== AttachmentStatus.Removed && a.status !== AttachmentStatus.Failed)
    .reduce((sum, a) => sum + a.size, 0);
  if (currentTotal + file.size > config.maxTotalSize) {
    return { valid: false, error: 'Total attachment size exceeds the allowed limit.' };
  }

  // Duplicate check
  const isDuplicate = existingAttachments.some(
    (a) =>
      a.fileName === file.name &&
      a.size === file.size &&
      a.status !== AttachmentStatus.Removed &&
      a.status !== AttachmentStatus.Failed,
  );
  if (isDuplicate) {
    return { valid: false, error: 'This file is already attached.' };
  }

  return { valid: true };
}

/**
 * Format a byte count as a human-readable size string (e.g. 2.4 MB).
 */
export function formatFileSize(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  let size = bytes;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex++;
  }
  const decimals = unitIndex === 0 ? 0 : 1;
  return `${size.toFixed(decimals)} ${units[unitIndex]}`;
}
