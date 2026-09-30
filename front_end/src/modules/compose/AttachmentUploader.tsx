// filepath: my-app/src/modules/compose/AttachmentUploader.tsx
// Outlook-like attachment experience: drag/drop, browse, validate, show
// progress, allow remove/retry/cancel, and emit ready-to-send metadata.

import React, { useEffect, useRef, useState } from 'react';
import { UploaderComponent } from '@syncfusion/ej2-react-inputs';
import { Paperclip, X, Upload, RotateCw, FileText } from 'lucide-react';
import {
  AttachmentStatus,
  type EmailAttachment,
} from '../../types/attachments';
import {
  formatFileSize,
  getFileExtension,
  guessContentType,
  validateFile,
} from '../../utils/attachmentValidation';
import { uploadFileToDraft } from '../../services/attachmentService';
import { fileRegistry } from './attachmentFileRegistry';

// Minimal shape we read from Syncfusion's Uploader event args.
interface UploaderFileInfo {
  rawFile?: File;
}
type UploaderSelectedEvent = {
  filesData?: UploaderFileInfo[];
  modifiedFilesData?: UploaderFileInfo[];
  cancel: boolean;
  event: Event;
  isModified: boolean;
  isCanceled?: boolean;
  currentRequest?: { [key: string]: string }[];
  customFormData: { [key: string]: object }[];
  progressInterval: string;
};

interface AttachmentUploaderProps {
  /** Token to upload against (the draft ID created in Microsoft Graph). */
  draftId: string | null;
  /** The full attachment list owned by the parent. */
  attachments: EmailAttachment[];
  /** Called when the user has selected a file (after validation). */
  onAttachmentsChanged: (next: EmailAttachment[]) => void;
  /** Called when validation fails so the parent can show a toast. */
  onValidationError?: (message: string) => void;
  /** Disables all interactions (e.g. while sending). */
  disabled?: boolean;
}

const AttachmentUploader: React.FC<AttachmentUploaderProps> = ({
  draftId,
  attachments,
  onAttachmentsChanged,
  onValidationError,
  disabled = false,
}) => {
  // The Syncfusion React wrapper intercepts the `ref` prop and stores
  // the DOM element on its own instance, so React refs to the wrapper
  // never resolve. We instead use a callback ref on a wrapper <div> and
  // reach the underlying widget through the DOM (the wrapper renders
  // an <input> that the Uploader widget is attached to).
  const wrapperRef = useRef<HTMLDivElement | null>(null);
  const [isDragging, setIsDragging] = useState(false);

  /**
   * Validate and append a file. The parent owns the attachments state.
   */
  const handleFile = (file: File) => {
    const result = validateFile({ file, existingAttachments: attachments });
    if (!result.valid) {
      onValidationError?.(result.error || 'Invalid file.');
      return;
    }

    const id = `att-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const attachment: EmailAttachment = {
      id,
      fileName: file.name,
      contentType:
        file.type || guessContentType(file.name) || 'application/octet-stream',
      size: file.size,
      status: AttachmentStatus.Selected,
      progress: 0,
    };

    fileRegistry.set(id, file);
    onAttachmentsChanged([...attachments, attachment]);
  };

  // Syncfusion Uploader events. We bind them imperatively via the
  // component's `on()` method (the React event-prop form causes React 19
  // to complain about function-valued DOM attributes that Syncfusion
  // spreads onto its hidden <input>).
  const extractFiles = (args: UploaderSelectedEvent): File[] =>
    (args.filesData || [])
      .map((f) => f.rawFile)
      .filter((f): f is File => Boolean(f));

  /** Find the rendered <input> Syncfusion created. */
  const getUploaderInput = (): HTMLInputElement | null => {
    const wrapper = wrapperRef.current;
    if (!wrapper) return null;
    // The Syncfusion React wrapper renders the widget's <input>
    // directly inside the wrapper element.
    return wrapper.querySelector('input');
  };

  /** Shape of the Syncfusion widget instance we need. */
  type UploaderWidgetInstance = UploaderComponent & {
    on: (event: string, handler: (e: UploaderSelectedEvent) => void) => void;
    off: (event: string, handler: (e: UploaderSelectedEvent) => void) => void;
    clearAll: () => void;
  };

  /** Get the underlying Uploader widget instance via the DOM. */
  const getUploaderInstance = (): UploaderWidgetInstance | null => {
    const input = getUploaderInput();
    if (!input) return null;
    // Syncfusion stores the widget instance on the element as
    // `ej2_instances[0]`. See @syncfusion/ej2-base `addInstance`.
    const instances = (input as HTMLElement & {
      ej2_instances?: UploaderWidgetInstance[];
    }).ej2_instances;
    return instances?.[0] ?? null;
  };

  const onSelected = (args: UploaderSelectedEvent) => {
    const files = extractFiles(args);
    files.forEach((f) => handleFile(f));
    getUploaderInstance()?.clearAll?.();
  };

  const onFileDrop = (args: UploaderSelectedEvent) => {
    setIsDragging(false);
    const files = extractFiles(args);
    files.forEach((f) => handleFile(f));
    getUploaderInstance()?.clearAll?.();
  };

  // Bind/unbind Syncfusion events once the underlying widget is ready.
  // Syncfusion initialises the widget in `componentDidMount`, so by the
  // time this effect runs the input is in the DOM. We still need to
  // poll briefly because the Uploader may attach a tick or two later.
  useEffect(() => {
    let cancelled = false;
    let cleanup: (() => void) | null = null;

    const tryBind = () => {
      if (cancelled) return;
      const inst = getUploaderInstance();
      if (!inst) {
        // Widget not ready yet - try again on the next frame.
        const handle = window.requestAnimationFrame(tryBind);
        cleanup = () => window.cancelAnimationFrame(handle);
        return;
      }
      inst.on('selected', onSelected as (e: UploaderSelectedEvent) => void);
      inst.on('dropping', onFileDrop as (e: UploaderSelectedEvent) => void);
      cleanup = () => {
        // inst.off('selected', onSelected as (e: UploaderSelectedEvent) => void);
        // inst.off('dropping', onFileDrop as (e: UploaderSelectedEvent) => void);
      };
    };

    tryBind();
    return () => {
      cancelled = true;
      cleanup?.();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Manual drag/drop fallback
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = Array.from(e.dataTransfer.files || []);
    files.forEach((f) => handleFile(f));
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => setIsDragging(false);

  // Trigger the hidden file input
  const triggerBrowse = () => {
    const input = getUploaderInput();
    if (input) {
      input.click();
    }
  };

  // Actions
  const handleRemove = (id: string) => {
    fileRegistry.delete(id);
    onAttachmentsChanged(
      attachments.map((a) =>
        a.id === id ? { ...a, status: AttachmentStatus.Removed } : a
      )
    );
  };

  const handleCancel = (id: string) => {
    onAttachmentsChanged(
      attachments.map((a) =>
        a.id === id ? { ...a, status: AttachmentStatus.Cancelled } : a
      )
    );
  };

  const handleRetry = (id: string) => {
    onAttachmentsChanged(
      attachments.map((a) =>
        a.id === id
          ? { ...a, status: AttachmentStatus.Selected, progress: 0, error: undefined }
          : a
      )
    );
  };

  const visibleAttachments = attachments.filter(
    (a) => a.status !== AttachmentStatus.Removed
  );

  return (
    <div className="w-full">
      {/* Hidden Syncfusion Uploader - we use it as the file picker + drop target.
          Events are bound imperatively via the .on() API to avoid React 19
          warnings about function-valued DOM attributes. */}
      <div ref={wrapperRef} aria-hidden="true" style={{ position: 'absolute', width: 0, height: 0, overflow: 'hidden' }}>
        <UploaderComponent
          multiple
          type='file'
          autoUpload={false}
          showFileList={false}
          locale="en-US"
          cssClass="hidden-uploader"
          enabled={!disabled}
          htmlAttributes={{ 'aria-label': 'Attach files' }}
          selected={onSelected}
        />
      </div>

      {/* Visible Outlook-style drop zone + browse button */}
      <div
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onClick={triggerBrowse}
        role="button"
        tabIndex={0}
        aria-label="Drop files here or click to browse"
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            triggerBrowse();
          }
        }}
        className={[
          'flex items-center gap-3 px-4 py-3 rounded-md border-2 border-dashed cursor-pointer transition-colors',
          isDragging
            ? 'border-blue-500 bg-blue-50'
            : 'border-gray-300 bg-gray-50 hover:border-blue-400 hover:bg-blue-50',
          disabled ? 'opacity-50 pointer-events-none' : '',
        ].join(' ')}
      >
        <Paperclip size={20} className="text-gray-600" aria-hidden="true" />
        <div className="flex-1">
          <p className="text-sm text-gray-700 font-medium">
            {isDragging ? 'Drop files to attach' : 'Drop files here or click to browse'}
          </p>
          <p className="text-xs text-gray-500 mt-0.5">
            Max 150 MB per file &middot; total attachments up to 150 MB
          </p>
        </div>
        <Upload size={18} className="text-gray-500" aria-hidden="true" />
      </div>

      {/* Attachment list (rendered above the editor body) */}
      {visibleAttachments.length > 0 && (
        <ul
          className="mt-3 space-y-2"
          aria-label="Attached files"
          role="list"
        >
          {visibleAttachments.map((att) => (
            <AttachmentRow
              key={att.id}
              attachment={att}
              draftId={draftId}
              disabled={disabled}
              onRemove={() => handleRemove(att.id)}
              onRetry={() => handleRetry(att.id)}
              onCancel={() => handleCancel(att.id)}
              attachments={attachments}
              onAttachmentsChanged={onAttachmentsChanged}
            />
          ))}
        </ul>
      )}
    </div>
  );
};

export default AttachmentUploader;

// ---------------------------------------------------------------
// AttachmentRow - renders one attachment and owns the upload pipeline
// ---------------------------------------------------------------
interface AttachmentRowProps {
  attachment: EmailAttachment;
  draftId: string | null;
  disabled: boolean;
  onRemove: () => void;
  onRetry: () => void;
  onCancel: () => void;
  attachments: EmailAttachment[];
  onAttachmentsChanged: (next: EmailAttachment[]) => void;
}

const AttachmentRow: React.FC<AttachmentRowProps> = ({
  attachment,
  draftId,
  disabled,
  onRemove,
  onRetry,
  onCancel,
  attachments,
  onAttachmentsChanged,
}) => {
  const ext = getFileExtension(attachment.fileName);
  // Mark as cancelled ref so we can stop the in-flight upload
  const cancelledRef = useRef(false);

  // Keep a ref to the latest `attachments` array so the `update`
  // callback inside the upload's onProgress handler always sees the
  // freshest state (not the stale closure from when the useEffect ran).
  // Without this, progress updates would overwrite each other with
  // the original (stale) array and progress would never advance.
  const attachmentsRef = useRef(attachments);
  useEffect(() => {
    attachmentsRef.current = attachments;
  }, [attachments]);

  // Helper: update a single attachment in the parent's list. Uses
  // the ref so it always maps over the LATEST attachments array.
  const update = (patch: Partial<EmailAttachment>) => {
    const latest = attachmentsRef.current;
    onAttachmentsChanged(
      latest.map((a) => (a.id === attachment.id ? { ...a, ...patch } : a))
    );
  };

  // Auto-trigger upload when row is in `Selected` state and we have a draft
  React.useEffect(() => {
    if (
      attachment.status === AttachmentStatus.Selected &&
      draftId &&
      fileRegistry.has(attachment.id)
    ) {
      const file = fileRegistry.get(attachment.id)!;
      cancelledRef.current = false;

      update({ status: AttachmentStatus.Uploading, progress: 0 });

      uploadFileToDraft(draftId, file, {
        onProgress: (p) => {
          if (cancelledRef.current) return;
          update({
            progress: p.percent,
            currentChunk: p.currentChunk,
            totalChunks: p.totalChunks,
          });
        },
      })
        .then((graphAttachmentId) => {
          if (cancelledRef.current) return;
          update({
            status: AttachmentStatus.Uploaded,
            graphAttachmentId,
            progress: 100,
          });
        })
        .catch((err: unknown) => {
          if (cancelledRef.current) return;
          const raw = err instanceof Error ? err.message : String(err);
          // Provide a friendlier message for throttling (HTTP 429).
          const friendly = raw.includes('429')
            ? 'Upload was throttled by Microsoft Graph. Please retry in a moment.'
            : raw || 'Upload failed';
          update({
            status: AttachmentStatus.Failed,
            error: friendly,
          });
        });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [attachment.status, draftId]);

  const handleCancel = () => {
    cancelledRef.current = true;
    onCancel();
  };

  return (
    <li
      className="flex items-center gap-3 px-3 py-2 bg-white border border-gray-200 rounded-md"
      role="listitem"
      aria-label={`${attachment.fileName}, ${attachment.status}`}
    >
      <div className="w-9 h-9 rounded bg-blue-50 flex items-center justify-center shrink-0">
        <FileText size={18} className="text-blue-600" aria-hidden="true" />
      </div>

      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-medium text-gray-800 truncate" title={attachment.fileName}>
            {attachment.fileName}
          </span>
          {ext && (
            <span className="text-[10px] uppercase font-semibold text-gray-500 bg-gray-100 px-1.5 py-0.5 rounded">
              {ext}
            </span>
          )}
        </div>

        <div className="flex items-center gap-2 mt-1">
          <span className="text-xs text-gray-500">{formatFileSize(attachment.size)}</span>
          {attachment.status === AttachmentStatus.Uploading && (
            <div className="flex-1 flex items-center gap-2">
              <div className="flex-1 h-1.5 bg-gray-200 rounded-full overflow-hidden">
                <div
                  className="h-full bg-blue-600 transition-all"
                  style={{ width: `${attachment.progress || 0}%` }}
                  role="progressbar"
                  aria-valuenow={attachment.progress || 0}
                  aria-valuemin={0}
                  aria-valuemax={100}
                  aria-label={`Upload progress for ${attachment.fileName}`}
                />
              </div>
              <span className="text-[11px] text-gray-500 whitespace-nowrap">
                {attachment.progress || 0}%
                {attachment.totalChunks
                  ? ` (Chunk ${attachment.currentChunk || 1} of ${attachment.totalChunks})`
                  : ''}
              </span>
            </div>
          )}
          {attachment.status === AttachmentStatus.Uploaded && (
            <span className="text-[11px] text-green-700 font-medium">Uploaded</span>
          )}
          {attachment.status === AttachmentStatus.Failed && (
            <span className="text-[11px] text-red-600 font-medium">
              {attachment.error || 'Upload failed'}
            </span>
          )}
          {attachment.status === AttachmentStatus.Cancelled && (
            <span className="text-[11px] text-gray-500">Cancelled</span>
          )}
        </div>
      </div>

      <div className="flex items-center gap-1 shrink-0">
        {attachment.status === AttachmentStatus.Failed && (
          <button
            type="button"
            onClick={onRetry}
            disabled={disabled || !draftId}
            className="p-1.5 rounded hover:bg-gray-100 disabled:opacity-50"
            aria-label={`Retry upload of ${attachment.fileName}`}
            title="Retry"
          >
            <RotateCw size={16} className="text-blue-600" />
          </button>
        )}
        {(attachment.status === AttachmentStatus.Uploading ||
          attachment.status === AttachmentStatus.Selected) && (
          <button
            type="button"
            onClick={handleCancel}
            disabled={disabled}
            className="p-1.5 rounded hover:bg-gray-100 disabled:opacity-50"
            aria-label={`Cancel upload of ${attachment.fileName}`}
            title="Cancel"
          >
            <X size={16} className="text-gray-600" />
          </button>
        )}
        <button
          type="button"
          onClick={onRemove}
          disabled={disabled}
          className="p-1.5 rounded hover:bg-gray-100 disabled:opacity-50"
          aria-label={`Remove ${attachment.fileName}`}
          title="Remove"
        >
          <X size={16} className="text-red-600" />
        </button>
      </div>
    </li>
  );
};
