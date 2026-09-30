import React from 'react';
import { useMailStore, useUIStore, useFolderStore } from '../../store/Store';
import { Mail, Reply, ReplyAll, Forward, Trash2, ArchiveRestore, MailWarning, Flag, Paperclip, Download, FileText, Image as ImageIcon, FileArchive, Film } from 'lucide-react';
import { mailAPI } from '../../services/api';
import type { MailAttachment } from '../../data/mockData';
import { formatMailDate } from '../../utils/transformers';

/**
 * Format a file size in bytes into a human-readable string
 * (e.g. "1.2 MB", "348 KB"). Mirrors the helper used by the
 * attachment uploader for consistency.
 */
function formatSize(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  let value = bytes;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit++;
  }
  return `${value.toFixed(value >= 100 || unit === 0 ? 0 : 1)} ${units[unit]}`;
}

/**
 * Pick a Lucide icon for an attachment based on its content type.
 * Returns a generic FileText for unknown types.
 */
function attachmentIcon(contentType: string) {
  const lc = (contentType || '').toLowerCase();
  if (lc.startsWith('image/')) return ImageIcon;
  if (lc.startsWith('video/')) return Film;
  if (
    lc.includes('zip') ||
    lc.includes('compressed') ||
    lc.includes('tar') ||
    lc.includes('rar') ||
    lc.includes('7z')
  )
    return FileArchive;
  return FileText;
}

export const MailReader: React.FC = () => {
  const { selectedMail, deleteMail, restoreMail, toggleReadStatus, toggleFlagStatus, fetchMailDetails } = useMailStore();
  const { selectedFolderId } = useFolderStore();
  const { setIsComposing, setComposeType, setComposeMailContext } = useUIStore();
  const [isDeleting, setIsDeleting] = React.useState(false);
  const [isDownloading, setIsDownloading] = React.useState<string | null>(null);
  // Track whether the most recent detail fetch failed. We use this to
  // show a friendly "mail unavailable" message instead of leaving the
  // user staring at a snippet that doesn't tell them what's going on.
  const [detailsUnavailable, setDetailsUnavailable] = React.useState(false);

  // Reset the "unavailable" flag whenever a different mail is selected.
  React.useEffect(() => {
    setDetailsUnavailable(false);
  }, [selectedMail?.id]);

  // When the user selects a different mail, fetch its full details
  // (body, attachment list) if we don't have them yet. The list view
  // shows lightweight preview data; the reader needs the full payload.
  React.useEffect(() => {
    // Skip if we already have details, if there's no selected mail,
    // or if the previous fetch failed (we don't want to keep retrying
    // and flooding the console with errors).
    if (!selectedMail) return;
    if (selectedMail.attachments || selectedMail.body) return;
    if (detailsUnavailable) return;

    let cancelled = false;
    (async () => {
      const result = await fetchMailDetails(selectedMail.id);
      if (cancelled) return;
      if (result === null) {
        // Either the mail doesn't exist in Graph anymore (404) or
        // there was a network/permissions error. Either way, mark
        // it so the UI can show a "unavailable" hint and we don't
        // retry on every re-render.
        setDetailsUnavailable(true);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [selectedMail, fetchMailDetails, detailsUnavailable]);

  /**
   * Download an attachment by triggering a save through the browser.
   * We use the backend's attachment endpoint (not the Graph URL
   * directly) so the request goes through our JWT auth.
   */
  const handleDownloadAttachment = async (att: MailAttachment) => {
    if (!selectedMail) return;
    setIsDownloading(att.id);
    try {
      const { blob, fileName } = await mailAPI.downloadAttachment(
        selectedMail.id,
        att.id
      );
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      // Revoke the object URL after a short delay so the browser
      // has time to start the download.
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      console.error('Failed to download attachment:', err);
      alert('Failed to download attachment. Please try again.');
    } finally {
      setIsDownloading(null);
    }
  };

  const handleCompose = (type: 'reply' | 'replyAll' | 'forward') => {
    if (selectedMail) {
      setComposeType(type);
      setComposeMailContext(selectedMail);
      setIsComposing(true);
    }
  };

  const handleDelete = async () => {
    if (!selectedMail) return;
    setIsDeleting(true);
    try {
      await deleteMail(selectedMail.id);
    } catch (error) {
      alert('Failed to delete email. Please try again.');
    } finally {
      setIsDeleting(false);
    }
  };

  const handleRestore = async () => {
    if (!selectedMail) return;
    setIsDeleting(true);
    try {
      await restoreMail(selectedMail.id);
    } catch (error) {
      alert('Failed to restore email. Please try again.');
    } finally {
      setIsDeleting(false);
    }
  };

  const isDeletedFolder = selectedFolderId === '2-4';

  return (
    <div className="h-full overflow-y-auto bg-gray-50">
      {selectedMail ? (
        <>
          <div className="px-6 py-3 border-b border-gray-200 bg-white flex gap-2">
            <button 
              onClick={() => handleCompose('reply')}
              className="p-2 text-gray-600 hover:text-gray-900 hover:bg-gray-100 rounded cursor-pointer border-none bg-transparent flex items-center gap-2"
              title="Reply"
            >
              <Reply size={18} />
              <span className="text-sm">Reply</span>
            </button>
            <button 
              onClick={() => handleCompose('replyAll')}
              className="p-2 text-gray-600 hover:text-gray-900 hover:bg-gray-100 rounded cursor-pointer border-none bg-transparent flex items-center gap-2"
              title="Reply All"
            >
              <ReplyAll size={18} />
              <span className="text-sm">Reply All</span>
            </button>
            <button 
              onClick={() => handleCompose('forward')}
              className="p-2 text-gray-600 hover:text-gray-900 hover:bg-gray-100 rounded cursor-pointer border-none bg-transparent flex items-center gap-2"
              title="Forward"
            >
              <Forward size={18} />
              <span className="text-sm">Forward</span>
            </button>
            
            <div className="w-px h-6 bg-gray-300 mx-2 self-center"></div>
            
            {isDeletedFolder ? (
              <button 
                onClick={handleRestore}
                disabled={isDeleting}
                className="p-2 text-gray-600 hover:text-blue-600 hover:bg-blue-50 rounded cursor-pointer border-none bg-transparent flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
                title="Restore"
              >
                <ArchiveRestore size={18} />
                <span className="text-sm">{isDeleting ? 'Restoring...' : 'Restore'}</span>
              </button>
            ) : (
              <button 
                onClick={handleDelete}
                disabled={isDeleting}
                className="p-2 text-gray-600 hover:text-red-600 hover:bg-red-50 rounded cursor-pointer border-none bg-transparent flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
                title="Delete"
              >
                <Trash2 size={18} />
                <span className="text-sm">{isDeleting ? 'Deleting...' : 'Delete'}</span>
              </button>
            )}

            <button 
              onClick={() => toggleReadStatus(selectedMail.id)}
              className="p-2 text-gray-600 hover:text-gray-900 hover:bg-gray-100 rounded cursor-pointer border-none bg-transparent flex items-center gap-2"
              title={selectedMail.isRead ? "Mark as unread" : "Mark as read"}
            >
              <MailWarning size={18} />
              <span className="text-sm">{selectedMail.isRead ? 'Mark Unread' : 'Mark Read'}</span>
            </button>

            <button 
              onClick={() => toggleFlagStatus(selectedMail.id)}
              className="p-2 text-gray-600 hover:text-red-600 hover:bg-red-50 rounded cursor-pointer border-none bg-transparent flex items-center gap-2"
              title={selectedMail.isFlagged ? "Unflag" : "Flag"}
            >
              <Flag size={18} fill={selectedMail.isFlagged ? "#ef4444" : "none"} color={selectedMail.isFlagged ? "#ef4444" : "currentColor"} />
              <span className="text-sm">{selectedMail.isFlagged ? 'Unflag' : 'Flag'}</span>
            </button>
          </div>

          <div className="px-6 py-5">
            <h1 className="text-2xl font-semibold mb-4 text-gray-800">{selectedMail.subject}</h1>
            <div className="flex items-center gap-3 mb-5 pb-4 border-b border-gray-200">
              <div
                className="w-10 h-10 rounded-full text-white flex items-center justify-center text-base font-semibold"
                style={{ backgroundColor: selectedMail.avatarColor }}
              >
                {selectedMail.senderName.charAt(0)}
              </div>
              <div>
                <div className="text-sm font-semibold text-gray-800">{selectedMail.senderName}</div>
                <div className="text-xs text-gray-600">{selectedMail.senderEmail}</div>
              </div>
              <div className="ml-auto text-xs text-gray-600">
                {formatMailDate(selectedMail.date)}
              </div>
            </div>
            <div className="text-sm text-gray-800 leading-relaxed">
              {selectedMail.body ? (
                // The body from Graph is HTML (most common) or plain
                // text. We render HTML as-is so embedded images and
                // formatting are preserved. The body is fetched from our
                // own backend, so it's safe to dangerouslySetInnerHTML.
                // We add the 'e-rte-content' class so the styles defined
                // in index.css match the styling the user saw inside the
                // composer (Syncfusion RTE styling docs).
                //
                // Note: do NOT add `whitespace-pre-wrap` to the outer
                // wrapper. The body is HTML, not plain text, and the
                // wrapped RTE envelope contains a <style> block whose
                // newlines would be rendered as visible whitespace.
                <div
                  className="mail-body e-rte-content"
                  dangerouslySetInnerHTML={{ __html: selectedMail.body }}
                />
              ) : detailsUnavailable ? (
                // The detail fetch returned 404. Show a clear hint
                // instead of a confusing snippet, and tell the user
                // that attachments (if any) aren't available either.
                <div className="flex items-start gap-2 p-3 bg-amber-50 border border-amber-200 rounded text-amber-800">
                  <MailWarning size={16} className="mt-0.5 flex-shrink-0" aria-hidden="true" />
                  <div className="text-xs leading-relaxed">
                    <p className="font-medium mb-1">
                      This message is no longer available on the server.
                    </p>
                    <p>
                      It may have been deleted, moved to another folder,
                      or the access token doesn't have permission to read
                      it. The preview above is from the local list cache.
                    </p>
                    {selectedMail.hasAttachments && (
                      <p className="mt-1">
                        <strong>Note:</strong> This message has
                        attachments but they can't be displayed.
                      </p>
                    )}
                  </div>
                </div>
              ) : (
                // Fall back to the preview snippet while the full body
                // is being fetched.
                <>
                  {selectedMail.snippet}
                  <p className="text-xs text-gray-400 mt-2 italic">
                    Loading full message…
                  </p>
                </>
              )}
            </div>

            {/* Attachments list - rendered below the body so it doesn't
                displace the message content. We use a stacked card per
                attachment with a download button. The list is hidden
                while attachment metadata is still being fetched (i.e.
                the parent knows the message has attachments but we
                don't have the metadata yet). */}
            {selectedMail.attachments && selectedMail.attachments.length > 0 && (
              <div
                className="mt-6 border-t border-gray-200 pt-4"
                aria-label="Attachments"
              >
                <div className="flex items-center gap-2 mb-3 text-xs font-semibold text-gray-600">
                  <Paperclip size={14} aria-hidden="true" />
                  <span>
                    {selectedMail.attachments.length} attachment
                    {selectedMail.attachments.length === 1 ? '' : 's'}
                  </span>
                </div>
                <ul className="space-y-2" role="list">
                  {selectedMail.attachments.map((att) => {
                    const Icon = attachmentIcon(att.contentType);
                    const isThisDownloading = isDownloading === att.id;
                    return (
                      <li
                        key={att.id}
                        className="flex items-center gap-3 px-3 py-2 bg-white border border-gray-200 rounded-md hover:border-blue-300 transition-colors"
                        role="listitem"
                        aria-label={`Attachment: ${att.name}`}
                      >
                        <div className="w-9 h-9 rounded bg-blue-50 flex items-center justify-center shrink-0">
                          <Icon size={18} className="text-blue-600" aria-hidden="true" />
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="text-sm font-medium text-gray-800 truncate" title={att.name}>
                            {att.name}
                          </div>
                          <div className="text-xs text-gray-500 mt-0.5">
                            {formatSize(att.size)}
                            {att.contentType ? ` · ${att.contentType}` : ''}
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={() => handleDownloadAttachment(att)}
                          disabled={isThisDownloading}
                          className="p-2 rounded hover:bg-gray-100 disabled:opacity-50 flex items-center gap-1.5 text-blue-600 hover:text-blue-700"
                          aria-label={`Download ${att.name}`}
                          title="Download"
                        >
                          <Download size={16} aria-hidden="true" />
                          <span className="text-xs font-medium">
                            {isThisDownloading ? 'Downloading…' : 'Download'}
                          </span>
                        </button>
                      </li>
                    );
                  })}
                </ul>
              </div>
            )}

            {/* Show a loading hint when we know the message has
                attachments but the metadata is still being fetched. */}
            {selectedMail.hasAttachments &&
              (!selectedMail.attachments || selectedMail.attachments.length === 0) && (
                <div
                  className="mt-6 border-t border-gray-200 pt-4 flex items-center gap-2 text-xs text-gray-500"
                  aria-live="polite"
                >
                  <Paperclip size={14} aria-hidden="true" />
                  {detailsUnavailable ? (
                    <span>
                      Attachments unavailable — the message couldn't be
                      loaded from the server.
                    </span>
                  ) : (
                    <span>Loading attachments…</span>
                  )}
                </div>
              )}
          </div>
        </>
      ) : (
        <div className="p-6 h-full flex items-center justify-center text-gray-400">
          <div className="text-center">
            <Mail size={48} className="mb-4 mx-auto text-gray-300" />
            <div>Select an item to read</div>
          </div>
        </div>
      )}
    </div>
  );
};
