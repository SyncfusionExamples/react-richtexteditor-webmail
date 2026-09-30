import React, { useRef, useState, useEffect } from 'react';
import {
  RichTextEditorComponent,
  Toolbar,
  Link,
  Image,
  HtmlEditor,
  QuickToolbar,
  Inject
} from '@syncfusion/ej2-react-richtexteditor';
import { MultiSelectComponent } from '@syncfusion/ej2-react-dropdowns';
import { useUIStore, useMailStore, useFolderStore } from '../../store/Store';
import { contactsData } from '../../data/mockData';
import type { MailData } from '../../data/mockData';
import type { MailAttachment } from '../../data/mockData';
import { X, Send, Paperclip, Trash2, FileText } from 'lucide-react';
import { mailAPI } from '../../services/api';
import AttachmentUploader from './AttachmentUploader';
import { AttachmentStatus, type EmailAttachment } from '../../types/attachments';
import { fileRegistry } from './attachmentFileRegistry';
import { wrapRteBody } from '../../utils/rteStyles';
import { formatMailDate } from '../../utils/transformers';

export const ComposeEditor: React.FC = () => {
  const { composeType, composeMailContext, clearCompose } = useUIStore();
  const { sendMail, setMails } = useMailStore();
  const { getSystemFolderId } = useFolderStore();
  const rteRef = useRef<RichTextEditorComponent>(null);

  const [to, setTo] = useState<string[]>([]);
  const [cc, setCc] = useState<string[]>([]);
  const [subject, setSubject] = useState('');
  const [showDiscardDialog, setShowDiscardDialog] = useState(false);
  const [attachments, setAttachments] = useState<EmailAttachment[]>([]);
  const [draftId, setDraftId] = useState<string | null>(null);
  const [isCreatingDraft, setIsCreatingDraft] = useState(false);
  const [toast, setToast] = useState<{ type: 'error' | 'info'; message: string } | null>(null);

  const extractText = (html: string) => {
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = html;
    return tempDiv.textContent || tempDiv.innerText || 'No content';
  };

  const escapeHtml = (value: string) =>
    value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');

  const buildQuotedBody = () => {
    if (!composeMailContext) {
      return '<p><br/></p>';
    }

    const sourceBody = composeMailContext.body || `<p>${escapeHtml(composeMailContext.snippet)}</p>`;
    const quotedMeta = `<p><b>From:</b> ${escapeHtml(composeMailContext.senderName)} &lt;${escapeHtml(composeMailContext.senderEmail)}&gt;<br/><b>Date:</b> ${escapeHtml(formatMailDate(composeMailContext.date))}<br/><b>Subject:</b> ${escapeHtml(composeMailContext.subject)}</p>`;
    const spacer = '<p><br/></p><p><br/></p>';

    return `${spacer}<hr/>${quotedMeta}<blockquote style="margin: 0 0 0 12px; padding-left: 12px; border-left: 2px solid #d1d5db;">${sourceBody}</blockquote>`;
  };

  useEffect(() => {
    // Whenever the compose session changes, reset attachments state.
    // These state updates are intentional - they sync the component's
    // local state with the new composeType/composeMailContext props
    // (an external source of truth managed by the global UI store).
    /* eslint-disable react-hooks/set-state-in-effect */
    setAttachments([]);
    setDraftId(null);
    /* eslint-enable react-hooks/set-state-in-effect */
    fileRegistry.clear();

    if (composeType && composeMailContext) {
      if (composeType === 'reply') {
        setTo([composeMailContext.senderEmail]);
        setSubject(composeMailContext.subject.startsWith('Re:') ? composeMailContext.subject : `Re: ${composeMailContext.subject}`);
        setCc([]);
        if (rteRef.current) {
          rteRef.current.value = buildQuotedBody();
        }
      } else if (composeType === 'replyAll') {
        setTo([composeMailContext.senderEmail]); // Add logic to handle other recipients if they existed
        setSubject(composeMailContext.subject.startsWith('Re:') ? composeMailContext.subject : `Re: ${composeMailContext.subject}`);
        setCc([]);
        if (rteRef.current) {
          rteRef.current.value = buildQuotedBody();
        }
      } else if (composeType === 'forward') {
        setTo([]);
        setCc([]);
        setSubject(composeMailContext.subject.startsWith('Fw:') ? composeMailContext.subject : `Fw: ${composeMailContext.subject}`);
        if (rteRef.current) {
          rteRef.current.value = buildQuotedBody();
        }
      }
    } else {
      setTo([]);
      setCc([]);
      setSubject('');
      if (rteRef.current) {
        rteRef.current.value = '<p><br/></p>';
      }
    }
  }, [composeType, composeMailContext]);

  // Auto-dismiss the validation toast
  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 4000);
    return () => clearTimeout(t);
  }, [toast]);

  // Esc-key support for closing the compose window. Mirrors the
  // behavior of the X button (`handleDiscardClick`), with one
  // extra rule so popups behave naturally:
  //   - If a Syncfusion popup is open (To/Cc MultiSelect
  //     suggestion list, RTE toolbar dropdowns like Font /
  //     Size / Format / Color, image/link dialog, full-screen
  //     RTE dialog, etc.) → Esc only closes the popup; the
  //     compose stays open. We do NOT call preventDefault so
  //     Syncfusion's own Esc handler can do its thing.
  //   - Otherwise, if the Discard dialog is already open →
  //     the first Esc press dismisses it; only a second Esc
  //     would then act on the compose.
  //   - Otherwise, if any field has content (To, Cc, Subject,
  //     or RTE body) → open the Discard draft? dialog.
  //   - Otherwise → close the compose directly.
  //
  // We bind the listener exactly once and read the latest
  // `clearCompose` through a ref so the effect doesn't re-run
  // on every render (the UI store's `clearCompose` is a plain
  // function whose identity changes on every render — that
  // would otherwise re-bind the listener on every state change).
  // The "has content" check is inlined here (instead of calling
  // `hasContent` from the closure) so we always read the
  // current value of every field via refs, not the value
  // captured by the listener at first-render time.
  //
  // The RTE-empty detection uses the same approach as
  // `hasContent`: strip HTML to text and trim. String-comparing
  // the raw HTML against '<p><br/></p>' is unreliable because
  // Syncfusion can emit '<p><br></p>', '<p>&nbsp;</p>', or just
  // '<br/>' depending on the browser and what the user has
  // done, and any of those would incorrectly look like content.
  const subjectRef = useRef(subject);
  const toRef = useRef(to);
  const ccRef = useRef(cc);
  const clearComposeRef = useRef(clearCompose);
  const showDiscardDialogRef = useRef(showDiscardDialog);
  useEffect(() => {
    subjectRef.current = subject;
  }, [subject]);
  useEffect(() => {
    toRef.current = to;
  }, [to]);
  useEffect(() => {
    ccRef.current = cc;
  }, [cc]);
  useEffect(() => {
    clearComposeRef.current = clearCompose;
  }, [clearCompose]);
  useEffect(() => {
    showDiscardDialogRef.current = showDiscardDialog;
  }, [showDiscardDialog]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      try {
        // If a Syncfusion popup is currently open (To/Cc
        // MultiSelect suggestion list, RTE toolbar dropdowns
        // like Font / Size / Format / Color, image/link dialog,
        // full-screen RTE dialog, etc.), let the popup handle
        // Esc on its own — the user wants to close the popup,
        // NOT the entire compose pane.
        //
        // Syncfusion's `Popup` class adds the `e-popup-open`
        // class to the popup root element while it's open
        // (see @syncfusion/ej2-popups/src/popup/popup.js,
        // CLASSNAMES.OPEN = 'e-popup-open'). Querying the DOM
        // for this class is the most reliable cross-component
        // way to detect an open popup without holding a ref
        // to every Syncfusion widget.
        //
        // We DO NOT call preventDefault/stopPropagation in
        // this branch so Syncfusion's own Esc handler can
        // close the popup. The compose stays open.
        if (document.querySelector('.e-popup-open')) {
          return;
        }
        e.preventDefault();
        e.stopPropagation();
        // If the Discard dialog is already open, the first Esc
        // press dismisses it (Outlook-style "Esc closes the
        // top-most modal"). Only a second Esc would then act
        // on the compose.
        if (showDiscardDialogRef.current) {
          setShowDiscardDialog(false);
          return;
        }
        // Read latest field values from refs. The RTE check
        // strips HTML to text first so empty-state variants
        // ('<p><br/></p>', '<p><br></p>', '<p>&nbsp;</p>',
        // '<br/>', whitespace, etc.) all correctly evaluate
        // to "no content".
        const htmlContent = rteRef.current?.value || '';
        const rteText = extractText(htmlContent)
          .replace(/\u00A0/g, ' ')
          .trim();
        const hasText =
          subjectRef.current.trim().length > 0 ||
          toRef.current.length > 0 ||
          ccRef.current.length > 0 ||
          rteText.length > 0;
        if (hasText) {
          setShowDiscardDialog(true);
        } else {
          // Nothing to save — close the compose directly.
          clearComposeRef.current();
        }
      } catch (err) {
        // Surface any failure to the dev console so a future
        // regression doesn't fail silently.
        // eslint-disable-next-line no-console
        console.error('[ComposeEditor] Esc handler failed:', err);
      }
    };
    // Listen on `document` in the CAPTURE phase so we run
    // before any inner handler (e.g. Syncfusion widgets) can
    // call stopPropagation(). `document` is the lowest-level
    // node that always receives keydown for any focused
    // element in the page.
    document.addEventListener('keydown', handleKeyDown, true);
    return () => document.removeEventListener('keydown', handleKeyDown, true);
  }, []); // Bind once. All dynamic state is read via refs.

  const rteToolbarSettings = {
    items: ['Bold', 'Italic', 'Underline', 'StrikeThrough',
      'FontName', 'FontSize', 'FontColor', 'BackgroundColor',
      'LowerCase', 'UpperCase', '|',
      'Formats', 'Alignments', 'OrderedList', 'UnorderedList',
      'Outdent', 'Indent', '|',
      'CreateLink', 'Image', '|', 'ClearFormat', 'Print',
      'SourceCode', 'FullScreen', '|', 'Undo', 'Redo']
  };

  const fields = { text: 'Name', value: 'Email' };

  const itemTemplate = (data: any) => {
    return (
      <div className="flex items-center gap-3 py-1">
        <div 
          className="w-8 h-8 rounded-full text-white flex items-center text-sm font-semibold shrink-0 leading-none"
          style={{ backgroundColor: data.Color || '#ccc' }}
        >
          {data.Name.charAt(0)}
        </div>
        <div className="flex flex-col justify-center">
          <span className="text-sm font-semibold text-gray-800 leading-tight">{data.Name}</span>
          <span className="text-xs text-gray-500 leading-tight mt-0.5">{data.Email}</span>
        </div>
      </div>
    );
  };

  const valueTemplate = (data: any) => {
    return (
      <div className="flex items-center gap-2 px-1 py-0.5">
        <div 
          className="w-5 h-5 rounded-full text-white flex items-center justify-center text-[10px] font-semibold shrink-0 leading-none"
          style={{ backgroundColor: data.Color || '#ccc' }}
        >
          {data.Name.charAt(0)}
        </div>
        <span className="text-sm leading-tight">{data.Name}</span>
      </div>
    );
  };

  /**
   * Lazily create a Microsoft Graph draft when an attachment is added.
   * The draft is the destination for all attachment uploads. If the user
   * sends the email, the draft is what gets sent (so all attachments
   * travel with it).
   *
   * Returns null on failure; callers are expected to mark the affected
   * attachment row as Failed with a clear reason.
   */
  const ensureDraft = async (): Promise<string | null> => {
    if (draftId) return draftId;
    if (isCreatingDraft) return null;
    try {
      setIsCreatingDraft(true);
      const res = await mailAPI.createAttachmentDraft({
        subject: subject || '(No subject)',
        to,
        cc,
      });
      setDraftId(res.draftId);
      return res.draftId;
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to create draft.';
      const isAuthError = /No authentication token available|Unauthorized/i.test(message);
      if (isAuthError) {
        // The user is unauthenticated (or the session has expired). This
        // is a known, expected state when the compose screen is reached
        // before sign-in - log at warn level so the console isn't noisy,
        // but keep it discoverable for developers.
        console.warn('Skipping attachment upload: no active session.', message);
        setToast({
          type: 'error',
          message: 'Please sign in to attach files.',
        });
        // Mark every Selected attachment as Failed with a clear reason so
        // the row UI doesn't look "stuck" in Selected state.
        setAttachments((prev) =>
          prev.map((a) =>
            a.status === AttachmentStatus.Selected
              ? {
                  ...a,
                  status: AttachmentStatus.Failed,
                  error: 'Please sign in to attach files.',
                }
              : a
          )
        );
      } else {
        console.error('Failed to create draft:', err);
        setToast({
          type: 'error',
          message: 'Could not initialize draft for attachments.',
        });
      }
      return null;
    } finally {
      setIsCreatingDraft(false);
    }
  };

  const handleAttachmentsChanged = async (next: EmailAttachment[]) => {
    setAttachments(next);

    // If the user is adding an attachment and we don't yet have a draft,
    // create one on demand. The first selected file triggers the draft.
    const hasNew = next.some((a) => a.status === AttachmentStatus.Selected);
    if (hasNew && !draftId && !isCreatingDraft) {
      await ensureDraft();
    }
  };

  const handleSend = async () => {
    const rawHtmlContent = rteRef.current?.value || '<p><br/></p>';
    const textContent = extractText(rawHtmlContent);
    // Wrap the body so the RTE content styles travel with the email
    // to the recipient's mail client (and to our own Sent Mail view).
    // The wrapper contains a <style> block + the same `e-rte-content`
    // class the in-app reader uses.
    const htmlContent = wrapRteBody(rawHtmlContent);

    // Recipients are required
    if (!to || to.length === 0) {
      setToast({ type: 'error', message: 'Please add at least one recipient.' });
      return;
    }

    // If there are attachments, they must all be uploaded
    const visible = attachments.filter((a) => a.status !== AttachmentStatus.Removed);
    const hasUploading = visible.some(
      (a) =>
        a.status === AttachmentStatus.Selected ||
        a.status === AttachmentStatus.Uploading
    );
    const hasFailed = visible.some((a) => a.status === AttachmentStatus.Failed);
    if (hasUploading) {
      setToast({ type: 'error', message: 'Please wait for all attachments to finish uploading.' });
      return;
    }
    if (hasFailed) {
      setToast({ type: 'error', message: 'Please retry or remove failed attachments before sending.' });
      return;
    }

    const attachmentIds = visible
      .filter((a) => a.status === AttachmentStatus.Uploaded && a.graphAttachmentId)
      .map((a) => a.graphAttachmentId!);

    try {
      if (attachmentIds.length > 0) {
        // Use the attachment-aware send path
        if (!draftId) {
          setToast({ type: 'error', message: 'Draft not initialized. Please try again.' });
          return;
        }
        const sendResponse = await mailAPI.sendMailWithAttachments({
          draftId,
          to,
          subject: subject || '(No subject)',
          body: htmlContent,
          cc,
          attachmentIds,
        });

        // The backend's send endpoint now returns the new sent message
        // id. The local copy MUST use this id (not a synthetic one and
        // not the draft id) so that subsequent fetches
        // (`GET /mail/messages/:id`) can find the message in Graph and
        // return its body + attachments. Without this, the user would
        // see the mail in the list but clicking it would 404.
        const realSentId =
          sendResponse?.sentMessageId ||
          // Fallback: if the backend didn't return one (older builds,
          // network blip, or the lookup query returned no candidates),
          // keep the draft id — the user will see the 404 warning but
          // at least the body/attachments on the draft are still
          // reachable for a moment before Graph fully purges it.
          draftId;

        // The backend successfully sent the mail via Graph. The local
        // Sent Items view doesn't auto-refresh on folder change, so we
        // append the new mail to local state (mirroring what the
        // non-attachment `sendMail` path does). We use the real
        // `sentMessageId` from the response so opening this mail
        // later succeeds.
        //
        // We also include the attachment metadata (name, contentType,
        // size, and the Graph attachment id) so the reader pane can
        // render the attachments immediately. Without this, the reader
        // would skip the `fetchMailDetails` call (because `body` is
        // already set) and the attachments would never appear.
        const sentFolderId = getSystemFolderId('sent') || '2-3';
        const senderEmail = localStorage.getItem('authEmail') || 'me@example.com';
        const sentAttachments: MailAttachment[] = visible
          .filter((a) => a.status === AttachmentStatus.Uploaded && a.graphAttachmentId)
          .map((a) => ({
            id: a.graphAttachmentId!,
            name: a.fileName,
            contentType: a.contentType,
            size: a.size,
          }));
        const newMail: MailData = {
          id: realSentId,
          folderId: sentFolderId,
          senderName: 'Me',
          senderEmail,
          subject: subject || '(No subject)',
          snippet:
            (textContent || '').substring(0, 100) +
            ((textContent || '').length > 100 ? '...' : ''),
          // Use the ISO UTC timestamp so the mail list's date
          // grouping ("Today", "Yesterday", ...) and the
          // `formatMailDate` formatter can parse it correctly.
          // A pre-formatted "3:59 PM" string fails to parse as a
          // Date, which is why newly-sent mails landed under the
          // "Older" group with only a time stamp instead of under
          // "Today" with the proper date+time.
          date: new Date().toISOString(),
          isRead: true,
          isStarred: false,
          isFlagged: false,
          avatarColor: '#0078d4',
          body: htmlContent,
          hasAttachments: sentAttachments.length > 0,
          attachments: sentAttachments,
          to,
          cc,
          bcc: [],
        };
        setMails((prev) => {
          // If a synthetic placeholder was already in the list (e.g.
          // because the user clicked Send twice quickly), replace it
          // by matching on draftId or by the new id.
          const filtered = prev.filter(
            (m) => m.id !== draftId && m.id !== realSentId
          );
          return [newMail, ...filtered];
        });
      } else {
        // Fall back to the regular send path. The store returns the
        // id of the new mail. We don't use it for navigation anymore
        // — the user stays on the current view after sending.
        await sendMail({
          to,
          cc,
          subject: subject || '(No subject)',
          bodyHtml: htmlContent,
          bodyText: textContent,
          senderEmail: localStorage.getItem('authEmail') || undefined,
          senderName: 'Me',
        });
      }

      // After a successful send, close the compose editor and keep the
      // user on whatever folder/mail they were viewing before. We no
      // longer navigate to Sent Items or auto-select the new mail.
      // Clean up registry
      fileRegistry.clear();
      setAttachments([]);
      setDraftId(null);
      clearCompose();
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to send the email.';
      const isAuthError = /No authentication token available|Unauthorized/i.test(message);
      if (isAuthError) {
        // Same handling as the attachment path: this is a known, expected
        // state when the user reaches the compose screen before sign-in.
        // Log at warn level so the console isn't noisy.
        console.warn('Skipping send mail: no active session.', message);
        setToast({ type: 'error', message: 'Please sign in to send email.' });
      } else {
        console.error('Failed to send mail:', error);
        setToast({ type: 'error', message: 'Failed to send the email. Please try again.' });
      }
    }
  };

  const handleSaveDraft = async () => {
    const rawHtmlContent = rteRef.current?.value || '<p><br/></p>';
    const textContent = extractText(rawHtmlContent);
    // Same self-contained envelope used for Send — see handleSend.
    const htmlContent = wrapRteBody(rawHtmlContent);
    const draftsFolderId = getSystemFolderId('drafts') || '2-2';

    // If attachments exist, ensure the user-visible draft is the same one
    // we already created. We just keep it in Drafts.
    const visible = attachments.filter((a) => a.status !== AttachmentStatus.Removed);
    const hasUploading = visible.some(
      (a) =>
        a.status === AttachmentStatus.Selected ||
        a.status === AttachmentStatus.Uploading
    );
    if (hasUploading) {
      setToast({ type: 'error', message: 'Please wait for all attachments to finish uploading.' });
      return;
    }

    try {
      // Call backend API to save draft to Microsoft Graph
      const draftResponse = (await mailAPI.saveDraft({
        to: to.length > 0 ? to : [],
        subject: subject || '(No subject)',
        body: htmlContent,
        cc,
      })) as { draft?: { id?: string } } | undefined;

      // Then update local state with the draft
      setMails(prev => [{
        id: draftResponse?.draft?.id || `draft-${Date.now()}`,
        folderId: draftsFolderId,
        senderName: 'Me',
        senderEmail: localStorage.getItem('authEmail') || 'me@example.com',
        subject: subject || '(No subject)',
        snippet: textContent.substring(0, 100) + (textContent.length > 100 ? '...' : ''),
        // ISO UTC timestamp — see the Send path above for why.
        date: new Date().toISOString(),
        isRead: true,
        isStarred: false,
        isFlagged: false,
        avatarColor: '#0078d4',
        body: htmlContent,
        to,
        cc,
      }, ...prev]);

      // Close the compose editor and keep the user on whatever folder/
      // mail they were viewing before. We no longer navigate to Drafts.
      // Clean up
      fileRegistry.clear();
      setAttachments([]);
      setDraftId(null);
      clearCompose();
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to save draft.';
      const isAuthError = /No authentication token available|Unauthorized/i.test(message);
      if (isAuthError) {
        console.warn('Skipping save draft: no active session.', message);
        setToast({ type: 'error', message: 'Please sign in to save drafts.' });
      } else {
        console.error('Failed to save draft:', error);
        setToast({ type: 'error', message: 'Failed to save draft. Please try again.' });
      }
    }
  };

  const hasContent = () => {
    // Check the four fields the user can fill in:
    //   - To, Cc        → array length
    //   - Subject       → trimmed string length
    //   - RTE body      → strip HTML and check for any non-empty
    //                     text. The previous version string-compared
    //                     the raw HTML against '<p><br/></p>', which
    //                     didn't match other empty states Syncfusion
    //                     emits ('<p><br></p>', '<p>&nbsp;</p>',
    //                     a single '<br/>', etc.) and incorrectly
    //                     classified an empty RTE as having content.
    //                     Stripping HTML to text and trimming is
    //                     robust against any of those variations.
    const htmlContent = rteRef.current?.value || '';
    const rteText = extractText(htmlContent).replace(/\u00A0/g, ' ').trim();
    const hasText =
      subject.trim().length > 0 ||
      to.length > 0 ||
      cc.length > 0 ||
      rteText.length > 0;
    return hasText;
  };

  const handleDiscardClick = () => {
    // The X button in the toolbar is the only "discard" affordance
    // now. If the compose has any content, open the confirmation
    // dialog so the user can choose between Keep Draft and Discard.
    // An empty compose closes immediately with no prompt.
    if (hasContent()) {
      setShowDiscardDialog(true);
    } else {
      clearCompose();
    }
  };

  const handleConfirmDiscard = () => {
    setShowDiscardDialog(false);
    clearCompose();
  };

  const handleKeepDraft = () => {
    setShowDiscardDialog(false);
    handleSaveDraft();
  };

  let title = 'New Message';
  if (composeType === 'reply') title = 'Reply';
  if (composeType === 'replyAll') title = 'Reply All';
  if (composeType === 'forward') title = 'Forward';

  return (
    <div className="bg-white flex flex-col h-full overflow-hidden">
      <div className="px-6 py-4 border-b border-gray-200 flex justify-between items-center gap-4">
        <h2 className="m-0 text-lg font-semibold text-gray-800 shrink-0">{title}</h2>
        <div className="flex items-center gap-2 ml-auto">
          <button
            onClick={handleSend}
            className="bg-blue-600 text-white border-none rounded px-4 py-1.5 text-sm font-semibold cursor-pointer flex items-center gap-2 hover:bg-blue-700 transition-colors"
            aria-label="Send email"
          >
            <Send size={16} />
            Send
          </button>

          <button
            onClick={handleSaveDraft}
            className="bg-amber-50 text-amber-600 border border-amber-200 rounded px-4 py-1.5 text-sm font-semibold cursor-pointer flex items-center gap-2 hover:bg-amber-100 transition-colors"
            title="Save this as draft"
            aria-label="Save draft"
          >
            <FileText size={16} />
            Draft
          </button>

          <button
            onClick={handleDiscardClick}
            className="bg-transparent border-none cursor-pointer flex items-center text-gray-600 hover:text-gray-900 transition-colors p-1"
            title="Discard draft"
            aria-label="Discard draft"
          >
            <X size={20} />
          </button>
        </div>
      </div>

      <div className="px-6 flex flex-col flex-1 overflow-y-auto">
        <div className="border-b border-gray-200 py-3 flex items-center">
          <span className="w-20 text-gray-600 text-sm">To:</span>
          <div className="flex-1 ml-2">
            <MultiSelectComponent 
              id="to-multiselect"
              dataSource={contactsData}
              fields={fields}
              placeholder="Recipients"
              mode="Box"
              itemTemplate={itemTemplate}
              valueTemplate={valueTemplate}
              value={to}
              change={(e: any) => setTo(e.value || [])}
              allowCustomValue={true}
              cssClass="border-none e-multiselect-compose"
            />
          </div>
        </div>
        <div className="border-b border-gray-200 py-3 flex items-center">
          <span className="w-20 text-gray-600 text-sm">Cc:</span>
          <div className="flex-1 ml-2">
            <MultiSelectComponent 
              id="cc-multiselect"
              dataSource={contactsData}
              fields={fields}
              placeholder="Cc"
              mode="Box"
              itemTemplate={itemTemplate}
              valueTemplate={valueTemplate}
              value={cc}
              change={(e: any) => setCc(e.value || [])}
              allowCustomValue={true}
              cssClass="border-none e-multiselect-compose"
            />
          </div>
        </div>
        <div className="border-b border-gray-200 py-3 flex items-center">
          <span className="w-20 text-gray-600 text-sm whitespace-nowrap">Subject:</span>
          <input 
            type="text" 
            placeholder="Subject" 
            className="border-none outline-none w-full text-sm p-1 ml-2 font-semibold"
            value={subject}
            onChange={(e) => setSubject(e.target.value)}
          />
        </div>
        
        <div className="flex-1 mt-4 flex flex-col min-h-[300px]">
          <RichTextEditorComponent
            ref={rteRef}
            height="100%"
            toolbarSettings={rteToolbarSettings}
            cssClass="custom-rte"
          >
            <p><br/></p>
            <Inject services={[Toolbar, Link, Image, HtmlEditor, QuickToolbar]} />
          </RichTextEditorComponent>
        </div>

        {/* Attachment uploader section */}
        <div className="py-3 border-t border-gray-200">
          <AttachmentUploader
            draftId={draftId}
            attachments={attachments}
            onAttachmentsChanged={handleAttachmentsChanged}
            onValidationError={(msg) => setToast({ type: 'error', message: msg })}
            disabled={isCreatingDraft}
          />
        </div>

        {/* Attachment summary badge - shown at the bottom of the compose area */}
        {attachments.filter(
          (a) => a.status === AttachmentStatus.Uploaded
        ).length > 0 && (
          <div
            className="py-3 mt-2 flex items-center justify-end gap-2 text-gray-700 text-sm border-t border-gray-200"
            aria-live="polite"
          >
            <Paperclip size={16} aria-hidden="true" />
            <span>
              {attachments.filter(
                (a) => a.status === AttachmentStatus.Uploaded
              ).length}{' '}
              attached
            </span>
          </div>
        )}
      </div>

      {/* Discard Confirmation Dialog */}
      {showDiscardDialog && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg shadow-lg p-6 w-96">
            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 bg-red-100 rounded-full flex items-center justify-center">
                <Trash2 size={20} className="text-red-600" />
              </div>
              <h2 className="text-lg font-semibold text-gray-800">Discard draft?</h2>
            </div>
            <p className="text-sm text-gray-600 mb-6">
              Are you sure you want to discard this draft? This action cannot be undone.
            </p>
            <div className="flex gap-2 justify-end">
              <button
                onClick={handleKeepDraft}
                className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded cursor-pointer border-none bg-transparent transition-colors font-medium"
              >
                Keep Draft
              </button>
              <button
                onClick={handleConfirmDiscard}
                className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded cursor-pointer border-none transition-colors font-medium flex items-center gap-2"
              >
                <Trash2 size={16} />
                Discard
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Inline validation toast (aria-live so screen readers announce it) */}
      {toast && (
        <div
          role="alert"
          aria-live="assertive"
          className="fixed bottom-6 right-6 z-50 max-w-sm bg-red-50 border border-red-200 text-red-800 text-sm rounded-md shadow-lg px-4 py-3 flex items-start gap-2"
        >
          <X
            size={16}
            className="mt-0.5 cursor-pointer text-red-600"
            onClick={() => setToast(null)}
            aria-label="Dismiss notification"
          />
          <span>{toast.message}</span>
        </div>
      )}
    </div>
  );
};
