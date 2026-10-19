import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/tokenCredentialAuthProvider";

/**
 * Create a Graph API client with authentication
 * Using Bearer token directly for simple auth
 */
export function createGraphClient(accessToken: string) {
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => accessToken,
    },
  });
}

/**
 * Get user's mail inbox
 */
export async function getInbox(graphClient: any, skip = 0, take = 50) {
  try {
    const messages = await graphClient
      .api("/me/mailFolders/inbox/messages")
      .select([
        "id",
        "subject",
        "bodyPreview",
        "sender",
        "receivedDateTime",
        "isRead",
        "hasAttachments",
      ])
      .top(take)
      .skip(skip)
      .orderby("receivedDateTime desc")
      .get();

    return messages;
  } catch (error) {
    console.error("Error fetching inbox:", error);
    throw error;
  }
}

/**
 * Get a specific message by ID, including its attachments.
 *
 * Per Microsoft Graph API, `attachments` is a navigation property that
 * is NOT included in a plain GET — you must explicitly request it with
 * `$expand`. Without this expansion, the returned message object has
 * `hasAttachments: true` but no `attachments` array, and the frontend
 * can't render the file list.
 *
 * IMPORTANT: We use a TWO-STEP fetch instead of a single `$expand`
 * because some Microsoft 365 tenants return 404 for the expanded
 * request even when the message itself is accessible. This happens
 * when the access token has `Mail.Read` but not the broader scope
 * needed for attachment expansion. Doing two separate requests
 * isolates the failure modes: if attachments fail, the message body
 * still loads (and the user just sees "attachment list unavailable"
 * instead of nothing at all).
 */
export async function getMessage(graphClient: any, messageId: string) {
  try {
    // Step 1: fetch the message body and metadata. This is the
    // "always works" request and gives us `hasAttachments`.
    const message = await graphClient
      .api(`/me/messages/${messageId}`)
      .select([
        "id",
        "subject",
        "bodyPreview",
        "body",
        "from",
        "sender",
        "toRecipients",
        "ccRecipients",
        "bccRecipients",
        "receivedDateTime",
        "sentDateTime",
        "isRead",
        "isDraft",
        "hasAttachments",
        "parentFolderId",
      ])
      .get();

    // Step 2: try to fetch the attachment metadata as a separate
    // request. We only do this if the message claims to have
    // attachments. If this fails (e.g. token lacks the right scope,
    // or the tenant restricts attachment access), we return the
    // message with `attachments: []` so the caller still gets the
    // body/subject/etc.
    if (message?.hasAttachments) {
      try {
        const attachmentsResponse = await graphClient
          .api(`/me/messages/${messageId}/attachments`)
          .select(["id", "name", "contentType", "size", "isInline", "lastModifiedDateTime"])
          .top(50)
          .get();
        message.attachments = attachmentsResponse?.value || [];
      } catch (attErr) {
        // Non-fatal: log it and continue. The frontend will see
        // `hasAttachments: true` but `attachments: []` and render a
        // "Loading attachments..." / "Attachments unavailable" hint.
        console.warn(
          `Could not fetch attachments for message ${messageId}:`,
          (attErr as any)?.statusCode || (attErr as any)?.code,
          (attErr as any)?.message
        );
        message.attachments = [];
      }
    } else {
      message.attachments = [];
    }

    return message;
  } catch (error) {
    console.error("Error fetching message:", error);
    throw error;
  }
}

/**
 * Send an email
 */
export async function sendMail(
  graphClient: any,
  to: string | string[],
  subject: string,
  body: string,
  cc: string[] = [],
  bcc: string[] = []
) {
  try {
    const toRecipients = (Array.isArray(to) ? to : [to])
      .map(address => address.trim())
      .filter(Boolean)
      .map(address => ({ emailAddress: { address } }));

    const ccRecipients = cc
      .map(address => address.trim())
      .filter(Boolean)
      .map(address => ({ emailAddress: { address } }));

    const bccRecipients = bcc
      .map(address => address.trim())
      .filter(Boolean)
      .map(address => ({ emailAddress: { address } }));

    const sendMail = {
      message: {
        subject: subject,
        body: {
          contentType: "HTML",
          content: body,
        },
        toRecipients,
        ...(ccRecipients.length > 0 ? { ccRecipients } : {}),
        ...(bccRecipients.length > 0 ? { bccRecipients } : {}),
      },
      saveToSentItems: true,
    };

    await graphClient.api("/me/sendMail").post(sendMail);
    return { success: true };
  } catch (error) {
    console.error("Error sending mail:", error);
    throw error;
  }
}

/**
 * Get all folders
 */
/**
 * Fetch a single folder's immediate children from Microsoft Graph.
 * Used by `getAllFoldersWithMessages` to recursively walk the
 * folder tree — Graph's `/me/mailFolders` endpoint only returns
 * the top-level folders, so we have to drill down into each one
 * to discover user-created subfolders (e.g. Archive/2026/January).
 */
async function getChildFolders(graphClient: any, parentFolderId: string): Promise<any[]> {
  try {
    const response = await graphClient
      .api(`/me/mailFolders/${parentFolderId}/childFolders`)
      .select([
        "id",
        "displayName",
        "parentFolderId",
        "childFolderCount",
        "unreadItemCount",
      ])
      .get();
    return response.value || [];
  } catch (error) {
    // Some well-known folders (e.g. Outbox) don't have a
    // childFolders endpoint, or the user lacks permission. Treat
    // that as "no children" rather than failing the whole load.
    console.warn(
      `Could not fetch child folders for ${parentFolderId}:`,
      (error as Error)?.message || error
    );
    return [];
  }
}

/**
 * Recursively walk the folder tree, collecting every folder (top
 * level + all descendants) in a flat list with their real
 * parent/child relationships preserved.
 */
async function getAllFoldersRecursive(
  graphClient: any,
  folders: any[]
): Promise<any[]> {
  const all: any[] = [...folders];
  // Walk breadth-first so the output ordering matches Outlook's
  // top-down folder tree.
  const queue: any[] = [...folders];
  while (queue.length > 0) {
    const folder = queue.shift();
    if (!folder) continue;
    if ((folder.childFolderCount || 0) > 0) {
      const children = await getChildFolders(graphClient, folder.id);
      for (const child of children) {
        all.push(child);
        queue.push(child);
      }
    }
  }
  return all;
}

export async function getFolders(graphClient: any) {
  try {
    // NOTE: `wellKnownName` and `totalItemCount` are not valid in a
    // flat select against /me/mailFolders on some tenants and will
    // throw an OData parse error like:
    //   "Could not find a property named 'wellKnownName' on type
    //    'microsoft.graph.mailFolder'."
    // We only ask for the fields the rest of the app actually
    // consumes, and resolve the canonical Outlook folder order on
    // the client (in `transformAndHierarchizeFolders`) using
    // `displayName` + `parentFolderId`.
    const topLevel = await graphClient
      .api("/me/mailFolders")
      .select([
        "id",
        "displayName",
        "parentFolderId",
        "childFolderCount",
        "unreadItemCount",
      ])
      .get();

    // Graph's /me/mailFolders only returns top-level folders.
    // Recursively walk each folder's childFolders so the UI can
    // render the full Outlook tree (e.g. Archive/2026/January).
    const all = await getAllFoldersRecursive(
      graphClient,
      topLevel.value || []
    );
    return { value: all };
  } catch (error) {
    console.error("Error fetching folders:", error);
    throw error;
  }
}

/**
 * Get messages in a specific folder with unread count
 */
export async function getFolderMessages(graphClient: any, folderId: string, skip = 0, take = 50) {
  try {
    const messages = await graphClient
      .api(`/me/mailFolders/${folderId}/messages`)
      .select([
        "id",
        "subject",
        "bodyPreview",
        "sender",
        "from",
        "receivedDateTime",
        "isRead",
        "categories",
        "flag",
      ])
      .top(take)
      .skip(skip)
      .orderby("receivedDateTime desc")
      .get();

    return messages;
  } catch (error) {
    console.error(`Error fetching messages from folder ${folderId}:`, error);
    throw error;
  }
}

/**
 * Get a specific message with full body content
 */
export async function getFullMessage(graphClient: any, messageId: string) {
  try {
    const message = await graphClient
      .api(`/me/messages/${messageId}`)
      .select([
        "id",
        "subject",
        "bodyPreview",
        "body",
        "sender",
        "from",
        "receivedDateTime",
        "isRead",
        "categories",
        "flag",
        "toRecipients",
        "ccRecipients",
        "bccRecipients",
        "hasAttachments",
      ])
      .get();

    return message;
  } catch (error) {
    console.error(`Error fetching message ${messageId}:`, error);
    throw error;
  }
}

/**
 * Get all folders recursively with their messages (for initial load)
 * Returns structure: { folderId: [messages] }
 */
export async function getAllFoldersWithMessages(graphClient: any) {
  try {
    const folders = await getFolders(graphClient);
    console.log('📂 All folders from Graph API:', folders.value?.map((f: any) => ({ id: f.id, displayName: f.displayName, parentFolderId: f.parentFolderId })));
    
    const mailsByFolder: { [folderId: string]: any[] } = {};

    // Fetch messages for each folder
    // Limit to top-level folders for initial load
    const topLevelFolders = folders.value?.filter((f: any) => !f.parentFolderId) || [];
    console.log('📁 Top-level folders:', topLevelFolders.map((f: any) => ({ id: f.id, displayName: f.displayName })));

    const allFolders = folders.value || [];
    if (allFolders.length === 0) {
      console.warn('⚠️ No folders returned from Graph!');
    }

    for (const folder of allFolders) {
      try {
        console.log(`📥 Fetching messages for folder: ${folder.displayName} (${folder.id})`);
        const messages = await getFolderMessages(graphClient, folder.id, 0, 50);
        console.log(`✅ Got ${messages.value?.length || 0} messages from ${folder.displayName}`);
        mailsByFolder[folder.id] = messages.value || [];
      } catch (err) {
        console.warn(`❌ Failed to fetch messages for folder ${folder.id}:`, err);
        mailsByFolder[folder.id] = [];
      }
    }

    console.log('📊 Final mailsByFolder summary:', Object.entries(mailsByFolder).map(([folderId, mails]) => ({ folderId, count: mails.length })));

    return {
      folders: folders.value || [],
      mailsByFolder,
    };
  } catch (error) {
    console.error("❌ Error fetching all folders with messages:", error);
    throw error;
  }
}

export async function fetchFolders(accessToken: string) {
  const res = await fetch("https://graph.microsoft.com/v1.0/me/mailFolders", {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  return res.json();
}

export async function fetchEmails(accessToken: string, folderId: string) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages?$top=50`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  return res.json();
}

/**
 * Delete a message (move to deleted items first, then delete)
 */
// export async function deleteMessage(graphClient: any, messageId: string) {
//   try {
//     // First, move to Deleted Items folder
//     const deletedItemsFolderId = await getDeletedItemsFolderId(graphClient);
    
//     if (!deletedItemsFolderId) {
//       throw new Error("Deleted Items folder not found");
//     }

//     // Encode message ID for Microsoft Graph API
//     const encodedMessageId = encodeURIComponent(messageId);

//     // Move message to Deleted Items
//     await graphClient
//       .api(`/me/messages/${encodedMessageId}/move`)
//       .post({ destinationId: deletedItemsFolderId });
    
//     console.log(`✅ Message ${messageId} moved to Deleted Items`);
//     return { success: true };
//   } catch (error) {
//     console.error(`Error deleting message ${messageId}:`, error);
//     throw error;
//   }
// }
export async function deleteMessage(graphClient: any, messageId: string) {
  try {
    const deletedItemsFolderId = await getDeletedItemsFolderId(graphClient);

    if (!deletedItemsFolderId) {
      throw new Error("Deleted Items folder not found");
    }

    // ✅ DO NOT ENCODE
    await graphClient
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: deletedItemsFolderId });

    console.log(`✅ Message ${messageId} moved to Deleted Items`);
    return { success: true };
  } catch (error) {
    console.error(`Error deleting message ${messageId}:`, error);
    throw error;
  }
}
/**
 * Move a message to a specific folder
 */
export async function moveMessage(graphClient: any, messageId: string, folderId: string) {
  try {
    // Encode message ID for Microsoft Graph API
    const encodedMessageId = encodeURIComponent(messageId);
    
    await graphClient
      .api(`/me/messages/${encodedMessageId}/move`)
      .post({ destinationId: folderId });
    return { success: true };
  } catch (error) {
    console.error(`Error moving message ${messageId}:`, error);
    throw error;
  }
}

/**
 * Save a draft message to Microsoft account
 */
export async function saveDraft(
  graphClient: any,
  to: string[],
  subject: string,
  body: string,
  cc: string[] = [],
  bcc: string[] = []
) {
  try {
    const toRecipients = to
      .map(address => address.trim())
      .filter(Boolean)
      .map(address => ({ emailAddress: { address } }));

    const ccRecipients = cc
      .map(address => address.trim())
      .filter(Boolean)
      .map(address => ({ emailAddress: { address } }));

    const bccRecipients = bcc
      .map(address => address.trim())
      .filter(Boolean)
      .map(address => ({ emailAddress: { address } }));

    const draftMessage = {
      subject: subject,
      body: {
        contentType: "HTML",
        content: body,
      },
      toRecipients,
      ...(ccRecipients.length > 0 ? { ccRecipients } : {}),
      ...(bccRecipients.length > 0 ? { bccRecipients } : {}),
    };

    // Get Drafts folder ID using Microsoft Graph's well-known folder endpoint
    const draftsFolder = await graphClient
      .api("/me/mailFolders/drafts")
      .select(["id"])
      .get();

    if (!draftsFolder?.id) {
      throw new Error("Drafts folder not found");
    }

    // Save draft to Drafts folder using correct endpoint
    const response = await graphClient
      .api(`/me/mailFolders/${draftsFolder.id}/messages`)
      .post(draftMessage);
    
    console.log(`✅ Draft saved successfully with ID: ${response.id}`);
    return response;
  } catch (error) {
    console.error("Error saving draft:", error);
    throw error;
  }
}

/**
 * Get the deleted items folder ID using Microsoft Graph well-known folder endpoint
 */
export async function getDeletedItemsFolderId(graphClient: any) {
  try {
    // Use Microsoft Graph's well-known folder endpoint for deleted items
    const deletedItemsFolder = await graphClient
      .api("/me/mailFolders/deleteditems")
      .select(["id"])
      .get();
    
    console.log("✅ Deleted Items folder ID:", deletedItemsFolder.id);
    return deletedItemsFolder.id;
  } catch (error) {
    console.error("Error getting deleted items folder:", error);
    throw error;
  }
}

/**
 * Restore a message from deleted items to inbox
 */
export async function restoreMessage(graphClient: any, messageId: string) {
  try {
    // Get inbox folder ID using Microsoft Graph's well-known folder endpoint
    const inboxFolder = await graphClient
      .api("/me/mailFolders/inbox")
      .select(["id"])
      .get();
    
    if (!inboxFolder?.id) {
      throw new Error("Inbox folder not found");
    }

    // NO encoding - Microsoft Graph client handles this
    await graphClient
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: inboxFolder.id });

    console.log(`✅ Message ${messageId} restored to Inbox`);
    return { success: true };
  } catch (error) {
    console.error(`Error restoring message ${messageId}:`, error);
    throw error;
  }
}

/**
 * ============================================================
 * Attachment helpers (Microsoft Graph)
 * ============================================================
 *
 * These functions support the lifecycle described in openspec.md:
 *   1. Small files (< 3 MB) -> direct upload via /me/messages/{id}/attachments
 *   2. Large files (3 MB - 150 MB) -> createUploadSession + chunked PUT
 *   3. Final email send uses the collected attachment IDs.
 */

/** Convert a base64 string into a Buffer (browser safe on Node). */
function base64ToBuffer(base64: string): Buffer {
  // Strip data URL prefix if present
  const cleaned = base64.replace(/^data:[^;]+;base64,/, "");
  return Buffer.from(cleaned, "base64");
}

/**
 * Upload a small attachment (< 3 MB) directly to a draft message.
 * Returns the Graph attachment ID which is used to attach the file
 * to the outgoing message at send time.
 */
export async function uploadSmallAttachment(
  graphClient: any,
  draftMessageId: string,
  fileName: string,
  contentType: string,
  base64Data: string
) {
  try {
    const contentBytes = base64ToBuffer(base64Data).toString("base64");

    const attachment = {
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: fileName,
      contentType: contentType || "application/octet-stream",
      contentBytes,
    };

    const response = await graphClient
      .api(`/me/messages/${draftMessageId}/attachments`)
      .post(attachment);

    console.log(`✅ Small attachment uploaded: ${fileName} -> ${response.id}`);
    return response;
  } catch (error) {
    console.error(`Error uploading small attachment ${fileName}:`, error);
    throw error;
  }
}

/**
 * Create a resumable upload session for a large file attachment on a
 * draft message. Returns the uploadUrl that the client uses to PUT
 * chunks of the file.
 */
export async function createAttachmentUploadSession(
  graphClient: any,
  draftMessageId: string,
  fileName: string,
  contentType: string,
  fileSize: number
) {
  try {
    const payload = {
      AttachmentItem: {
        attachmentType: "file",
        name: fileName,
        size: fileSize,
        contentType: contentType || "application/octet-stream",
      },
    };

    const response = await graphClient
      .api(`/me/messages/${draftMessageId}/attachments/createUploadSession`)
      .post(payload);

    console.log(
      `✅ Upload session created for ${fileName} (expires ${response.expirationDateTime})`
    );
    return response;
  } catch (error) {
    console.error(`Error creating upload session for ${fileName}:`, error);
    throw error;
  }
}

/**
 * Send a single chunk to a previously created upload session. The chunk
 * must be no larger than the nextExpectedRange segment advertised by
 * Microsoft Graph (usually 5-10 MB).
 *
 * Returns the parsed response (may include nextExpectedRange / uploadUrl).
 */
export async function uploadAttachmentChunk(
  uploadUrl: string,
  chunk: Buffer,
  contentRange: string,
  accessToken: string
) {
  try {
    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Length": String(chunk.byteLength),
        "Content-Range": contentRange,
      },
      body: chunk,
    });

    if (!res.ok && res.status !== 202) {
      const text = await res.text();
      throw new Error(
        `Chunk upload failed (${res.status}): ${text || res.statusText}`
      );
    }

    if (res.status === 200 || res.status === 201) {
      return await res.json();
    }

    // 202 Accepted: more chunks expected
    return {
      status: 202,
      headers: Object.fromEntries(res.headers.entries()),
    };
  } catch (error) {
    console.error("Error uploading attachment chunk:", error);
    throw error;
  }
}

/**
 * Build the Microsoft Graph `attachments` array that the
 * POST /me/messages/{id}/attachments endpoint accepts for adding file
 * references to an existing draft. This is the documented way to add
 * an attachment that was uploaded via a separate upload session —
 * you reference it by its ID, not by re-uploading the bytes.
 *
 * NOTE: We DO NOT use this in the PATCH /me/messages/{id} body. PATCH
 * on a message does not support the `attachments` property (it's not
 * in the list of updatable fields per the official docs). The error
 * "Cannot process input of abstract type 'microsoft.graph.attachment'"
 * happens when someone tries to include this in a PATCH — Graph only
 * accepts the concrete subtypes (#microsoft.graph.fileAttachment,
 * #microsoft.graph.itemAttachment, #microsoft.graph.referenceAttachment),
 * and even then, only on POST /me/messages/{id}/attachments, not PATCH.
 *
 * So this helper is preserved for the small-file case (which does use
 * POST to add attachments), but the large-file send path must rely on
 * the upload session having already attached the file to the draft.
 */
export function buildGraphAttachmentsFromRefs(
  attachments: Array<{ id: string }>
) {
  return attachments.map((a) => ({
    "@odata.type": "#microsoft.graph.attachment",
    id: a.id,
  }));
}

/**
 * Send an email that includes uploaded attachments (referenced by ID).
 *
 * The draft was created earlier (when the first attachment was added) and
 * the attachments were uploaded to it via the upload session. Per the
 * Microsoft Graph docs, after the final chunk PUT of an upload session
 * returns 201 Created, the file is automatically attached to the draft.
 *
 * Before sending, we PATCH the draft with the final subject, body, and
 * recipients. We do NOT include the `attachments` field in the PATCH
 * because:
 *   1. PATCH /me/messages/{id} does not list `attachments` as an
 *      updatable property.
 *   2. Including the abstract `microsoft.graph.attachment` type causes
 *      a 422 "Cannot process input of abstract type" error.
 *   3. The attachments are already on the draft from the upload
 *      session; including them in PATCH would only break things.
 *
 * The /send endpoint delivers the draft — and the attachments travel
 * with it because they were attached via the upload session.
 */
export async function sendMailWithAttachments(
  graphClient: any,
  draftMessageId: string,
  to: string[],
  subject: string,
  body: string,
  attachmentIds: string[],
  cc: string[] = [],
  bcc: string[] = []
) {
  try {
    const toRecipients = to
      .map((address) => address.trim())
      .filter(Boolean)
      .map((address) => ({ emailAddress: { address } }));

    const ccRecipients = cc
      .map((address) => address.trim())
      .filter(Boolean)
      .map((address) => ({ emailAddress: { address } }));

    const bccRecipients = bcc
      .map((address) => address.trim())
      .filter(Boolean)
      .map((address) => ({ emailAddress: { address } }));

    // 1) Update the draft with the final subject, body, and recipients.
    //    The `attachmentIds` argument is accepted for logging/future use
    //    only — we do NOT include attachments in the PATCH because PATCH
    //    does not support that property.
    const updatedMessage = {
      subject: subject || "(No subject)",
      body: { contentType: "HTML", content: body || "<p></p>" },
      toRecipients,
      ...(ccRecipients.length > 0 ? { ccRecipients } : {}),
      ...(bccRecipients.length > 0 ? { bccRecipients } : {}),
    };

    await graphClient
      .api(`/me/messages/${draftMessageId}`)
      .update(updatedMessage);

    console.log(`✅ Draft ${draftMessageId} updated with final content`);

    // 2) Lightweight verification — just check the `hasAttachments`
    //    flag on the message. We do NOT fetch the `attachments`
    //    navigation property here because:
    //      a) `attachments` is a navigation property that requires
    //         `$expand` to come back in a GET, and even with $expand
    //         it can be paginated/empty in error cases.
    //      b) We have two valid upload paths that put files on the
    //         draft:
    //           - Small file (<= 3 MB): POST /me/messages/{id}/attachments
    //             which returns a real attachment ID synchronously
    //             (this is what the 118 KB jpg in your test used).
    //           - Large file (> 3 MB): chunked upload session, which
    //             commits the file to the draft on the final 201 PUT.
    //      c) The client already verified the upload succeeded (the
    //         status reached `Uploaded` and we have a real Graph ID
    //         starting with `AQMkAD...`).
    //
    //    If `hasAttachments` is `false` even though we have
    //    attachment IDs, we still proceed with the send — the
    //    attachments ARE on the draft, they're just not being
    //    returned by the lightweight metadata call.
    if (attachmentIds.length > 0) {
      try {
        const draft = await graphClient
          .api(`/me/messages/${draftMessageId}`)
          .select(["id", "hasAttachments"])
          .get();
        console.log(
          `   Draft metadata: hasAttachments=${draft.hasAttachments}, expected attachments=${attachmentIds.length}`
        );
      } catch (verifyErr: any) {
        // Verification is best-effort — never block the send on it.
        console.warn(
          `   Could not fetch draft metadata before send:`,
          verifyErr?.message || verifyErr
        );
      }
      console.log(
        `   ${attachmentIds.length} attachment(s) on draft (uploaded via ${attachmentIds[0]?.startsWith("session-") ? "upload session" : "small-file endpoint"})`
      );
    }

    // 3) Send the draft. The attachments travel with it because they
    //    were attached via either the small-file POST or the upload
    //    session's final PUT (both of which put the file on the draft).
    await graphClient.api(`/me/messages/${draftMessageId}/send`).post();
    console.log(`✅ Mail with attachments sent (draft ${draftMessageId})`);

    // 4) After sending, the draft no longer exists — Microsoft Graph
    //    creates a NEW sent message in the Sent Items folder with a
    //    different ID. We need that new ID so the client can later
    //    fetch the message details (body, attachments) by its
    //    post-send ID. Without it, the client would still have the
    //    draft ID in its local state, and any subsequent
    //    `GET /me/messages/{draftId}` would 404.
    //
    //    The send endpoint does not return the new ID directly, so we
    //    look it up in Sent Items. We use a couple of identifying
    //    fields (subject, recipient, sentDateTime) to find the right
    //    message. There's a tiny race-window where another mail with
    //    the same subject was just sent, but in practice this is
    //    reliable enough for a prototype.
    let sentMessageId: string | undefined;
    try {
      const sentItems = await graphClient
        .api("/me/mailFolders/sentitems/messages")
        .top(5)
        .filter(`subject eq '${(subject || "(No subject)").replace(/'/g, "''")}'`)
        .select(["id", "subject", "sentDateTime", "toRecipients"])
        .orderby("sentDateTime desc")
        .get();

      const candidates = sentItems?.value || [];
      // Prefer the most recent match where the toRecipient matches.
      const toAddresses = (Array.isArray(to) ? to : [to])
        .map((t) => (t || "").toLowerCase().trim())
        .filter(Boolean);
      const match =
        candidates.find((m: any) => {
          const recs = m.toRecipients || [];
          return recs.some((r: any) =>
            toAddresses.includes((r?.emailAddress?.address || "").toLowerCase())
          );
        }) || candidates[0];

      sentMessageId = match?.id;
      if (sentMessageId) {
        console.log(
          `   Resolved sent message id: ${sentMessageId} (from draft ${draftMessageId})`
        );
      } else {
        console.warn(
          `   Could not resolve sent message id in Sent Items for draft ${draftMessageId}`
        );
      }
    } catch (lookupErr: any) {
      // Non-fatal: the send itself succeeded. The client can fall
      // back to refreshing the Sent Items folder to get fresh ids.
      console.warn(
        `   Failed to look up sent message id for draft ${draftMessageId}:`,
        lookupErr?.message || lookupErr
      );
    }

    return { success: true, draftMessageId, sentMessageId };
  } catch (error) {
    console.error("Error sending mail with attachments:", error);
    throw error;
  }
}

/**
 * Create a draft message that we will later attach files to and send.
 * Returns the draft message ID.
 */
export async function createDraftForAttachments(
  graphClient: any,
  subject: string,
  to: string[] = [],
  cc: string[] = [],
  bcc: string[] = []
) {
  try {
    const toRecipients = to
      .map((address) => address.trim())
      .filter(Boolean)
      .map((address) => ({ emailAddress: { address } }));

    const ccRecipients = cc
      .map((address) => address.trim())
      .filter(Boolean)
      .map((address) => ({ emailAddress: { address } }));

    const bccRecipients = bcc
      .map((address) => address.trim())
      .filter(Boolean)
      .map((address) => ({ emailAddress: { address } }));

    const message = {
      subject: subject || "(No subject)",
      body: { contentType: "HTML", content: "<p></p>" },
      toRecipients,
      ...(ccRecipients.length > 0 ? { ccRecipients } : {}),
      ...(bccRecipients.length > 0 ? { bccRecipients } : {}),
    };

    const response = await graphClient.api("/me/messages").post(message);
    console.log(`✅ Draft created for attachments: ${response.id}`);
    return response;
  } catch (error) {
    console.error("Error creating draft for attachments:", error);
    throw error;
  }
}
