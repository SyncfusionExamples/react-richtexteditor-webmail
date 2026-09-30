import { FastifyInstance, FastifyRequest, FastifyReply } from "fastify";
import { 
  createGraphClient, 
  getInbox, 
  getMessage, 
  sendMail, 
  getFolders,
  getFolderMessages,
  getFullMessage,
  getAllFoldersWithMessages,
  deleteMessage,
  moveMessage,
  saveDraft,
  restoreMessage
} from "../services/graphService.js";
import { sessionStore } from "../services/sessionStore.js";
import { db } from "../db/index.js";
import { accounts } from "../db/schema.js";
import { decrypt } from "../utils/encryption.js";
import { eq } from "drizzle-orm";
import {
  transformAndHierarchizeFolders,
  transformGraphMessagesToFrontend,
  transformGraphMessageToFrontend,
  formatInitialDataResponse,
  type FrontendFolder,
  type FrontendMail,
} from "../utils/transformers.js";

/**
 * Verify JWT and get user access token from database
 */
async function getUserAccessToken(request: FastifyRequest) {
  try {
    // Get user ID from JWT claims
    const decoded = request.user as any;
    const userId = decoded?.sub;

    if (!userId) {
      throw new Error("No user ID in JWT token");
    }

    console.log(`🔍 Looking up access token for user: ${userId}`);

    // Query database for user's account (encrypted access token)
    const account = await db.query.accounts.findFirst({
      where: eq(accounts.userId, userId),
    });

    if (!account || !account.accessToken) {
      console.error(`❌ No account found for user: ${userId}`);
      throw new Error("Account not found in database");
    }

    // Decrypt the access token
    const accessToken = decrypt(account.accessToken);
    console.log(`✅ Access token decrypted successfully`);

    return {
      session: null, // No longer using session store
      accessToken,
    };
  } catch (error) {
    console.error("getUserAccessToken error:", error);
    throw new Error("Unauthorized");
  }
}

export async function mailRoutes(fastify: FastifyInstance) {
  // Add a hook to log all incoming requests to /mail routes
  fastify.addHook("preHandler", (request, reply, done) => {
    if (request.url.includes("delete")) {
      console.log(`\n📍 INCOMING REQUEST TO MAIL ROUTES:`);
      console.log(`   Method: ${request.method}`);
      console.log(`   URL: ${request.url}`);
      console.log(`   Path: ${request.url.split('?')[0]}`);
      console.log(`   Full request path: ${request.routerPath}`);
    }
    done();
  });

  /**
   * GET /mail/messages
   * Get inbox messages with pagination
   */
  fastify.get(
    "/messages",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { skip = "0", take = "50" } = request.query as {
          skip?: string;
          take?: string;
        };

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        const messages = await getInbox(
          graphClient,
          parseInt(skip),
          parseInt(take)
        );

        reply.send({
          messages: messages.value || [],
          total: messages["@odata.count"] || 0,
        });
      } catch (error: any) {
        console.error("Get messages error:", error);
        reply.code(500).send({
          error: error.message || "Failed to fetch messages",
        });
      }
    }
  );

  /**
   * GET /mail/messages/:id
   * Get a specific message, including its attachments. The
   * `getMessage` helper in graphService uses `$expand=attachments`
   * so the response includes the full attachment list (id, name,
   * contentType, size, isInline) without the (potentially huge)
   * contentBytes payload. The client fetches attachment content
   * on demand via the endpoint below.
   */
  fastify.get(
    "/messages/:id",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      const { id } = request.params as { id: string };
      // Diagnostic: log the ID we're about to fetch so we can spot
      // cases where the local list contains a synthetic ID (e.g. a
      // newly-sent mail created with `m${Date.now()}`) that doesn't
      // exist in Graph yet.
      console.log(`📨 GET /mail/messages/:id (id length=${id.length}, starts with=${id.slice(0, 8)})`);

      try {
        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        const message = await getMessage(graphClient, id);

        reply.send(message);
      } catch (error: any) {
        console.error(
          `Get message error for id ${id.slice(0, 16)}…:`,
          error?.statusCode || error?.code,
          error?.message
        );

        // Pass through Microsoft Graph's status code (404, 401, 403,
        // etc.) instead of always sending 500. This lets the client
        // distinguish a real "not found" from a server-side bug and
        // react appropriately (e.g. show a "mail no longer available"
        // message instead of a generic error).
        const status = error?.statusCode || error?.status;
        if (status && typeof status === 'number') {
          reply.code(status).send({
            error: error.message || "Failed to fetch message",
          });
          return;
        }

        reply.code(500).send({
          error: error.message || "Failed to fetch message",
        });
      }
    }
  );

  /**
   * GET /mail/messages/:id/attachments/:attachmentId
   * Download the raw bytes of a specific attachment. The client
   * calls this when the user clicks the "Download" button on an
   * attachment row in the mail reader.
   *
   * Returns the attachment bytes with the original content type so
   * the browser can offer a "Save as" dialog.
   *
   * IMPORTANT: We fetch the raw bytes using a DIRECT `fetch()` call
   * (NOT the Graph SDK) because:
   *   1) The Graph SDK's `.get()` returns an opaque object for binary
   *      responses on newer SDK versions — Fastify's `reply.send()`
   *      rejects it with "Attempted to send payload of invalid type
   *      'object'".
   *   2) A direct fetch gives us a standard `Response` object whose
   *      `.arrayBuffer()` we can convert to a Node `Buffer` and pipe
   *      through to the client.
   *   3) We avoid loading the SDK's JSON parser on binary data.
   *
   * We DO use the Graph SDK for the metadata fetch (step 1) since
   * that returns JSON.
   */
  fastify.get(
    "/messages/:id/attachments/:attachmentId",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { id, attachmentId } = request.params as {
          id: string;
          attachmentId: string;
        };

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        // Step 1: Fetch attachment metadata (name, contentType, size)
        // via the Graph SDK. This is a JSON response so the SDK works.
        const attachmentMeta = await graphClient
          .api(`/me/messages/${id}/attachments/${attachmentId}`)
          .select(["id", "name", "contentType", "size"])
          .get();

        const contentType =
          attachmentMeta?.contentType || "application/octet-stream";
        let fileName = attachmentMeta?.name || "attachment";

        // Sanitize the filename so it's safe to embed in an HTTP
        // header. The `Content-Disposition` header is strict about
        // which characters are allowed — anything outside ASCII,
        // plus control characters like \r and \n, will trigger
        // Fastify's `ERR_INVALID_CHAR` error (HTTP 500).
        const asciiFallback = fileName
          .replace(/[^\x20-\x7E]/g, "_")
          .replace(/["\\]/g, "")
          .replace(/\s+/g, " ")
          .trim() || "attachment";

        const encodedName = encodeURIComponent(fileName);
        reply.header("Content-Type", contentType);
        reply.header(
          "Content-Disposition",
          `attachment; filename="${asciiFallback}"; filename*=UTF-8''${encodedName}`
        );

        // Step 2: Fetch the RAW attachment bytes via a direct fetch()
        // to the Graph `/$value` endpoint. We bypass the Graph SDK
        // here because newer SDK versions return an opaque object for
        // binary responses that Fastify can't `send()` directly.
        //
        // We intentionally do NOT set a `Content-Length` header here
        // because we're going to pipe the entire body through; the
        // reply.send(buffer) below will set it automatically.
        const valueUrl = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(
          id
        )}/attachments/${encodeURIComponent(attachmentId)}/$value`;

        const rawResponse = await fetch(valueUrl, {
          method: "GET",
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        });

        if (!rawResponse.ok) {
          const errText = await rawResponse.text().catch(() => "");
          console.error(
            `Graph $value fetch failed (${rawResponse.status}): ${errText}`
          );
          reply.code(rawResponse.status || 500).send({
            error: `Failed to download attachment: ${
              errText || rawResponse.statusText
            }`,
          });
          return;
        }

        // Convert the response body to a Node Buffer so Fastify can
        // send it as the response body. For large attachments, this
        // loads the whole file into server memory — acceptable for
        // the 150 MB Graph attachment cap, though a streaming
        // approach would be better for production at truly large
        // sizes.
        const arrayBuffer = await rawResponse.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);
        reply.send(buffer);
      } catch (error: any) {
        console.error("Download attachment error:", error);
        reply.code(500).send({
          error: error.message || "Failed to download attachment",
        });
      }
    }
  );

  /**
   * POST /mail/send
   * Send an email
   */
  fastify.post(
    "/send",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { to, subject, body, cc = [], bcc = [] } = request.body as {
          to: string | string[];
          subject: string;
          body: string;
          cc?: string[];
          bcc?: string[];
        };

        const recipients = Array.isArray(to) ? to : [to];

        if (!recipients.length || !subject || !body) {
          return reply.code(400).send({
            error: "Missing required fields: to, subject, body",
          });
        }

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        await sendMail(graphClient, recipients, subject, body, cc, bcc);

        reply.send({
          success: true,
          message: "Email sent successfully",
        });
      } catch (error: any) {
        console.error("Send mail error:", error);
        reply.code(500).send({
          error: error.message || "Failed to send email",
        });
      }
    }
  );

  /**
   * GET /mail/folders
   * Get all mail folders (transformed for frontend)
   */
  fastify.get(
    "/folders",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        const graphFolders = await getFolders(graphClient);
        const transformedFolders = transformAndHierarchizeFolders(graphFolders.value || []);

        reply.send({
          folders: transformedFolders,
        });
      } catch (error: any) {
        console.error("Get folders error:", error);
        reply.code(500).send({
          error: error.message || "Failed to fetch folders",
        });
      }
    }
  );

  /**
   * GET /mail/folders/:folderId/messages
   * Get messages in a specific folder
   */
  fastify.get(
    "/folders/:folderId/messages",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { folderId } = request.params as { folderId: string };
        const { skip = "0", take = "50" } = request.query as {
          skip?: string;
          take?: string;
        };

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        const messages = await getFolderMessages(
          graphClient,
          folderId,
          parseInt(skip),
          parseInt(take)
        );

        const transformedMails = transformGraphMessagesToFrontend(
          messages.value || [],
          folderId
        );

        reply.send({
          mails: transformedMails,
          total: messages["@odata.count"] || 0,
        });
      } catch (error: any) {
        console.error("Get folder messages error:", error);
        reply.code(500).send({
          error: error.message || "Failed to fetch folder messages",
        });
      }
    }
  );

  /**
   * GET /mail/messages/:id/full
   * Get full message content with body
   */
  fastify.get(
    "/messages/:id/full",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { id } = request.params as { id: string };

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        const message = await getFullMessage(graphClient, id);
        // Note: This returns raw Graph message, frontend can decide how to use body
        reply.send(message);
      } catch (error: any) {
        console.error("Get full message error:", error);
        reply.code(500).send({
          error: error.message || "Failed to fetch message",
        });
      }
    }
  );

  /**
   * GET /mail/initial-data
   * Get all folders and initial batch of messages (for initial load after login)
   * This is the main endpoint used after user logs in
   */
  fastify.get(
    "/initial-data",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        console.log('📡 /mail/initial-data endpoint called');
        
        const { accessToken } = await getUserAccessToken(request);
        console.log('✅ Access token retrieved from session');

        const graphClient = createGraphClient(accessToken);
        console.log('✅ Graph client created');

        console.log('🔄 Fetching folders and messages from Microsoft Graph...');
        const { folders: graphFolders, mailsByFolder: rawMailsByFolder } =
          await getAllFoldersWithMessages(graphClient);
        const transformedFolders = transformAndHierarchizeFolders(graphFolders || []);

        const mailsByFolder: { [folderId: string]: FrontendMail[] } = {};
        for (const [folderId, messages] of Object.entries(rawMailsByFolder)) {
          mailsByFolder[folderId] = transformGraphMessagesToFrontend(
            (messages as any[]) || [],
            folderId
          );
        }

        const response = formatInitialDataResponse(transformedFolders, mailsByFolder);
        console.log('📤 Sending response with', response.folders.length, 'folders and', response.totalMailCount, 'total mails');
        reply.send(response);
      } catch (error: any) {
        console.error("❌ Get initial data error:", error);
        reply.code(500).send({
          error: error.message || "Failed to fetch initial data",
        });
      }
    }
  );

  /**
   * POST /mail/messages/delete
   * Delete a message permanently
   * Body: { messageId: "..." }
   */
  fastify.post(
    "/messages/delete",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { messageId } = request.body as { messageId: string };
        
        if (!messageId) {
          return reply.code(400).send({
            error: "Missing required field: messageId",
          });
        }

        console.log(`🔍 DELETE endpoint reached for message: ${messageId}`);

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        console.log(`🔄 Calling deleteMessage() for ID: ${messageId}`);
        await deleteMessage(graphClient, messageId);

        console.log(`✅ Message ${messageId} deleted successfully`);
        reply.send({
          success: true,
          message: "Message deleted successfully",
        });
      } catch (error: any) {
        console.error("❌ Delete message error:", error);
        reply.code(500).send({
          error: error.message || "Failed to delete message",
        });
      }
    }
  );

  /**
   * POST /mail/messages/move
   * Move a message to a specific folder
   * Body: { messageId: "...", folderId: "..." }
   */
  fastify.post(
    "/messages/move",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { messageId, folderId } = request.body as { messageId: string; folderId: string };

        if (!messageId) {
          return reply.code(400).send({
            error: "Missing required field: messageId",
          });
        }

        if (!folderId) {
          return reply.code(400).send({
            error: "Missing required field: folderId",
          });
        }

        console.log(`🔍 MOVE endpoint reached for message: ${messageId} to folder: ${folderId}`);

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        console.log(`🔄 Calling moveMessage() for ID: ${messageId} to folder: ${folderId}`);
        await moveMessage(graphClient, messageId, folderId);

        console.log(`✅ Message ${messageId} moved successfully`);
        reply.send({
          success: true,
          message: "Message moved successfully",
        });
      } catch (error: any) {
        console.error("❌ Move message error:", error);
        reply.code(500).send({
          error: error.message || "Failed to move message",
        });
      }
    }
  );

  /**
   * POST /mail/messages/restore
   * Restore a message from deleted items to inbox
   * Body: { messageId: "..." }
   */
  fastify.post(
    "/messages/restore",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { messageId } = request.body as { messageId: string };
        
        if (!messageId) {
          return reply.code(400).send({
            error: "Missing required field: messageId",
          });
        }

        console.log(`🔍 RESTORE endpoint reached for message: ${messageId}`);

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        console.log(`🔄 Calling restoreMessage() for ID: ${messageId}`);
        await restoreMessage(graphClient, messageId);

        console.log(`✅ Message ${messageId} restored successfully`);
        reply.send({
          success: true,
          message: "Message restored successfully",
        });
      } catch (error: any) {
        console.error("❌ Restore message error:", error);
        reply.code(500).send({
          error: error.message || "Failed to restore message",
        });
      }
    }
  );

  /**
   * POST /mail/drafts/save
   * Save a draft message
   */
  fastify.post(
    "/drafts/save",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { to, subject, body, cc = [], bcc = [] } = request.body as {
          to: string[];
          subject: string;
          body: string;
          cc?: string[];
          bcc?: string[];
        };

        if (!to || !Array.isArray(to) || to.length === 0 || !subject || !body) {
          return reply.code(400).send({
            error: "Missing required fields: to (array), subject, body",
          });
        }

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        const draft = await saveDraft(graphClient, to, subject, body, cc, bcc);

        reply.send({
          success: true,
          message: "Draft saved successfully",
          draft,
        });
      } catch (error: any) {
        console.error("Save draft error:", error);
        reply.code(500).send({
          error: error.message || "Failed to save draft",
        });
      }
    }
  );
}
