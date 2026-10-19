// filepath: my-app-backend/src/routes/attachments.ts
// Routes for the email attachment lifecycle (create draft, upload small/large
// files, send the composed email). All endpoints require a valid JWT.

import { FastifyInstance, FastifyRequest, FastifyReply } from "fastify";
import { db } from "../db/index.js";
import { accounts } from "../db/schema.js";
import { eq } from "drizzle-orm";
import { decrypt } from "../utils/encryption.js";
import {
  createDraftForAttachments,
  createAttachmentUploadSession,
  uploadSmallAttachment,
  sendMailWithAttachments,
} from "../services/graphService.js";

/** Resolve the user's Microsoft Graph access token from the DB. */
async function getUserAccessToken(request: FastifyRequest) {
  const decoded = request.user as any;
  const userId = decoded?.sub;
  if (!userId) throw new Error("No user ID in JWT token");

  const account = await db.query.accounts.findFirst({
    where: eq(accounts.userId, userId),
  });
  if (!account || !account.accessToken) {
    throw new Error("Account not found in database");
  }
  return decrypt(account.accessToken);
}

// Lazy import to avoid circular deps and to keep graphClient factory in one place
async function getGraphClient(accessToken: string) {
  const { Client } = await import("@microsoft/microsoft-graph-client");
  return Client.initWithMiddleware({
    authProvider: { getAccessToken: async () => accessToken },
  });
}

const SMALL_FILE_LIMIT = 3 * 1024 * 1024; // 3 MB
const MAX_FILE_SIZE = 150 * 1024 * 1024; // 150 MB

export async function attachmentRoutes(fastify: FastifyInstance) {
  /**
   * POST /attachments/draft
   * Create an empty draft message. The draft will receive all attachments
   * and is the message that gets sent.
   * Body: { subject, to, cc, bcc }
   */
  fastify.post(
    "/attachments/draft",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { subject, to = [], cc = [], bcc = [] } = request.body as {
          subject?: string;
          to?: string[];
          cc?: string[];
          bcc?: string[];
        };

        const accessToken = await getUserAccessToken(request);
        const graphClient = await getGraphClient(accessToken);

        const draft = await createDraftForAttachments(
          graphClient,
          subject || "(No subject)",
          to,
          cc,
          bcc
        );

        reply.send({ success: true, draftId: draft.id });
      } catch (error: any) {
        console.error("Create draft error:", error);
        reply.code(500).send({ error: error.message || "Failed to create draft" });
      }
    }
  );

  /**
   * POST /attachments/small
   * Direct upload of a small file (< 3 MB) to a draft message.
   * Body: { draftId, fileName, contentType, base64 }
   */
  fastify.post(
    "/attachments/small",
    {
      onRequest: [fastify.authenticate],
      bodyLimit: 4 * 1024 * 1024, // 4 MB - small attachments only
    },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { draftId, fileName, contentType, base64 } = request.body as {
          draftId: string;
          fileName: string;
          contentType: string;
          base64: string;
        };

        if (!draftId || !fileName || !base64) {
          return reply.code(400).send({ error: "Missing required fields" });
        }

        // Approximate size from base64 (4 chars per 3 bytes)
        const size = Math.floor((base64.length * 3) / 4);
        if (size > SMALL_FILE_LIMIT) {
          return reply.code(413).send({
            error: "File too large for direct upload. Use upload-session instead.",
          });
        }
        if (size > MAX_FILE_SIZE) {
          return reply.code(413).send({ error: "File exceeds 150 MB limit." });
        }

        const accessToken = await getUserAccessToken(request);
        const graphClient = await getGraphClient(accessToken);

        const attachment = await uploadSmallAttachment(
          graphClient,
          draftId,
          fileName,
          contentType || "application/octet-stream",
          base64
        );

        reply.send({
          success: true,
          attachmentId: attachment.id,
          size: attachment.size,
        });
      } catch (error: any) {
        console.error("Small attachment upload error:", error);
        reply.code(500).send({ error: error.message || "Upload failed" });
      }
    }
  );

  /**
   * POST /attachments/upload-session
   * Create a chunked upload session for a large file (3 MB - 150 MB).
   * Returns the Graph upload URL that the client uses to PUT chunks
   * directly to Microsoft Graph (bypassing the backend's body limit).
   * Body: { draftId, fileName, contentType, fileSize }
   */
  fastify.post(
    "/attachments/upload-session",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { draftId, fileName, contentType, fileSize } = request.body as {
          draftId: string;
          fileName: string;
          contentType: string;
          fileSize: number;
        };

        if (!draftId || !fileName || !fileSize) {
          return reply.code(400).send({ error: "Missing required fields" });
        }
        if (fileSize > MAX_FILE_SIZE) {
          return reply.code(413).send({ error: "File exceeds 150 MB limit." });
        }
        if (fileSize <= SMALL_FILE_LIMIT) {
          return reply.code(400).send({
            error: "File is small enough for direct upload. Use /attachments/small.",
          });
        }

        const accessToken = await getUserAccessToken(request);
        const graphClient = await getGraphClient(accessToken);

        const session = await createAttachmentUploadSession(
          graphClient,
          draftId,
          fileName,
          contentType || "application/octet-stream",
          fileSize
        );

        reply.send({
          success: true,
          // The uploadUrl is pre-authenticated by Microsoft Graph.
          // The client uses it to PUT chunks directly to Graph
          // without needing an Authorization header.
          uploadUrl: session.uploadUrl,
          expirationDateTime: session.expirationDateTime,
        });
      } catch (error: any) {
        console.error("Upload session error:", error);
        reply.code(500).send({ error: error.message || "Failed to create upload session" });
      }
    }
  );

  /**
   * Chunk uploads for large files happen DIRECTLY from the client to
   * Microsoft Graph using the uploadUrl returned by
   * POST /attachments/upload-session. The client receives the Graph
   * access token in that response and uses it to PUT chunks to Graph.
   *
   * This avoids the backend's body size limit and is the recommended
   * pattern in the Microsoft Graph documentation for large file
   * attachments. See:
   * https://learn.microsoft.com/en-us/graph/api/attachment-createuploadsession
   */

  /**
   * POST /attachments/send
   * Send a draft message (with all attachments already uploaded).
   * Body: { draftId, to, subject, body, cc, bcc, attachmentIds }
   */
  fastify.post(
    "/attachments/send",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const {
          draftId,
          to,
          subject,
          body,
          cc = [],
          bcc = [],
          attachmentIds = [],
        } = request.body as {
          draftId: string;
          to: string[];
          subject: string;
          body: string;
          cc?: string[];
          bcc?: string[];
          attachmentIds?: string[];
        };

        if (!draftId || !to || !Array.isArray(to) || to.length === 0) {
          return reply.code(400).send({ error: "Missing required fields" });
        }

        console.log(`📤 Sending email with attachments:`);
        console.log(`   Draft ID: ${draftId}`);
        console.log(`   To: ${to.join(", ")}`);
        console.log(`   Subject: ${subject}`);
        console.log(`   Attachment IDs: ${attachmentIds.join(", ") || "(none)"}`);

        const accessToken = await getUserAccessToken(request);
        const graphClient = await getGraphClient(accessToken);

        // Send the draft (its attachments are already attached in Graph).
        // The backend first PATCHes the draft with the final subject,
        // body, and recipients before sending it. We also capture the
        // new sent-message ID (which is different from the draft ID —
        // Graph creates a fresh sent message) and return it so the
        // client can refresh its local view with a valid id.
        const sendResult = await sendMailWithAttachments(
          graphClient,
          draftId,
          to,
          subject || "(No subject)",
          body || "",
          attachmentIds,
          cc,
          bcc
        );

        reply.send({
          success: true,
          message: "Email with attachments sent",
          // Include both ids for client convenience. `draftId` is the
          // id the client already knows; `sentMessageId` is the new
          // id of the message in Sent Items. The client SHOULD use
          // `sentMessageId` going forward; we keep `draftId` for
          // backward compatibility.
          draftId: sendResult.draftMessageId,
          sentMessageId: sendResult.sentMessageId,
        });
      } catch (error: any) {
        console.error("Send with attachments error:", error);
        reply.code(500).send({ error: error.message || "Failed to send" });
      }
    }
  );
}
