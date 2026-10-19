import { FastifyInstance, FastifyRequest, FastifyReply } from "fastify";
import { 
  createGraphClient, 
  getInbox, 
  getMessage, 
  sendMail, 
  getFolders,
  getFolderMessages,
  getFullMessage,
  getAllFoldersWithMessages
} from "../services/graphService.js";
import { sessionStore } from "../services/sessionStore.js";
import {
  transformAndHierarchizeFolders,
  transformGraphMessagesToFrontend,
  transformGraphMessageToFrontend,
  formatInitialDataResponse,
  type FrontendFolder,
  type FrontendMail,
} from "../utils/transformers.js";

/**
 * Verify JWT and get user access token from session store
 */
async function getUserAccessToken(request: FastifyRequest) {
  try {
    // Get user email from JWT claims
    const decoded = request.user as any;
    const email = decoded?.email;

    if (!email) {
      throw new Error("No email in JWT token");
    }

    // Get session from in-memory session store
    const session = sessionStore.getSession(email);
    if (!session) {
      throw new Error("Session not found or expired");
    }

    return {
      session,
      accessToken: session.accessToken,
    };
  } catch (error) {
    console.error("getUserAccessToken error:", error);
    throw new Error("Unauthorized");
  }
}

export async function mailRoutes(fastify: FastifyInstance) {
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
   * Get a specific message
   */
  fastify.get(
    "/messages/:id",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { id } = request.params as { id: string };

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        const message = await getMessage(graphClient, id);

        reply.send(message);
      } catch (error: any) {
        console.error("Get message error:", error);
        reply.code(500).send({
          error: error.message || "Failed to fetch message",
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
        const { to, subject, body } = request.body as {
          to: string;
          subject: string;
          body: string;
        };

        if (!to || !subject || !body) {
          return reply.code(400).send({
            error: "Missing required fields: to, subject, body",
          });
        }

        const { accessToken } = await getUserAccessToken(request);
        const graphClient = createGraphClient(accessToken);

        await sendMail(graphClient, to, subject, body);

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

        // Fetch all folders and their messages
        console.log('🔄 Fetching folders and messages from Microsoft Graph...');
        const { folders: graphFolders, mailsByFolder: rawMailsByFolder } =
          await getAllFoldersWithMessages(graphClient);
        console.log('📁 Retrieved folders:', graphFolders.length, 'folders');
        console.log('📧 Retrieved messages:', Object.keys(rawMailsByFolder).length, 'folders with messages');

        // Transform folders
        const transformedFolders = transformAndHierarchizeFolders(graphFolders);
        console.log('✅ Transformed folders:', transformedFolders.length);

        // Transform messages and organize by folder
        const mailsByFolder: { [folderId: string]: FrontendMail[] } = {};

        for (const [folderId, messages] of Object.entries(rawMailsByFolder)) {
          mailsByFolder[folderId] = transformGraphMessagesToFrontend(
            (messages as any[]) || [],
            folderId
          );
        }
        console.log('✅ Transformed messages in all folders');

        // Format response
        const response = formatInitialDataResponse(transformedFolders, mailsByFolder);
        console.log('📤 Sending response with', response.folders.length, 'folders and', response.totalMailCount, 'total mails');
        
        // Detailed logging of structure
        console.log('📋 FOLDERS in response:', JSON.stringify(response.folders.map(f => ({ id: f.id.substring(0, 40), name: f.name })), null, 2));
        
        console.log('📧 MAILS by folder:');
        for (const [folderId, mails] of Object.entries(response.mailsByFolder)) {
          if (mails.length > 0) {
            console.log(`   - ${folderId.substring(0, 40)}: ${mails.length} mails`);
            console.log(`     Sample mail - id: ${mails[0].id.substring(0, 20)}, folderId: ${mails[0].folderId.substring(0, 40)}`);
          }
        }
        
        // Verify structure integrity
        let mismatchCount = 0;
        for (const [folderId, mails] of Object.entries(response.mailsByFolder)) {
          for (const mail of mails) {
            if (mail.folderId !== folderId) {
              mismatchCount++;
              console.warn(`⚠️ MISMATCH: folder key=${folderId.substring(0, 20)}, mail.folderId=${mail.folderId.substring(0, 20)}`);
            }
          }
        }
        if (mismatchCount === 0) {
          console.log('✅ All mails have correct folderId');
        } else {
          console.warn(`❌ Found ${mismatchCount} folder ID mismatches`);
        }

        reply.send(response);
      } catch (error: any) {
        console.error("❌ Get initial data error:", error);
        reply.code(500).send({
          error: error.message || "Failed to fetch initial data",
        });
      }
    }
  );
}
