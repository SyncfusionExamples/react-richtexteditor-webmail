import { FastifyInstance, FastifyRequest, FastifyReply } from "fastify";
import { db } from "../db/index.js";
import { subscriptions, accounts, emails } from "../db/schema.js";
import { desc, eq } from "drizzle-orm";
import { runDeltaSync } from "../services/syncServices.js";
import { createSubscription } from "../services/subscriptionService.js";
import { decrypt } from "../utils/encryption.js";
import { getActiveConnectionCount, getActiveUsers, triggerNewEmailEvent } from "../services/websocketService.js";

export async function debugRoutes(fastify: FastifyInstance) {
  // List all subscriptions
  fastify.get("/debug/subscriptions", async (request: FastifyRequest, reply: FastifyReply) => {
    try {
      const rows = await db.query.subscriptions.findMany();
      reply.send(rows);
    } catch (err) {
      console.error("Error fetching subscriptions:", err);
      reply.code(500).send({ error: "Failed to fetch subscriptions" });
    }
  });

  // Create a Graph subscription for an existing user
  fastify.post("/debug/create-subscription", async (request: FastifyRequest, reply: FastifyReply) => {
    try {
      const { userId } = request.body as { userId?: string };
      if (!userId) return reply.code(400).send({ error: "userId required" });

      const account = await db.query.accounts.findFirst({
        where: eq(accounts.userId, userId),
      });
      if (!account || !account.accessToken) {
        return reply.code(404).send({ error: "Account not found for user" });
      }

      const accessToken = decrypt(account.accessToken);
      const subscription = await createSubscription(userId, accessToken);

      if (!subscription) {
        return reply.code(500).send({ error: "Failed to create subscription" });
      }

      reply.send({ success: true, subscription });
    } catch (err) {
      console.error("debug create-subscription error:", err);
      reply.code(500).send({ error: "Failed to create subscription" });
    }
  });

  // Trigger a delta sync for a given userId (body: { userId: "..." })
  fastify.post("/debug/trigger-sync", async (request: FastifyRequest, reply: FastifyReply) => {
    try {
      const { userId } = request.body as { userId?: string };
      if (!userId) return reply.code(400).send({ error: "userId required" });

      const account = await db.query.accounts.findFirst({ where: eq(accounts.userId, userId) });
      if (!account) return reply.code(404).send({ error: "Account not found for user" });

      const accessToken = decrypt(account.accessToken);
      await runDeltaSync(userId, accessToken);

      reply.send({ success: true });
    } catch (err) {
      console.error("trigger-sync error:", err);
      reply.code(500).send({ error: "Failed to trigger sync" });
    }
  });

  // Debug: Check active WebSocket connections
  fastify.get("/debug/ws-status", async (request: FastifyRequest, reply: FastifyReply) => {
    const totalConnections = getActiveConnectionCount();
    const activeUsers = getActiveUsers();
    reply.send({
      totalConnections,
      activeUsers,
      message: totalConnections === 0
        ? "No WebSocket connections. Frontend may not be connected."
        : `${totalConnections} connection(s) active for ${activeUsers.length} user(s).`,
    });
  });

  // Debug: Manually trigger a new_email event for a user
  fastify.post("/debug/trigger-event", async (request: FastifyRequest, reply: FastifyReply) => {
    const { userId } = request.body as { userId?: string };
    if (!userId) return reply.code(400).send({ error: "userId required" });
    const sent = triggerNewEmailEvent(userId);
    reply.send({
      success: sent,
      message: sent
        ? `Event sent to user ${userId}`
        : `No active WebSocket connections for user ${userId}`,
    });
  });

  // Debug: List the latest emails persisted in the DB for a user
  fastify.get("/debug/emails", async (request: FastifyRequest, reply: FastifyReply) => {
    try {
      const { userId, limit } = request.query as { userId?: string; limit?: string };
      if (!userId) return reply.code(400).send({ error: "userId query param required" });
      const rowLimit = Math.min(Math.max(parseInt(limit || "20", 10) || 20, 1), 100);
      const rows = await db
        .select({
          id: emails.id,
          folderId: emails.folderId,
          subject: emails.subject,
          fromEmail: emails.fromEmail,
          receivedAt: emails.receivedAt,
          isRead: emails.isRead,
        })
        .from(emails)
        .where(eq(emails.userId, userId))
        .orderBy(desc(emails.receivedAt))
        .limit(rowLimit);
      reply.send({ count: rows.length, emails: rows });
    } catch (err) {
      console.error("debug emails error:", err);
      reply.code(500).send({ error: "Failed to fetch emails" });
    }
  });
}
