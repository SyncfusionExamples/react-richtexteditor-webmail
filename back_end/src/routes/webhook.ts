import { FastifyInstance, FastifyRequest, FastifyReply } from "fastify";
import { findSubscriptionById } from "../services/subscriptionService.js";
import { runDeltaSync } from "../services/syncServices.js";
import { db } from "../db/index.js";
import { accounts } from "../db/schema.js";
import { eq } from "drizzle-orm";
import { decrypt } from "../utils/encryption.js";

export async function webhookRoutes(fastify: FastifyInstance) {
  // Validation handshake from Microsoft Graph
  fastify.get("/webhook", async (request: FastifyRequest, reply: FastifyReply) => {
    const q = request.query as any;
    if (q && q.validationToken) {
      // Respond with the validation token as plain text
      console.log("🔔 Received validation request via GET", q.validationToken);
      reply.type("text/plain").send(q.validationToken);
      return;
    }

    reply.code(400).send({ error: "Missing validationToken" });
  });

  // Notification endpoint
  fastify.post("/webhook", async (request: FastifyRequest, reply: FastifyReply) => {
    try {
        const isValidationRequest = Boolean((request.query as any)?.validationToken);
        if (isValidationRequest) {
          const validationToken = (request.query as any).validationToken;
          console.log("🔔 Received validation request via POST", validationToken);
          reply.type("text/plain").send(validationToken);
          return;
        }

        const body = request.body as any;
        if (!body || !body.value) {
          return reply.code(400).send({ error: "Invalid notification payload" });
        }
      for (const notification of body.value) {
        const subscriptionId = notification.subscriptionId || notification.subscription; // some payloads use different keys
        if (!subscriptionId) continue;

        console.log(`🔔 Received notification for subscription ${subscriptionId}`);

        // Find subscription and associated user
        const sub = await findSubscriptionById(subscriptionId);
        if (!sub) {
          console.warn(`No subscription found for id ${subscriptionId}`);
          continue;
        }

        // Get user's access token
        console.log(`🔎 Finding account for user ${sub.userId}`);
        const account = await db.query.accounts.findFirst({ where: eq(accounts.userId, sub.userId) });
        if (!account || !account.accessToken) {
          console.warn(`No account found for user ${sub.userId}`);
          continue;
        }
        console.log(`✅ Found account for user ${sub.userId}`);
        const accessToken = decrypt(account.accessToken);

        // Run delta sync for this user (async, don't block response)
        runDeltaSync(sub.userId, accessToken).catch(err => console.error("runDeltaSync error:", err));
      }

      reply.send({ received: true });
    } catch (err) {
      console.error("Webhook handler error:", err);
      reply.code(500).send({ error: "Internal error" });
    }
  });
}
