import { FastifyInstance, FastifyRequest, FastifyReply } from "fastify";
import { subscribeMailUpdates } from "../services/eventBus.js";

export async function eventRoutes(fastify: FastifyInstance) {
  fastify.get("/events", async (request: FastifyRequest, reply: FastifyReply) => {
    const token = (request.query as any)?.token;
    if (!token) {
      return reply.code(401).send({ error: "Missing token" });
    }

    try {
      fastify.jwt.verify(token);
    } catch (err) {
      return reply.code(401).send({ error: "Invalid token" });
    }

    reply.raw.setHeader("Content-Type", "text/event-stream");
    reply.raw.setHeader("Cache-Control", "no-cache");
    reply.raw.setHeader("Connection", "keep-alive");
    reply.raw.write("retry: 10000\n\n");

    const onMailUpdate = (event: any) => {
      if (reply.raw.writableEnded) return;
      reply.raw.write(`event: mailUpdated\n`);
      reply.raw.write(`data: ${JSON.stringify(event)}\n\n`);
    };

    const unsubscribe = subscribeMailUpdates(onMailUpdate);

    request.raw.on("close", () => {
      unsubscribe();
      if (!reply.raw.writableEnded) {
        reply.raw.end();
      }
    });

    // Keep the connection open
    return reply;
  });
}
