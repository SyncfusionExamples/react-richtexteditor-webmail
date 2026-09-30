import { FastifyInstance } from "fastify";
import { db } from "../db";
import { emails } from "../db/schema";
import { desc, lt } from "drizzle-orm";

export async function emailRoutes(fastify: FastifyInstance) {

  fastify.get("/api/emails", async (req, reply) => {
    const { folderId, cursor } = req.query as any;

    let query = db.select().from(emails)
      .where(emails.folderId.eq(folderId))
      .orderBy(desc(emails.receivedAt))
      .limit(50);

    if (cursor) {
      query = query.where(lt(emails.receivedAt, new Date(cursor)));
    }

    const results = await query;

    return {
      data: results,
      nextCursor: results.length
        ? results[results.length - 1].receivedAt
        : null,
    };
  });
}
``