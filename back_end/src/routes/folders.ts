import { FastifyInstance } from "fastify";
import { db } from "../db";
import { folders } from "../db/schema";

export async function folderRoutes(fastify: FastifyInstance) {

  fastify.get("/api/folders", async (req, reply) => {
    const { userId } = req.query as any;

    const data = await db.select().from(folders)
      .where(folders.userId.eq(userId));

    return data;
  });
}