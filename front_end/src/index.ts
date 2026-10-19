import Fastify from "fastify";
import fastifyJwt from "@fastify/jwt";
import fastifyCors from "@fastify/cors";
import { config } from "./config/env.js";
import { authRoutes } from "./routes/auth.js";
import { mailRoutes } from "./routes/mail.js";

const fastify = Fastify({
  logger: true,
});

// Register plugins
await fastify.register(fastifyCors, {
  origin: config.frontend.url,
  credentials: true,
});

await fastify.register(fastifyJwt, {
  secret: config.jwt.secret,
});

// Decorator for authentication
fastify.decorate("authenticate", async function (request: any, reply: any) {
  try {
    await request.jwtVerify();
  } catch (err) {
    reply.code(401).send({ error: "Unauthorized" });
  }
});

// Health check
fastify.get("/health", async (request, reply) => {
  return { status: "ok", timestamp: new Date().toISOString() };
});

// Register routes
await fastify.register(authRoutes, { prefix: "/auth" });
await fastify.register(mailRoutes, { prefix: "/mail" });

// Start server
const start = async () => {
  try {
    await fastify.listen({ port: config.server.port, host: "0.0.0.0" });
    console.log(`✅ Server running at http://localhost:${config.server.port}`);
    console.log(`📧 Environment: ${config.server.env}`);
  } catch (err) {
    fastify.log.error(err);
    process.exit(1);
  }
};

start();
