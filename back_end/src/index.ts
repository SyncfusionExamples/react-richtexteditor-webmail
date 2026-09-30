import Fastify from "fastify";
import fastifyJwt from "@fastify/jwt";
import fastifyCors from "@fastify/cors";
import { config } from "./config/env.js";
import { authRoutes } from "./routes/auth.js";
import { mailRoutes } from "./routes/mail.js";
import { webhookRoutes } from "./routes/webhook.js";
import { debugRoutes } from "./routes/debug.js";
import { attachmentRoutes } from "./routes/attachments.js";
import { registerWebSocketServer } from "./services/websocketService.js";
import { startSyncScheduler } from "./services/syncScheduler.js";

const fastify = Fastify({
  logger: true,
  // The default 1 MB body limit is too small for small attachment
  // uploads (a 3 MB binary file base64-encodes to ~4 MB, plus the
  // JSON wrapper). 6 MB gives us room for the largest small file
  // with overhead. Large file chunks are uploaded directly from the
  // client to Microsoft Graph and never traverse this server.
  bodyLimit: 6 * 1024 * 1024,
  // ----------------------------------------------------------------
  // maxParamLength
  // ----------------------------------------------------------------
  // Fastify's internal router (find-my-way) defaults to a max param
  // length of 100 characters. Microsoft Graph message IDs are often
  // 150-200 characters long (base64-encoded, with embedded slashes
  // and `=` padding). Requests with such IDs were silently returning
  // 404 "route not found" before the auth hook even ran, because
  // the router refused to match `:id`.
  //
  // Setting this to 500 gives us plenty of headroom for Graph IDs
  // (which max out around 200 chars), while still preventing
  // unbounded path segments from being processed.
  maxParamLength: 500,
});

// Global request logging for debugging
fastify.addHook("preHandler", async (request, reply) => {
  // FIX 1: Safely skip the WebSocket endpoint to avoid breaking the protocol upgrade lifecycle
  if (request.routerPath === "/ws" || request.url.startsWith("/ws")) {
    return; 
  }

  console.log(`\n${'='.repeat(60)}`);
  console.log(`📨 INCOMING REQUEST`);
  console.log(`   Method: ${request.method}`);
  console.log(`   URL: ${request.url}`);
  console.log(`   Headers: Authorization=${request.headers.authorization ? '✅ Present' : '❌ Missing'}`);
  console.log(`${'='.repeat(60)}`);
});

// Register plugins
async function registerAll() {
  // FIX 2: Register CORS and JWT globally using standard registration chaining
  // Removing 'await' on individual items lets Fastify load them into the global scope together correctly
  fastify.register(fastifyCors, {
    origin: config.frontend.url,
    credentials: true,
  });

  fastify.register(fastifyJwt, {
    secret: config.jwt.secret,
  });

  // Register routes
  fastify.register(authRoutes, { prefix: "/auth" });
  fastify.register(mailRoutes, { prefix: "/mail" });
  // Attachment upload routes (small + chunked)
  fastify.register(attachmentRoutes);
  // Webhook endpoint (public - used by Microsoft Graph)
  fastify.register(webhookRoutes);
  // Debug endpoints (for local testing)
  fastify.register(debugRoutes);

  // Register WebSocket server (at /ws)
  // fastify.register(async (instance) => {
  //   registerWebSocketServer(instance);
  // });
  fastify.register(registerWebSocketServer);
}

// Decorator for authentication
fastify.decorate("authenticate", async function (request: any, reply: any) {
  try {
    await request.jwtVerify();
  } catch (err: any) {
    console.error(`❌ Authentication failed for ${request.method} ${request.url}:`, err.message);
    reply.code(401).send({ error: "Unauthorized", details: err.message });
  }
});

// Global error handler
fastify.setErrorHandler((error, request, reply) => {
  console.error(`\n${'='.repeat(60)}`);
  console.error(`❌ ERROR HANDLER CAUGHT`);
  console.error(`   Method: ${request.method}`);
  console.error(`   URL: ${request.url}`);
  console.error(`   Error: ${error.message}`);
  console.error(`   Status Code: ${error.statusCode || 500}`);
  console.error(`   Stack: ${error.stack}`);
  console.error(`${'='.repeat(60)}\n`);
  
  reply.code(error.statusCode || 500).send({
    error: error.message,
    statusCode: error.statusCode || 500,
  });
});

// Health check
fastify.get("/health", async (request, reply) => {
  return { status: "ok", timestamp: new Date().toISOString() };
});

// Start server
const start = async () => {
  try {
    // Await the entire registration pool to finish compiling before listening
    await registerAll();
    
    await fastify.listen({ port: config.server.port, host: "0.0.0.0" });
    console.log(`✅ Server running at http://localhost:${config.server.port}`);
    console.log(`📧 Environment: ${config.server.env}`);

    // Start the periodic sync scheduler (fallback when webhooks are not available)
    startSyncScheduler();
  } catch (err) {
    fastify.log.error(err);
    process.exit(1);
  }
};

start();
