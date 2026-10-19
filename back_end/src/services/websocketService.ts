import { FastifyInstance, FastifyRequest } from "fastify";
import fastifyWebsocket, { SocketStream } from "@fastify/websocket";
import WebSocket from "ws"; 
import { verifyJWT } from "./authService.js";
import { subscribeMailUpdates, MailUpdateEvent } from "./eventBus.js";

// Extend FastifyRequest type interface to hold our authenticated userId
interface AuthenticatedWSRequest extends FastifyRequest {
	userId?: string;
}

// In-memory userId -> Set<WebSocket> mapping
const userSockets: Map<string, Set<WebSocket>> = new Map();

// Singleton guard to prevent duplicate EventBus listeners on plugin reload
let isEventBusSubscribed = false;

export async  function registerWebSocketServer(fastify: FastifyInstance) {
	// Register the official Fastify WebSocket plugin
	await fastify.register(fastifyWebsocket);
	
	fastify.addHook("onReady", async () => {
		console.log("✅ WebSocket plugin ready");
	});

	// WebSocket Route with integrated authentication pre-validation hook
	fastify.get("/ws", { 
		websocket: true,
		preValidation: async (request: AuthenticatedWSRequest, reply) => {
			const query = request.query as Record<string, string | undefined>;
			const token = query.token;

			console.log(`🔌 WebSocket connection attempt. token=${token ? token.substring(0, 20) + '...' : 'MISSING'}`);

			if (!token) {
				console.log('❌ Rejecting WebSocket upgrade: Missing token');
				reply.code(401).send({ error: "Unauthorized: Missing Token" });
				return;
			}

			const payload = verifyJWT(token);
			console.log("JWT Payload:", payload);

			if (!payload || typeof payload !== "object" || !payload.sub) {
				console.log('❌ Rejecting WebSocket upgrade: Invalid token');
				reply.code(401).send({ error: "Unauthorized: Invalid Token" });
				return;
			}

			// Store the extracted userId in the request object for the handler below to access
			request.userId = payload.sub as string;
			console.log("PREVALIDATION USER:",request.userId);
		}
	}, (connection: SocketStream, request: AuthenticatedWSRequest) => {
		// FIX: Correctly extract the native WebSocket instance from Fastify's SocketStream wrapper
		// console.log("connection.ws:", (connection as any).ws);
		// console.log("connection.constructor:", connection?.constructor?.name);
		const ws = connection.socket;
		console.log("WS constructor:",      ws.constructor.name);
		console.log("readyState:",     ws.readyState  );
		 
		const userId = request.userId!; // Guaranteed to exist by preValidation hook
		console.log(`✅ WebSocket authenticated and established for user ${userId}`);

		// Register the raw socket for broadcasting updates
		if (!userSockets.has(userId)) {
			userSockets.set(userId, new Set());
		}
		userSockets.get(userId)!.add(ws);
		console.log(`✅ WebSocket registered. Total active sockets for user ${userId}: ${userSockets.get(userId)!.size}`);

		// Listen to events on the individual client connection
		ws.on('close', () => {
			const sockets = userSockets.get(userId);
			if (sockets) {
				sockets.delete(ws);
				// FIX: Read the remaining count safely before removing the map key to avoid "Cannot read properties of undefined" crashing the server
				console.log(`🔴 WebSocket closed for user ${userId}. Total remaining active sockets: ${sockets.size}`);
				if (sockets.size === 0) {
					userSockets.delete(userId);
				}
			}
		});

		ws.on('error', (err: any) => {
			console.error(`⚠️ WebSocket error for user ${userId}:`, err);
		});

		ws.on('message', (rawMessage) => {
			try {
				const message = JSON.parse(rawMessage.toString());
				// Handle client-side heartbeat pings to keep connections alive
				if (message.type === 'ping') {
					ws.send(JSON.stringify({ type: 'pong' }));
					return;
				}
				console.log(`📨 WebSocket payload received from user ${userId}:`, message);
			} catch {
				console.log(`📨 Raw WebSocket text/buffer received from user ${userId}`);
			}
		});
	});

	// FIX: Central EventBus subscription singleton guard prevents memory leaks and multiple listener build-ups
	if (!isEventBusSubscribed) {
		subscribeMailUpdates((event: MailUpdateEvent) => {
			const { userId } = event;
			const sockets = userSockets.get(userId);
			console.log(sockets)
			console.log(`📤 Sending new_email event to user ${userId}, active sockets: ${sockets ? sockets.size : 0}`);
			
			if (sockets && sockets.size > 0) {
				const payload = JSON.stringify({ 
					type: "new_email", 
					id: event.id || undefined, // Recommended to deduplicate on client
					timestamp: event.timestamp 
				});
				
				for (const ws of sockets) {
					// Verify standard WebSocket connection state before sending data
					if (ws.readyState === WebSocket.OPEN) {
						try {
							ws.send(payload);
							console.log(`   ✅ Sent via ws.send()`);
						} catch (err) {
							console.error(`   ❌ ws.send() failed for user ${userId}:`, err);
						}
					} else {
						console.warn( `⚠️ Socket state: ${ws.readyState}`);
						console.log("OPEN:", WebSocket.OPEN, "CURRENT:", ws.readyState);
					}
				}
			}
		});
		isEventBusSubscribed = true;
		console.log("🌐 Central EventBus Mail Subscriber activated");
	}
}

// --- Debug & Multi-Instance Utilities ---

export function getActiveConnectionCount(): number {
	let total = 0;
	for (const sockets of userSockets.values()) {
		total += sockets.size;
	}
	return total;
}

export function getActiveUsers(): string[] {
	return Array.from(userSockets.keys());
}

export function triggerNewEmailEvent(userId: string) {
	const sockets = userSockets.get(userId);
	console.log(`🧪 Manual trigger: Sending new_email event to user ${userId}, sockets: ${sockets ? sockets.size : 0}`);
	if (sockets && sockets.size > 0) {
		const payload = JSON.stringify({ type: "new_email", timestamp: new Date().toISOString() });
		for (const ws of sockets) {
			if (ws.readyState === WebSocket.OPEN) ws.send(payload);
		}
		return true;
	}
	return false;
}

export function getActiveUserIds(): string[] {
	return Array.from(userSockets.keys());
}
