import { FastifyInstance, FastifyRequest, FastifyReply } from "fastify";
import {
  getAuthorizationUrl,
  handleAuthCallback,
  refreshAccessToken,
  generateJWT,
} from "../services/authService.js";
import { config } from "../config/env.js";
import { sessionStore } from "../services/sessionStore.js";

export async function authRoutes(fastify: FastifyInstance) {
  /**
   * GET /auth/login
   * Returns the authorization URL for MSAL login
   */
  fastify.get("/login", async (request: FastifyRequest, reply: FastifyReply) => {
    try {
      const authUrl = await getAuthorizationUrl();
      reply.send({
        authUrl,
        message: "Redirect user to this URL to login",
      });
    } catch (error) {
      console.error("Auth login error:", error);
      reply.code(500).send({
        error: "Failed to generate login URL",
      });
    }
  });

  /**
   * GET /auth/callback
   * Handle OAuth callback from Microsoft
   */
  fastify.get(
    "/callback",
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { code, error, error_description } = request.query as {
          code?: string;
          error?: string;
          error_description?: string;
        };

        if (error) {
          return reply.code(400).send({
            error,
            description: error_description,
          });
        }

        if (!code) {
          return reply.code(400).send({
            error: "No authorization code provided",
          });
        }

        // Handle OAuth callback and get user data + access token
        const authResponse = await handleAuthCallback(code);

        // Extract email from profile
        const email = authResponse.profile.mail || authResponse.profile.userPrincipalName;
        if (!email) {
          return reply.code(400).send({
            error: "Could not retrieve email from Microsoft account",
          });
        }

        // Generate unique user ID (based on email hash or use profile ID)
        const userId = `user_${Date.now()}_${Math.random().toString(36).substring(7)}`;

        // Store user session in memory with Microsoft access token
        sessionStore.saveSession(email, {
          id: userId,
          displayName: authResponse.profile.displayName,
          accessToken: authResponse.accessToken,
        });

        // Generate JWT token for frontend
        const jwtToken = generateJWT(email, userId);

        // Redirect to frontend with JWT token and email in URL
        const redirectUrl = `${config.frontend.url}?token=${encodeURIComponent(jwtToken)}&email=${encodeURIComponent(email)}`;
        reply.redirect(redirectUrl);

      } catch (error) {
        console.error("Auth callback error:", error);
        reply.code(500).send({
          error: "Authentication failed",
        });
      }
    }
  );

  /**
   * POST /auth/refresh
   * Refresh the access token
   */
  fastify.post(
    "/refresh",
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const { email } = request.body as { email: string };

        if (!email) {
          return reply.code(400).send({
            error: "Email is required",
          });
        }

        const newAccessToken = await refreshAccessToken(email);

        reply.send({
          accessToken: newAccessToken,
        });
      } catch (error) {
        console.error("Token refresh error:", error);
        reply.code(500).send({
          error: "Failed to refresh token",
        });
      }
    }
  );
}
