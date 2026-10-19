import { FastifyInstance, FastifyRequest, FastifyReply } from "fastify";
import {
  getAuthorizationUrl,
  handleAuthCallback,
  refreshAccessToken,
  generateJWT,
} from "../services/authService.js";
import { config } from "../config/env.js";
import { sessionStore } from "../services/sessionStore.js";

import { db } from "../db";
import { users, accounts } from "../db/schema";
import { encrypt, decrypt } from "../utils/encryption";
import { createSubscription } from "../services/subscriptionService";
import { eq } from "drizzle-orm";

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
        //Commented for reason
        //const userId = `user_${Date.now()}_${Math.random().toString(36).substring(7)}`;

        // Store user session in memory with Microsoft access token
        //Commented for reason
        // sessionStore.saveSession(email, {
        //   id: userId,
        //   displayName: authResponse.profile.displayName,
        //   accessToken: authResponse.accessToken,
        // });

        const userIdBase = authResponse.profile.id || email;
        const userId = `user_${userIdBase.toString().replace(/[^a-zA-Z0-9_-]/g, "_")}`;

        await db.insert(users).values({
          id: userId,
          email,
          displayName: authResponse.profile.displayName,
        }).onConflictDoNothing();

        const encryptedAccessToken = encrypt(authResponse.accessToken);
        const encryptedRefreshToken = authResponse.refreshToken ? encrypt(authResponse.refreshToken) : null;
        const expiresAt = new Date(Date.now() + 3600 * 1000);

        const existingAccount = await db.query.accounts.findFirst({
          where: eq(accounts.userId, userId),
        });

        if (existingAccount) {
          await db.update(accounts)
            .set({
              accessToken: encryptedAccessToken,
              refreshToken: encryptedRefreshToken,
              expiresAt,
            })
            .where(eq(accounts.userId, userId));
        } else {
          await db.insert(accounts).values({
            userId,
            accessToken: encryptedAccessToken,
            refreshToken: encryptedRefreshToken,
            expiresAt,
          });
        }

        // Try to create a Graph subscription for this user's inbox
        try {
          const subscription = await createSubscription(userId, authResponse.accessToken);
          if (!subscription) {
            console.warn("Subscription creation returned null.");
          }
        } catch (err) {
          console.warn("Failed to create subscription after login:", err);
        }

        // Generate JWT token for frontend
        const jwtToken = generateJWT(email, userId);

        // Redirect to frontend with JWT token, email, and displayName in URL
        const displayName = authResponse.profile.displayName || email;
        const redirectUrl = `${config.frontend.url}?token=${encodeURIComponent(jwtToken)}&email=${encodeURIComponent(email)}&name=${encodeURIComponent(displayName)}`;
        reply.redirect(redirectUrl);

      } catch (error) {
        console.error("Auth callback error:", error);
        reply.code(500).send({
          error:
            error instanceof Error ? error.message : "Authentication failed",
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

  /**
   * GET /auth/me
   * Return the current authenticated user's profile (from DB +
   * a fresh Microsoft Graph /me lookup so displayName/jobTitle stay
   * up to date). Requires a valid JWT in the Authorization header.
   */
  fastify.get(
    "/me",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const user = request.user as { email?: string; sub?: string };
        const userId = user?.sub;
        const email = user?.email;

        if (!userId || !email) {
          return reply.code(401).send({ error: "Invalid token payload" });
        }

        // Try to fetch a fresh profile from Microsoft Graph using the
        // stored access token; fall back to the DB row if the token is
        // missing/expired (we don't refresh here to keep this endpoint
        // fast and side-effect free).
        const account = await db.query.accounts.findFirst({
          where: eq(accounts.userId, userId),
        });

        const dbUser = await db.query.users.findFirst({
          where: eq(users.id, userId),
        });

        let displayName = dbUser?.displayName || email;
        let jobTitle: string | null = null;
        let graphPhotoAvailable = false;

        if (account?.accessToken) {
          try {
            const accessToken = decrypt(account.accessToken);
            const profileRes = await fetch(
              "https://graph.microsoft.com/v1.0/me?$select=displayName,givenName,surname,userPrincipalName,mail,jobTitle",
              { headers: { Authorization: `Bearer ${accessToken}` } }
            );
            if (profileRes.ok) {
              const profile = (await profileRes.json()) as {
                displayName?: string;
                jobTitle?: string;
              };
              if (profile.displayName) displayName = profile.displayName;
              if (profile.jobTitle) jobTitle = profile.jobTitle;

              // Also check if a photo exists (HEAD-style: use /me/photo
              // which returns 200 with metadata or 404 if not set).
              const photoMetaRes = await fetch(
                "https://graph.microsoft.com/v1.0/me/photo",
                { headers: { Authorization: `Bearer ${accessToken}` } }
              );
              graphPhotoAvailable = photoMetaRes.ok;
            }
          } catch (err) {
            console.warn("Failed to refresh /me from Graph:", err);
          }
        }

        return reply.send({
          id: userId,
          email,
          displayName,
          jobTitle,
          hasPhoto: graphPhotoAvailable,
        });
      } catch (error) {
        console.error("Auth /me error:", error);
        return reply.code(500).send({
          error: "Failed to load profile",
        });
      }
    }
  );

  /**
   * GET /auth/me/photo
   * Stream the current user's Microsoft Graph photo as a binary
   * response. Returns 404 if the user has no photo set in Microsoft
   * Graph (the frontend falls back to an initials avatar in that case).
   */
  fastify.get(
    "/me/photo",
    { onRequest: [fastify.authenticate] },
    async (request: FastifyRequest, reply: FastifyReply) => {
      try {
        const user = request.user as { sub?: string };
        const userId = user?.sub;

        if (!userId) {
          return reply.code(401).send({ error: "Invalid token payload" });
        }

        const account = await db.query.accounts.findFirst({
          where: eq(accounts.userId, userId),
        });

        if (!account?.accessToken) {
          return reply.code(404).send({ error: "No account found" });
        }

        const accessToken = decrypt(account.accessToken);
        const photoRes = await fetch(
          "https://graph.microsoft.com/v1.0/me/photo/$value",
          { headers: { Authorization: `Bearer ${accessToken}` } }
        );

        if (!photoRes.ok) {
          return reply.code(404).send({ error: "Photo not available" });
        }

        const contentType =
          photoRes.headers.get("content-type") || "image/jpeg";
        const arrayBuffer = await photoRes.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);

        return reply
          .header("Content-Type", contentType)
          .header("Cache-Control", "private, max-age=3600")
          .send(buffer);
      } catch (error) {
        console.error("Auth /me/photo error:", error);
        return reply.code(500).send({ error: "Failed to load photo" });
      }
    }
  );
}
