import jwt from "jsonwebtoken";
import { config } from "../config/env.js";
import { prisma } from "./db.js";
import { cca, graphScopes } from "../config/msal.js";
import { access } from "fs";
import axios from "axios"

/**
 * Generate JWT token for session
 */
export function generateJWT(email: string, sub: string) {
  return jwt.sign(
    {
      email,
      sub,
    },
    config.jwt.secret,
    {
      expiresIn: "7d",
    }
  );
}

/**
 * Verify JWT token
 */
export function verifyJWT(token: string) {
  try {
    return jwt.verify(token, config.jwt.secret);
  } catch (error) {
    return null;
  }
}

/**
 * Get authorization URL for MSAL login
 */
export async function getAuthorizationUrl() {
  try {
    const authCodeUrlParameters = {
      scopes: graphScopes,
      redirectUri: config.azure.redirectUri,
    };

    const response = await cca.getAuthCodeUrl(authCodeUrlParameters);
    return response;
  } catch (error) {
    console.error("Error getting authorization URL:", error);
    throw error;
  }
}

/**
 * Exchange authorization code for tokens
 */
export async function handleAuthCallback(code: string) {
  try {
    const tokenRequest = {
      code: code,
      scopes: graphScopes,
      redirectUri: config.azure.redirectUri,
    };

    const response = await cca.acquireTokenByCode(tokenRequest);

    if (!response || !response.accessToken) {
      throw new Error("No access token received");
    }

    
    // return {
    //   accessToken: response.accessToken,
    //   idToken: response.idToken,
    //   account: response.account,
    //   expiresOn: response.expiresOn,
    // };
    
    const accessToken = response.accessToken;

    // ✅ 1. Get logged-in user profile
    const profileRes = await axios.get(
      "https://graph.microsoft.com/v1.0/me",
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    // ✅ 2. Get user emails (last 10)
    const mailRes = await axios.get(
      "https://graph.microsoft.com/v1.0/me/messages?$top=10",
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    return {
      profile: profileRes.data,
      emails: mailRes.data.value,
      accessToken,
    };



    // Get user info using access token
    const userInfo = response.account;
    const email = userInfo?.username || userInfo?.localAccountId;

    if (!email) {
      throw new Error("Could not get email from token response");
    }

    // Save or update user in database
    const user = await prisma.user.upsert({
      where: { email },
      update: {
        accessToken: response.accessToken,
        refreshToken: response.refreshToken || undefined,
        tokenExpiry: response.expiresOn || undefined,
        displayName: userInfo?.name || undefined,
      },
      create: {
        email,
        accessToken: response.accessToken,
        refreshToken: response.refreshToken || undefined,
        tokenExpiry: response.expiresOn || undefined,
        displayName: userInfo?.name || undefined,
      },
    });

    // Generate JWT for frontend
    const jwtToken = generateJWT(user.email, user.id);

    return {
      user,
      jwtToken,
      accessToken: response.accessToken,
    };
  } catch (error) {
    console.error("Error handling auth callback:", error);
    throw error;
  }
}


/**
 * Refresh access token
 */
export async function refreshAccessToken(email: string) {
  try {
    const user = await prisma.user.findUnique({
      where: { email },
    });

    if (!user || !user.refreshToken) {
      throw new Error("User or refresh token not found");
    }

    const refreshTokenRequest = {
      refreshToken: user.refreshToken,
      scopes: graphScopes,
    };

    const response = await cca.acquireTokenByRefreshToken(refreshTokenRequest);

    if (!response || !response.accessToken) {
      throw new Error("Failed to refresh token");
    }

    // Update user with new tokens
    await prisma.user.update({
      where: { email },
      data: {
        accessToken: response.accessToken,
        refreshToken: response.refreshToken || user.refreshToken,
        tokenExpiry: response.expiresOn || undefined,
      },
    });

    return response.accessToken;
  } catch (error) {
    console.error("Error refreshing token:", error);
    throw error;
  }
}

/**
 * Get user by email
 */
export async function getUserByEmail(email: string) {
  return prisma.user.findUnique({
    where: { email },
  });
}
