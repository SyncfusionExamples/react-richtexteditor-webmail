import { config } from "../config/env.js";
import { db } from "../db/index.js";
import { subscriptions } from "../db/schema.js";
import { and, eq, or } from "drizzle-orm";

export async function createSubscription(userId: string, accessToken: string) {
  try {
    if (!config.webhook || !config.webhook.notificationUrl) {
      console.warn(
        "WEBHOOK_URL not configured; skipping subscription creation.",
        "Set WEBHOOK_URL or WEBHOOK_NOTIFICATION_URL to your public https://<ngrok>.dev/webhook endpoint."
      );
      console.warn("WEBHOOK_URL=", process.env.WEBHOOK_URL, "WEBHOOK_NOTIFICATION_URL=", process.env.WEBHOOK_NOTIFICATION_URL);
      return null;
    }

    const notificationUrl = config.webhook.notificationUrl;
    const resource = "me/mailFolders('Inbox')/messages";

    const existingSubscription = await db.query.subscriptions.findFirst({
      where: and(
        eq(subscriptions.userId, userId),
        eq(subscriptions.resource, resource)
      ),
    });

    // If we already have an active subscription for this user/resource, keep it
    if (existingSubscription && existingSubscription.expirationDateTime > new Date()) {
      console.log(`🔔 Reusing active subscription ${existingSubscription.subscriptionId} for user ${userId}`);
      return existingSubscription;
    }

    const expiration = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString(); // 3 days

    const body = {
      changeType: "created,updated,deleted",
      notificationUrl,
      resource,
      expirationDateTime: expiration,
      clientState: "steed-client",
    };

    console.log("🔔 Creating subscription with notificationUrl:", notificationUrl);
    console.log("🔔 Subscription request body:", JSON.stringify(body));
    const res = await fetch("https://graph.microsoft.com/v1.0/subscriptions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    const data = await res.json();

    console.log(`Graph subscriptions response status: ${res.status}`);
    console.log("Graph subscription create response:", data);
    if (!res.ok) {
      console.warn("Failed to create subscription:", data);
      return null;
    }

    if (existingSubscription) {
      await db.update(subscriptions)
        .set({
          subscriptionId: data.id,
          notificationUrl: data.notificationUrl,
          expirationDateTime: new Date(data.expirationDateTime),
          clientState: data.clientState,
          changeType: data.changeType,
          deltaLink: null,
        })
        .where(eq(subscriptions.id, existingSubscription.id));
      return {
        ...existingSubscription,
        subscriptionId: data.id,
        notificationUrl: data.notificationUrl,
        expirationDateTime: new Date(data.expirationDateTime),
        clientState: data.clientState,
        changeType: data.changeType,
        deltaLink: null,
      };
    }

    await db.insert(subscriptions).values({
      id: data.id,
      subscriptionId: data.id,
      userId,
      resource: data.resource,
      changeType: data.changeType,
      notificationUrl: data.notificationUrl,
      expirationDateTime: new Date(data.expirationDateTime),
      clientState: data.clientState,
      deltaLink: null,
    });

    return data;

    console.log("✅ Created or updated Graph subscription row:", data.id);
    return data;
  } catch (err) {
    console.error("createSubscription error:", err);
    return null;
  }
}

export async function updateSubscriptionDeltaLink(subscriptionId: string, deltaLink: string) {
  try {
    await db.update(subscriptions).set({ deltaLink }).where(eq(subscriptions.subscriptionId, subscriptionId));
  } catch (err) {
    console.error("updateSubscriptionDeltaLink error:", err);
  }
}

export async function findSubscriptionById(subscriptionId: string) {
  return db.query.subscriptions.findFirst({
    where: or(
      eq(subscriptions.id, subscriptionId),
      eq(subscriptions.subscriptionId, subscriptionId)
    ),
  });
}
