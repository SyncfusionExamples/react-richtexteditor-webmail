import { db } from "../db";
import { folders, emails, subscriptions } from "../db/schema";
import { fetchFolders, fetchEmails } from "./graphService";
import { updateSubscriptionDeltaLink } from "./subscriptionService";
import { notifyMailUpdated } from "./eventBus.js";
import { desc, eq } from "drizzle-orm";

/**
 * Run a delta sync for a user's inbox. This will call the delta endpoint and
 * insert new/updated emails into the emails table and update the stored deltaLink.
 */
export async function runDeltaSync(userId: string, accessToken: string) {
  try {
    console.log(`🔁 Running delta sync for user ${userId}`);

    // Pick the latest subscription row for this user (prefer the most recent expiration)
    const sub = await db.query.subscriptions.findFirst({
      where: eq(subscriptions.userId, userId),
      orderBy: [desc(subscriptions.expirationDateTime)],
    });

    if (sub) {
      console.log(`🔎 Using subscription ${sub.subscriptionId} for delta sync`);
      console.log(`   deltaLink present: ${Boolean(sub.deltaLink)}`);
    } else {
      console.warn(`⚠️ No subscription row found for user ${userId}; using initial delta URL`);
    }

    // Build delta URL
    let deltaUrl = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta";
    if (sub && sub.deltaLink) {
      deltaUrl = sub.deltaLink;
    }

    const res = await fetch(deltaUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    if (!res.ok) {
      const txt = await res.text();
      throw new Error(`Delta request failed: ${res.status} ${txt}`);
    }

    const data = await res.json();
    let hasChanges = false;
    // Insert or update emails
    if (Array.isArray(data.value)) {
      for (const m of data.value) {
        try {
          hasChanges = true;
          await db.insert(emails).values({
            id: m.id,
            userId,
            folderId: m.parentFolderId || "inbox",
            subject: m.subject || "",
            fromEmail: m.from?.emailAddress?.address || m.sender?.emailAddress?.address || null,
            receivedAt: m.receivedDateTime ? new Date(m.receivedDateTime) : new Date(),
            isRead: !!m.isRead,
          }).onConflictDoUpdate({
            target: emails.id,
            set: {
              subject: (m.subject || ""),
              isRead: !!m.isRead,
            }
          });
        } catch (err) {
          console.warn("Failed to upsert email", m.id, err);
        }
      }
    }

    // Persist nextLink or deltaLink
    const newDelta = data['@odata.deltaLink'] || data['@odata.nextLink'] || null;
    if (newDelta && sub) {
      await updateSubscriptionDeltaLink(sub.subscriptionId, newDelta);
    }

    //notifyMailUpdated(userId);
    if (hasChanges) {
      notifyMailUpdated(userId);
    }
    console.log(`✅ Delta sync completed for user ${userId}`);
    return true;
  } catch (err) {
    console.error("runDeltaSync error:", err);
    throw err;
  }
}

export async function syncFolders(userId: string, accessToken: string) {
  const data = await fetchFolders(accessToken);

  for (const f of data.value) {
    await db.insert(folders).values({
      id: f.id,
      userId,
      displayName: f.displayName,
      parentId: f.parentFolderId,
    }).onConflictDoNothing();
  }
}

export async function syncEmails(userId: string, folderId: string, accessToken: string) {
  const data = await fetchEmails(accessToken, folderId);

  for (const m of data.value) {
    await db.insert(emails).values({
      id: m.id,
      userId,
      folderId,
      subject: m.subject,
      fromEmail: m.from?.emailAddress?.address,
      receivedAt: new Date(m.receivedDateTime),
      isRead: m.isRead,
    }).onConflictDoNothing();
  }
}