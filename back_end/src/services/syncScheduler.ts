import { db } from "../db";
import { accounts } from "../db/schema";
import { runDeltaSync } from "./syncServices.js";
import { decrypt } from "../utils/encryption.js";
import { getActiveUserIds } from "./websocketService.js";
import { eq, inArray } from "drizzle-orm";

/**
 * Periodic sync scheduler
 * Runs delta sync for all active users every N seconds
 * This is a fallback when webhooks are not available
 */

const SYNC_INTERVAL_MS = parseInt(process.env.SYNC_INTERVAL_MS || "30000"); // Default: 30 seconds
let intervalId: NodeJS.Timeout | null = null;

export function startSyncScheduler() {
  if (intervalId) {
    console.log("⚠️ Sync scheduler already running");
    return;
  }

  console.log(`🔄 Starting periodic sync scheduler (interval: ${SYNC_INTERVAL_MS}ms)`);

  intervalId = setInterval(async () => {
    try {
      // Only sync users with active WebSocket connections
      const activeUserIds = getActiveUserIds();

      if (activeUserIds.length === 0) {
        // No active users, skip
        return;
      }

      console.log(`🔄 Running periodic sync for ${activeUserIds.length} active user(s): ${activeUserIds.join(', ')}`);

      for (const userId of activeUserIds) {
        try {
          // Get the account for this user
          const account = await db.query.accounts.findFirst({
            where: eq(accounts.userId, userId),
          });

          if (!account || !account.accessToken) {
            console.warn(`⚠️ No access token for user ${userId}, skipping`);
            continue;
          }

          const accessToken = decrypt(account.accessToken);
          await runDeltaSync(userId, accessToken);
          console.log(`✅ Sync completed for user ${userId}`);
        } catch (err) {
          console.error(`❌ Sync failed for user ${userId}:`, err);
        }
      }
    } catch (err) {
      console.error("❌ Sync scheduler error:", err);
    }
  }, SYNC_INTERVAL_MS);
}

export function stopSyncScheduler() {
  if (intervalId) {
    clearInterval(intervalId);
    intervalId = null;
    console.log("🛑 Sync scheduler stopped");
  }
}
