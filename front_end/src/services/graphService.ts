import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/tokenCredentialAuthProvider";

/**
 * Create a Graph API client with authentication
 * Using Bearer token directly for simple auth
 */
export function createGraphClient(accessToken: string) {
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => accessToken,
    },
  });
}

/**
 * Get user's mail inbox
 */
export async function getInbox(graphClient: any, skip = 0, take = 50) {
  try {
    const messages = await graphClient
      .api("/me/mailFolders/inbox/messages")
      .select([
        "id",
        "subject",
        "bodyPreview",
        "sender",
        "receivedDateTime",
        "isRead",
        "hasAttachments",
      ])
      .top(take)
      .skip(skip)
      .orderby("receivedDateTime desc")
      .get();

    return messages;
  } catch (error) {
    console.error("Error fetching inbox:", error);
    throw error;
  }
}

/**
 * Get a specific message by ID
 */
export async function getMessage(graphClient: any, messageId: string) {
  try {
    const message = await graphClient.api(`/me/messages/${messageId}`).get();
    return message;
  } catch (error) {
    console.error("Error fetching message:", error);
    throw error;
  }
}

/**
 * Send an email
 */
export async function sendMail(
  graphClient: any,
  to: string,
  subject: string,
  body: string
) {
  try {
    const sendMail = {
      message: {
        subject: subject,
        body: {
          contentType: "HTML",
          content: body,
        },
        toRecipients: [
          {
            emailAddress: {
              address: to,
            },
          },
        ],
      },
      saveToSentItems: true,
    };

    await graphClient.api("/me/sendMail").post(sendMail);
    return { success: true };
  } catch (error) {
    console.error("Error sending mail:", error);
    throw error;
  }
}

/**
 * Get all folders
 */
export async function getFolders(graphClient: any) {
  try {
    const folders = await graphClient
      .api("/me/mailFolders")
      .select(["id", "displayName", "parentFolderId", "childFolderCount", "unreadItemCount"])
      .get();

    return folders;
  } catch (error) {
    console.error("Error fetching folders:", error);
    throw error;
  }
}

/**
 * Get messages in a specific folder with unread count
 */
export async function getFolderMessages(graphClient: any, folderId: string, skip = 0, take = 50) {
  try {
    const messages = await graphClient
      .api(`/me/mailFolders/${folderId}/messages`)
      .select([
        "id",
        "subject",
        "bodyPreview",
        "sender",
        "from",
        "receivedDateTime",
        "isRead",
        "categories",
        "flag",
      ])
      .top(take)
      .skip(skip)
      .orderby("receivedDateTime desc")
      .get();

    return messages;
  } catch (error) {
    console.error(`Error fetching messages from folder ${folderId}:`, error);
    throw error;
  }
}

/**
 * Get a specific message with full body content
 */
export async function getFullMessage(graphClient: any, messageId: string) {
  try {
    const message = await graphClient
      .api(`/me/messages/${messageId}`)
      .select([
        "id",
        "subject",
        "bodyPreview",
        "body",
        "sender",
        "from",
        "receivedDateTime",
        "isRead",
        "categories",
        "flag",
        "toRecipients",
        "ccRecipients",
        "bccRecipients",
        "hasAttachments",
      ])
      .get();

    return message;
  } catch (error) {
    console.error(`Error fetching message ${messageId}:`, error);
    throw error;
  }
}

/**
 * Get all folders recursively with their messages (for initial load)
 * Returns structure: { folderId: [messages] }
 */
export async function getAllFoldersWithMessages(graphClient: any) {
  try {
    const folders = await getFolders(graphClient);
    console.log('📂 All folders from Graph API:', folders.value?.map((f: any) => ({ id: f.id, displayName: f.displayName, parentFolderId: f.parentFolderId })));
    
    const mailsByFolder: { [folderId: string]: any[] } = {};

    // Fetch messages for each folder
    // Limit to top-level folders for initial load
    const topLevelFolders = folders.value?.filter((f: any) => !f.parentFolderId) || [];
    console.log('📁 Top-level folders:', topLevelFolders.map((f: any) => ({ id: f.id, displayName: f.displayName })));

    if (topLevelFolders.length === 0) {
      console.warn('⚠️ No top-level folders found! Trying all folders instead...');
      for (const folder of (folders.value || [])) {
        try {
          console.log(`📥 Fetching messages for folder: ${folder.displayName} (${folder.id})`);
          const messages = await getFolderMessages(graphClient, folder.id, 0, 50);
          console.log(`✅ Got ${messages.value?.length || 0} messages from ${folder.displayName}`);
          mailsByFolder[folder.id] = messages.value || [];
        } catch (err) {
          console.warn(`❌ Failed to fetch messages for folder ${folder.id}:`, err);
          mailsByFolder[folder.id] = [];
        }
      }
    } else {
      for (const folder of topLevelFolders) {
        try {
          console.log(`📥 Fetching messages for folder: ${folder.displayName} (${folder.id})`);
          const messages = await getFolderMessages(graphClient, folder.id, 0, 50);
          console.log(`✅ Got ${messages.value?.length || 0} messages from ${folder.displayName}`);
          mailsByFolder[folder.id] = messages.value || [];
        } catch (err) {
          console.warn(`❌ Failed to fetch messages for folder ${folder.id}:`, err);
          mailsByFolder[folder.id] = [];
        }
      }
    }

    console.log('📊 Final mailsByFolder summary:', Object.entries(mailsByFolder).map(([folderId, mails]) => ({ folderId, count: mails.length })));

    return {
      folders: folders.value || [],
      mailsByFolder,
    };
  } catch (error) {
    console.error("❌ Error fetching all folders with messages:", error);
    throw error;
  }
}
