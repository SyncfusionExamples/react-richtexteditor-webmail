import { pgTable, text, timestamp, boolean } from "drizzle-orm/pg-core";

export const users = pgTable("users", {
  id: text("id").primaryKey(),
  email: text("email").unique(),
  displayName: text("display_name"),
  createdAt: timestamp("created_at").defaultNow(),
});

export const accounts = pgTable("accounts", {
  userId: text("user_id"),
  accessToken: text("access_token"),
  refreshToken: text("refresh_token"),
  expiresAt: timestamp("expires_at"),
});

export const folders = pgTable("folders", {
  id: text("id").primaryKey(),
  userId: text("user_id"),
  displayName: text("display_name"),
  parentId: text("parent_id"),
});

export const emails = pgTable("emails", {
  id: text("id").primaryKey(),
  userId: text("user_id"),
  folderId: text("folder_id"),
  subject: text("subject"),
  fromEmail: text("from_email"),
  receivedAt: timestamp("received_at"),
  isRead: boolean("is_read"),
});

export const subscriptions = pgTable("subscriptions", {
  id: text("id").primaryKey(),
  subscriptionId: text("subscription_id"),
  userId: text("user_id"),
  resource: text("resource"),
  changeType: text("change_type"),
  notificationUrl: text("notification_url"),
  expirationDateTime: timestamp("expiration_date_time"),
  clientState: text("client_state"),
  deltaLink: text("delta_link"),
});