import dotenv from "dotenv";

dotenv.config();

export const config = {
  azure: {
    clientId: process.env.AZURE_CLIENT_ID!,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
    tenantId: process.env.AZURE_TENANT_ID!,
    redirectUri: process.env.REDIRECT_URI!,
  },
  database: {
    url: process.env.DATABASE_URL!,
  },
  jwt: {
    secret: process.env.JWT_SECRET || "your-secret-key",
  },
  server: {
    port: parseInt(process.env.PORT || "3001"),
    env: process.env.NODE_ENV || "development",
  },
  frontend: {
    url: process.env.FRONTEND_URL || "http://localhost:5173",
  },
  webhook: {
    notificationUrl: process.env.WEBHOOK_URL || process.env.WEBHOOK_NOTIFICATION_URL || undefined,
  },
};

// Validate required env variables
const requiredVars = [
  "AZURE_CLIENT_ID",
  "AZURE_CLIENT_SECRET",
  "AZURE_TENANT_ID",
  "DATABASE_URL",
];

for (const varName of requiredVars) {
  if (!process.env[varName]) {
    console.warn(`⚠️  Missing environment variable: ${varName}`);
  }
}
