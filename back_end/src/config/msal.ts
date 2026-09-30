import { PublicClientApplication, ConfidentialClientApplication } from "@azure/msal-node";
import { config } from "./env.js";

export const msalConfig = {
  auth: {
    clientId: config.azure.clientId,
    // Use a tenant-specific authority when you only want users from one Azure AD tenant.
    // Use "organizations" or "common" in AZURE_TENANT_ID for multi-tenant / personal account support.
    authority: `https://login.microsoftonline.com/common`,
    clientSecret: config.azure.clientSecret,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel: any, message: string, containsPii: boolean) {
        console.log(message);
      },
      piiLoggingEnabled: false,
      logLevel: "verbose" as any,
    },
  },
};

// Server-side confidential client (for backend)
export const cca = new ConfidentialClientApplication(msalConfig);

// Graph API scopes
// Recommended scopes for interactive auth code flow — include openid/profile/offline_access
export const graphScopes = [
  "openid",
  "profile",
  "offline_access",
  "User.Read",
  "Mail.Read",
  "Mail.Send",
  "Mail.ReadWrite"
];
