// API Request/Response Types

export interface AuthLoginResponse {
  authUrl: string;
  message: string;
}

export interface AuthCallbackQuery {
  code?: string;
  error?: string;
  error_description?: string;
}

export interface RefreshTokenRequest {
  email: string;
}

export interface SendMailRequest {
  to: string;
  subject: string;
  body: string;
}

export interface MailMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  sender: {
    emailAddress: {
      address: string;
      name: string;
    };
  };
  receivedDateTime: string;
  isRead: boolean;
  hasAttachments: boolean;
}

export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
  childFolderCount: number;
}

export interface MessagesResponse {
  messages: MailMessage[];
  total: number;
}

export interface FoldersResponse {
  folders: MailFolder[];
}

export interface SendMailResponse {
  success: boolean;
  message: string;
}

export interface User {
  id: string;
  email: string;
  displayName?: string;
  accessToken: string;
  refreshToken?: string;
  tokenExpiry?: Date;
  createdAt: Date;
  updatedAt: Date;
}

export interface JWTPayload {
  email: string;
  sub: string;
  iat?: number;
  exp?: number;
}
