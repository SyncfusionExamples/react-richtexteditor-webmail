/**
 * API Client for Mail Application
 * Handles all communication with the backend
 */

const API_URL = import.meta.env.VITE_API_URL || 'http://localhost:3001';

export interface InitialDataResponse {
  folders: any[];
  mailsByFolder: { [folderId: string]: any[] };
  totalMailCount: number;
}

export interface FolderMessagesResponse {
  mails: any[];
  total: number;
}

class MailAPI {
  private token: string | null = null;

  constructor() {
    this.token = localStorage.getItem('authToken');
  }

  /**
   * Update the JWT token (called after login)
   */
  setToken(token: string) {
    this.token = token;
    localStorage.setItem('authToken', token);
  }

  /**
   * Clear the token (called on logout)
   */
  clearToken() {
    this.token = null;
    localStorage.removeItem('authToken');
  }

  /**
   * Read the current token from localStorage. This is the source of
   * truth for the JWT - the in-memory `this.token` is just a cache
   * that may go stale if the token is set by a different path (e.g.
   * the OAuth redirect handler) after the singleton was constructed.
   */
  private getToken(): string | null {
    if (this.token) return this.token;
    const fromStorage = localStorage.getItem('authToken');
    if (fromStorage) {
      // Repopulate the cache so subsequent calls are fast.
      this.token = fromStorage;
    }
    return fromStorage;
  }

  /**
   * Get authorization headers
   */
  private getHeaders(): Record<string, string> {
    const token = this.getToken();
    return {
      'Content-Type': 'application/json',
      ...(token ? { 'Authorization': `Bearer ${token}` } : {}),
    };
  }

  /**
   * Make API request with error handling
   */
  private async request<T>(
    endpoint: string,
    options: RequestInit = {}
  ): Promise<T> {
    const url = `${API_URL}${endpoint}`;

    try {
      const response = await fetch(url, {
        ...options,
        headers: {
          ...this.getHeaders(),
          ...options.headers,
        },
      });

      if (!response.ok) {
        if (response.status === 401) {
          // Token expired or invalid
          this.clearToken();
          throw new Error('Unauthorized. Please login again.');
        }

        // Try to read the error body for a useful message; fall back to
        // the status text. We deliberately do NOT use `await` here on
        // a 204/empty body to avoid a confusing SyntaxError.
        const errorText = await response.text().catch(() => '');
        let errorData: { error?: string } = {};
        if (errorText) {
          try {
            errorData = JSON.parse(errorText) as { error?: string };
          } catch {
            // body wasn't JSON - ignore
          }
        }
        throw new Error(
          errorData.error || `API Error: ${response.statusText || response.status}`
        );
      }

      return await response.json();
    } catch (error) {
      // Log network/fetch failures at warn level so they don't look like
      // application bugs. The caller (ComposeEditor) is responsible for
      // surfacing a user-facing message via the toast.
      console.warn(`API Request Failed [${endpoint}]:`, error);
      throw error;
    }
  }

  /**
   * GET /mail/initial-data
   * Fetch all folders and initial messages after login
   */
  async getInitialData(): Promise<InitialDataResponse> {
    if (!this.token) {
      throw new Error('No authentication token available');
    }

    return this.request<InitialDataResponse>('/mail/initial-data');
  }

  /**
   * GET /mail/folders
   * Fetch all mail folders
   */
  async getFolders() {
    if (!this.token) {
      throw new Error('No authentication token available');
    }

    return this.request('/mail/folders');
  }

  /**
   * GET /mail/folders/:folderId/messages
   * Fetch messages in a specific folder
   */
  async getFolderMessages(
    folderId: string,
    skip: number = 0,
    take: number = 50
  ): Promise<FolderMessagesResponse> {
    if (!this.token) {
      throw new Error('No authentication token available');
    }

    return this.request<FolderMessagesResponse>(
      `/mail/folders/${encodeURIComponent(folderId)}/messages?skip=${skip}&take=${take}`
    );
  }

  /**
   * GET /mail/messages/:id
   * Fetch a specific message, including its attachments (the backend
   * uses `$expand=attachments` so the response includes the full
   * attachment list without the base64 content bytes).
   *
   * IMPORTANT: Graph message IDs contain characters like `+`, `/`, and
   * `=` that MUST be percent-encoded in URLs. If we pass the raw id,
   * the browser/HTTP layer interprets the `=` as a query-string
   * separator and the `+` as a literal space, which truncates the
   * path and Fastify responds with 404 (route not found). This is
   * the root cause of the "Not Found" error that started appearing
   * after we added attachments rendering.
   */
  async getMessage(messageId: string): Promise<any> {
    if (!this.token) {
      throw new Error('No authentication token available');
    }

    return this.request(`/mail/messages/${encodeURIComponent(messageId)}`);
  }

  /**
   * GET /mail/messages/:id/full
   * Fetch full message with body content
   */
  async getFullMessage(messageId: string) {
    if (!this.token) {
      throw new Error('No authentication token available');
    }

    return this.request(`/mail/messages/${encodeURIComponent(messageId)}/full`);
  }

  /**
   * GET /mail/messages/:id/attachments/:attachmentId
   * Download the raw bytes of an attachment. Returns a Blob with the
   * correct content type so the browser can save it locally or open
   * it in a new tab.
   *
   * We do NOT use the generic `request()` helper because that helper
   * always parses JSON. The attachment endpoint returns binary data
   * (image/png, application/pdf, etc.), so we call `fetch` directly
   * and use `response.blob()`.
   */
  async downloadAttachment(
    messageId: string,
    attachmentId: string
  ): Promise<{ blob: Blob; contentType: string; fileName: string }> {
    if (!this.token) {
      throw new Error('No authentication token available');
    }

    const url = `${API_URL}/mail/messages/${encodeURIComponent(
      messageId
    )}/attachments/${encodeURIComponent(attachmentId)}`;

    const response = await fetch(url, {
      method: 'GET',
      headers: this.getHeaders(),
    });

    if (!response.ok) {
      if (response.status === 401) {
        this.clearToken();
        throw new Error('Unauthorized. Please login again.');
      }
      const errorText = await response.text().catch(() => '');
      throw new Error(
        errorText || `Download failed: ${response.status} ${response.statusText}`
      );
    }

    // The Content-Disposition header carries the original filename,
    // encoded per RFC 5987. We try to extract the plain filename from
    // it; if that fails we fall back to the bare attachment id.
    const disposition = response.headers.get('Content-Disposition') || '';
    const fileNameMatch = disposition.match(/filename\*=UTF-8''([^;]+)/i);
    const plainMatch = disposition.match(/filename="?([^";]+)"?/i);
    let fileName = '';
    if (fileNameMatch) {
      try {
        fileName = decodeURIComponent(fileNameMatch[1]);
      } catch {
        fileName = fileNameMatch[1];
      }
    } else if (plainMatch) {
      fileName = plainMatch[1];
    }
    if (!fileName) fileName = `attachment-${attachmentId}`;

    const blob = await response.blob();
    const contentType =
      response.headers.get('Content-Type') || 'application/octet-stream';
    return { blob, contentType, fileName };
  }

  /**
   * POST /mail/send
   * Send an email
   */
  async sendMail(payload: {
    to: string | string[];
    subject: string;
    body: string;
    cc?: string[];
    bcc?: string[];
  }) {
    if (!this.token) {
      throw new Error('No authentication token available');
    }

    return this.request('/mail/send', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  }

  /**
   * POST /mail/messages/:id/delete
   * Delete a message permanently
   */
  // async deleteMail(messageId: string) {
  //   if (!this.token) {
  //     throw new Error('No authentication token available');
  //   }

  //   // URL encode the message ID since it contains special characters like = and /
  //   const encodedId = encodeURIComponent(messageId);
  //   return this.request(`/mail/messages/${encodedId}/delete`, {
  //     method: 'POST',
  //   });
  // }

async deleteMail(messageId: string) {
  if (!this.token) {
    throw new Error('No authentication token available');
  }

  // Send messageId in request body instead of URL path
  // This avoids routing issues with special characters in the ID
  return this.request('/mail/messages/delete', {
    method: 'POST',
    body: JSON.stringify({ messageId }),
  });
}

  /**
   * POST /mail/messages/move
   * Move a message to a specific folder
   */
  async moveMail(messageId: string, folderId: string) {
    if (!this.token) {
      throw new Error('No authentication token available');
    }

    // Send both messageId and folderId in request body
    return this.request('/mail/messages/move', {
      method: 'POST',
      body: JSON.stringify({ messageId, folderId }),
    });
  }

  /**
   * POST /mail/messages/restore
   * Restore a message from deleted items to inbox
   */
  async restoreMail(messageId: string) {
    if (!this.token) {
      throw new Error('No authentication token available');
    }

    // Send messageId in request body instead of URL path
    // This avoids routing issues with special characters in the ID
    return this.request('/mail/messages/restore', {
      method: 'POST',
      body: JSON.stringify({ messageId }),
    });
  }

  /**
   * POST /mail/drafts/save
   * Save a draft message
   */
  async saveDraft(payload: {
    to: string[];
    subject: string;
    body: string;
    cc?: string[];
    bcc?: string[];
  }) {
    if (!this.token) {
      throw new Error('No authentication token available');
    }

    return this.request('/mail/drafts/save', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  }

  /**
   * GET /auth/refresh
   * Refresh the access token
   */
  async refreshToken(email: string) {
    return this.request('/auth/refresh', {
      method: 'POST',
      body: JSON.stringify({ email }),
    });
  }

  // ---------------------------------------------------------------
  // Attachment endpoints (Microsoft Graph upload session flow)
  // ---------------------------------------------------------------

  /**
   * Convert a File or Blob to a base64 string (without the data URL prefix).
   * Used by both small and chunked uploads so the backend can accept
   * JSON payloads instead of multipart/form-data.
   */
  async fileToBase64(file: File | Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        const base64 = result.includes(',') ? result.split(',')[1] : result;
        resolve(base64);
      };
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsDataURL(file);
    });
  }

  /**
   * POST /attachments/draft
   * Create an empty draft to which attachments will be added.
   */
  async createAttachmentDraft(payload: {
    subject: string;
    to: string[];
    cc?: string[];
    bcc?: string[];
  }) {
    if (!this.getToken()) throw new Error('No authentication token available');
    return this.request<{ success: boolean; draftId: string }>(
      '/attachments/draft',
      { method: 'POST', body: JSON.stringify(payload) }
    );
  }

  /**
   * POST /attachments/small
   * Direct upload of a small file (< 3 MB) to a draft message.
   */
  async uploadSmallAttachment(payload: {
    draftId: string;
    fileName: string;
    contentType: string;
    base64: string;
  }) {
    if (!this.getToken()) throw new Error('No authentication token available');
    return this.request<{ success: boolean; attachmentId: string; size: number }>(
      '/attachments/small',
      { method: 'POST', body: JSON.stringify(payload) }
    );
  }

  /**
   * POST /attachments/upload-session
   * Create a chunked upload session for a large file (>= 3 MB).
   * The response includes a pre-authenticated uploadUrl that the
   * client uses to PUT chunks directly to Microsoft Graph (bypassing
   * the backend's body size limit). The URL itself contains the
   * authentication - no Authorization header is needed for chunk PUTs.
   */
  async createAttachmentUploadSession(payload: {
    draftId: string;
    fileName: string;
    contentType: string;
    fileSize: number;
  }) {
    if (!this.getToken()) throw new Error('No authentication token available');
    return this.request<{
      success: boolean;
      uploadUrl: string;
      expirationDateTime: string;
    }>('/attachments/upload-session', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  }

  /**
   * POST /attachments/send
   * Send a draft message that already has its attachments attached.
   *
   * Returns the response from the backend, which now includes the
   * `sentMessageId` — the id of the new message in the Sent Items
   * folder. This is DIFFERENT from `draftId` because Microsoft Graph
   * creates a fresh sent message when /send is called; the draft is
   * removed. The client should use `sentMessageId` going forward so
   * that subsequent fetches (`/mail/messages/:id`) succeed — fetching
   * by the old draftId would 404.
   */
  async sendMailWithAttachments(payload: {
    draftId: string;
    to: string[];
    subject: string;
    body: string;
    cc?: string[];
    bcc?: string[];
    attachmentIds?: string[];
  }) {
    if (!this.token) throw new Error('No authentication token available');
    return this.request<{
      success: boolean;
      message: string;
      draftId?: string;
      sentMessageId?: string;
    }>('/attachments/send', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  }

  /**
   * GET /auth/me
   * Fetch the authenticated user's profile from the backend, which
   * pulls a fresh displayName/jobTitle from Microsoft Graph and
   * reports whether the user has a profile photo set.
   */
  async getMe(): Promise<{
    id: string;
    email: string;
    displayName: string;
    jobTitle?: string | null;
    hasPhoto: boolean;
  }> {
    if (!this.token) {
      throw new Error('No authentication token available');
    }
    return this.request('/auth/me');
  }

  /**
   * GET /auth/me/photo
   * Fetch the user's Microsoft Graph profile photo as a Blob plus
   * its content-type. Throws when the user has no photo (caller
   * falls back to an initials avatar).
   */
  async getMyPhoto(): Promise<{ blob: Blob; contentType: string } | null> {
    if (!this.token) return null;

    const url = `${API_URL}/auth/me/photo`;
    try {
      const response = await fetch(url, {
        method: 'GET',
        headers: this.getHeaders(),
      });
      if (response.status === 404) return null;
      if (!response.ok) {
        throw new Error(
          `Photo request failed: ${response.status} ${response.statusText}`
        );
      }
      const contentType =
        response.headers.get('Content-Type') || 'image/jpeg';
      const blob = await response.blob();
      return { blob, contentType };
    } catch (err) {
      console.warn('Failed to load user photo:', err);
      return null;
    }
  }
}

// Export singleton instance
export const mailAPI = new MailAPI();
