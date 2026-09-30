/**
 * In-Memory Session Store
 * Stores user sessions and tokens in memory during runtime
 * Standard approach for session management
 */

interface UserSession {
  id: string;
  email: string;
  displayName?: string;
  accessToken: string;
  createdAt: Date;
  expiresAt: Date;
}

class SessionStore {
  private sessions: Map<string, UserSession> = new Map();

  /**
   * Store user session
   */
  saveSession(email: string, data: {
    id: string;
    displayName?: string;
    accessToken: string;
  }): UserSession {
    const session: UserSession = {
      id: data.id,
      email,
      displayName: data.displayName,
      accessToken: data.accessToken,
      createdAt: new Date(),
      expiresAt: new Date(Date.now() + 24 * 60 * 60 * 1000), // 24 hours
    };

    this.sessions.set(email, session);
    return session;
  }

  /**
   * Get session by email
   */
  getSession(email: string): UserSession | undefined {
    const session = this.sessions.get(email);
    
    // Check if session expired
    if (session && session.expiresAt < new Date()) {
      this.sessions.delete(email);
      return undefined;
    }

    return session;
  }

  /**
   * Update access token for user
   */
  updateAccessToken(email: string, accessToken: string): boolean {
    const session = this.sessions.get(email);
    if (!session) return false;

    session.accessToken = accessToken;
    session.expiresAt = new Date(Date.now() + 24 * 60 * 60 * 1000);
    return true;
  }

  /**
   * Delete session (logout)
   */
  deleteSession(email: string): boolean {
    return this.sessions.delete(email);
  }

  /**
   * Cleanup expired sessions (run periodically)
   */
  cleanupExpiredSessions(): number {
    const now = new Date();
    let deleted = 0;

    for (const [email, session] of this.sessions.entries()) {
      if (session.expiresAt < now) {
        this.sessions.delete(email);
        deleted++;
      }
    }

    return deleted;
  }

  /**
   * Get all active sessions (useful for debugging)
   */
  getAllSessions(): UserSession[] {
    return Array.from(this.sessions.values()).filter(
      (session) => session.expiresAt > new Date()
    );
  }
}

// Export singleton instance
export const sessionStore = new SessionStore();

// Cleanup expired sessions every 10 minutes
setInterval(() => {
  const cleaned = sessionStore.cleanupExpiredSessions();
  if (cleaned > 0) {
    console.log(`Cleaned up ${cleaned} expired sessions`);
  }
}, 10 * 60 * 1000);
