/**
 * OAuth Manager for handling Microsoft OAuth 2.0 authentication
 */

import { config } from '../config/index.js';
import { logger } from '../utils/logger.js';
import { OAuthTokens, SessionData } from '../types/index.js';

export class OAuthManager {
  private sessions: Map<string, SessionData>;

  constructor() {
    this.sessions = new Map();
  }

  /**
   * Generate OAuth authorization URL
   */
  getAuthorizationUrl(state: string): string {
    const { clientId, tenantId, redirectUri } = config.microsoft;
    const scopes = config.graph.scopes;

    const params = new URLSearchParams({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      response_mode: 'query',
      scope: scopes.join(' '),
      state,
      prompt: 'consent', // Force consent to get refresh token
    });

    const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${params.toString()}`;

    logger.info('[OAuthManager] Generated authorization URL', { state });
    return authUrl;
  }

  /**
   * Exchange authorization code for tokens
   */
  async exchangeCodeForTokens(code: string): Promise<OAuthTokens> {
    try {
      logger.info('[OAuthManager] Exchanging code for tokens');

      const { clientId, clientSecret, tenantId, redirectUri } = config.microsoft;

      const params = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        code,
        redirect_uri: redirectUri,
        grant_type: 'authorization_code',
      });

      const response = await fetch(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
          body: params.toString(),
        }
      );

      if (!response.ok) {
        const error = await response.text();
        logger.error('[OAuthManager] Token exchange failed:', error);
        throw new Error('Failed to exchange authorization code for tokens');
      }

      const data = await response.json() as {
        access_token: string;
        refresh_token?: string;
        expires_in?: number;
        token_type?: string;
        scope?: string;
        id_token?: string;
      };

      logger.info('[OAuthManager] Successfully exchanged code for tokens');

      return {
        access_token: data.access_token,
        refresh_token: data.refresh_token || undefined,
        expires_in: data.expires_in || 3600,
        token_type: data.token_type || 'Bearer',
        scope: data.scope || undefined,
        id_token: data.id_token || undefined,
      };
    } catch (error) {
      logger.error('[OAuthManager] Error exchanging code for tokens:', error);
      throw new Error('Failed to exchange authorization code for tokens');
    }
  }

  /**
   * Refresh access token using refresh token
   */
  async refreshAccessToken(refreshToken: string): Promise<OAuthTokens> {
    try {
      logger.info('[OAuthManager] Refreshing access token');

      const { clientId, clientSecret, tenantId } = config.microsoft;

      const params = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
      });

      const response = await fetch(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
          body: params.toString(),
        }
      );

      if (!response.ok) {
        const error = await response.text();
        logger.error('[OAuthManager] Token refresh failed:', error);
        throw new Error('Failed to refresh access token');
      }

      const data = await response.json() as {
        access_token: string;
        refresh_token?: string;
        expires_in?: number;
        token_type?: string;
        scope?: string;
        id_token?: string;
      };

      logger.info('[OAuthManager] Successfully refreshed access token');

      return {
        access_token: data.access_token,
        refresh_token: data.refresh_token || refreshToken, // Keep old refresh token if new one not provided
        expires_in: data.expires_in || 3600,
        token_type: data.token_type || 'Bearer',
        scope: data.scope || undefined,
        id_token: data.id_token || undefined,
      };
    } catch (error) {
      logger.error('[OAuthManager] Error refreshing access token:', error);
      throw new Error('Failed to refresh access token');
    }
  }

  /**
   * Store session data
   */
  storeSession(userId: string, tokens: OAuthTokens): void {
    const expiresAt = tokens.expires_in
      ? new Date(Date.now() + tokens.expires_in * 1000)
      : undefined;

    const sessionData: SessionData = {
      userId,
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token,
      expiresAt,
      createdAt: new Date(),
    };

    this.sessions.set(userId, sessionData);
    logger.info('[OAuthManager] Stored session for user', { userId });
  }

  /**
   * Get session data
   */
  getSession(userId: string): SessionData | undefined {
    return this.sessions.get(userId);
  }

  /**
   * Check if session is valid (not expired)
   */
  isSessionValid(userId: string): boolean {
    const session = this.sessions.get(userId);
    if (!session) {
      return false;
    }

    if (!session.expiresAt) {
      return true; // No expiration set
    }

    return session.expiresAt > new Date();
  }

  /**
   * Get valid access token (refresh if needed)
   */
  async getValidAccessToken(userId: string): Promise<string | null> {
    const session = this.sessions.get(userId);
    if (!session) {
      logger.warn('[OAuthManager] No session found for user', { userId });
      return null;
    }

    // Check if token is still valid
    if (this.isSessionValid(userId)) {
      return session.accessToken;
    }

    // Token expired, try to refresh
    if (!session.refreshToken) {
      logger.warn('[OAuthManager] No refresh token available for user', { userId });
      return null;
    }

    try {
      const newTokens = await this.refreshAccessToken(session.refreshToken);
      this.storeSession(userId, newTokens);
      return newTokens.access_token;
    } catch (error) {
      logger.error('[OAuthManager] Failed to refresh token for user', { userId, error });
      return null;
    }
  }

  /**
   * Remove session
   */
  removeSession(userId: string): void {
    this.sessions.delete(userId);
    logger.info('[OAuthManager] Removed session for user', { userId });
  }

  /**
   * Get all active sessions
   */
  getActiveSessions(): string[] {
    return Array.from(this.sessions.keys());
  }
}

// Singleton instance
export const oauthManager = new OAuthManager();







