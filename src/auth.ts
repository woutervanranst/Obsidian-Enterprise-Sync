import { requestUrl } from "obsidian";
import { SyncLogger, defaultLogger } from "./logger";

export interface AuthTokens {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
  userEmail?: string;
  userDisplayName?: string;
  userId?: string;
}

export interface AuthConfig {
  /** Use "common" for Consumer OneDrive, or your Enterprise Directory (tenant) ID */
  tenantId: string;
  clientId: string;
  /** Automatically true for Enterprise (SharePoint & Microsoft Teams), false for Consumer */
  includeSitesScope: boolean;
  /** Optional Microsoft Graph base URL (defaults to https://graph.microsoft.com/v1.0) */
  graphBaseUrl?: string;
  /** Optional OAuth authority base URL for sovereign clouds or tests (defaults to https://login.microsoftonline.com) */
  authorityBaseUrl?: string;
}

interface OAuthTokenResponse {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
}

interface GraphMeResponse {
  id?: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

interface JwtClaimsPayload {
  name?: string;
  upn?: string;
  preferred_username?: string;
  email?: string;
  oid?: string;
  sub?: string;
}

const REDIRECT_URI = "obsidian://entra-sync-auth";

function base64UrlEncode(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

export class EntraPkceAuth {
  private codeVerifier: string | null = null;

  constructor(
    private config: AuthConfig,
    private logger: SyncLogger = defaultLogger
  ) {}

  private authorityBase(): string {
    return (this.config.authorityBaseUrl || "https://login.microsoftonline.com").replace(
      /\/+$/,
      ""
    );
  }

  private graphBase(): string {
    return (this.config.graphBaseUrl || "https://graph.microsoft.com/v1.0").replace(
      /\/+$/,
      ""
    );
  }

  isSharePointEnabled(): boolean {
    const tenant = (this.config.tenantId || "common").trim().toLowerCase();
    return Boolean(
      this.config.includeSitesScope ||
        (tenant && tenant !== "common" && tenant !== "consumers")
    );
  }

  getScopes(): string {
    const scopes = ["openid", "profile", "offline_access", "User.Read", "Files.ReadWrite.All"];
    if (this.isSharePointEnabled()) {
      scopes.push("Sites.ReadWrite.All");
    }
    return scopes.join(" ");
  }

  async startLogin(): Promise<string> {
    const randomBytes = crypto.getRandomValues(new Uint8Array(32));
    this.codeVerifier = base64UrlEncode(randomBytes.buffer);
    const digest = await crypto.subtle.digest(
      "SHA-256",
      new TextEncoder().encode(this.codeVerifier)
    );
    const codeChallenge = base64UrlEncode(digest);

    const params = new URLSearchParams({
      client_id: this.config.clientId,
      response_type: "code",
      redirect_uri: REDIRECT_URI,
      response_mode: "query",
      scope: this.getScopes(),
      code_challenge: codeChallenge,
      code_challenge_method: "S256",
    });

    const tenant = (this.config.tenantId || "common").trim();
    return `${this.authorityBase()}/${tenant}/oauth2/v2.0/authorize?${params.toString()}`;
  }

  async handleCallback(code: string): Promise<AuthTokens> {
    if (!this.codeVerifier) {
      throw new Error("Missing PKCE code_verifier. Restart sign-in.");
    }

    const body = new URLSearchParams({
      client_id: this.config.clientId,
      grant_type: "authorization_code",
      code,
      redirect_uri: REDIRECT_URI,
      code_verifier: this.codeVerifier,
      scope: this.getScopes(),
    });

    const tenant = (this.config.tenantId || "common").trim();
    const tokenUrl = `${this.authorityBase()}/${tenant}/oauth2/v2.0/token`;
    this.logger.info(`Exchanging PKCE authorization code at ${tokenUrl}`);

    const res = await requestUrl({
      url: tokenUrl,
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
      throw: false,
    });

    this.codeVerifier = null;
    const tokenJson = res.json as OAuthTokenResponse | undefined;
    if (res.status !== 200 || !tokenJson?.access_token) {
      const errDetail =
        tokenJson?.error_description || tokenJson?.error || res.text || `HTTP ${res.status}`;
      this.logger.error(`PKCE token exchange failed (${res.status})`, errDetail);
      throw new Error(`PKCE token exchange failed (${res.status}): ${errDetail}`);
    }

    this.logger.info("PKCE token exchange succeeded", {
      hasRefreshToken: Boolean(tokenJson.refresh_token),
      expiresInSec: tokenJson.expires_in,
    });
    return this.enrichWithProfile({
      accessToken: tokenJson.access_token,
      refreshToken: tokenJson.refresh_token ?? "",
      expiresAt: Date.now() + (tokenJson.expires_in ?? 3600) * 1000 - 60_000,
    });
  }

  async enrichWithProfile(tokens: AuthTokens): Promise<AuthTokens> {
    try {
      const meRes = await requestUrl({
        url: `${this.graphBase()}/me`,
        headers: { Authorization: `Bearer ${tokens.accessToken}` },
        throw: false,
      });
      const meJson = meRes.json as GraphMeResponse | undefined;
      if (meRes.status === 200 && meJson) {
        tokens.userDisplayName = meJson.displayName;
        tokens.userEmail = meJson.mail || meJson.userPrincipalName;
        tokens.userId = meJson.id;
        this.logger.info("Loaded user profile from /v1.0/me", {
          userDisplayName: tokens.userDisplayName,
          userEmail: tokens.userEmail,
        });
      } else {
        this.logger.warn(`GET /v1.0/me returned status ${meRes.status}`, meRes.text);
      }
    } catch (err) {
      this.logger.warn("Failed to fetch /v1.0/me profile, falling back to JWT claims", String(err));
    }
    if (!tokens.userEmail || !tokens.userDisplayName) {
      try {
        const payloadPart = tokens.accessToken.split(".")[1];
        if (payloadPart) {
          const padded = payloadPart.replace(/-/g, "+").replace(/_/g, "/");
          const decoded = JSON.parse(atob(padded)) as JwtClaimsPayload;
          tokens.userDisplayName = tokens.userDisplayName || decoded.name;
          tokens.userEmail =
            tokens.userEmail || decoded.upn || decoded.preferred_username || decoded.email;
          tokens.userId = tokens.userId || decoded.oid || decoded.sub;
        }
      } catch {
        // Ignore JWT parse errors
      }
    }
    return tokens;
  }

  async refresh(tokens: AuthTokens, force = false): Promise<AuthTokens> {
    if (!force && Date.now() < tokens.expiresAt) {
      return tokens;
    }

    if (!tokens.refreshToken) {
      this.logger.warn(
        "Token refresh requested but no refreshToken is present. Please sign in again."
      );
      if (force) {
        throw new Error("No refresh token available. Please click 'Sign in' again.");
      }
      return tokens;
    }

    // Keep configured tenantId ("common" works for both Consumer OneDrive MSA & Entra ID refresh tokens!)
    const authorityTenant = (this.config.tenantId || "common").trim();
    const tokenUrl = `${this.authorityBase()}/${authorityTenant}/oauth2/v2.0/token`;
    this.logger.info(`Refreshing OAuth2 access token via ${tokenUrl}`, { force });

    const body = new URLSearchParams({
      client_id: this.config.clientId,
      grant_type: "refresh_token",
      refresh_token: tokens.refreshToken,
      scope: this.getScopes(),
    });

    const res = await requestUrl({
      url: tokenUrl,
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
      throw: false,
    });

    const tokenJson = res.json as OAuthTokenResponse | undefined;
    if (res.status !== 200 || !tokenJson?.access_token) {
      const errMsg =
        tokenJson?.error_description ||
        tokenJson?.error ||
        res.text ||
        `HTTP ${res.status}`;
      this.logger.error(`Token refresh failed (${res.status}) at ${tokenUrl}`, errMsg);
      if (force || Date.now() >= tokens.expiresAt) {
        throw new Error(`Token refresh failed (${res.status}): ${errMsg}`);
      }
      return tokens;
    }

    const updated: AuthTokens = {
      ...tokens,
      accessToken: tokenJson.access_token,
      refreshToken: tokenJson.refresh_token ?? tokens.refreshToken,
      expiresAt: Date.now() + (tokenJson.expires_in ?? 3600) * 1000 - 60_000,
    };
    this.logger.info("Access token refreshed successfully", {
      expiresAt: new Date(updated.expiresAt).toISOString(),
    });
    return updated;
  }
}
