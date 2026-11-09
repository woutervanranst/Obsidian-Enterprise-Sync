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
  /** Use "common" for Consumer OneDrive + Multitenant Entra ID, or your Enterprise Tenant UUID */
  tenantId: string;
  clientId: string;
  /** Set true only when enterprise SharePoint sites discovery (/sites) is enabled */
  includeSitesScope: boolean;
}

export interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message: string;
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

  getScopes(): string {
    const scopes = ["openid", "profile", "offline_access", "User.Read", "Files.ReadWrite.All"];
    if (this.config.includeSitesScope) {
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

    return `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/authorize?${params.toString()}`;
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

    const tokenUrl = `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`;
    this.logger.info(`Exchanging PKCE authorization code at ${tokenUrl}`);

    const res = await requestUrl({
      url: tokenUrl,
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
      throw: false,
    });

    this.codeVerifier = null;
    if (res.status !== 200 || !res.json?.access_token) {
      const errDetail =
        res.json?.error_description || res.json?.error || res.text || `HTTP ${res.status}`;
      this.logger.error(`PKCE token exchange failed (${res.status})`, errDetail);
      throw new Error(`PKCE token exchange failed (${res.status}): ${errDetail}`);
    }

    this.logger.info("PKCE token exchange succeeded", {
      hasRefreshToken: Boolean(res.json.refresh_token),
      expiresInSec: res.json.expires_in,
    });
    return this.enrichWithProfile({
      accessToken: res.json.access_token,
      refreshToken: res.json.refresh_token ?? "",
      expiresAt: Date.now() + (res.json.expires_in ?? 3600) * 1000 - 60_000,
    });
  }

  /**
   * OAuth2 Device Code Flow — ideal for enterprise desktops and containerized Obsidian instances.
   */
  async startDeviceCode(): Promise<DeviceCodeResponse> {
    const body = new URLSearchParams({
      client_id: this.config.clientId,
      scope: this.getScopes(),
    });

    const res = await requestUrl({
      url: `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/devicecode`,
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    return res.json as DeviceCodeResponse;
  }

  async pollDeviceCode(deviceCode: string): Promise<AuthTokens | null> {
    const body = new URLSearchParams({
      client_id: this.config.clientId,
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      device_code: deviceCode,
    });

    const res = await requestUrl({
      url: `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`,
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
      throw: false,
    });

    if (res.status === 200 && res.json?.access_token) {
      return this.enrichWithProfile({
        accessToken: res.json.access_token,
        refreshToken: res.json.refresh_token ?? "",
        expiresAt: Date.now() + (res.json.expires_in ?? 3600) * 1000 - 60_000,
      });
    }

    if (res.json?.error === "authorization_pending" || res.json?.error === "slow_down") {
      return null;
    }

    throw new Error(res.json?.error_description || res.json?.error || "Device code login failed");
  }

  /**
   * Direct Entra ID organizational account sign-in (OAuth2 ROPC flow).
   * Note: Microsoft requires /organizations or /{tenantId} for direct password grant.
   */
  async signInWithCredentials(username: string, password: string): Promise<AuthTokens> {
    const authorityTenant =
      this.config.tenantId === "common" || this.config.tenantId === "consumers"
        ? "organizations"
        : this.config.tenantId;

    const body = new URLSearchParams({
      client_id: this.config.clientId,
      grant_type: "password",
      username: username.trim(),
      password,
      scope: this.getScopes(),
    });

    const res = await requestUrl({
      url: `https://login.microsoftonline.com/${authorityTenant}/oauth2/v2.0/token`,
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
      throw: false,
    });

    if (res.status !== 200 || !res.json?.access_token) {
      throw new Error(
        res.json?.error_description || res.json?.error || `Entra sign-in failed (${res.status})`
      );
    }

    return this.enrichWithProfile({
      accessToken: res.json.access_token,
      refreshToken: res.json.refresh_token ?? "",
      expiresAt: Date.now() + (res.json.expires_in ?? 3600) * 1000 - 60_000,
    });
  }

  async enrichWithProfile(tokens: AuthTokens): Promise<AuthTokens> {
    try {
      const meRes = await requestUrl({
        url: "https://graph.microsoft.com/v1.0/me",
        headers: { Authorization: `Bearer ${tokens.accessToken}` },
        throw: false,
      });
      if (meRes.status === 200 && meRes.json) {
        tokens.userDisplayName = meRes.json.displayName;
        tokens.userEmail = meRes.json.mail || meRes.json.userPrincipalName;
        tokens.userId = meRes.json.id;
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
          const decoded = JSON.parse(atob(padded));
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
    const tokenUrl = `https://login.microsoftonline.com/${authorityTenant}/oauth2/v2.0/token`;
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

    if (res.status !== 200 || !res.json?.access_token) {
      const errMsg =
        res.json?.error_description ||
        res.json?.error ||
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
      accessToken: res.json.access_token,
      refreshToken: res.json.refresh_token ?? tokens.refreshToken,
      expiresAt: Date.now() + (res.json.expires_in ?? 3600) * 1000 - 60_000,
    };
    this.logger.info("Access token refreshed successfully", {
      expiresAt: new Date(updated.expiresAt).toISOString(),
    });
    return updated;
  }
}
