import crypto from "crypto";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const AUTH_BASE = "https://login.microsoftonline.com/common/oauth2/v2.0";

const SCOPES = [
  "Mail.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "User.Read",
  "offline_access",
];

function getEncryptionKey(): string {
  const key = process.env.SESSION_SECRET;
  if (!key) throw new Error("SESSION_SECRET environment variable is required");
  return key;
}

function getRedirectUri(): string {
  return process.env.MS_REDIRECT_URI || "http://localhost:3000/api/auth/callback";
}

export function getAuthUrl(): string {
  const params = new URLSearchParams({
    client_id: process.env.MS_CLIENT_ID || "",
    response_type: "code",
    redirect_uri: getRedirectUri(),
    scope: SCOPES.join(" "),
    response_mode: "query",
    prompt: "consent",
  });
  return `${AUTH_BASE}/authorize?${params.toString()}`;
}

export async function exchangeCode(code: string) {
  const body = new URLSearchParams({
    client_id: process.env.MS_CLIENT_ID || "",
    client_secret: process.env.MS_CLIENT_SECRET || "",
    code,
    redirect_uri: getRedirectUri(),
    grant_type: "authorization_code",
    scope: SCOPES.join(" "),
  });

  const res = await fetch(`${AUTH_BASE}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Token exchange failed: ${err}`);
  }

  return res.json();
}

async function refreshAccessToken(refreshToken: string) {
  const body = new URLSearchParams({
    client_id: process.env.MS_CLIENT_ID || "",
    client_secret: process.env.MS_CLIENT_SECRET || "",
    refresh_token: refreshToken,
    grant_type: "refresh_token",
    scope: SCOPES.join(" "),
  });

  const res = await fetch(`${AUTH_BASE}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!res.ok) {
    throw new Error("Token refresh failed");
  }

  return res.json();
}

export function encryptTokens(tokens: any): string {
  const key = crypto.scryptSync(getEncryptionKey(), "salt", 32);
  const iv = crypto.randomBytes(16);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  let encrypted = cipher.update(JSON.stringify(tokens), "utf8", "base64");
  encrypted += cipher.final("base64");
  const tag = cipher.getAuthTag();
  return `${iv.toString("base64")}:${tag.toString("base64")}:${encrypted}`;
}

export function decryptTokens(encrypted: string): any {
  const key = crypto.scryptSync(getEncryptionKey(), "salt", 32);
  const [ivB64, tagB64, data] = encrypted.split(":");
  const iv = Buffer.from(ivB64, "base64");
  const tag = Buffer.from(tagB64, "base64");
  const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
  decipher.setAuthTag(tag);
  let decrypted = decipher.update(data, "base64", "utf8");
  decrypted += decipher.final("utf8");
  return JSON.parse(decrypted);
}

async function getAccessToken(tokens: any): Promise<{ accessToken: string; updatedTokens: any | null }> {
  // Check if token is expired (with 5 min buffer)
  const expiresAt = tokens.expires_at || 0;
  const now = Math.floor(Date.now() / 1000);

  if (now < expiresAt - 300) {
    return { accessToken: tokens.access_token, updatedTokens: null };
  }

  // Token expired or about to expire — refresh it
  if (!tokens.refresh_token) {
    throw new Error("No refresh token available");
  }

  const refreshed = await refreshAccessToken(tokens.refresh_token);
  const updatedTokens = {
    access_token: refreshed.access_token,
    refresh_token: refreshed.refresh_token || tokens.refresh_token,
    expires_at: now + (refreshed.expires_in || 3600),
  };

  return { accessToken: updatedTokens.access_token, updatedTokens };
}

async function graphFetch(tokens: any, path: string, options: RequestInit = {}) {
  const { accessToken } = await getAccessToken(tokens);
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...options.headers,
    },
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Graph API error (${res.status}): ${err}`);
  }

  // Some endpoints return 204 No Content or 202 Accepted with empty body
  const text = await res.text();
  if (!text) return {};
  return JSON.parse(text);
}

export interface EmailSummary {
  id: string;
  conversationId: string;
  from: string;
  subject: string;
  snippet: string;
  date: string;
}

export async function getUnreadEmails(
  tokens: any,
  maxResults = 10
): Promise<EmailSummary[]> {
  const data = await graphFetch(
    tokens,
    `/me/mailFolders/inbox/messages?$filter=isRead eq false&$top=${maxResults}&$orderby=receivedDateTime desc&$select=id,conversationId,from,subject,bodyPreview,receivedDateTime`
  );

  if (!data.value) return [];

  return data.value.map((msg: any) => ({
    id: msg.id,
    conversationId: msg.conversationId,
    from: msg.from?.emailAddress
      ? `${msg.from.emailAddress.name} <${msg.from.emailAddress.address}>`
      : "",
    subject: msg.subject || "",
    snippet: msg.bodyPreview || "",
    date: msg.receivedDateTime || "",
  }));
}

export async function getEmailBody(
  tokens: any,
  messageId: string
): Promise<string> {
  const msg = await graphFetch(
    tokens,
    `/me/messages/${messageId}?$select=body`
  );

  if (!msg.body) return "";

  // If HTML, strip tags for a text representation
  if (msg.body.contentType === "html") {
    return msg.body.content
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
      .replace(/<[^>]+>/g, " ")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/\s+/g, " ")
      .trim();
  }

  return msg.body.content || "";
}

export async function sendReply(
  tokens: any,
  messageId: string,
  body: string
): Promise<void> {
  await graphFetch(tokens, `/me/messages/${messageId}/reply`, {
    method: "POST",
    body: JSON.stringify({
      comment: body,
    }),
  });
}

export async function deleteEmail(
  tokens: any,
  messageId: string
): Promise<void> {
  await graphFetch(tokens, `/me/messages/${messageId}`, {
    method: "DELETE",
  });
}

export async function markAsRead(
  tokens: any,
  messageId: string
): Promise<void> {
  await graphFetch(tokens, `/me/messages/${messageId}`, {
    method: "PATCH",
    body: JSON.stringify({
      isRead: true,
    }),
  });
}

export async function getUserEmail(tokens: any): Promise<string> {
  const profile = await graphFetch(tokens, "/me?$select=mail,userPrincipalName");
  return profile.mail || profile.userPrincipalName || "unknown";
}
