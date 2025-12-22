#!/usr/bin/env bun
/**
 * Mailgrep - Office 365 Email Attachment Downloader
 *
 * A CLI tool to download attachments from Office 365 emails with:
 * - OAuth2 PKCE authentication (browser-based, no client secret needed)
 * - Configurable file types (images by default, or pdf, docx, zip, etc.)
 * - Incremental backup via manifest.json (skips already-downloaded files)
 * - Multi-account support (tracks by userEmail + senderEmail combo)
 * - Adaptive rate limiting for large mailboxes
 * - Graceful Ctrl+C handling (saves progress before exit)
 *
 * Setup:
 *   1. Create Azure AD app at https://portal.azure.com → App registrations
 *   2. Set redirect URI to "http://localhost:8400" (Mobile/Desktop platform)
 *   3. Enable "Allow public client flows"
 *   4. Create .env with AZURE_CLIENT_ID and AZURE_TENANT_ID
 *   5. Optionally configure in .env:
 *      - DEFAULT_SENDER: Default sender email for filtering
 *      - DEFAULT_OUTPUT_DIR: Output directory path
 *      - FILE_TYPES: Comma-separated extensions (e.g., "png,jpg,pdf") or "*" for all
 *
 * Usage:
 *   mailgrep                         # Interactive mode
 *   mailgrep --dry-run               # Preview without downloading
 *   mailgrep --no-cache              # Force re-download all
 *   mailgrep --show-accounts         # List cached accounts
 *   mailgrep -e sender@example.com   # Non-interactive with sender
 *
 * Repository: https://github.com/keskinonur/mailgrep
 */

import { join } from "path";
import { homedir } from "os";
import * as readline from "readline";
import { Command } from "commander";
import chalk from "chalk";
import ora, { type Ora } from "ora";
import cliProgress from "cli-progress";

// ============================================
// Constants
// ============================================
const VERSION = "1.0.8";
const OAUTH_PORT = 8400;              // Local server port for OAuth redirect
const PAGE_SIZE = 50;                 // Emails per Graph API page (max 50)

// Defaults from environment (optional - prompts if not set)
const DEFAULT_SENDER = process.env.DEFAULT_SENDER || "";
const DEFAULT_OUTPUT = process.env.DEFAULT_OUTPUT_DIR || "";

// File type filter - comma-separated extensions or "*" for all
// Default: common image formats
const DEFAULT_FILE_TYPES = "png,jpg,jpeg,gif,bmp,webp,tiff";
const FILE_TYPES_RAW = process.env.FILE_TYPES?.trim() || DEFAULT_FILE_TYPES;
const ALLOWED_EXTENSIONS = FILE_TYPES_RAW === "*"
  ? null  // null means allow all
  : FILE_TYPES_RAW.split(",").map(ext => `.${ext.trim().toLowerCase()}`);

// Rate limiting - adaptive delay for large mailboxes
// Kicks in when email count >= threshold to avoid hitting Graph API limits
// Graph API limit: ~10,000 requests per 10 minutes per app
const LARGE_MAILBOX_THRESHOLD = 500;  // emails
const DELAY_MS = 50;                  // ms between attachment downloads

// ============================================
// Exit Codes
// ============================================
enum ExitCode {
  Success = 0,
  ConfigError = 1,
  AuthError = 2,
  NetworkError = 3,
  FileSystemError = 4,
}

// ============================================
// ASCII Banner
// ============================================
function showBanner(): void {
  const banner = `
  ${chalk.cyan("┌─────────────────────────────────────────┐")}
  ${chalk.cyan("│")}                                         ${chalk.cyan("│")}
  ${chalk.cyan("│")}   ${chalk.bold.white("╔╦╗╔═╗╦╦  ╔═╗╦═╗╔═╗╔═╗")}                ${chalk.cyan("│")}
  ${chalk.cyan("│")}   ${chalk.bold.white("║║║╠═╣║║  ║ ╦╠╦╝║╣ ╠═╝")}                ${chalk.cyan("│")}
  ${chalk.cyan("│")}   ${chalk.bold.white("╩ ╩╩ ╩╩╩═╝╚═╝╩╚═╚═╝╩")}                  ${chalk.cyan("│")}
  ${chalk.cyan("│")}                                         ${chalk.cyan("│")}
  ${chalk.cyan("│")}   ${chalk.dim("Like grep, but for your mailbox.")}      ${chalk.cyan("│")}
  ${chalk.cyan("│")}   ${chalk.dim(`v${VERSION}`)}                                ${chalk.cyan("│")}
  ${chalk.cyan("│")}                                         ${chalk.cyan("│")}
  ${chalk.cyan("└─────────────────────────────────────────┘")}
`;
  console.log(banner);
}

// ============================================
// Types
// ============================================
interface Attachment {
  id: string;
  name: string;
  contentType: string;
  contentBytes?: string;
  size: number;
  isInline: boolean;
}

interface EmailAddress {
  address: string;
  name: string;
}

interface Message {
  id: string;
  subject: string;
  receivedDateTime: string;
  from: {
    emailAddress: EmailAddress;
  };
  hasAttachments: boolean;
  attachments?: Attachment[];
}

interface GraphResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

interface Config {
  senderEmail: string;
  startDate: string;
  endDate: string;
  outputDir: string;
}

interface CLIOptions {
  email?: string;
  start?: string;
  end?: string;
  output?: string;
  user?: string;  // --user to select cached account
  dryRun: boolean;
  verbose: boolean;
  quiet: boolean;
  cache: boolean;  // --no-cache sets this to false
  showAccounts: boolean;
  checkDuplicates: boolean;
  rebuildHashes: boolean;
  dedupe: boolean;
  logout: boolean;
  reauth: boolean;
}

// ============================================
// Manifest Types - Incremental Backup System
// ============================================
// The manifest tracks downloaded images to enable incremental backups.
// Key design: messageId|attachmentId composite key ensures uniqueness
// even if the same image is attached to multiple emails.

interface ManifestEntry {
  key: string;              // Composite key: "messageId|attachmentId"
  filename: string;         // Saved filename (e.g., "2025-01-15_1_photo.jpg")
  originalName: string;     // Original attachment name
  size: number;             // File size in bytes
  hash: string;             // SHA256 hash for duplicate detection
  emailSubject: string;     // For reference/debugging
  emailDate: string;        // ISO date of the email
  downloadedAt: string;     // ISO date when downloaded
}

interface UserSenderManifest {
  userEmail: string;        // Authenticated user (from OAuth token)
  senderEmail: string;      // Email sender being tracked
  lastSync: string;         // Last sync timestamp
  entries: ManifestEntry[]; // All downloaded images for this user+sender
  processedEmailIds: string[];  // Email IDs fully processed (skip API call on repeat runs)
}

interface Manifest {
  version: number;          // Schema version for future migrations
  updatedAt: string;        // Last manifest update
  accounts: UserSenderManifest[];  // Supports multiple user+sender combos
}

// ============================================
// Token Cache Types - Persistent Authentication
// ============================================
// Stores OAuth tokens to avoid re-authentication on every run
// Tokens are cached per-tenant in ~/.mailgrep/tokens.json

interface TokenCache {
  tenantId: string;
  accessToken: string;
  refreshToken: string;
  expiresAt: number;        // Unix timestamp (ms) when access token expires
  userEmail: string;
  cachedAt: string;         // ISO date when cached
}

interface TokenCacheFile {
  version: number;
  tokens: TokenCache[];     // Multiple accounts supported
}

// ============================================
// OAuth Configuration
// ============================================
// Uses OAuth2 Authorization Code flow with PKCE (Proof Key for Code Exchange)
// PKCE allows public clients (like CLI tools) to authenticate without a client secret
// Flow: Browser login → redirect to localhost → exchange code for token
const OAUTH_CONFIG = {
  clientId: process.env.AZURE_CLIENT_ID || "",
  tenantId: process.env.AZURE_TENANT_ID || "organizations",  // "organizations" = any Azure AD tenant
  redirectUri: `http://localhost:${OAUTH_PORT}`,
  scopes: ["https://graph.microsoft.com/Mail.Read", "offline_access"],  // Mail.Read for email access
  get authorizeUrl() {
    return `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/authorize`;
  },
  get tokenUrl() {
    return `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
  },
};

// ============================================
// Logger
// ============================================
class Logger {
  constructor(private verbose: boolean, private quiet: boolean) {}

  info(...args: unknown[]) {
    if (!this.quiet) console.log(chalk.blue("info"), ...args);
  }

  success(...args: unknown[]) {
    if (!this.quiet) console.log(chalk.green("✓"), ...args);
  }

  warn(...args: unknown[]) {
    console.log(chalk.yellow("warn"), ...args);
  }

  error(...args: unknown[]) {
    console.error(chalk.red("error"), ...args);
  }

  debug(...args: unknown[]) {
    if (this.verbose) console.log(chalk.gray("debug"), ...args);
  }

  dim(...args: unknown[]) {
    if (!this.quiet) console.log(chalk.dim(...args.map(String)));
  }
}

let logger: Logger;

// ============================================
// CLI Input Helper
// ============================================
function createPrompt(): readline.Interface {
  return readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
}

async function ask(rl: readline.Interface, question: string, defaultValue: string): Promise<string> {
  return new Promise((resolve) => {
    rl.question(chalk.cyan(`${question}`) + chalk.dim(` [${defaultValue}]: `), (answer) => {
      resolve(answer.trim() || defaultValue);
    });
  });
}

async function askRequired(rl: readline.Interface, question: string): Promise<string> {
  return new Promise((resolve) => {
    const prompt = () => {
      rl.question(chalk.cyan(`${question}: `), (answer) => {
        const value = answer.trim();
        if (value) {
          resolve(value);
        } else {
          console.log(chalk.red("  This field is required"));
          prompt();
        }
      });
    };
    prompt();
  });
}

// ============================================
// Date Helpers
// ============================================
function getDefaultDates(): { startDate: string; endDate: string } {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");

  return {
    startDate: `${year}-01-01`,
    endDate: `${year}-${month}-${day}`,
  };
}

function getTimestampFolder(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  return `${year}${month}${day}_${hours}${minutes}${seconds}`;
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDuration(ms: number): string {
  const seconds = Math.floor(ms / 1000);
  const minutes = Math.floor(seconds / 60);
  const hours = Math.floor(minutes / 60);

  if (hours > 0) {
    return `${hours}h ${minutes % 60}m ${seconds % 60}s`;
  }
  if (minutes > 0) {
    return `${minutes}m ${seconds % 60}s`;
  }
  return `${seconds}s`;
}

function formatTime(date: Date): string {
  return date.toLocaleTimeString("en-US", { hour12: false });
}

// Create clickable terminal hyperlink (OSC 8)
function terminalLink(text: string, url: string): string {
  return `\x1b]8;;${url}\x07${text}\x1b]8;;\x07`;
}

function folderLink(path: string): string {
  const fileUrl = `file://${path}`;
  return terminalLink(path, fileUrl);
}

// Cross-platform helpers
async function openBrowser(url: string): Promise<void> {
  const platform = process.platform;
  try {
    if (platform === "darwin") {
      await Bun.$`open ${url}`.quiet();
    } else if (platform === "win32") {
      await Bun.$`cmd /c start ${url}`.quiet();
    } else {
      // Linux and others
      await Bun.$`xdg-open ${url}`.quiet();
    }
  } catch {
    // Silently fail - we'll show URL in console
  }
}

async function deleteFile(filepath: string): Promise<boolean> {
  try {
    const { unlink } = await import("fs/promises");
    await unlink(filepath);
    return true;
  } catch {
    return false;
  }
}

// ============================================
// PKCE Helpers
// ============================================
// PKCE (Proof Key for Code Exchange) prevents authorization code interception attacks
// 1. Generate random code_verifier (kept secret)
// 2. Create code_challenge = SHA256(code_verifier) in base64url
// 3. Send code_challenge with auth request
// 4. Send code_verifier with token request (server verifies hash matches)

function generateCodeVerifier(): string {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  return base64UrlEncode(array);
}

function generateCodeChallenge(verifier: string): string {
  const encoder = new TextEncoder();
  const data = encoder.encode(verifier);
  const hash = new Bun.CryptoHasher("sha256").update(data).digest();
  return base64UrlEncode(new Uint8Array(hash));
}

// Base64URL encoding (RFC 4648) - URL-safe variant of base64
function base64UrlEncode(buffer: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < buffer.length; i++) {
    binary += String.fromCharCode(buffer[i]);
  }
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

// ============================================
// JWT Helpers
// ============================================
// Extract user email from JWT access token without validation
// JWT structure: header.payload.signature (base64 encoded)
// We only need the payload to get the user's email for manifest tracking

function getUserEmailFromToken(accessToken: string): string {
  try {
    const parts = accessToken.split(".");
    if (parts.length !== 3) return "unknown";

    // Decode payload (middle part) - no signature validation needed
    // Token is already validated by Microsoft when we received it
    const payload = JSON.parse(atob(parts[1]));
    return payload.preferred_username || payload.upn || payload.email || "unknown";
  } catch {
    return "unknown";
  }
}

function getTenantIdFromToken(accessToken: string): string | null {
  // Extract actual tenant ID from JWT 'tid' claim
  // Important when AZURE_TENANT_ID is "organizations" (multi-tenant)
  try {
    const parts = accessToken.split(".");
    if (parts.length !== 3) return null;
    const payload = JSON.parse(atob(parts[1]));
    return payload.tid || null;
  } catch {
    return null;
  }
}

// ============================================
// Manifest Helpers
// ============================================
const MANIFEST_VERSION = 1;

async function loadManifest(path: string): Promise<Manifest> {
  try {
    const file = Bun.file(path);
    if (await file.exists()) {
      const content = await file.json();
      return content as Manifest;
    }
  } catch (error) {
    logger?.warn(`Could not load manifest, starting fresh: ${error}`);
  }

  return {
    version: MANIFEST_VERSION,
    updatedAt: new Date().toISOString(),
    accounts: [],
  };
}

async function saveManifest(path: string, manifest: Manifest): Promise<void> {
  manifest.updatedAt = new Date().toISOString();
  await Bun.write(path, JSON.stringify(manifest, null, 2));
}

function getOrCreateUserSenderManifest(
  manifest: Manifest,
  userEmail: string,
  senderEmail: string
): UserSenderManifest {
  let account = manifest.accounts.find(
    (a) => a.userEmail.toLowerCase() === userEmail.toLowerCase() &&
           a.senderEmail.toLowerCase() === senderEmail.toLowerCase()
  );

  if (!account) {
    account = {
      userEmail,
      senderEmail,
      lastSync: new Date().toISOString(),
      entries: [],
      processedEmailIds: [],
    };
    manifest.accounts.push(account);
  }

  // Migration: add processedEmailIds if missing (from older manifest)
  if (!account.processedEmailIds) {
    account.processedEmailIds = [];
  }

  return account;
}

function getDownloadedKeys(account: UserSenderManifest): Set<string> {
  return new Set(account.entries.map((e) => e.key));
}

function getProcessedEmailIds(account: UserSenderManifest): Set<string> {
  return new Set(account.processedEmailIds || []);
}

function createManifestKey(messageId: string, attachmentId: string): string {
  return `${messageId}|${attachmentId}`;
}

function hashBuffer(buffer: Buffer): string {
  return new Bun.CryptoHasher("sha256").update(buffer).digest("hex");
}

interface DuplicateGroup {
  hash: string;
  files: ManifestEntry[];
}

function findDuplicates(entries: ManifestEntry[]): DuplicateGroup[] {
  const hashMap = new Map<string, ManifestEntry[]>();

  for (const entry of entries) {
    if (!entry.hash) continue;
    const existing = hashMap.get(entry.hash) || [];
    existing.push(entry);
    hashMap.set(entry.hash, existing);
  }

  return Array.from(hashMap.entries())
    .filter(([_, files]) => files.length > 1)
    .map(([hash, files]) => ({ hash, files }));
}

// ============================================
// Token Cache Helpers
// ============================================
// Caches OAuth tokens to avoid browser login on every run
// Stored in ~/.mailgrep/tokens.json for security (not in project dir)

const TOKEN_CACHE_DIR = join(homedir(), ".mailgrep");
const TOKEN_CACHE_PATH = join(TOKEN_CACHE_DIR, "tokens.json");
const TOKEN_CACHE_VERSION = 3;  // v3: actual tid from token, not "organizations"

async function loadTokenCache(): Promise<TokenCacheFile> {
  try {
    const file = Bun.file(TOKEN_CACHE_PATH);
    if (await file.exists()) {
      const cache = await file.json() as TokenCacheFile;
      return migrateTokenCache(cache);
    }
  } catch {
    // Ignore errors, return empty cache
  }
  return { version: TOKEN_CACHE_VERSION, tokens: [] };
}

function migrateTokenCache(cache: TokenCacheFile): TokenCacheFile {
  // Migration from v1 to v2: normalize userEmail casing, remove invalid entries
  if (cache.version < 2) {
    cache.tokens = cache.tokens
      .filter(t => t.userEmail && t.userEmail !== "unknown")  // Remove invalid entries
      .map(t => ({
        ...t,
        userEmail: t.userEmail.toLowerCase(),  // Normalize casing
      }));
    cache.version = 2;
  }

  // Migration v2 to v3: replace "organizations" tenantId with actual tid from token
  if (cache.version < 3) {
    cache.tokens = cache.tokens.map(t => {
      if (t.tenantId === "organizations" && t.accessToken) {
        const actualTid = getTenantIdFromToken(t.accessToken);
        if (actualTid) {
          return { ...t, tenantId: actualTid };
        }
      }
      return t;
    });
    cache.version = 3;
  }

  return cache;
}

async function saveTokenCache(cache: TokenCacheFile): Promise<void> {
  await Bun.$`mkdir -p ${TOKEN_CACHE_DIR}`.quiet();
  await Bun.write(TOKEN_CACHE_PATH, JSON.stringify(cache, null, 2));
}

function getCachedTokensForTenant(cache: TokenCacheFile, tenantId: string): TokenCache[] {
  // Get all tokens for this tenant, sorted by most recently cached
  // When tenantId is "organizations" (multi-tenant), return all cached tokens
  const isMultiTenant = tenantId === "organizations";
  return cache.tokens
    .filter(t => isMultiTenant || t.tenantId === tenantId)
    .sort((a, b) => new Date(b.cachedAt).getTime() - new Date(a.cachedAt).getTime());
}

function getCachedToken(cache: TokenCacheFile, tenantId: string, userEmail?: string): TokenCache | undefined {
  const tenantTokens = getCachedTokensForTenant(cache, tenantId);

  if (userEmail) {
    // Explicit user selection via --user flag
    return tenantTokens.find(t => t.userEmail.toLowerCase() === userEmail.toLowerCase());
  }

  // If only one token for tenant, use it; otherwise return undefined to trigger prompt
  return tenantTokens.length === 1 ? tenantTokens[0] : undefined;
}

function setCachedToken(cache: TokenCacheFile, token: TokenCache): void {
  // Key by (tenantId, userEmail) to avoid cross-tenant clobbering
  const normalizedEmail = token.userEmail.toLowerCase();
  const index = cache.tokens.findIndex(
    t => t.tenantId === token.tenantId && t.userEmail.toLowerCase() === normalizedEmail
  );
  if (index >= 0) {
    cache.tokens[index] = { ...token, userEmail: normalizedEmail };
  } else {
    cache.tokens.push({ ...token, userEmail: normalizedEmail });
  }
}

function isTokenExpired(token: TokenCache): boolean {
  // Add 60 second buffer to avoid edge cases
  return Date.now() >= (token.expiresAt - 60000);
}

function getTokenExpiry(accessToken: string): number {
  try {
    const parts = accessToken.split(".");
    if (parts.length !== 3) return Date.now();
    const payload = JSON.parse(atob(parts[1]));
    // JWT exp is in seconds, convert to ms
    return (payload.exp || 0) * 1000;
  } catch {
    return Date.now();
  }
}

async function refreshAccessToken(refreshToken: string): Promise<{ accessToken: string; refreshToken: string; expiresAt: number } | null> {
  try {
    const response = await fetch(OAUTH_CONFIG.tokenUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: OAUTH_CONFIG.clientId,
        grant_type: "refresh_token",
        refresh_token: refreshToken,
        scope: OAUTH_CONFIG.scopes.join(" "),
      }),
    });

    const data = await response.json() as {
      access_token?: string;
      refresh_token?: string;
      error?: string;
    };

    if (data.error || !data.access_token) {
      return null;
    }

    return {
      accessToken: data.access_token,
      refreshToken: data.refresh_token || refreshToken, // May return new refresh token
      expiresAt: getTokenExpiry(data.access_token),
    };
  } catch {
    return null;
  }
}

// ============================================
// Interactive Browser Authentication
// ============================================
interface AuthResult {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
}

async function authenticate(spinner: Ora): Promise<AuthResult> {
  return new Promise((resolve, reject) => {
    // Use crypto.randomUUID for secure state
    const state = crypto.randomUUID();
    const codeVerifier = generateCodeVerifier();
    const codeChallenge = generateCodeChallenge(codeVerifier);

    const authUrl = new URL(OAUTH_CONFIG.authorizeUrl);
    authUrl.searchParams.set("client_id", OAUTH_CONFIG.clientId);
    authUrl.searchParams.set("response_type", "code");
    authUrl.searchParams.set("redirect_uri", OAUTH_CONFIG.redirectUri);
    authUrl.searchParams.set("scope", OAUTH_CONFIG.scopes.join(" "));
    authUrl.searchParams.set("state", state);
    authUrl.searchParams.set("code_challenge", codeChallenge);
    authUrl.searchParams.set("code_challenge_method", "S256");
    authUrl.searchParams.set("prompt", "select_account");

    spinner.text = "Opening browser for login...";

    const server = Bun.serve({
      port: OAUTH_PORT,
      async fetch(req) {
        const url = new URL(req.url);

        if (url.pathname === "/") {
          const code = url.searchParams.get("code");
          const returnedState = url.searchParams.get("state");
          const error = url.searchParams.get("error");
          const errorDescription = url.searchParams.get("error_description");

          if (error) {
            setTimeout(() => {
              server.stop();
              reject(new Error(`OAuth error: ${errorDescription}`));
            }, 100);
            return new Response(errorPage(errorDescription || error), {
              headers: { "Content-Type": "text/html; charset=utf-8" },
            });
          }

          if (!code || returnedState !== state) {
            setTimeout(() => {
              server.stop();
              reject(new Error("Invalid OAuth response - state mismatch"));
            }, 100);
            return new Response(errorPage("Invalid response"), {
              headers: { "Content-Type": "text/html; charset=utf-8" },
            });
          }

          try {
            spinner.text = "Exchanging token...";

            const tokenResponse = await fetch(OAUTH_CONFIG.tokenUrl, {
              method: "POST",
              headers: { "Content-Type": "application/x-www-form-urlencoded" },
              body: new URLSearchParams({
                client_id: OAUTH_CONFIG.clientId,
                grant_type: "authorization_code",
                code: code,
                redirect_uri: OAUTH_CONFIG.redirectUri,
                code_verifier: codeVerifier,
              }),
            });

            const tokenData = (await tokenResponse.json()) as {
              access_token?: string;
              refresh_token?: string;
              error?: string;
              error_description?: string;
            };

            if (tokenData.error || !tokenData.access_token) {
              setTimeout(() => {
                server.stop();
                reject(new Error(`Token error: ${tokenData.error_description}`));
              }, 100);
              return new Response(errorPage(tokenData.error_description || "Token error"), {
                headers: { "Content-Type": "text/html; charset=utf-8" },
              });
            }

            setTimeout(() => {
              server.stop();
              resolve({
                accessToken: tokenData.access_token!,
                refreshToken: tokenData.refresh_token || "",
                expiresAt: getTokenExpiry(tokenData.access_token!),
              });
            }, 100);

            return new Response(successPage(), {
              headers: { "Content-Type": "text/html; charset=utf-8" },
            });
          } catch (err) {
            setTimeout(() => {
              server.stop();
              reject(err);
            }, 100);
            return new Response(errorPage("Authentication failed"), {
              headers: { "Content-Type": "text/html; charset=utf-8" },
            });
          }
        }

        return new Response("Not found", { status: 404 });
      },
    });

    // Open browser (cross-platform)
    openBrowser(authUrl.toString()).catch(() => {
      spinner.warn("Could not open browser automatically");
      console.log(chalk.dim("\nOpen this URL in your browser:\n"));
      console.log(chalk.cyan(authUrl.toString()));
      console.log();
    });
  });
}

function successPage(): string {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Success</title>
  <style>
    body { font-family: -apple-system, system-ui, sans-serif; text-align: center; padding: 50px; background: #0a0a0a; color: #fff; }
    h1 { color: #22c55e; }
    p { color: #a1a1aa; }
  </style>
</head>
<body>
  <h1>Authentication Successful!</h1>
  <p>You can close this window and return to the terminal.</p>
</body>
</html>`;
}

function errorPage(message: string): string {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Error</title>
  <style>
    body { font-family: -apple-system, system-ui, sans-serif; text-align: center; padding: 50px; background: #0a0a0a; color: #fff; }
    h1 { color: #ef4444; }
    p { color: #a1a1aa; }
  </style>
</head>
<body>
  <h1>Authentication Failed</h1>
  <p>${message}</p>
</body>
</html>`;
}

// ============================================
// Graph API Helpers
// ============================================
// Microsoft Graph API is the unified endpoint for Microsoft 365 services
// Docs: https://learn.microsoft.com/en-us/graph/api/overview

async function graphFetch<T>(url: string, accessToken: string): Promise<GraphResponse<T>> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);  // 30s timeout

  try {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      signal: controller.signal,
    });

    clearTimeout(timeout);

    if (!response.ok) {
      // Handle rate limiting (429 Too Many Requests)
      // Graph API returns Retry-After header with seconds to wait
      if (response.status === 429) {
        const retryAfter = response.headers.get("Retry-After");
        logger.warn(`Rate limited. Waiting ${retryAfter || 5}s...`);
        await Bun.sleep((retryAfter ? parseInt(retryAfter) : 5) * 1000);
        return graphFetch<T>(url, accessToken);  // Retry automatically
      }
      const errorText = await response.text();
      throw new Error(`Graph API error: ${response.status} - ${errorText}`);
    }

    return response.json();
  } catch (error) {
    clearTimeout(timeout);
    throw error;
  }
}

async function getAttachment(messageId: string, attachmentId: string, accessToken: string): Promise<Attachment> {
  const url = `https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments/${attachmentId}`;
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch attachment: ${response.status}`);
  }

  return response.json();
}

// ============================================
// File Type Detection
// ============================================
// Checks if attachment matches allowed file types from FILE_TYPES env var
// If ALLOWED_EXTENSIONS is null (FILE_TYPES=*), all files are allowed

// Map content types to extensions for inline image detection
const CONTENT_TYPE_MAP: Record<string, string> = {
  "image/png": ".png",
  "image/jpeg": ".jpg",
  "image/jpg": ".jpg",
  "image/gif": ".gif",
  "image/bmp": ".bmp",
  "image/webp": ".webp",
  "image/tiff": ".tiff",
};

function isAllowedFile(attachment: Attachment): boolean {
  // Allow all files if FILE_TYPES=*
  if (ALLOWED_EXTENSIONS === null) {
    return true;
  }

  const filename = attachment.name.toLowerCase();

  // Check by filename extension first
  if (ALLOWED_EXTENSIONS.some((ext) => filename.endsWith(ext))) {
    return true;
  }

  // Also check contentType for inline images (often have no extension in name)
  // e.g., "image001" with contentType "image/png"
  const contentType = attachment.contentType?.toLowerCase() || "";
  const mappedExt = CONTENT_TYPE_MAP[contentType];
  if (mappedExt && ALLOWED_EXTENSIONS.includes(mappedExt)) {
    return true;
  }

  return false;
}

// ============================================
// Configuration
// ============================================
function isValidDate(dateStr: string): boolean {
  // Check format YYYY-MM-DD
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    return false;
  }
  // Check if date is actually valid (e.g., not 2025-02-30)
  const date = new Date(dateStr + "T00:00:00Z");
  if (isNaN(date.getTime())) {
    return false;
  }
  // Verify the parsed date matches input (catches invalid dates like 02-30)
  const [year, month, day] = dateStr.split("-").map(Number);
  return date.getUTCFullYear() === year &&
         date.getUTCMonth() + 1 === month &&
         date.getUTCDate() === day;
}

class DateValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "DateValidationError";
  }
}

function validateDates(startDate: string, endDate: string): void {
  if (!isValidDate(startDate)) {
    throw new DateValidationError(`Invalid start date: "${startDate}". Use YYYY-MM-DD format.`);
  }
  if (!isValidDate(endDate)) {
    throw new DateValidationError(`Invalid end date: "${endDate}". Use YYYY-MM-DD format.`);
  }
  if (startDate > endDate) {
    throw new DateValidationError(`Start date (${startDate}) cannot be after end date (${endDate}).`);
  }
}

async function getConfig(options: CLIOptions): Promise<Config> {
  const defaults = getDefaultDates();
  const defaultOutputDir = DEFAULT_OUTPUT || join(process.cwd(), "downloads");

  // Non-interactive mode if all options provided
  if (options.email) {
    const startDate = options.start || defaults.startDate;
    const endDate = options.end || defaults.endDate;
    validateDates(startDate, endDate);

    return {
      senderEmail: options.email,
      startDate,
      endDate,
      outputDir: expandPath(options.output || defaultOutputDir),
    };
  }

  // Interactive mode
  console.log();
  console.log(chalk.bold("Mailgrep - Office 365 Email Attachment Downloader"));
  console.log(chalk.dim("─".repeat(50)));
  console.log();

  const rl = createPrompt();

  try {
    // Prompt for sender email (required if no default)
    let senderEmail: string;
    if (DEFAULT_SENDER) {
      senderEmail = await ask(rl, "Sender email", DEFAULT_SENDER);
    } else {
      senderEmail = await askRequired(rl, "Sender email");
    }

    const startDate = await ask(rl, "Start date (YYYY-MM-DD)", defaults.startDate);
    const endDate = await ask(rl, "End date (YYYY-MM-DD)", defaults.endDate);
    const outputDir = await ask(rl, "Output folder", defaultOutputDir);

    rl.close();

    // Validate dates before proceeding
    validateDates(startDate, endDate);

    return {
      senderEmail,
      startDate,
      endDate,
      outputDir: expandPath(outputDir),
    };
  } catch (error) {
    rl.close();
    throw error;
  }
}

function expandPath(path: string): string {
  if (path.startsWith("~")) {
    return join(homedir(), path.slice(1));
  }
  if (!path.startsWith("/")) {
    return join(process.cwd(), path);
  }
  return path;
}

function getOutputDir(options: CLIOptions): string {
  const defaultOutputDir = DEFAULT_OUTPUT || join(process.cwd(), "downloads");
  return expandPath(options.output || defaultOutputDir);
}

// ============================================
// Main Logic
// ============================================
// Main execution flow:
// 1. Get config (interactive or CLI args)
// 2. Authenticate via browser OAuth
// 3. Load manifest for incremental backup
// 4. Fetch emails from Graph API (paginated)
// 5. For each email with attachments:
//    - Skip if image already in manifest
//    - Download and save new images
//    - Add to manifest entries
// 6. Save manifest (also on Ctrl+C)
// 7. Display summary

async function run(options: CLIOptions, forceReauth: boolean = false): Promise<void> {
  logger = new Logger(options.verbose, options.quiet);
  const startTime = new Date();

  const config = await getConfig(options);

  console.log();
  logger.dim(`Sender:     ${config.senderEmail}`);
  logger.dim(`Date range: ${config.startDate} to ${config.endDate}`);
  logger.dim(`Output:     ${config.outputDir}`);
  logger.dim(`File types: ${ALLOWED_EXTENSIONS ? FILE_TYPES_RAW : "all (*)"}`);
  logger.dim(`Started:    ${formatTime(startTime)}`);


  if (options.dryRun) {
    console.log();
    console.log(chalk.yellow("DRY RUN MODE - No files will be downloaded"));
  }

  console.log();

  // Create output directory
  if (!options.dryRun) {
    await Bun.$`mkdir -p ${config.outputDir}`.quiet();
  }

  // Authenticate - try cached token first, then refresh, then browser
  const authSpinner = ora("Authenticating with Office 365...").start();
  let accessToken: string;
  let userEmail: string;

  try {
    // Skip cache if --reauth flag is set
    if (forceReauth) {
      throw new Error("force_reauth");
    }

    const tokenCache = await loadTokenCache();

    // Check if multiple users exist for this tenant and no --user specified
    const tenantTokens = getCachedTokensForTenant(tokenCache, OAUTH_CONFIG.tenantId);
    let selectedUser = options.user;

    if (tenantTokens.length > 1 && !selectedUser) {
      // Multiple accounts cached, prompt user to select
      authSpinner.stop();
      console.log();
      console.log(chalk.bold("Multiple cached accounts found:"));
      tenantTokens.forEach((t, i) => {
        console.log(`  ${chalk.cyan(i + 1)}. ${t.userEmail}`);
      });
      console.log(`  ${chalk.cyan(tenantTokens.length + 1)}. Login as different user`);
      console.log();

      const rl = createPrompt();
      const answer = await new Promise<string>((resolve) => {
        rl.question(chalk.cyan("Select account (number): "), resolve);
      });
      rl.close();

      const choice = parseInt(answer, 10);
      if (choice >= 1 && choice <= tenantTokens.length) {
        selectedUser = tenantTokens[choice - 1].userEmail;
      } else {
        // User wants to login as different user
        throw new Error("user_selected_new_login");
      }

      authSpinner.start("Authenticating with Office 365...");
    }

    const cachedToken = getCachedToken(tokenCache, OAUTH_CONFIG.tenantId, selectedUser);

    // Log hint when --user specified but not found in cache
    if (!cachedToken && options.user) {
      authSpinner.info(`No cached token for user "${options.user}" in tenant "${OAUTH_CONFIG.tenantId}"`);
      authSpinner.start("Opening browser for login...");
      throw new Error("user_not_in_cache");
    }

    if (cachedToken) {
      if (!isTokenExpired(cachedToken)) {
        // Use cached token directly
        accessToken = cachedToken.accessToken;
        userEmail = cachedToken.userEmail;
        authSpinner.succeed(`Authenticated as ${chalk.cyan(userEmail)} ${chalk.dim("(cached)")}`);
      } else if (cachedToken.refreshToken) {
        // Token expired, try refresh
        authSpinner.text = "Refreshing authentication...";
        const refreshed = await refreshAccessToken(cachedToken.refreshToken);

        if (refreshed) {
          accessToken = refreshed.accessToken;
          userEmail = getUserEmailFromToken(accessToken);
          const actualTenantId = getTenantIdFromToken(accessToken) || OAUTH_CONFIG.tenantId;

          // Update cache with new tokens (use actual tenant ID from token)
          setCachedToken(tokenCache, {
            tenantId: actualTenantId,
            accessToken: refreshed.accessToken,
            refreshToken: refreshed.refreshToken,
            expiresAt: refreshed.expiresAt,
            userEmail,
            cachedAt: new Date().toISOString(),
          });
          await saveTokenCache(tokenCache);

          authSpinner.succeed(`Authenticated as ${chalk.cyan(userEmail)} ${chalk.dim("(refreshed)")}`);
        } else {
          // Refresh failed, need browser auth
          throw new Error("refresh_failed");
        }
      } else {
        // No refresh token, need browser auth
        throw new Error("no_refresh_token");
      }
    } else {
      // No cached token (or --user specified non-existent user), need browser auth
      throw new Error("no_cached_token");
    }
  } catch (error) {
    // Fall back to browser authentication
    if (error instanceof Error && !["refresh_failed", "no_refresh_token", "no_cached_token", "force_reauth", "user_selected_new_login", "user_not_in_cache"].includes(error.message)) {
      // Real error, re-throw
      authSpinner.fail("Authentication failed");
      throw error;
    }

    try {
      const authResult = await authenticate(authSpinner);
      accessToken = authResult.accessToken;
      userEmail = getUserEmailFromToken(accessToken);
      const actualTenantId = getTenantIdFromToken(accessToken) || OAUTH_CONFIG.tenantId;

      // Save tokens to cache (use actual tenant ID from token)
      const tokenCache = await loadTokenCache();
      setCachedToken(tokenCache, {
        tenantId: actualTenantId,
        accessToken: authResult.accessToken,
        refreshToken: authResult.refreshToken,
        expiresAt: authResult.expiresAt,
        userEmail,
        cachedAt: new Date().toISOString(),
      });
      await saveTokenCache(tokenCache);

      authSpinner.succeed(`Authenticated as ${chalk.cyan(userEmail)}`);
    } catch (authError) {
      authSpinner.fail("Authentication failed");
      throw authError;
    }
  }

  // Load manifest for incremental backup
  const manifestPath = join(config.outputDir, "manifest.json");
  const manifest = await loadManifest(manifestPath);
  const userSenderManifest = getOrCreateUserSenderManifest(manifest, userEmail, config.senderEmail);
  const downloadedKeys = !options.cache ? new Set<string>() : getDownloadedKeys(userSenderManifest);
  const processedEmailIds = !options.cache ? new Set<string>() : getProcessedEmailIds(userSenderManifest);

  if (options.cache && (downloadedKeys.size > 0 || processedEmailIds.size > 0)) {
    logger.dim(`Cache:      ${downloadedKeys.size} files, ${processedEmailIds.size} emails processed`);
  }
  if (!options.cache) {
    logger.dim(`Cache:      disabled (--no-cache)`);
  }

  // Fetch emails using Graph API with OData $filter
  // Using $filter with date range is more reliable than $search which has result limits
  const fetchSpinner = ora("Fetching emails...").start();

  // Build OData filter: sender email + date range
  // Note: from/emailAddress/address filter may cause InefficientFilter on some tenants,
  // so we also do client-side sender validation as fallback
  const filterParts = [
    `receivedDateTime ge ${config.startDate}T00:00:00Z`,
    `receivedDateTime le ${config.endDate}T23:59:59Z`,
  ];
  const filterQuery = encodeURIComponent(filterParts.join(" and "));

  // Graph API pagination: follow @odata.nextLink until exhausted
  let url: string | undefined = `https://graph.microsoft.com/v1.0/me/messages?$filter=${filterQuery}&$orderby=receivedDateTime desc&$select=id,subject,receivedDateTime,from,hasAttachments&$top=${PAGE_SIZE}`;

  const emails: Message[] = [];

  while (url) {
    const response: GraphResponse<Message> = await graphFetch<Message>(url, accessToken);

    for (const message of response.value) {
      // Client-side sender filter (more reliable than OData filter on sender)
      const fromEmail = message.from?.emailAddress?.address?.toLowerCase() || "";
      if (fromEmail !== config.senderEmail.toLowerCase()) continue;

      emails.push(message);
    }

    url = response["@odata.nextLink"];
    fetchSpinner.text = `Fetching emails... (${emails.length} found)`;
  }

  fetchSpinner.succeed(`Found ${emails.length} emails from ${config.senderEmail}`);

  if (emails.length === 0) {
    logger.info("No emails found matching criteria");
    return;
  }

  // Check if we need rate limiting for large mailboxes
  const useRateLimiting = emails.length >= LARGE_MAILBOX_THRESHOLD;
  if (useRateLimiting) {
    logger.dim(`Rate limit: ${DELAY_MS}ms delay enabled (large mailbox)`);
  }

  // Process emails and download images
  console.log();

  // Spinner frames (Claude Code style)
  const spinnerFrames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"];
  let spinnerIndex = 0;

  // Truncate text to fit terminal
  const truncate = (str: string, maxLen: number): string => {
    if (str.length <= maxLen) return str;
    return str.slice(0, maxLen - 1) + "…";
  };

  // Disable progress bar in verbose mode to avoid interference with debug logs
  const useProgressBar = !options.verbose;

  const progressBar = new cliProgress.SingleBar(
    {
      format: `{spinner} ${chalk.cyan("{bar}")} {percentage}% | {value}/{total} emails | {images} new | {skipped} cached | {status}`,
      hideCursor: true,
      barsize: 25,
    },
    cliProgress.Presets.shades_classic
  );

  if (useProgressBar) {
    progressBar.start(emails.length, 0, {
      images: 0,
      skipped: 0,
      spinner: chalk.cyan(spinnerFrames[0]),
      status: chalk.dim("Starting...")
    });
  }

  let totalImages = 0;
  let totalSkipped = 0;
  let totalSkippedEmails = 0;
  let totalSize = 0;
  const newEntries: ManifestEntry[] = [];
  const newProcessedEmailIds: string[] = [];

  // Handle Ctrl+C - save manifest before exit
  const cleanup = async () => {
    if (useProgressBar) progressBar.stop();
    console.log();
    console.log(chalk.yellow("\nInterrupted! Saving progress..."));

    if (newEntries.length > 0 || newProcessedEmailIds.length > 0) {
      userSenderManifest.entries.push(...newEntries);
      userSenderManifest.processedEmailIds.push(...newProcessedEmailIds);
      userSenderManifest.lastSync = new Date().toISOString();
      await saveManifest(manifestPath, manifest);
      console.log(chalk.green(`✓ Saved ${newEntries.length} files, ${newProcessedEmailIds.length} emails to manifest`));
    }

    process.exit(0);
  };

  process.on("SIGINT", cleanup);
  process.on("SIGTERM", cleanup);

  for (let i = 0; i < emails.length; i++) {
    const message = emails[i];
    const dateStr = message.receivedDateTime.slice(0, 10);
    const subject = message.subject || "(no subject)";

    // Update spinner and status
    spinnerIndex = (spinnerIndex + 1) % spinnerFrames.length;
    const statusText = truncate(subject, 30);

    if (useProgressBar) {
      progressBar.update(i, {
        images: totalImages,
        skipped: totalSkipped,
        spinner: chalk.cyan(spinnerFrames[spinnerIndex]),
        status: statusText
      });
    }

    logger.debug(`Processing: ${dateStr} - ${subject} (hasAttachments=${message.hasAttachments})`);

    // Skip already processed emails (no API call needed)
    if (processedEmailIds.has(message.id)) {
      totalSkippedEmails++;
      logger.debug(`  Skipped (already processed)`);
      if (useProgressBar) progressBar.update(i + 1, { images: totalImages, skipped: totalSkipped });
      continue;
    }

    // Note: hasAttachments can be false for inline images (embedded via CID)
    // We check all emails to catch inline images too, unless already processed
    // This costs one extra API call per email but ensures we don't miss inline images

    logger.debug(`  Fetching attachments...`);
    const attachmentsUrl = `https://graph.microsoft.com/v1.0/me/messages/${message.id}/attachments`;
    const attachmentsResponse: GraphResponse<Attachment> = await graphFetch<Attachment>(attachmentsUrl, accessToken);

    logger.debug(`  Found ${attachmentsResponse.value.length} attachment(s)`);

    for (const attachment of attachmentsResponse.value) {
      logger.debug(`    Attachment: "${attachment.name}" (${attachment.contentType}, inline=${attachment.isInline})`);

      if (!isAllowedFile(attachment)) {
        logger.debug(`    Skipped: not an allowed file type`);
        continue;
      }

      // Check if already downloaded
      const manifestKey = createManifestKey(message.id, attachment.id);
      if (downloadedKeys.has(manifestKey)) {
        totalSkipped++;
        logger.debug(`  Skipped (cached): ${attachment.name}`);
        continue;
      }

      // Update spinner on each attachment
      spinnerIndex = (spinnerIndex + 1) % spinnerFrames.length;
      if (useProgressBar) {
        progressBar.update(i, {
          spinner: chalk.cyan(spinnerFrames[spinnerIndex]),
          skipped: totalSkipped,
          status: chalk.green(truncate(attachment.name, 25))
        });
      }

      if (options.dryRun) {
        totalImages++;
        totalSize += attachment.size || 0;
        logger.debug(`  Would download: ${attachment.name} (${formatSize(attachment.size || 0)})`);
      } else {
        const fullAttachment = await getAttachment(message.id, attachment.id, accessToken);

        if (fullAttachment.contentBytes) {
          let safeName = attachment.name.replace(/[^a-zA-Z0-9._-]/g, "_");

          // Add extension for inline images that don't have one (e.g., "image001")
          const hasExtension = /\.[a-zA-Z0-9]+$/.test(safeName);
          if (!hasExtension && attachment.contentType) {
            const ext = CONTENT_TYPE_MAP[attachment.contentType.toLowerCase()];
            if (ext) safeName += ext;
          }

          const imageIndex = userSenderManifest.entries.length + newEntries.length + 1;
          const filename = `${dateStr}_${imageIndex}_${safeName}`;
          const filepath = join(config.outputDir, filename);

          const buffer = Buffer.from(fullAttachment.contentBytes, "base64");
          const fileHash = hashBuffer(buffer);
          await Bun.write(filepath, buffer);

          // Add to new entries for manifest
          newEntries.push({
            key: manifestKey,
            filename,
            originalName: attachment.name,
            size: buffer.length,
            hash: fileHash,
            emailSubject: subject,
            emailDate: message.receivedDateTime,
            downloadedAt: new Date().toISOString(),
          });

          totalImages++;
          totalSize += buffer.length;
          logger.debug(`  Saved: ${filename} (${formatSize(buffer.length)})`);

          // Rate limiting delay for large mailboxes
          if (useRateLimiting) {
            await Bun.sleep(DELAY_MS);
          }
        }
      }
    }

    // Mark email as fully processed
    newProcessedEmailIds.push(message.id);

    if (useProgressBar) progressBar.update(i + 1, { images: totalImages, skipped: totalSkipped });
  }

  if (useProgressBar) {
    progressBar.update(emails.length, {
      spinner: chalk.green("✓"),
      skipped: totalSkipped,
      status: chalk.green("Complete!")
    });
    progressBar.stop();
  } else {
    logger.info(`Complete! ${totalImages} new, ${totalSkipped} cached`);
  }

  // Remove signal handlers
  process.off("SIGINT", cleanup);
  process.off("SIGTERM", cleanup);

  // Save manifest with new entries and processed email IDs
  if (!options.dryRun && (newEntries.length > 0 || newProcessedEmailIds.length > 0)) {
    userSenderManifest.entries.push(...newEntries);
    userSenderManifest.processedEmailIds.push(...newProcessedEmailIds);
    userSenderManifest.lastSync = new Date().toISOString();
    await saveManifest(manifestPath, manifest);
  }

  // Summary
  const endTime = new Date();
  const duration = endTime.getTime() - startTime.getTime();

  // Check for duplicates
  const allEntries = [...userSenderManifest.entries, ...newEntries];
  const duplicates = findDuplicates(allEntries);
  const duplicateCount = duplicates.reduce((sum, g) => sum + g.files.length - 1, 0);

  console.log();
  console.log(chalk.bold("Summary"));
  console.log(chalk.dim("─".repeat(40)));
  console.log(`  Account:          ${chalk.cyan(userEmail)}`);
  console.log(`  Sender:           ${chalk.green(config.senderEmail)}`);
  console.log(`  Emails processed: ${chalk.cyan(emails.length)}`);
  console.log(`  Files ${options.dryRun ? "found" : "saved"}:      ${chalk.green(totalImages)}`);
  if (totalSkipped > 0) {
    console.log(`  Files skipped:    ${chalk.dim(totalSkipped)} ${chalk.dim("(cached)")}`);
  }
  console.log(`  Total size:       ${chalk.yellow(formatSize(totalSize))}`);
  if (duplicateCount > 0) {
    const wastedBytes = duplicates.reduce((sum, g) => sum + g.files[0].size * (g.files.length - 1), 0);
    console.log(`  Duplicates:       ${chalk.yellow(duplicateCount)} files ${chalk.dim(`(${formatSize(wastedBytes)} wasted)`)}`);
  }
  console.log(`  Duration:         ${chalk.cyan(formatDuration(duration))}`);

  if (!options.dryRun && totalImages > 0) {
    console.log();
    console.log(`  ${chalk.dim("Saved to:")} ${folderLink(config.outputDir)}`);
  }

  // Smart flow: offer to handle duplicates
  if (!options.dryRun && !options.quiet && duplicateCount > 0) {
    console.log();
    console.log(chalk.dim("─".repeat(40)));
    console.log();
    console.log(`  ${chalk.yellow("Duplicates detected!")} Common in email reply threads.`);
    console.log();

    const rl = createPrompt();
    const answer = await new Promise<string>((resolve) => {
      rl.question(chalk.cyan("  Remove duplicates to free space? (y/n): "), resolve);
    });
    rl.close();

    if (answer.toLowerCase() === "y" || answer.toLowerCase() === "yes") {
      console.log();

      // Check if hashes exist
      const needsHashes = allEntries.some(e => !e.hash);
      if (needsHashes) {
        console.log(chalk.dim("  Building file hashes..."));
        for (const entry of userSenderManifest.entries) {
          if (entry.hash) continue;
          const filepath = join(config.outputDir, entry.filename);
          const file = Bun.file(filepath);
          if (await file.exists()) {
            const buffer = Buffer.from(await file.arrayBuffer());
            entry.hash = hashBuffer(buffer);
          }
        }
      }

      // Dedupe
      const freshDuplicates = findDuplicates(userSenderManifest.entries);
      let deleted = 0;
      let freed = 0;

      for (const group of freshDuplicates) {
        const sorted = group.files.sort((a, b) =>
          new Date(a.emailDate).getTime() - new Date(b.emailDate).getTime()
        );
        const keep = sorted[0];

        for (const entry of sorted.slice(1)) {
          const filepath = join(config.outputDir, entry.filename);
          const file = Bun.file(filepath);
          if (await file.exists()) {
            if (await deleteFile(filepath)) {
              deleted++;
              freed += entry.size;
            }
          }
          entry.filename = `[deleted] -> ${keep.filename}`;
        }
      }

      await saveManifest(manifestPath, manifest);
      console.log(chalk.green(`  ✓ Removed ${deleted} duplicates, freed ${formatSize(freed)}`));
    }
  }

  console.log();
}

// ============================================
// CLI Setup
// ============================================
const program = new Command();

program
  .name("mailgrep")
  .description("Download email attachments from Office 365 (images by default, configurable via FILE_TYPES)")
  .version(VERSION)
  .option("-e, --email <address>", "sender email address")
  .option("-s, --start <date>", "start date (YYYY-MM-DD)")
  .option("-n, --end <date>", "end date (YYYY-MM-DD)")
  .option("-o, --output <dir>", "output directory")
  .option("--dry-run", "preview downloads without saving", false)
  .option("--no-cache", "force re-download all images (ignore manifest)")
  .option("--show-accounts", "list all cached user/sender accounts", false)
  .option("--check-duplicates", "analyze and report duplicate files", false)
  .option("--rebuild-hashes", "calculate hashes for existing files (for duplicate detection)", false)
  .option("--dedupe", "remove duplicate files (keeps oldest, preserves cache)", false)
  .option("--logout", "clear cached authentication tokens", false)
  .option("--reauth", "force re-authentication (ignore cached tokens)", false)
  .option("-u, --user <email>", "use specific cached account (skip prompt)")
  .option("-v, --verbose", "show detailed logs", false)
  .option("-q, --quiet", "minimal output", false)
  .action(async (options: CLIOptions) => {
    logger = new Logger(options.verbose, options.quiet);

    // Show banner unless in quiet mode
    if (!options.quiet) {
      showBanner();
    }

    // Handle --logout
    if (options.logout) {
      const tokenCache = await loadTokenCache();

      console.log();
      console.log(chalk.bold("Logout"));
      console.log(chalk.dim("─".repeat(50)));

      if (tokenCache.tokens.length === 0) {
        console.log(chalk.dim("  No cached tokens found."));
      } else {
        for (const token of tokenCache.tokens) {
          console.log(`  ${chalk.cyan(token.userEmail)} (${token.tenantId})`);
        }
        tokenCache.tokens = [];
        await saveTokenCache(tokenCache);
        console.log();
        console.log(chalk.green("✓ All cached tokens cleared"));
      }

      console.log();
      process.exit(ExitCode.Success);
    }

    // Handle --show-accounts
    if (options.showAccounts) {
      const outputDir = getOutputDir(options);
      const manifestPath = join(outputDir, "manifest.json");
      const manifest = await loadManifest(manifestPath);

      console.log();
      console.log(chalk.bold("Cached Accounts"));
      console.log(chalk.dim("─".repeat(50)));
      console.log(chalk.dim(`  Manifest: ${manifestPath}`));
      console.log();

      if (manifest.accounts.length === 0) {
        console.log(chalk.dim("  No cached accounts found."));
      } else {
        for (const account of manifest.accounts) {
          console.log(`  ${chalk.cyan(account.userEmail)} ← ${chalk.green(account.senderEmail)}`);
          console.log(`    ${chalk.dim("Files:")} ${account.entries.length}`);
          console.log(`    ${chalk.dim("Last sync:")} ${account.lastSync}`);
          console.log();
        }
      }

      console.log();
      process.exit(ExitCode.Success);
    }

    // Handle --check-duplicates
    if (options.checkDuplicates) {
      const outputDir = getOutputDir(options);
      const manifestPath = join(outputDir, "manifest.json");
      const manifest = await loadManifest(manifestPath);

      console.log();
      console.log(chalk.bold("Duplicate Analysis"));
      console.log(chalk.dim("─".repeat(50)));
      console.log(chalk.dim(`  Manifest: ${manifestPath}`));

      let totalDuplicates = 0;
      let totalWastedBytes = 0;

      for (const account of manifest.accounts) {
        const duplicates = findDuplicates(account.entries);
        if (duplicates.length === 0) continue;

        console.log(`\n  ${chalk.cyan(account.userEmail)} ← ${chalk.green(account.senderEmail)}`);

        for (const group of duplicates) {
          const wastedBytes = group.files[0].size * (group.files.length - 1);
          totalDuplicates += group.files.length - 1;
          totalWastedBytes += wastedBytes;

          console.log(`\n    ${chalk.yellow("Duplicate group")} (${group.files.length} files, ${formatSize(wastedBytes)} wasted):`);
          for (const file of group.files) {
            console.log(`      ${chalk.dim("•")} ${file.filename}`);
            console.log(`        ${chalk.dim(file.emailSubject.slice(0, 50))}`);
          }
        }
      }

      console.log();
      console.log(chalk.dim("─".repeat(50)));
      if (totalDuplicates > 0) {
        console.log(`  ${chalk.yellow("Total duplicates:")} ${totalDuplicates} files`);
        console.log(`  ${chalk.yellow("Wasted space:")}     ${formatSize(totalWastedBytes)}`);
        console.log();
        console.log(chalk.dim("  Note: Duplicates often come from email reply threads."));
        console.log(chalk.dim("  A future version may add --dedupe to remove them."));
      } else {
        console.log(chalk.green("  No duplicates found!"));
      }

      console.log();
      process.exit(ExitCode.Success);
    }

    // Handle --rebuild-hashes
    if (options.rebuildHashes) {
      const outputDir = getOutputDir(options);
      const manifestPath = join(outputDir, "manifest.json");
      const manifest = await loadManifest(manifestPath);

      console.log();
      console.log(chalk.bold("Rebuilding Hashes"));
      console.log(chalk.dim("─".repeat(50)));
      console.log(chalk.dim(`  Manifest: ${manifestPath}`));
      console.log();

      let updated = 0;
      let missing = 0;
      let total = 0;

      for (const account of manifest.accounts) {
        for (const entry of account.entries) {
          total++;
          if (entry.hash) continue;  // Already has hash

          const filepath = join(outputDir, entry.filename);
          const file = Bun.file(filepath);

          if (await file.exists()) {
            const buffer = Buffer.from(await file.arrayBuffer());
            entry.hash = hashBuffer(buffer);
            updated++;
            process.stdout.write(`\r  Processing: ${updated} files hashed...`);
          } else {
            missing++;
          }
        }
      }

      console.log(`\r  Processed:  ${total} files                    `);
      console.log(`  Updated:    ${chalk.green(updated)} hashes added`);
      if (missing > 0) {
        console.log(`  Missing:    ${chalk.yellow(missing)} files not found`);
      }

      if (updated > 0) {
        await saveManifest(manifestPath, manifest);
        console.log();
        console.log(chalk.green("✓ Manifest updated with hashes"));
      }

      console.log();
      process.exit(ExitCode.Success);
    }

    // Handle --dedupe
    if (options.dedupe) {
      const outputDir = getOutputDir(options);
      const manifestPath = join(outputDir, "manifest.json");
      const manifest = await loadManifest(manifestPath);

      console.log();
      console.log(chalk.bold("Deduplicating Files"));
      console.log(chalk.dim("─".repeat(50)));
      console.log(chalk.dim(`  Manifest: ${manifestPath}`));

      let totalDeleted = 0;
      let totalFreed = 0;

      for (const account of manifest.accounts) {
        const duplicates = findDuplicates(account.entries);
        if (duplicates.length === 0) continue;

        console.log(`\n  ${chalk.cyan(account.userEmail)} ← ${chalk.green(account.senderEmail)}`);

        for (const group of duplicates) {
          // Sort by date to keep the oldest
          const sorted = group.files.sort((a, b) =>
            new Date(a.emailDate).getTime() - new Date(b.emailDate).getTime()
          );

          const keep = sorted[0];
          const toDelete = sorted.slice(1);

          console.log(`\n    ${chalk.dim("Keeping:")} ${keep.filename}`);

          for (const entry of toDelete) {
            const filepath = join(outputDir, entry.filename);
            const file = Bun.file(filepath);

            if (await file.exists()) {
              const deleted = await deleteFile(filepath);
              if (deleted) {
                totalDeleted++;
                totalFreed += entry.size;
                console.log(`    ${chalk.red("Deleted:")} ${entry.filename}`);
              }
            }

            // Mark as deleted in manifest but keep entry for cache
            entry.filename = `[deleted] -> ${keep.filename}`;
          }
        }
      }

      console.log();
      console.log(chalk.dim("─".repeat(50)));
      console.log(`  ${chalk.green("Files deleted:")}  ${totalDeleted}`);
      console.log(`  ${chalk.green("Space freed:")}    ${formatSize(totalFreed)}`);
      console.log();
      console.log(chalk.dim("  Cache preserved - duplicates won't be re-downloaded."));

      await saveManifest(manifestPath, manifest);
      console.log(chalk.green("✓ Manifest updated"));

      console.log();
      process.exit(ExitCode.Success);
    }

    // Check for Azure credentials
    if (!OAUTH_CONFIG.clientId) {
      console.error(chalk.red("\nError: AZURE_CLIENT_ID not configured\n"));
      console.log("Create a .env file with:");
      console.log(chalk.cyan("  AZURE_CLIENT_ID=your-client-id"));
      console.log(chalk.cyan("  AZURE_TENANT_ID=your-tenant-id"));
      console.log();
      console.log("See: https://portal.azure.com → App registrations");
      process.exit(ExitCode.ConfigError);
    }

    try {
      await run(options, options.reauth);
      process.exit(ExitCode.Success);
    } catch (error) {
      if (error instanceof DateValidationError) {
        // Friendly error for date validation (no stack trace needed)
        console.error(chalk.red(`\nError: ${error.message}`));
        process.exit(ExitCode.ConfigError);
      }
      if (error instanceof Error) {
        console.error(chalk.red(`\nError: ${error.message}`));
        if (options.verbose) {
          console.error(error.stack);
        }
      }
      process.exit(ExitCode.NetworkError);
    }
  });

program.parse();
