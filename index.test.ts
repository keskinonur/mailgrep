/**
 * Mailgrep Test Suite
 *
 * Tests for manifest operations, hash functions, and utility helpers.
 * Run with: bun test
 */

import { describe, test, expect, beforeEach, afterEach } from "bun:test";
import { join } from "path";
import { tmpdir } from "os";
import { mkdtemp, rm } from "fs/promises";

// ============================================
// Test Utilities
// ============================================

/**
 * Creates a temporary directory for test isolation
 */
async function createTempDir(): Promise<string> {
  return await mkdtemp(join(tmpdir(), "mailgrep-test-"));
}

/**
 * Cleans up a temporary directory
 */
async function cleanupTempDir(dir: string): Promise<void> {
  await rm(dir, { recursive: true, force: true });
}

// ============================================
// Manifest Types (duplicated for testing)
// ============================================

interface ManifestEntry {
  key: string;
  filename: string;
  originalName: string;
  size: number;
  hash: string;
  emailSubject: string;
  emailDate: string;
  downloadedAt: string;
}

interface UserSenderManifest {
  userEmail: string;
  senderEmail: string;
  lastSync: string;
  entries: ManifestEntry[];
  processedEmailIds: string[];
}

interface Manifest {
  version: number;
  updatedAt: string;
  accounts: UserSenderManifest[];
}

// ============================================
// Manifest Helper Functions (for testing)
// ============================================

const MANIFEST_VERSION = 1;

function isValidManifest(obj: unknown): obj is Manifest {
  if (typeof obj !== "object" || obj === null) return false;
  const manifest = obj as Record<string, unknown>;
  return (
    typeof manifest.version === "number" &&
    typeof manifest.updatedAt === "string" &&
    Array.isArray(manifest.accounts)
  );
}

async function loadManifest(path: string): Promise<Manifest> {
  try {
    const file = Bun.file(path);
    if (await file.exists()) {
      const content = await file.json();
      if (isValidManifest(content)) {
        return content;
      }
    }
  } catch {
    // Ignore errors, return fresh manifest
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
    (a) =>
      a.userEmail.toLowerCase() === userEmail.toLowerCase() &&
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
// Security Helpers (for testing)
// ============================================

function sanitizeFilename(filename: string): string {
  return filename
    .replace(/[^a-zA-Z0-9._-]/g, "_")
    .replace(/\.{2,}/g, ".")
    .substring(0, 200);
}

function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) && email.length < 254;
}

// ============================================
// Tests: Manifest Operations
// ============================================

describe("Manifest Operations", () => {
  let tempDir: string;

  beforeEach(async () => {
    tempDir = await createTempDir();
  });

  afterEach(async () => {
    await cleanupTempDir(tempDir);
  });

  describe("loadManifest", () => {
    test("returns fresh manifest when file doesn't exist", async () => {
      const manifestPath = join(tempDir, "manifest.json");
      const manifest = await loadManifest(manifestPath);

      expect(manifest.version).toBe(MANIFEST_VERSION);
      expect(manifest.accounts).toEqual([]);
      expect(typeof manifest.updatedAt).toBe("string");
    });

    test("loads existing valid manifest", async () => {
      const manifestPath = join(tempDir, "manifest.json");
      const existingManifest: Manifest = {
        version: 1,
        updatedAt: "2025-01-01T00:00:00Z",
        accounts: [
          {
            userEmail: "user@example.com",
            senderEmail: "sender@example.com",
            lastSync: "2025-01-01T00:00:00Z",
            entries: [],
            processedEmailIds: [],
          },
        ],
      };
      await Bun.write(manifestPath, JSON.stringify(existingManifest));

      const manifest = await loadManifest(manifestPath);

      expect(manifest.version).toBe(1);
      expect(manifest.accounts).toHaveLength(1);
      expect(manifest.accounts[0].userEmail).toBe("user@example.com");
    });

    test("returns fresh manifest for invalid JSON", async () => {
      const manifestPath = join(tempDir, "manifest.json");
      await Bun.write(manifestPath, "{ invalid json }");

      const manifest = await loadManifest(manifestPath);

      expect(manifest.version).toBe(MANIFEST_VERSION);
      expect(manifest.accounts).toEqual([]);
    });

    test("returns fresh manifest for wrong schema", async () => {
      const manifestPath = join(tempDir, "manifest.json");
      await Bun.write(manifestPath, JSON.stringify({ foo: "bar" }));

      const manifest = await loadManifest(manifestPath);

      expect(manifest.version).toBe(MANIFEST_VERSION);
      expect(manifest.accounts).toEqual([]);
    });
  });

  describe("saveManifest", () => {
    test("saves manifest to file", async () => {
      const manifestPath = join(tempDir, "manifest.json");
      const manifest: Manifest = {
        version: 1,
        updatedAt: "",
        accounts: [
          {
            userEmail: "user@test.com",
            senderEmail: "sender@test.com",
            lastSync: "2025-01-01T00:00:00Z",
            entries: [],
            processedEmailIds: ["email-1", "email-2"],
          },
        ],
      };

      await saveManifest(manifestPath, manifest);

      const file = Bun.file(manifestPath);
      expect(await file.exists()).toBe(true);

      const loaded = await file.json();
      expect(loaded.version).toBe(1);
      expect(loaded.accounts).toHaveLength(1);
      expect(loaded.accounts[0].processedEmailIds).toEqual(["email-1", "email-2"]);
    });

    test("updates updatedAt timestamp", async () => {
      const manifestPath = join(tempDir, "manifest.json");
      const manifest: Manifest = {
        version: 1,
        updatedAt: "old-timestamp",
        accounts: [],
      };

      const beforeSave = Date.now();
      await saveManifest(manifestPath, manifest);
      const afterSave = Date.now();

      const saved = await Bun.file(manifestPath).json();
      const savedTime = new Date(saved.updatedAt).getTime();

      expect(savedTime).toBeGreaterThanOrEqual(beforeSave);
      expect(savedTime).toBeLessThanOrEqual(afterSave);
    });
  });

  describe("getOrCreateUserSenderManifest", () => {
    test("creates new account if not exists", () => {
      const manifest: Manifest = {
        version: 1,
        updatedAt: "",
        accounts: [],
      };

      const account = getOrCreateUserSenderManifest(
        manifest,
        "user@example.com",
        "sender@example.com"
      );

      expect(account.userEmail).toBe("user@example.com");
      expect(account.senderEmail).toBe("sender@example.com");
      expect(account.entries).toEqual([]);
      expect(account.processedEmailIds).toEqual([]);
      expect(manifest.accounts).toHaveLength(1);
    });

    test("returns existing account if found", () => {
      const existingAccount: UserSenderManifest = {
        userEmail: "user@example.com",
        senderEmail: "sender@example.com",
        lastSync: "2025-01-01T00:00:00Z",
        entries: [
          {
            key: "msg|att",
            filename: "test.jpg",
            originalName: "test.jpg",
            size: 1000,
            hash: "abc123",
            emailSubject: "Test",
            emailDate: "2025-01-01T00:00:00Z",
            downloadedAt: "2025-01-01T00:00:00Z",
          },
        ],
        processedEmailIds: ["email-1"],
      };
      const manifest: Manifest = {
        version: 1,
        updatedAt: "",
        accounts: [existingAccount],
      };

      const account = getOrCreateUserSenderManifest(
        manifest,
        "USER@EXAMPLE.COM", // Different case
        "SENDER@EXAMPLE.COM" // Different case
      );

      expect(account).toBe(existingAccount);
      expect(manifest.accounts).toHaveLength(1);
    });

    test("handles multiple accounts", () => {
      const manifest: Manifest = {
        version: 1,
        updatedAt: "",
        accounts: [],
      };

      getOrCreateUserSenderManifest(manifest, "user1@example.com", "sender1@example.com");
      getOrCreateUserSenderManifest(manifest, "user2@example.com", "sender2@example.com");
      getOrCreateUserSenderManifest(manifest, "user1@example.com", "sender2@example.com");

      expect(manifest.accounts).toHaveLength(3);
    });
  });

  describe("getDownloadedKeys", () => {
    test("returns set of entry keys", () => {
      const account: UserSenderManifest = {
        userEmail: "user@example.com",
        senderEmail: "sender@example.com",
        lastSync: "",
        entries: [
          { key: "msg1|att1", filename: "", originalName: "", size: 0, hash: "", emailSubject: "", emailDate: "", downloadedAt: "" },
          { key: "msg1|att2", filename: "", originalName: "", size: 0, hash: "", emailSubject: "", emailDate: "", downloadedAt: "" },
          { key: "msg2|att1", filename: "", originalName: "", size: 0, hash: "", emailSubject: "", emailDate: "", downloadedAt: "" },
        ],
        processedEmailIds: [],
      };

      const keys = getDownloadedKeys(account);

      expect(keys.size).toBe(3);
      expect(keys.has("msg1|att1")).toBe(true);
      expect(keys.has("msg1|att2")).toBe(true);
      expect(keys.has("msg2|att1")).toBe(true);
      expect(keys.has("msg3|att1")).toBe(false);
    });

    test("returns empty set for empty entries", () => {
      const account: UserSenderManifest = {
        userEmail: "user@example.com",
        senderEmail: "sender@example.com",
        lastSync: "",
        entries: [],
        processedEmailIds: [],
      };

      const keys = getDownloadedKeys(account);

      expect(keys.size).toBe(0);
    });
  });

  describe("getProcessedEmailIds", () => {
    test("returns set of processed email IDs", () => {
      const account: UserSenderManifest = {
        userEmail: "user@example.com",
        senderEmail: "sender@example.com",
        lastSync: "",
        entries: [],
        processedEmailIds: ["email-1", "email-2", "email-3"],
      };

      const ids = getProcessedEmailIds(account);

      expect(ids.size).toBe(3);
      expect(ids.has("email-1")).toBe(true);
      expect(ids.has("email-2")).toBe(true);
      expect(ids.has("email-3")).toBe(true);
      expect(ids.has("email-4")).toBe(false);
    });

    test("handles undefined processedEmailIds", () => {
      // Simulate legacy manifest without processedEmailIds field
      const account = {
        userEmail: "user@example.com",
        senderEmail: "sender@example.com",
        lastSync: "",
        entries: [],
      } as unknown as UserSenderManifest;

      const ids = getProcessedEmailIds(account);

      expect(ids.size).toBe(0);
    });
  });
});

// ============================================
// Tests: Hash Functions
// ============================================

describe("Hash Functions", () => {
  describe("createManifestKey", () => {
    test("creates composite key from messageId and attachmentId", () => {
      const key = createManifestKey("msg-123", "att-456");
      expect(key).toBe("msg-123|att-456");
    });

    test("handles special characters in IDs", () => {
      const key = createManifestKey("msg=abc/123", "att=xyz/789");
      expect(key).toBe("msg=abc/123|att=xyz/789");
    });

    test("handles empty strings", () => {
      const key = createManifestKey("", "");
      expect(key).toBe("|");
    });
  });

  describe("hashBuffer", () => {
    test("returns SHA-256 hash of buffer", () => {
      const buffer = Buffer.from("Hello, World!");
      const hash = hashBuffer(buffer);

      // SHA-256 hash is 64 hex characters
      expect(hash).toHaveLength(64);
      expect(/^[a-f0-9]+$/.test(hash)).toBe(true);
    });

    test("returns consistent hash for same content", () => {
      const buffer1 = Buffer.from("test content");
      const buffer2 = Buffer.from("test content");

      expect(hashBuffer(buffer1)).toBe(hashBuffer(buffer2));
    });

    test("returns different hash for different content", () => {
      const buffer1 = Buffer.from("content 1");
      const buffer2 = Buffer.from("content 2");

      expect(hashBuffer(buffer1)).not.toBe(hashBuffer(buffer2));
    });

    test("handles empty buffer", () => {
      const buffer = Buffer.from("");
      const hash = hashBuffer(buffer);

      expect(hash).toHaveLength(64);
      // Known SHA-256 of empty string
      expect(hash).toBe("e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855");
    });
  });
});

// ============================================
// Tests: Duplicate Detection
// ============================================

describe("Duplicate Detection", () => {
  describe("findDuplicates", () => {
    test("finds duplicate files by hash", () => {
      const entries: ManifestEntry[] = [
        { key: "1", filename: "a.jpg", originalName: "a.jpg", size: 100, hash: "hash1", emailSubject: "", emailDate: "2025-01-01", downloadedAt: "" },
        { key: "2", filename: "b.jpg", originalName: "b.jpg", size: 100, hash: "hash1", emailSubject: "", emailDate: "2025-01-02", downloadedAt: "" },
        { key: "3", filename: "c.jpg", originalName: "c.jpg", size: 200, hash: "hash2", emailSubject: "", emailDate: "2025-01-03", downloadedAt: "" },
      ];

      const duplicates = findDuplicates(entries);

      expect(duplicates).toHaveLength(1);
      expect(duplicates[0].hash).toBe("hash1");
      expect(duplicates[0].files).toHaveLength(2);
    });

    test("returns empty array when no duplicates", () => {
      const entries: ManifestEntry[] = [
        { key: "1", filename: "a.jpg", originalName: "a.jpg", size: 100, hash: "hash1", emailSubject: "", emailDate: "", downloadedAt: "" },
        { key: "2", filename: "b.jpg", originalName: "b.jpg", size: 200, hash: "hash2", emailSubject: "", emailDate: "", downloadedAt: "" },
        { key: "3", filename: "c.jpg", originalName: "c.jpg", size: 300, hash: "hash3", emailSubject: "", emailDate: "", downloadedAt: "" },
      ];

      const duplicates = findDuplicates(entries);

      expect(duplicates).toHaveLength(0);
    });

    test("handles entries without hash", () => {
      const entries: ManifestEntry[] = [
        { key: "1", filename: "a.jpg", originalName: "a.jpg", size: 100, hash: "", emailSubject: "", emailDate: "", downloadedAt: "" },
        { key: "2", filename: "b.jpg", originalName: "b.jpg", size: 100, hash: "", emailSubject: "", emailDate: "", downloadedAt: "" },
      ];

      const duplicates = findDuplicates(entries);

      expect(duplicates).toHaveLength(0);
    });

    test("groups multiple duplicates correctly", () => {
      const entries: ManifestEntry[] = [
        { key: "1", filename: "a.jpg", originalName: "", size: 100, hash: "hash1", emailSubject: "", emailDate: "", downloadedAt: "" },
        { key: "2", filename: "b.jpg", originalName: "", size: 100, hash: "hash1", emailSubject: "", emailDate: "", downloadedAt: "" },
        { key: "3", filename: "c.jpg", originalName: "", size: 100, hash: "hash1", emailSubject: "", emailDate: "", downloadedAt: "" },
        { key: "4", filename: "d.jpg", originalName: "", size: 200, hash: "hash2", emailSubject: "", emailDate: "", downloadedAt: "" },
        { key: "5", filename: "e.jpg", originalName: "", size: 200, hash: "hash2", emailSubject: "", emailDate: "", downloadedAt: "" },
      ];

      const duplicates = findDuplicates(entries);

      expect(duplicates).toHaveLength(2);

      const group1 = duplicates.find(d => d.hash === "hash1");
      const group2 = duplicates.find(d => d.hash === "hash2");

      expect(group1?.files).toHaveLength(3);
      expect(group2?.files).toHaveLength(2);
    });

    test("handles empty entries array", () => {
      const duplicates = findDuplicates([]);
      expect(duplicates).toHaveLength(0);
    });
  });
});

// ============================================
// Tests: Manifest Validation
// ============================================

describe("Manifest Validation", () => {
  describe("isValidManifest", () => {
    test("validates correct manifest structure", () => {
      const valid: Manifest = {
        version: 1,
        updatedAt: "2025-01-01T00:00:00Z",
        accounts: [],
      };

      expect(isValidManifest(valid)).toBe(true);
    });

    test("rejects null", () => {
      expect(isValidManifest(null)).toBe(false);
    });

    test("rejects undefined", () => {
      expect(isValidManifest(undefined)).toBe(false);
    });

    test("rejects non-object", () => {
      expect(isValidManifest("string")).toBe(false);
      expect(isValidManifest(123)).toBe(false);
      expect(isValidManifest([])).toBe(false);
    });

    test("rejects missing version", () => {
      expect(isValidManifest({ updatedAt: "", accounts: [] })).toBe(false);
    });

    test("rejects missing updatedAt", () => {
      expect(isValidManifest({ version: 1, accounts: [] })).toBe(false);
    });

    test("rejects missing accounts", () => {
      expect(isValidManifest({ version: 1, updatedAt: "" })).toBe(false);
    });

    test("rejects wrong types", () => {
      expect(isValidManifest({ version: "1", updatedAt: "", accounts: [] })).toBe(false);
      expect(isValidManifest({ version: 1, updatedAt: 123, accounts: [] })).toBe(false);
      expect(isValidManifest({ version: 1, updatedAt: "", accounts: "not array" })).toBe(false);
    });
  });
});

// ============================================
// Tests: Security Helpers
// ============================================

describe("Security Helpers", () => {
  describe("sanitizeFilename", () => {
    test("allows safe characters", () => {
      expect(sanitizeFilename("file.jpg")).toBe("file.jpg");
      expect(sanitizeFilename("my-image_01.png")).toBe("my-image_01.png");
    });

    test("replaces unsafe characters with underscore", () => {
      expect(sanitizeFilename("file<script>.jpg")).toBe("file_script_.jpg");
      expect(sanitizeFilename("path/to/file.jpg")).toBe("path_to_file.jpg");
      expect(sanitizeFilename("file name.jpg")).toBe("file_name.jpg");
    });

    test("collapses multiple dots", () => {
      expect(sanitizeFilename("file...jpg")).toBe("file.jpg");
      expect(sanitizeFilename("file....hidden")).toBe("file.hidden");
    });

    test("truncates long filenames", () => {
      const longName = "a".repeat(300) + ".jpg";
      const result = sanitizeFilename(longName);

      expect(result.length).toBeLessThanOrEqual(200);
    });

    test("handles empty string", () => {
      expect(sanitizeFilename("")).toBe("");
    });
  });

  describe("isValidEmail", () => {
    test("accepts valid emails", () => {
      expect(isValidEmail("user@example.com")).toBe(true);
      expect(isValidEmail("user.name@example.co.uk")).toBe(true);
      expect(isValidEmail("user+tag@example.com")).toBe(true);
    });

    test("rejects invalid emails", () => {
      expect(isValidEmail("")).toBe(false);
      expect(isValidEmail("notanemail")).toBe(false);
      expect(isValidEmail("user@")).toBe(false);
      expect(isValidEmail("@example.com")).toBe(false);
      expect(isValidEmail("user @example.com")).toBe(false);
    });

    test("rejects very long emails", () => {
      const longEmail = "a".repeat(250) + "@example.com";
      expect(isValidEmail(longEmail)).toBe(false);
    });
  });
});

// ============================================
// Tests: Integration
// ============================================

describe("Integration", () => {
  let tempDir: string;

  beforeEach(async () => {
    tempDir = await createTempDir();
  });

  afterEach(async () => {
    await cleanupTempDir(tempDir);
  });

  test("full manifest lifecycle", async () => {
    const manifestPath = join(tempDir, "manifest.json");

    // Load fresh manifest
    let manifest = await loadManifest(manifestPath);
    expect(manifest.accounts).toHaveLength(0);

    // Create account
    const account = getOrCreateUserSenderManifest(
      manifest,
      "user@test.com",
      "sender@test.com"
    );

    // Add entry
    const buffer = Buffer.from("test image content");
    account.entries.push({
      key: createManifestKey("msg-1", "att-1"),
      filename: "2025-01-01_1_image.jpg",
      originalName: "image.jpg",
      size: buffer.length,
      hash: hashBuffer(buffer),
      emailSubject: "Test Email",
      emailDate: "2025-01-01T10:00:00Z",
      downloadedAt: new Date().toISOString(),
    });
    account.processedEmailIds.push("msg-1");

    // Save
    await saveManifest(manifestPath, manifest);

    // Reload
    manifest = await loadManifest(manifestPath);
    const reloadedAccount = getOrCreateUserSenderManifest(
      manifest,
      "user@test.com",
      "sender@test.com"
    );

    // Verify
    expect(reloadedAccount.entries).toHaveLength(1);
    expect(reloadedAccount.entries[0].filename).toBe("2025-01-01_1_image.jpg");
    expect(getDownloadedKeys(reloadedAccount).has("msg-1|att-1")).toBe(true);
    expect(getProcessedEmailIds(reloadedAccount).has("msg-1")).toBe(true);
  });

  test("duplicate detection across multiple entries", async () => {
    const manifest: Manifest = {
      version: 1,
      updatedAt: "",
      accounts: [],
    };

    const account = getOrCreateUserSenderManifest(
      manifest,
      "user@test.com",
      "sender@test.com"
    );

    // Simulate email thread with same image attached multiple times
    const imageHash = hashBuffer(Buffer.from("shared image"));

    account.entries.push(
      {
        key: "msg-1|att-1",
        filename: "2025-01-01_1_logo.png",
        originalName: "logo.png",
        size: 5000,
        hash: imageHash,
        emailSubject: "Original email",
        emailDate: "2025-01-01T10:00:00Z",
        downloadedAt: "",
      },
      {
        key: "msg-2|att-1",
        filename: "2025-01-02_2_logo.png",
        originalName: "logo.png",
        size: 5000,
        hash: imageHash,
        emailSubject: "RE: Original email",
        emailDate: "2025-01-02T10:00:00Z",
        downloadedAt: "",
      },
      {
        key: "msg-3|att-1",
        filename: "2025-01-03_3_logo.png",
        originalName: "logo.png",
        size: 5000,
        hash: imageHash,
        emailSubject: "RE: RE: Original email",
        emailDate: "2025-01-03T10:00:00Z",
        downloadedAt: "",
      }
    );

    const duplicates = findDuplicates(account.entries);

    expect(duplicates).toHaveLength(1);
    expect(duplicates[0].files).toHaveLength(3);

    // Wasted space = (3-1) * 5000 = 10000 bytes
    const wastedBytes = duplicates[0].files[0].size * (duplicates[0].files.length - 1);
    expect(wastedBytes).toBe(10000);
  });
});
