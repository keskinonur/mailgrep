# Mailgrep

Download email attachments from Office 365 with a single command.

```
  ┌─────────────────────────────────────────┐
  │                                         │
  │   ╔╦╗╔═╗╦╦  ╔═╗╦═╗╔═╗╔═╗                │
  │   ║║║╠═╣║║  ║ ╦╠╦╝║╣ ╠═╝                │
  │   ╩ ╩╩ ╩╩╩═╝╚═╝╩╚═╚═╝╩                  │
  │                                         │
  │   Like grep, but for your mailbox.      │
  │                                         │
  └─────────────────────────────────────────┘
```

## Features

- **Persistent Login** - Caches auth tokens so you don't need to login every time
- **Incremental Backup** - Only downloads new files, resumes where you left off
- **Duplicate Detection** - Finds and removes duplicate attachments from email threads
- **Configurable Filters** - Download images, PDFs, documents, or all attachments
- **Multi-Account** - Supports multiple Office 365 accounts
- **Cross-Platform** - Runs on macOS, Linux, and Windows
- **JSON Output** - CI/CD friendly output for automation
- **Statistics Mode** - Quick email/attachment counts without downloading

## Quick Start

```bash
# Download and install
git clone https://github.com/keskinonur/mailgrep.git
cd mailgrep && bun install

# Configure (see Azure Setup below)
cp .env.example .env

# Run
./mailgrep
```

## Usage

```bash
# Interactive mode
mailgrep

# Download attachments from a specific sender
mailgrep -e sender@company.com

# Specify date range
mailgrep -e sender@company.com -s 2025-01-01 -n 2025-12-31

# Preview without downloading
mailgrep -e sender@company.com --dry-run

# Only download emails since last sync
mailgrep -e sender@company.com --since last

# Quick statistics (no download)
mailgrep --stats -e sender@company.com

# JSON output for scripting
mailgrep --json -e sender@company.com

# Find and remove duplicates
mailgrep --check-duplicates
mailgrep --dedupe

# Show version
mailgrep -V
```

### All Options

| Option | Description |
|--------|-------------|
| `-e, --email <address>` | Filter by sender email |
| `-s, --start <date>` | Start date (YYYY-MM-DD) |
| `-n, --end <date>` | End date (YYYY-MM-DD) |
| `-o, --output <dir>` | Output directory |
| `--since <date>` | Only process emails after date (YYYY-MM-DD or 'last') |
| `--dry-run` | Preview without downloading |
| `--stats` | Show email/attachment statistics without downloading |
| `--json` | Output results in JSON format |
| `--no-cache` | Force re-download all files |
| `--show-accounts` | List cached accounts |
| `--check-duplicates` | Analyze duplicate files |
| `--dedupe` | Remove duplicates (keeps oldest) |
| `--rebuild-hashes` | Recalculate file hashes |
| `--logout` | Clear cached auth tokens |
| `--reauth` | Force browser re-authentication |
| `-u, --user <email>` | Use specific cached account |
| `-v, --verbose` | Detailed output (includes API response times) |
| `-q, --quiet` | Minimal output |
| `-V, --version` | Show version number |

### Incremental Sync with `--since`

The `--since` option enables efficient incremental syncing:

```bash
# Sync only emails from a specific date
mailgrep -e sender@company.com --since 2025-01-01

# Sync only emails since last successful run
mailgrep -e sender@company.com --since last
```

When using `--since last`, mailgrep reads the last sync timestamp from the manifest. If no prior sync exists, it falls back to the default start date (January 1st of current year).

### JSON Output for Automation

The `--json` flag outputs results in a structured format, ideal for CI/CD pipelines:

```bash
mailgrep --json -e sender@company.com
```

Output:
```json
{
  "account": "user@company.com",
  "sender": "sender@company.com",
  "emailsProcessed": 150,
  "emailsSkipped": 45,
  "filesDownloaded": 23,
  "filesSkipped": 89,
  "totalSize": 15728640,
  "duplicates": 5,
  "duration": 12345,
  "outputDir": "/path/to/downloads"
}
```

### Statistics Mode

Get quick counts without downloading anything:

```bash
mailgrep --stats -e sender@company.com
```

Output:
```
Statistics
────────────────────────────────────────
  Account:              user@company.com
  Sender:               sender@company.com
  Date range:           2025-01-01 to 2025-12-26
  Total emails:         150
  With attachments:     42
```

Combine with `--json` for scripting:
```bash
mailgrep --stats --json -e sender@company.com
```

## Configuration

Create a `.env` file:

```bash
# Required - Azure AD credentials (see setup below)
AZURE_CLIENT_ID=your-client-id
AZURE_TENANT_ID=your-tenant-id

# Optional - Default values
DEFAULT_SENDER=alerts@company.com
DEFAULT_OUTPUT_DIR=./downloads

# Optional - File types to download (default: images)
FILE_TYPES=png,jpg,jpeg,gif,bmp,webp,tiff
```

### File Type Examples

```bash
FILE_TYPES=png,jpg,jpeg           # Images only
FILE_TYPES=pdf,docx,xlsx          # Documents only
FILE_TYPES=*                      # All attachments
```

### Multi-Tenant vs Single-Tenant

| Setting | Use Case |
|---------|----------|
| `AZURE_TENANT_ID=your-tenant-id` | Single organization (recommended) |
| `AZURE_TENANT_ID=organizations` | Any Azure AD account (multi-tenant) |
| Omit `AZURE_TENANT_ID` | Defaults to `organizations` |

Multi-tenant mode allows logging in with any work account but requires the Azure app to be configured for multi-tenant.

### Multiple Accounts

When multiple accounts are cached, mailgrep prompts you to select:

```
Multiple cached accounts found:
  1. user1@company.com
  2. user2@company.com
  3. Login as different user

Select account (number):
```

Or skip the prompt with `--user`:

```bash
mailgrep --user user1@company.com -e sender@example.com
```

### Cache Locations

| Cache | Location | Purpose |
|-------|----------|---------|
| Auth tokens | `~/.mailgrep/tokens.json` | OAuth access/refresh tokens (chmod 600) |
| Download manifest | `<output-dir>/manifest.json` | Tracks downloaded files |
| Manifest backup | `<output-dir>/manifest.json.bak` | Backup before each update |

Use `--output` to change the download directory (and manifest location):

```bash
mailgrep -o ~/backups/mail --show-accounts
```

## Security

Mailgrep implements several security measures:

- **Token Storage** - Auth tokens stored with restrictive permissions (chmod 600)
- **Path Validation** - Prevents directory traversal attacks in filenames
- **Input Sanitization** - All user inputs validated before use
- **XSS Prevention** - OAuth error pages escape HTML content
- **Manifest Backup** - Automatic backup before each manifest update

## Azure AD Setup

1. Go to [Azure Portal](https://portal.azure.com) → **App registrations** → **New registration**
2. Name it "Mailgrep", select **Single tenant**, click **Register**
3. Go to **Authentication** → **Add platform** → **Mobile and desktop**
4. Add redirect URI: `http://localhost:8400`
5. Enable **Allow public client flows** → **Save**
6. Copy **Application ID** and **Directory ID** to your `.env` file

## Building

```bash
# Build for current platform
bun run build

# Build for all platforms
bun run build:all
```

| Platform | Command | Output |
|----------|---------|--------|
| macOS | `bun run build` | `mailgrep` |
| Linux | `bun run build:linux` | `mailgrep-linux-x64` |
| Windows | `bun run build:windows` | `mailgrep-windows-x64.exe` |

## Troubleshooting

### Port 8400 already in use

If you see "Port 8400 is already in use", another mailgrep instance may be running. Close it or wait for the OAuth timeout (5 minutes).

### Authentication timeout

The OAuth browser flow times out after 5 minutes. If you don't complete login in time, run the command again.

### Network errors

Mailgrep automatically retries failed requests with exponential backoff. If you see persistent network errors, check your internet connection.

### Rate limiting

For large mailboxes (500+ emails), mailgrep automatically enables rate limiting to avoid hitting Microsoft Graph API limits.

## License

MIT
