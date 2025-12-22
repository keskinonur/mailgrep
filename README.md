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

# Find and remove duplicates
mailgrep --check-duplicates
mailgrep --dedupe
```

### All Options

| Option | Description |
|--------|-------------|
| `-e, --email <address>` | Filter by sender email |
| `-s, --start <date>` | Start date (YYYY-MM-DD) |
| `-n, --end <date>` | End date (YYYY-MM-DD) |
| `-o, --output <dir>` | Output directory |
| `--dry-run` | Preview without downloading |
| `--no-cache` | Force re-download all files |
| `--show-accounts` | List cached accounts |
| `--check-duplicates` | Analyze duplicate files |
| `--dedupe` | Remove duplicates (keeps oldest) |
| `--rebuild-hashes` | Recalculate file hashes |
| `--logout` | Clear cached auth tokens |
| `--reauth` | Force browser re-authentication |
| `-v, --verbose` | Detailed output |
| `-q, --quiet` | Minimal output |

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

## License

MIT
