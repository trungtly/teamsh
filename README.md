# teamsh

A terminal-based Microsoft Teams client with both CLI and TUI modes.

## Install

```bash
cargo build --release
cp target/release/teamsh ~/.local/bin/   # or anywhere in PATH
```

Optional tools for enhanced experience:
- [bat](https://github.com/sharkdp/bat) -- syntax-highlighted message rendering
- [tv](https://github.com/alexpasmantier/television) -- full-text search across synced messages

## Authentication

teamsh uses two OAuth tokens to access Microsoft Teams and Outlook:

| Token | Purpose | Lifetime | Setup |
|-------|---------|----------|-------|
| Teams | Chat, messages, presence | ~24 hours | Paste refresh token from browser |
| Graph | Email via Microsoft Graph | ~90 days | Device code flow (browser opens) |

### Initial Setup

```bash
teamsh auth init
```

This walks you through two steps:

**Step 1 -- Teams token (manual, one-time):**
1. Open https://teams.cloud.microsoft in your browser
2. Open DevTools (F12) > Network tab
3. Filter for `login.microsoftonline.com`
4. Find a POST to `oauth2/v2.0/token`
5. In the Response tab, copy the `refresh_token` value
6. Paste it when prompted

**Step 2 -- Graph token (automatic):**
1. A browser opens to Microsoft's device login page
2. Enter the code shown in the terminal
3. Approve the login

### Keeping Tokens Fresh

The Teams token expires after ~24h of inactivity. The Graph token lasts ~90 days. Both auto-rotate on each use, so regular usage keeps them alive.

For unattended operation, use the auto-refresh script:

```bash
# First run: opens a browser for manual login (saves SSO cookies)
./scripts/refresh-token.py --headed

# Subsequent runs: headless, uses saved SSO cookies
./scripts/refresh-token.py

# Cron: refresh both tokens every 12 hours
0 */12 * * * /path/to/scripts/refresh-token.py 2>> /tmp/teamsh-refresh.log
```

Requires: `npx` (Node.js). The script uses `agent-browser` (auto-installed via npx).

Options: `--teams-only`, `--graph-only`, `--timeout N`

## Usage

### TUI Mode (default)

```bash
teamsh
```

Full-screen terminal UI with sidebar navigation, message view, and compose input.

#### Key Bindings

| Key | Action |
|-----|--------|
| `j`/`k` or arrows | Navigate sidebar / scroll messages |
| `Enter` | Select conversation / send message |
| `s` or `f` | Search sidebar |
| `n`/`N` | Next/previous search match |
| `/` | Search within current thread |
| `T` | Full-text search via tv |
| `i` | Focus message input |
| `Tab` | Jump between sidebar sections |
| `Space` | Toggle favourite |
| `<`/`>` | Resize sidebar |
| `g`/`G` | Jump to top/bottom |
| `r` | Refresh conversations |
| `L` | Re-login (when token expires) |
| `Esc` | Back to sidebar |
| `q` | Quit |

### CLI Mode

```bash
teamsh chats                      # List conversations
teamsh chats --channels           # Channels only
teamsh chats --dms                # DMs only
teamsh messages <conv-id>         # Fetch messages
teamsh messages <conv-id> --plain # Plain text (no HTML)
teamsh send <conv-id> -m "hello"  # Send a message
teamsh search <query>             # Search people
teamsh emails                     # List inbox
teamsh sync                       # Sync all to local files
teamsh green [--hours N] [--keep] # Set status to Available
teamsh list                       # List synced conversations
```

Add `--format json` for JSON output.

### Local File Store

`teamsh sync` saves all conversations and emails as plain text files under `~/.config/teamsh/data/`. These are grep-friendly and searchable with standard Unix tools or the `T` key in TUI mode.

## Project Structure

```
src/           Rust source (compiled binary)
scripts/       Token refresh and tv integration scripts
```
