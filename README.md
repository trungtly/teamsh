# teamsh

A terminal-based Microsoft Teams client with both CLI and TUI modes.

## Features

- **TUI mode** (default): Full-screen ratatui interface with sidebar, message view, and input
  - Collapsible sections: Favourites, Activity, DMs, Channels, Meetings, Emails
  - Sidebar search with `s`/`f`, navigate matches with `n`/`N`
  - In-thread search with `/`
  - Full-text search via [tv](https://github.com/alexpasmantier/television) integration (`T`)
  - Mouse support (click, scroll, text selection + copy)
  - Syntax-highlighted messages via bat
  - Keyboard-driven: vim-style navigation (`j`/`k`/`g`/`G`), accelerated scrolling
  - New message polling with bell notification
  - Re-login with `L` when token expires

- **CLI mode**: Subcommands for scripting and automation
  - `teamsh chats` -- list conversations (JSON or plain)
  - `teamsh messages <id>` -- fetch messages from a conversation
  - `teamsh send <id> -m "text"` -- send a message
  - `teamsh search <query>` -- search people
  - `teamsh emails` -- list inbox emails
  - `teamsh sync` -- sync all data to local files (grep-friendly plain text)
  - `teamsh green` -- set presence to Available

- **Email** via Microsoft Graph API (separate 90-day token)

- **Local file store** for offline access and full-text search

## Authentication

teamsh uses two OAuth tokens:

| Token | Client | Lifetime | Method |
|-------|--------|----------|--------|
| Teams (chat) | Teams SPA (`5e3ce6c0`) | ~24 hours | Manual paste from browser DevTools |
| Graph (email) | Office native (`d3590ed6`) | ~90 days | Device code flow (browser opens automatically) |

### Setup

```bash
teamsh auth init
```

This runs a two-step login:
1. **Teams token**: Paste a refresh token from browser DevTools (see prompts)
2. **Email token**: Device code flow -- a browser opens, you approve, done

### Auto-refresh (WSL)

`refresh-token.py` can automatically refresh the Teams token by launching Edge with your existing profile:

```bash
# First run: opens Edge for login
python3 refresh-token.py --headed

# Subsequent runs: headless via CDP
python3 refresh-token.py

# Cron every 20 hours
0 */20 * * * python3 /path/to/refresh-token.py 2>> /tmp/teamsh-refresh.log
```

Requires: `pip install playwright && playwright install chromium`

## Build

```bash
cargo build --release
# Binary at target/release/teamsh
```

## Dependencies

- Rust 2021 edition
- [ratatui](https://ratatui.rs/) + crossterm (TUI)
- reqwest + tokio (async HTTP)
- clap (CLI parsing)
- Optional: [tv](https://github.com/alexpasmantier/television) (full-text search), [bat](https://github.com/sharkdp/bat) (syntax highlighting)

## Key Bindings (TUI)

| Key | Action |
|-----|--------|
| `j`/`k` or arrows | Navigate sidebar / scroll messages |
| `Enter` | Select conversation / send message |
| `s` or `f` | Sidebar search |
| `n`/`N` | Next/previous search match |
| `/` | Search within current thread |
| `T` | Full-text search (tv) |
| `i` | Focus input |
| `Esc` | Back to sidebar |
| `g`/`G` | Jump to top/bottom |
| `Tab` | Toggle sidebar sections |
| `F` | Toggle favourite |
| `r` | Refresh conversations |
| `L` | Re-login (token expired) |
| `q` | Quit |

## Architecture

```
src/
  main.rs    -- CLI entry point, subcommands
  auth.rs    -- Dual OAuth token management
  api.rs     -- Teams chatsvc + Graph API client
  types.rs   -- Shared data types (Conversation, Message, etc.)
  html.rs    -- HTML to plain text + styled spans
  cache.rs   -- Binary cache for fast TUI startup
  store.rs   -- Local file store (plain text, grep-friendly)
  tui/
    mod.rs   -- Terminal setup, run loop
    app.rs   -- TUI state, rendering, key handling
```
