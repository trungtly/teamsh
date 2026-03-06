# teamsh

Terminal-based Microsoft Teams client (CLI + TUI) in Rust.

**Repo:** `github.com/trungtly/teamsh`
**Config dir at runtime:** `~/.config/teamsh/` (stores tokens, cache, data)

## Build & Run

```bash
cargo build --release          # Binary: target/release/teamsh
teamsh                         # TUI mode (default)
teamsh auth init               # Two-step login setup
teamsh chats                   # CLI: list conversations
teamsh sync                    # CLI: sync all to local files
```

## Architecture

```
src/
  main.rs      (554 lines)  CLI entry, subcommands, auth init flow
  auth.rs      (299 lines)  Dual OAuth: Teams SPA + Graph native client
  api.rs       (323 lines)  HTTP client for Teams chatsvc + Graph API
  types.rs     (172 lines)  Conversation, Message, ConvKind, etc.
  html.rs      (734 lines)  HTML->plain text + styled ratatui spans
  cache.rs     (50 lines)   Binary cache for fast TUI startup
  store.rs     (526 lines)  Local file store (plain text, grep-friendly)
  tui/
    mod.rs     (24 lines)   Terminal setup
    app.rs     (2802 lines) TUI state machine, rendering, all key handling
```

### Key Data Flow

1. **Auth**: `auth.rs` manages two independent OAuth2 tokens
2. **API**: `api.rs` calls Teams chatsvc (chat/messages) and Graph (emails)
3. **TUI startup**: load cache -> rebuild sidebar -> load favourites -> rebuild sidebar -> fetch from API -> rebuild sidebar
4. **Polling**: every 60s (`tick_count % 600`), poll for new messages; early poll at tick 5

### Auth System (Dual Token)

| Token | Client ID | Purpose | Lifetime | Method |
|-------|-----------|---------|----------|--------|
| Teams | `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (SPA) | Chat/messages API | ~24h refresh | Manual paste from browser DevTools |
| Graph | `d3590ed6-52b3-4102-aeff-aad2292ab01c` (native) | Email via Graph API | ~90 days | Device code flow |

**Why two tokens:** Teams chat API (chatsvc) requires tokens from the Teams SPA client ID specifically (error 209 otherwise). SPA client can't do device code flow (needs client_secret). Graph API works with the Office native client which supports device code.

**Files in `~/.config/teamsh/`:**
- `refresh_token` — Teams SPA refresh token (24h)
- `graph_refresh_token` — Graph/Office native refresh token (~90 days)
- `tenant_id` — Azure AD tenant ID (extracted from JWT)
- `region` — Teams API region (e.g., `apac`)
- `favourites` — newline-separated conversation/email IDs

### TUI State Machine

Focus states: `Sidebar` | `Messages` | `Input` | `ThreadSearch` | `SidebarFilter`

Sidebar sections: Favourites, Activity (top 10 recent), DMs, Channels, Meetings, Emails.
Sections are collapsible (`collapsed_sections: HashSet<String>`).
Search mode (`s`/`f`) saves collapsed state, expands all, then restores on `Esc`.

Key patterns:
- `rebuild_sidebar()` — rebuilds `sidebar_items: Vec<SidebarItem>` from conversations/emails
- `preview_selected()` — loads messages for selected conversation (async)
- `render_messages()` — pipes through bat for syntax highlighting, caches result
- `poll_new_messages()` — checks for new messages via API, marks unread

### External Tools

- **bat**: Syntax highlighting for message rendering (uses custom `TeamshMessage.sublime-syntax` + `teamsh.tmTheme`)
- **tv** (television): Full-text search across local file store (`T` key)
- Both are optional — graceful degradation if not installed

## Known Issues & Debt

### Compilation Warnings (non-blocking)
- Dead code in `html.rs`: `RichSegment`, `strip_html_rich`, `rich_to_spans`, color constants — these are from an unused rich HTML rendering path, kept for future use
- Dead code: `api.search_emails`, `auth.login_teams`, `store.load_conv_meta` — potentially useful methods not yet wired up
- Unused struct fields in `types.rs`: `conv_type`, `lastimreceivedtime`, `last_join_at`, `member_count`, `from_given_name` — needed for serde deserialization
- Minor: unused `row`/`col` vars in `app.rs` text measurement, unnecessary `mut` on 2 variables

### Token Refresh Automation (WIP)
`refresh-token.py` — attempts to automate Teams token refresh via Edge CDP on WSL.
**Status:** Partially working. Issues:
- WSL2 can't reach Windows localhost (network isolation) — solved by running CDP commands through PowerShell
- Copied Edge profile cookies don't work (encrypted per-profile) — must use real profile, requires Edge closed
- PowerShell WebSocket CDP monitoring works but token intercept not yet confirmed end-to-end
- Consider alternative: run a small Node.js script on Windows side instead

### Code Quality (from /simplify review)
Items identified but not yet fixed:
- Duplicate `display_name()` calls in same loop iteration (`load_conversations`)
- Duplicate struct initialization between `new()` and `new_demo()`
- Duplicate plain text assembly from store entries (3 places)
- Style params passed as 8 individual args could be a struct
- `get_tenant_id()` returns owned `String`, could return `&str`

## Git

- NO emojis in source code
- NEVER commit automatically
- NEVER mention Claude/AI in commits
- Test locally (`cargo check` / `cargo build`) before committing
