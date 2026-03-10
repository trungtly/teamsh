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
  main.rs       CLI entry, subcommands, auth init flow
  auth.rs       Dual OAuth: Teams SPA + Graph native client
  api.rs        HTTP client for Teams chatsvc + Graph API
  types.rs      Conversation, Message, ConvKind, etc.
  html.rs       HTML->plain text + styled ratatui spans
  cache.rs      Binary cache for fast TUI startup
  store.rs      Local file store (plain text, grep-friendly)
  tui/
    mod.rs      Terminal setup
    app.rs      TUI state machine, rendering, all key handling
  TeamshMessage.sublime-syntax   bat syntax (compiled in)
  teamsh.tmTheme                 bat theme (compiled in)
scripts/
  refresh-token.py   Auto-refresh both tokens via browser automation
  tv-preview.sh      tv integration preview script (compiled in)
```

### Key Data Flow

1. **Auth**: `auth.rs` manages two independent OAuth2 tokens
2. **API**: `api.rs` calls Teams chatsvc (chat/messages) and Graph (emails)
3. **TUI startup**: load cache -> rebuild sidebar -> load favourites -> rebuild sidebar -> fetch from API -> rebuild sidebar
4. **Polling**: every 60s (`tick_count % 600`), poll for new messages; early poll at tick 5

### Auth System (Dual Token)

| Token | Client ID | Purpose | Lifetime | Method |
|-------|-----------|---------|----------|--------|
| Teams | `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (SPA) | Chat/messages API | ~24h refresh | Browser token paste or auto-refresh |
| Graph | `d3590ed6-52b3-4102-aeff-aad2292ab01c` (native) | Email via Graph API | ~90 days | Device code flow |

**Why two tokens:** Teams chat API (chatsvc) requires tokens from the Teams SPA client ID specifically (error 209 otherwise). SPA client can't do device code flow (needs client_secret). Graph API works with the Office native client which supports device code.

**Files in `~/.config/teamsh/`:**
- `refresh_token` -- Teams SPA refresh token (24h)
- `graph_refresh_token` -- Graph/Office native refresh token (~90 days)
- `tenant_id` -- Azure AD tenant ID (extracted from JWT)
- `region` -- Teams API region (e.g., `apac`)
- `favourites` -- newline-separated conversation/email IDs

### Token Auto-Refresh

`scripts/refresh-token.py` automates both tokens using `agent-browser` (Playwright CLI):
- **Teams**: Opens Teams in Chromium, extracts refresh token from MSAL browser storage
- **Graph**: Starts device code flow, navigates browser to approval page, SSO cookies auto-approve

First run requires `--headed` for manual login. Subsequent runs are headless via saved SSO cookies (~90 day lifetime). Cron every 12h keeps both tokens alive.

### TUI State Machine

Focus states: `Sidebar` | `Messages` | `Input` | `ThreadSearch` | `SidebarFilter`

Sidebar sections: Favourites, Activity (top 10 recent), DMs, Channels, Meetings, Emails.
Sections are collapsible (`collapsed_sections: HashSet<String>`).
Search mode (`s`/`f`) saves collapsed state, expands all, then restores on `Esc`.
Sidebar width adjustable with `<`/`>` keys (15-50%).

Key patterns:
- `rebuild_sidebar()` -- rebuilds `sidebar_items: Vec<SidebarItem>` from conversations/emails
- `preview_selected()` -- loads messages for selected conversation (async)
- `render_messages()` -- pipes through bat for syntax highlighting, caches result
- `poll_new_messages()` -- checks for new messages via API, marks unread

### External Tools

- **bat**: Syntax highlighting for message rendering (uses custom `src/TeamshMessage.sublime-syntax` + `src/teamsh.tmTheme`)
- **tv** (television): Full-text search across local file store (`T` key)
- Both are optional -- graceful degradation if not installed

## Known Issues & Debt

### Compilation Warnings (non-blocking)
- Dead code in `html.rs`: `RichSegment`, `strip_html_rich`, `rich_to_spans`, color constants -- unused rich HTML rendering path, kept for future use
- Dead code: `api.search_emails`, `auth.login_teams`, `store.load_conv_meta` -- potentially useful methods not yet wired up
- Unused struct fields in `types.rs`: `conv_type`, `lastimreceivedtime`, `last_join_at`, `member_count`, `from_given_name` -- needed for serde deserialization
- Minor: unused `row`/`col` vars in `app.rs` text measurement, unnecessary `mut` on 2 variables

### Code Quality
- Duplicate `display_name()` calls in same loop iteration (`load_conversations`)
- Duplicate struct initialization between `new()` and `new_demo()`
- Duplicate plain text assembly from store entries (3 places)
- Style params passed as 8 individual args could be a struct

## Git

- NO emojis in source code
- NEVER commit automatically
- NEVER mention Claude/AI in commits
- Test locally (`cargo check` / `cargo build`) before committing
