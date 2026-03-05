# teamsh - Microsoft Teams TUI

## Problem

Working in a remote CLI-only environment (WSL2) with no GUI access. Need to read/send Teams messages, monitor channels, and manage presence from the terminal. No Azure AD app registration available.

## Approach

Use the **Teams internal chatsvc API** with an OAuth2 refresh token extracted from the Teams web browser session. The refresh token is long-lived (months) and can be used to mint fresh access tokens from the CLI.

## Auth Model (Verified via HAR Analysis)

Teams web (`teams.cloud.microsoft`) uses MSAL.js with multiple OAuth2 tokens for different services. The key discovery:

### Token Chain

```
Browser login (one-time)
    |
    v
Refresh Token (long-lived, months)  <-- stored in ~/.config/teamsh/
    |
    | POST login.microsoftonline.com/.../oauth2/v2.0/token
    | + Origin: https://teams.cloud.microsoft  (required: SPA client)
    | + client_id: 5e3ce6c0-2b1f-4285-8d4b-75ee78787346  (Teams web app)
    v
IC3 Access Token (scope: ic3.teams.office.com/Teams.AccessAsUser.All, ~85min TTL)
    |
    | Authorization: Bearer <token>
    v
Chatsvc API (teams.cloud.microsoft/api/chatsvc/...)  --> messages, conversations
```

### Key Details

- **Client ID:** `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (Microsoft Teams web, public SPA client)
- **Tenant ID:** Org-specific (e.g. `f2dbeea5-157c-4463-91d3-d043b7416b6c`)
- **Token endpoint:** `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`
- **Origin header required:** `Origin: https://teams.cloud.microsoft` (SPA tokens need cross-origin)
- **Refresh token rotates:** Each refresh returns a new refresh token (must save it)
- **IC3 token TTL:** ~5000 seconds (~83 minutes)

### One-Time Setup

Extract refresh token from browser HAR:

1. Open https://teams.cloud.microsoft in browser
2. Open DevTools > Network tab
3. Filter for `login.microsoftonline.com`
4. Find a POST to `oauth2/v2.0/token` that returns `ic3.teams.office.com` in scope
5. Copy `refresh_token` from the response JSON
6. Save to `~/.config/teamsh/refresh_token`

After setup, teamsh refreshes tokens automatically.

### Graph API (Limited)

The Graph API token is also obtainable but **missing `Chat.Read` scope** (not preauthorized for the Teams SPA client). Graph API can be used for:

- User profiles (`/v1.0/me`, `/v1.0/users/{id}`)
- Presence (`/v1.0/me/presence`) -- needs separate token with `presence.teams.microsoft.com` scope

Chat operations must use the internal chatsvc API.

## Chatsvc API (Verified)

Base URL: `https://teams.cloud.microsoft/api/chatsvc/{region}/v1`

Region is `au` for Australia (from Skype token `rgn` field). Other regions: `us`, `eu`, etc.

### Conversations

| Action | Method | Endpoint |
|--------|--------|----------|
| List all conversations | GET | `/users/ME/conversations?view=msnp24Equivalent&pageSize=50` |
| Get messages in conversation | GET | `/users/ME/conversations/{conv-id}/messages?view=msnp24Equivalent&pageSize=20` |
| Send message | POST | `/users/ME/conversations/{conv-id}/messages` |

### Conversation Types (by ID pattern)

| Pattern | Type | Example |
|---------|------|---------|
| `19:...@thread.skype` | Team channel (old) | Core GRC - Registers Team |
| `19:...@thread.tacv2` | Team channel (new) | Product & Engineering |
| `19:...@thread.v2` | Group chat / channel | Level-2 support |
| `19:{user1}_{user2}` | 1:1 DM | Direct messages |
| `19:meeting_...` | Meeting chat | Sprint Demo |
| `48:notifications` | System: notifications | |
| `48:mentions` | System: mentions | |
| `48:notes` | System: notes | |

### Message Format

Messages arrive as `RichText/Html`. Key fields:

```json
{
  "id": "1772677275177",
  "originalarrivaltime": "2026-03-05T02:21:15.177Z",
  "messagetype": "RichText/Html",
  "imdisplayname": "Andrew Lawrence",
  "content": "<p>Hi Abby,</p><p>I believe the Developer API route...</p>",
  "properties": { ... }
}
```

### Required Headers

```
Authorization: Bearer {ic3_access_token}
behavioroverride: redirectAs404
x-ms-migration: True
```

## Architecture

```
teamsh (Rust binary)
  |
  +-- tui/          # Ratatui-based terminal UI
  |     +-- app.rs          # App state, event loop
  |     +-- views/          # Sidebar, message list, compose
  |     +-- keybindings.rs  # Vim-style navigation
  |
  +-- api/          # Teams API client
  |     +-- auth.rs         # OAuth2 refresh flow, token rotation
  |     +-- chatsvc.rs      # Conversations, messages (chatsvc API)
  |     +-- types.rs        # API response types (serde)
  |
  +-- config/       # Configuration
  |     +-- mod.rs          # Token storage (~/.config/teamsh/), settings
  |
  +-- html.rs       # Strip HTML tags from message content
  +-- polling.rs    # Background message polling (async)
```

### Key Dependencies

- `ratatui` + `crossterm` - TUI rendering
- `reqwest` - HTTP client (async, rustls)
- `tokio` - Async runtime
- `serde` / `serde_json` - JSON parsing
- `dirs` - Config directory (`~/.config/teamsh/`)
- `clap` - CLI argument parsing

## TUI Layout

```
+------------------+----------------------------------------+
| Channels/Chats   |  #general - Engineering                |
|                  |                                        |
| Channels         |  [10:30] Alice: Hey, PR is ready       |
|  > Level-2       |  [10:31] Bob: LGTM, merging            |
|    Coders        |  [10:35] Alice: Deployed to staging     |
|    Pipeline Test |    > Bob reacted: thumbsup             |
|                  |  [10:40] You: Nice work!               |
| Chats            |                                        |
|  Andrew L.       |                                        |
|  Julie, Saurabh  |                                        |
|  Abby            |                                        |
|                  |----------------------------------------+
|                  | > Type a message...                    |
+------------------+----------------------------------------+
| F1:Help  F2:Channels  F3:Chats  F5:Refresh  q:Quit      |
+----------------------------------------------------------+
```

### Keybindings

- `j/k` or arrow keys - Navigate up/down
- `Enter` - Select channel/chat, send message
- `Tab` - Switch between sidebar and message pane
- `/` - Search
- `r` - Reply to thread
- `e` - React to message
- `F5` or `Ctrl-R` - Refresh
- `q` or `Ctrl-C` - Quit

## Implementation Plan

### Phase 1: Foundation (MVP)

1. **Scaffold Rust project** - `cargo init`, add dependencies
2. **Auth module** - Refresh token loading, IC3 token minting, auto-rotation
3. **Chatsvc client** - List conversations, get messages, categorize by type
4. **HTML stripping** - Convert HTML messages to plain text
5. **Basic TUI shell** - Ratatui app loop, sidebar with channels/chats
6. **Read messages** - Display messages in selected conversation
7. **Send messages** - Compose and POST to chatsvc

### Phase 2: Full Features

8. **Thread replies** - View and post thread replies
9. **Reactions** - Display reactions, add reactions
10. **Presence** - Separate token for presence.teams.microsoft.com scope
11. **Notifications** - Terminal bell on new messages / mentions
12. **Background polling** - Async polling for new messages
13. **Search** - Search messages

### Phase 3: Polish

14. **Token management** - Detect expiry, helpful error messages
15. **Message formatting** - Better HTML-to-terminal rendering (code blocks, links)
16. **Unread indicators** - Track read/unread state
17. **Config file** - Customizable keybindings, poll interval, theme

## Risks

| Risk | Mitigation |
|------|------------|
| Refresh token expires | Long-lived (months observed), re-extract from browser if needed |
| SPA origin check tightened | Origin header currently sufficient, monitor for changes |
| Chatsvc API changes (undocumented) | Pin to known working paths, version responses |
| Rate limiting | Backoff, cache, reasonable poll intervals (30s+) |
| HTML message complexity | Start with tag stripping, improve incrementally |
| Region detection | Extract from Skype token `rgn` field or make configurable |
