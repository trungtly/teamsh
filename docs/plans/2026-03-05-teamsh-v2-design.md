# teamsh v2 Design

Date: 2026-03-05

## Summary

Redesign teamsh to use local file storage for all messages and emails, integrate with tv (television) CLI for fuzzy search, improve sidebar with sections and jump keys, and fix scrolling. The existing ratatui TUI stays as the primary interface for reading and replying to messages. tv is invoked as a subprocess from the TUI for search.

## Architecture

```
┌──────────────────────────────────────────────────────┐
│                    teamsh TUI                         │
│                                                       │
│  ┌─────────────┐  ┌──────────────────────────────┐   │
│  │  Sidebar     │  │  Main Panel                   │  │
│  │             │  │                               │   │
│  │ Favourites  │  │  Messages / Email body         │  │
│  │ Activity    │  │                               │   │
│  │ DMs         │←→│  (cached rendered lines)       │  │
│  │ Channels    │  │                               │   │
│  │ Meetings    │  │                               │   │
│  │ Emails      │  ├──────────────────────────────┤   │
│  │  Inbox      │  │  Input: compose reply          │  │
│  │  Sent       │  └──────────────────────────────┘   │
│  │  ...        │                                     │
│  └─────────────┘         press /                     │
│         │                    │                        │
└─────────┼────────────────────┼───────────────────────┘
          │                    │
          │                    ▼
          │         ┌─────────────────────┐
          │         │   tv (subprocess)    │
          │         │                     │
          │         │  rg over data dir   │
          │         │  preview: teamsh    │
          │         │  preview {id}       │
          │         │                     │
          │         │  returns: file:line  │
          │         └──────────┬──────────┘
          │                    │
          ▼                    ▼
   ┌──────────────────────────────────────┐
   │     ~/.config/teamsh/data/           │
   │                                      │
   │  index.json                          │
   │  favourites.json                     │
   │  conversations/{id}/meta.json        │
   │  conversations/{id}/messages/*.txt   │
   │  emails/{folder}/*.txt               │
   └──────────────────────────────────────┘
          ▲
          │
   ┌──────────────┐
   │  Background   │
   │  Sync         │
   │  (API poll)   │
   └──────────────┘
```

Key interactions:
- TUI reads from local files (fast, no API wait for rendering)
- Background sync writes to local files from API
- `/` suspends TUI, spawns tv over data dir, tv returns selection, TUI resumes at that message
- `teamsh list` and `teamsh preview` are CLI subcommands that tv calls
- Works with rg, bat, tv as external tools

## 1. Local File Storage

### Directory structure

```
~/.config/teamsh/data/
  index.json                              # quick lookup for sidebar + tv list
  favourites.json                         # pinned conversation/channel IDs
  conversations/
    {conv_id}/
      meta.json                           # name, kind, members, unread, version
      messages/
        {msg_id}.txt                      # plain text, one file per message
  emails/
    {folder_name}/
      {email_id}.txt                      # plain text email
```

### File formats

**index.json** -- fast lookup for sidebar rendering and `teamsh list`:
```json
{
  "conversations": [
    {"id": "...", "name": "Alice, Bob", "kind": "Chat", "last_activity": 1709654400, "unread": true}
  ],
  "email_folders": [
    {"name": "Inbox", "id": "...", "count": 12},
    {"name": "Sent Items", "id": "...", "count": 5}
  ]
}
```

**meta.json** -- per-conversation metadata:
```json
{
  "name": "Alice, Bob",
  "kind": "Chat",
  "members": ["Alice Smith", "Bob Jones"],
  "unread": false,
  "version": 1709654400,
  "last_message_id": "...",
  "consumptionhorizon": "..."
}
```

**messages/{msg_id}.txt** -- plain text, human-readable, grep-friendly:
```
14:32 Trung Ly
  Hey, did you see the PR? I think we should merge before Friday.
```

**emails/{folder}/{email_id}.txt** -- plain text email:
```
From: Bob Smith <bob@6clicks.com>
Date: 2026-03-05T10:30:00Z
Subject: Re: PR #51

Looks good to merge. I left one comment about the error handling.
```

### Sync strategy

| Data | Frequency | Trigger |
|------|-----------|---------|
| Conversation list + metadata | Every 15s | Poll |
| Messages for open conversation | Every 15s | Poll (if new msg detected) |
| Messages for all conversations | On startup + manual `r` | Background sync |
| Email folders + messages | On startup + manual `e` | Background, then periodic |
| Favourites | Immediately | On `f` key toggle |

Startup flow:
1. Load `index.json` -> render sidebar instantly
2. Background sync from API -> write new files
3. Update `index.json` when sync completes

## 2. tv Integration

### Keybinding

`/` in the TUI shells out to tv. No fallback -- tv is a required dependency.

### Flow

1. TUI exits alternate screen, restores terminal
2. Spawns tv in text mode: `tv --source-command "rg . --no-heading --line-number ~/.config/teamsh/data/" --preview-command "teamsh preview '{}'"`
3. User searches, navigates results -- preview panel shows conversation content scrolled to the match (like tv text in Helix)
4. User presses Enter on a result
5. tv exits, outputs `conversations/{conv_id}/messages/{msg_id}.txt:14:the matched text`
6. TUI parses this -> extracts conv_id + msg_id
7. TUI resumes -> opens that conversation, scrolls to that specific message, highlights the search term
8. User presses `i` to reply

### New CLI subcommands

- `teamsh list` -- reads `index.json`, outputs one line per conversation/email for tv's source command
- `teamsh preview {path}` -- formats message content for tv's preview panel (plain text with context)
- `teamsh sync` -- force a full sync to local files (can be run standalone)

### Custom tv channel (optional)

Installed by `teamsh setup` to `~/.config/television/cable/teamsh.toml`:
```toml
[source]
command = "teamsh list"

[preview]
command = "teamsh preview '{}'"
```

User can switch between the teamsh channel (browse by name) and text channel (search content) using tv's built-in channel switcher.

## 3. Sidebar Sections

### Layout

```
 teamsh ──────────────────
 Favourites (3)              ★
   @ Alice, Bob
   # engineering
   @ Charlie

 Activity
   @ Alice, Bob         2m ago
   # general            5m ago
   M Sprint Planning   12m ago

 Direct Messages (12)
   @ Dave
   @ Eve, Frank
   ...
 Channels (5)
   # deployments
   ...
 Meetings (3)
   ...
 Emails
   Inbox (12)
     GitHub - [EXTERNAL]...
   Sent Items (5)
     Re: Deploy schedule...
   Flagged (2)
     Important: Q4 review...
   Archive (18)
     ...
```

### Sections

| Section | Content | Sort |
|---------|---------|------|
| Favourites | Pinned conversations/channels (toggle with `f`) | Manual order |
| Activity | Last ~10 items with new messages, across all types | Most recent first |
| DMs | All direct/group chats | Last activity |
| Channels | All channels | Last activity |
| Meetings | Meeting chats | Last activity |
| Emails | All mail folders from Graph API, grouped by folder | Received date |

### Navigation

- `<-`/`->` or `h`/`l` -- move focus between sidebar and main panel
- `Tab` -- cycle forward through sidebar sections
- `Shift+Tab` -- cycle backward through sidebar sections
- `j`/`k` -- navigate items within and across sections (skip headers)
- `Enter` -- open selected item
- `f` -- toggle favourite on selected item

### Favourites

- Press `f` on any sidebar item to pin/unpin
- Stored in `favourites.json`: `["conv_id_1", "conv_id_2", ...]`
- Favourites appear in both the Favourites section and their original section (dimmed star marker)

### Email folders

- Fetch all folders via `GET /me/mailFolders` (Microsoft Graph API)
- Show all folders with their message counts
- Subfolders supported via `GET /me/mailFolders/{id}/childFolders`
- Each folder is a collapsible sub-section under Emails

## 4. Scrolling Fixes

### Laggy/choppy rendering

- Cache rendered lines -- only re-render when messages change or window resizes
- Store `rendered_lines: Vec<Line>` and a dirty flag
- On draw: if not dirty, reuse cached lines; only compute visible slice

### Slow to move through content

| Key | Action |
|-----|--------|
| `j`/`k` | 1 line |
| `Ctrl+d`/`Ctrl+u` | Half page |
| `PgUp`/`PgDn` | Full page |
| Mouse scroll | 3 lines per tick |
| `G` | Jump to bottom |
| `g` | Jump to top |
| Hold `j`/`k` | Accelerate after 3 repeated presses within 200ms (3 lines at a time) |

### Auto-follow new messages

- When at bottom (within 5 lines): auto-scroll on new messages
- When scrolled up reading history: do NOT auto-scroll
- Show `New messages` indicator at bottom when new messages arrive while scrolled up
- `G` jumps to bottom and dismisses the indicator

## 5. External Dependencies

Required tools (not Rust crates):
- `tv` (television) -- fuzzy finder, required for `/` search
- `rg` (ripgrep) -- used by tv for text search over data dir
- `bat` -- used by tv for preview with syntax highlighting (optional)

## 6. Migration Path

1. Add local file storage module (reads/writes data dir)
2. Modify sync to write files alongside existing cache.json
3. Add `teamsh list`, `teamsh preview`, `teamsh sync` CLI subcommands
4. Add tv subprocess spawning from TUI on `/`
5. Refactor sidebar into sections with Tab navigation
6. Add favourites support
7. Add Activity section
8. Fetch and display email folders
9. Fix scrolling (render cache, auto-follow, acceleration)
10. Remove old cache.json once file storage is stable
