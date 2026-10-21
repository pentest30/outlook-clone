# Outlook Clone — Plan

A classic-desktop Outlook-style client. It talks to real mail servers over IMAP/SMTP,
and a local SQLite cache plays the role of Outlook's `.ost` file. Mail ships first, then calendar.

## Stack

| Concern | Choice |
|---|---|
| UI | Avalonia 11.3 (WPF-style XAML, Fluent theme), MVVM with CommunityToolkit.Mvvm |
| Runtime | .NET 10 |
| Mail protocols | MailKit / MimeKit (IMAP, SMTP, MIME parsing) |
| Local cache | SQLite through EF Core |
| Credentials | Windows DPAPI (`ProtectedData`), per user. OAuth2 tokens later |
| Dev mail server | GreenMail in Docker (IMAP 3143, SMTP 3025, accepts any login) |
| Tests | xUnit |

## Solution layout

```
src/
  Outlook.Core   domain models + interfaces (no UI, no I/O)
  Outlook.Data   EF Core SQLite cache (MailStore), credential store
  Outlook.Mail   MailKit IMAP sync engine + SMTP sender
  Outlook.App    Avalonia desktop app (views, view models)
tools/
  Outlook.DevSeed  sends sample mail into GreenMail
tests/
  Outlook.Tests
```

## How sync works (the same model Outlook uses)

1. The UI only ever reads from the local SQLite cache, so it stays fast and works offline.
2. `ImapSyncEngine` keeps the cache in step with the server, folder by folder:
   - If `UIDVALIDITY` changed, the server renumbered the folder: wipe the local copy and refetch.
   - Fetch new messages (UIDs above the highest one cached): envelope, flags, size, body structure.
   - Refresh flags on messages already cached; delete local messages that vanished from the server.
3. Bodies download lazily the first time a message is opened.
4. User actions (read, flag, move, delete) update the cache first (optimistic UI), then go to the server.
5. Later: IMAP IDLE gives push for the Inbox, and CONDSTORE/QRESYNC give cheap delta syncs.

## Phases

### Phase 0 — Foundation (done)
- Solution, central package management, GreenMail docker-compose, seed tool
- Domain model: Account, MailFolder, MailMessage, addresses, special folders
- SQLite cache + DPAPI credential store
- IMAP sync for folders and message headers, body fetched on open, read/flag write-back
- Three-pane shell: folder pane, message list, reading pane (plain text), account setup dialog
- Thread ID resolution from `Message-ID` / `In-Reply-To` / `References`

### Phase 1 — Mail core (done)
- HTML reading pane with Avalonia.HtmlRenderer (cross-platform, never runs scripts): active content
  stripped, inline `cid:` pictures served from the cache, remote pictures/stylesheets blocked behind a
  "Download pictures" banner, links open in the browser, `mailto:` opens compose
- Attachments cached with the body; open (risky types like .exe are blocked) or save as
- Compose window: new / reply / reply all / forward / edit draft, Outlook-style quote header,
  Cc/Bcc, attachments (picker or drag and drop, 25 MB cap), Ctrl+Enter to send
- Drafts auto-save every 30 s and on close, replacing the same draft by Message-ID
- SMTP send plus a copy in Sent; the original gets the replied (\Answered) or forwarded ($Forwarded) mark
- Move / delete / archive with optimistic UI, multi-select, drag messages onto folders, context menu.
  Delete goes to Trash, or deletes permanently (with confirmation) from Trash. Archive creates the folder if needed.
- With UIDPLUS, moved messages keep their cached body and attachments
- IMAP IDLE push on the Inbox (separate connection, reconnects with backoff), plus the 2-minute sync timer
- Background refreshes merge into the list, so selection and scroll position survive
- Schema managed with EF Core migrations (`dotnet ef migrations add <Name> --project src/Outlook.Data`)

Deferred from Phase 1:
- Rich-text (HTML) compose. Avalonia has no built-in rich text editor, so compose is plain text for now.
- Opening a message in its own window (double-click currently only opens drafts)

### Phase 2a — Outlook-specific features (done)
- Conversation view (toggle, remembered): one row per thread with participant names and a count.
  The reading pane stacks every message of the thread newest-first, including your replies from
  Sent and messages you archived, each labelled with its folder. Unread messages open expanded.
  Actions on a conversation row apply to all its messages in the folder.
- Search (Ctrl+E): SQLite FTS5 index kept current by triggers, prefix matching, and the operators
  `from:` `to:` `subject:` `category:` `has:attachment` `is:unread` `is:read` `is:flagged`.
  Search all folders or just the open one; results show their folder.
- Categories: Outlook's six defaults plus "New category…", stored on the server as IMAP
  keywords (`$Cat_Name`) so they follow the mail; chips in the list and reading pane
- Undo send: the message waits 5 s (`UndoSendSeconds` in settings.json, 0–10) with an Undo bar.
  Undo reopens the same compose session. Closing the app sends pending messages, never drops them.
- Hardening found during on-screen testing: reconnect on dropped IMAP sessions; a send whose Sent
  copy fails is reported as sent (never "failed", which invites a duplicate); unexpected UI errors
  go to `errors.log` and the status bar instead of crashing
- Shortcuts: Ctrl+N, Ctrl+R, Ctrl+Shift+R, Ctrl+F, Ctrl+E, Delete, Backspace, Ctrl+Q/U, Insert, F9

### Phase 2b — still to do
- Server-side search for mail older than the cached window
- Flags with due dates, pin
- Focused / Other inbox, rules engine (on new mail: move / flag / categorise)
- Scheduled send, signatures, auto-reply
- Category management (rename, recolour, delete)
- Settings window (undo-send delay and the rest of settings.json)
- Multiple accounts, unified Inbox

### Phase 3 — Accounts & robustness
- OAuth2 (XOAUTH2) for Microsoft 365 / Outlook.com and Gmail. Both block plain-password IMAP.
- Autodiscover (Mozilla ISPDB / well-known hosts)
- CONDSTORE/QRESYNC deltas, retry/backoff, offline outbox queue
- Notifications (toasts), system tray

### Phase 4 — Calendar
- Day / work-week / week / month views, drag to create, move and resize events
- Recurrence (RRULE), time zones
- Meeting invites: parse `text/calendar` in mail, Accept / Tentative / Decline replies (iTIP)
- CalDAV sync for real calendars (Graph API later for M365)
- Reminders

### Phase 5 — People & To Do
- Contacts (CardDAV), address autocomplete from history
- Flagged mail shows up as tasks

## Dev setup

```
docker compose up -d                                   # GreenMail
dotnet run --project tools/Outlook.DevSeed             # put sample mail in dev@localhost
dotnet run --project src/Outlook.App                   # sign in: dev@localhost / any password,
                                                       # IMAP localhost:3143, SMTP localhost:3025, no TLS
dotnet test
```
