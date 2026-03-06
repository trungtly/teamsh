# Outlook Email API Notes

## Approach: Microsoft Graph API (not OWA)

The Outlook web client uses a proprietary OWA JSON protocol with cookie-based auth.
Instead, use **Microsoft Graph API** which is already partially integrated in teamsh:
- `auth.rs` has `graph_token()` (line 71-76, scope `https://graph.microsoft.com/.default`)
- `api.rs` `search_people()` already uses Graph API with Bearer auth

Requires `Mail.Read` permission on the Azure AD app registration.

## Key Graph Endpoints

### List emails in inbox
```
GET https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages
  ?$top=25
  &$select=subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview,importance
  &$orderby=receivedDateTime desc
Authorization: Bearer {graph_token}
```

### Read a single email
```
GET https://graph.microsoft.com/v1.0/me/messages/{message-id}
  ?$select=subject,from,toRecipients,ccRecipients,body,receivedDateTime,isRead,hasAttachments
Authorization: Bearer {graph_token}
```

### Search emails
```
GET https://graph.microsoft.com/v1.0/me/messages?$search="keyword"&$top=25
```

### Well-known folder IDs
`inbox`, `sentitems`, `drafts`, `deleteditems`, `junkemail`, `archive`

## Response Shape

### Message list item
```json
{
  "id": "AAMk...",
  "subject": "...",
  "from": {"emailAddress": {"name": "...", "address": "..."}},
  "receivedDateTime": "2026-03-05T06:09:33Z",
  "isRead": true,
  "hasAttachments": false,
  "bodyPreview": "First 255 chars...",
  "importance": "normal"
}
```

### Full message
```json
{
  "id": "AAMk...",
  "subject": "...",
  "from": {"emailAddress": {"name": "...", "address": "..."}},
  "toRecipients": [{"emailAddress": {"name": "...", "address": "..."}}],
  "body": {"contentType": "html", "content": "<html>..."},
  "receivedDateTime": "2026-03-05T06:09:33Z",
  "isRead": true
}
```

### Pagination
- Response includes `@odata.nextLink` for next page
- Use `$top=N` and `$skip=N` for manual pagination
