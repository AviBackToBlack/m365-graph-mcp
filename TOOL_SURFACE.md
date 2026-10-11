# m365-graph-mcp — Normative Tool Surface

This file is generated from the user-selected `endpoints.json` and is normative for **what capabilities exist**.

Rules:
- Do not add, remove, rename, or broaden tools without explicit human approval.
- `endpoints.json` is authoritative for the desired surface, but its permission metadata is reference data, not an authority for Microsoft Graph permissions.
- Validate endpoint behavior, request shapes, query constraints, and least-privilege permissions against current Microsoft documentation during implementation.
- Read-only means **no mutation semantics**. A POST is allowed only for the explicitly selected retrieval operations below.

Selected tools: **76**  
HTTP methods: **71 GET**, **5 POST**

## Tool catalog

| # | Tool | Method | Path | Reference scopes |
|---:|---|:---:|---|---|
| 1 | `graph-batch` | `POST` | `/$batch` | `—` |
| 2 | `list-mail-messages` | `GET` | `/me/messages` | `Mail.Read` |
| 3 | `list-mail-folders` | `GET` | `/me/mailFolders` | `Mail.Read` |
| 4 | `list-mail-child-folders` | `GET` | `/me/mailFolders/{mailFolder-id}/childFolders` | `Mail.Read` |
| 5 | `list-mail-folder-messages` | `GET` | `/me/mailFolders/{mailFolder-id}/messages` | `Mail.Read` |
| 6 | `get-mail-message` | `GET` | `/me/messages/{message-id}` | `Mail.Read` |
| 7 | `list-shared-mailbox-messages` | `GET` | `/users/{user-id}/messages` | `Mail.Read.Shared` |
| 8 | `list-shared-mailbox-folder-messages` | `GET` | `/users/{user-id}/mailFolders/{mailFolder-id}/messages` | `Mail.Read.Shared` |
| 9 | `get-shared-mailbox-message` | `GET` | `/users/{user-id}/messages/{message-id}` | `Mail.Read.Shared` |
| 10 | `list-users` | `GET` | `/users` | `User.Read.All` |
| 11 | `list-mail-attachments` | `GET` | `/me/messages/{message-id}/attachments` | `Mail.Read` |
| 12 | `list-calendar-events` | `GET` | `/me/events` | `Calendars.Read` |
| 13 | `get-calendar-event` | `GET` | `/me/events/{event-id}` | `Calendars.Read` |
| 14 | `list-specific-calendar-events` | `GET` | `/me/calendars/{calendar-id}/events` | `Calendars.Read` |
| 15 | `get-specific-calendar-event` | `GET` | `/me/calendars/{calendar-id}/events/{event-id}` | `Calendars.Read` |
| 16 | `get-calendar-view` | `GET` | `/me/calendarView` | `Calendars.Read` |
| 17 | `get-specific-calendar-view` | `GET` | `/me/calendars/{calendar-id}/calendarView` | `Calendars.Read` |
| 18 | `get-schedule` | `POST` | `/me/calendar/getSchedule` | `Calendars.ReadBasic` |
| 19 | `list-shared-calendar-events` | `GET` | `/users/{user-id}/calendar/events` | `Calendars.Read.Shared` |
| 20 | `get-shared-calendar-view` | `GET` | `/users/{user-id}/calendarView` | `Calendars.Read.Shared` |
| 21 | `list-calendar-event-instances` | `GET` | `/me/calendars/{calendar-id}/events/{event-id}/instances` | `Calendars.Read` |
| 22 | `list-calendars` | `GET` | `/me/calendars` | `Calendars.Read` |
| 23 | `find-meeting-times` | `POST` | `/me/findMeetingTimes` | `Calendars.Read.Shared` |
| 24 | `get-current-user` | `GET` | `/me` | `User.Read` |
| 25 | `get-my-profile` | `GET` | `/me/profile` | `User.Read` |
| 26 | `get-mailbox-settings` | `GET` | `/me/mailboxSettings` | `MailboxSettings.Read` |
| 27 | `get-my-manager` | `GET` | `/me/manager` | `User.Read.All` |
| 28 | `list-my-direct-reports` | `GET` | `/me/directReports` | `User.Read.All` |
| 29 | `get-user-manager` | `GET` | `/users/{user-id}/manager` | `User.Read.All` |
| 30 | `list-user-direct-reports` | `GET` | `/users/{user-id}/directReports` | `User.Read.All` |
| 31 | `list-relevant-people` | `GET` | `/me/people` | `People.Read` |
| 32 | `list-my-memberships` | `GET` | `/me/memberOf` | `User.Read` |
| 33 | `list-groups` | `GET` | `/groups` | `Group.Read.All` |
| 34 | `get-group` | `GET` | `/groups/{group-id}` | `Group.Read.All` |
| 35 | `get-group-calendar-view` | `GET` | `/groups/{group-id}/calendarView` | `Group.Read.All` |
| 36 | `list-group-events` | `GET` | `/groups/{group-id}/events` | `Group.Read.All` |
| 37 | `list-group-members` | `GET` | `/groups/{group-id}/members` | `GroupMember.ReadBasic.All` |
| 38 | `list-group-owners` | `GET` | `/groups/{group-id}/owners` | `GroupMember.Read.All` |
| 39 | `list-chats` | `GET` | `/me/chats` | `Chat.ReadBasic` |
| 40 | `get-chat` | `GET` | `/chats/{chat-id}` | `Chat.ReadBasic` |
| 41 | `list-chat-members` | `GET` | `/chats/{chat-id}/members` | `Chat.ReadBasic` |
| 42 | `list-chat-messages` | `GET` | `/chats/{chat-id}/messages` | `Chat.Read` |
| 43 | `get-chat-message` | `GET` | `/chats/{chat-id}/messages/{chatMessage-id}` | `Chat.Read` |
| 44 | `list-chat-message-hosted-contents` | `GET` | `/chats/{chat-id}/messages/{chatMessage-id}/hostedContents` | `Chat.Read` |
| 45 | `get-team` | `GET` | `/teams/{team-id}` | `Team.ReadBasic.All` |
| 46 | `list-team-channels` | `GET` | `/teams/{team-id}/channels` | `Channel.ReadBasic.All` |
| 47 | `get-team-channel` | `GET` | `/teams/{team-id}/channels/{channel-id}` | `Channel.ReadBasic.All` |
| 48 | `list-channel-messages` | `GET` | `/teams/{team-id}/channels/{channel-id}/messages` | `ChannelMessage.Read.All` |
| 49 | `get-channel-message` | `GET` | `/teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}` | `ChannelMessage.Read.All` |
| 50 | `list-channel-message-hosted-contents` | `GET` | `/teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}/hostedContents` | `ChannelMessage.Read.All` |
| 51 | `list-channel-message-replies` | `GET` | `/teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}/replies` | `ChannelMessage.Read.All` |
| 52 | `list-channel-tabs` | `GET` | `/teams/{team-id}/channels/{channel-id}/tabs` | `TeamsTab.Read.All` |
| 53 | `list-team-members` | `GET` | `/teams/{team-id}/members` | `TeamMember.Read.All` |
| 54 | `get-channel-files-folder` | `GET` | `/teams/{team-id}/channels/{channel-id}/filesFolder` | `Files.Read.All` |
| 55 | `list-pinned-chat-messages` | `GET` | `/chats/{chat-id}/pinnedMessages` | `Chat.Read` |
| 56 | `search-sharepoint-sites` | `GET` | `/sites` | `Sites.Read.All` |
| 57 | `get-sharepoint-site` | `GET` | `/sites/{site-id}` | `(Sites.Read.All), (Sites.Selected)` |
| 58 | `list-sharepoint-site-drives` | `GET` | `/sites/{site-id}/drives` | `(Files.Read.All), (Sites.Read.All), (Sites.Selected)` |
| 59 | `get-sharepoint-site-drive-by-id` | `GET` | `/sites/{site-id}/drives/{drive-id}` | `(Files.Read.All), (Sites.Read.All), (Sites.Selected)` |
| 60 | `list-sharepoint-site-lists` | `GET` | `/sites/{site-id}/lists` | `(Sites.Read.All), (Sites.Selected)` |
| 61 | `get-sharepoint-site-list` | `GET` | `/sites/{site-id}/lists/{list-id}` | `(Sites.Read.All), (Sites.Selected)` |
| 62 | `list-sharepoint-site-list-items` | `GET` | `/sites/{site-id}/lists/{list-id}/items` | `(Sites.Read.All), (Sites.Selected)` |
| 63 | `get-sharepoint-site-list-item` | `GET` | `/sites/{site-id}/lists/{list-id}/items/{listItem-id}` | `(Sites.Read.All), (Sites.Selected)` |
| 64 | `list-sharepoint-list-columns` | `GET` | `/sites/{site-id}/lists/{list-id}/columns` | `(Sites.Read.All), (Sites.Selected)` |
| 65 | `get-sharepoint-list-column` | `GET` | `/sites/{site-id}/lists/{list-id}/columns/{columnDefinition-id}` | `(Sites.Read.All), (Sites.Selected)` |
| 66 | `get-sharepoint-site-by-path` | `GET` | `/sites/{site-id}:/{path}` | `(Sites.Read.All), (Sites.Selected)` |
| 67 | `search-query` | `POST` | `/search/query` | `(Mail.Read), (Calendars.Read), (Files.Read.All), (Sites.Read.All), (Chat.Read), (ChannelMessage.Read.All)` |
| 68 | `copilot-retrieve` | `POST` | `/copilot/retrieval` | `(Files.Read.All + Sites.Read.All), (ExternalItem.Read.All)` |
| 69 | `list-group-conversations` | `GET` | `/groups/{group-id}/conversations` | `Group-Conversation.Read.All` |
| 70 | `list-group-threads` | `GET` | `/groups/{group-id}/threads` | `Group-Conversation.Read.All` |
| 71 | `get-my-presence` | `GET` | `/me/presence` | `Presence.Read` |
| 72 | `get-user-presence` | `GET` | `/users/{user-id}/presence` | `Presence.Read.All` |
| 73 | `list-trending-insights` | `GET` | `/me/insights/trending` | `Sites.Read.All` |
| 74 | `list-calendar-events-delta` | `GET` | `/me/events/delta()` | `Calendars.ReadBasic` |
| 75 | `list-calendar-view-delta` | `GET` | `/me/calendarView/delta()` | `Calendars.Read` |
| 76 | `list-mail-folder-messages-delta` | `GET` | `/me/mailFolders/{mailFolder-id}/messages/delta()` | `Mail.Read` |

## Explicit POST allowlist

These POST operations are selected because their semantics are retrieval/read-only. No other POST is permitted.

- `graph-batch` → `POST /$batch` — special hardening required
- `get-schedule` → `POST /me/calendar/getSchedule` — readOnly=true
- `find-meeting-times` → `POST /me/findMeetingTimes` — readOnly=true
- `search-query` → `POST /search/query` — readOnly=true
- `copilot-retrieve` → `POST /copilot/retrieval` — readOnly=true

### `$batch` special rule

`graph-batch` is transport-level POST but MUST be constrained by code:
- every sub-request method MUST be GET;
- every sub-request URL MUST be relative;
- every sub-request path MUST match an implemented allowlisted GET endpoint;
- nested `$batch` MUST be rejected;
- caller-supplied Authorization/Host headers MUST be rejected;
- maximum 20 sub-requests;
- no write verb may be accepted by the schema, parser, validation, or implementation; negative tests MUST include POST/PATCH/PUT/DELETE etc. as hostile inputs and prove they are rejected.
