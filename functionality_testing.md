### Mail — own mailbox
- [10] list-mail-messages (GET /me/messages)
- [10] list-mail-folders (GET /me/mailFolders)
- [10] list-mail-folder-messages (GET /me/mailFolders/{mailFolder-id}/messages)
- [10] get-mail-message (GET /me/messages/{message-id})
- [ ] create-draft-email (POST /me/messages)
- [ ] delete-mail-message (DELETE /me/messages/{message-id})
- [ ] move-mail-message (POST /me/messages/{message-id}/move)
- [ ] add-mail-attachment (POST /me/messages/{message-id}/attachments)
- [ ] list-mail-attachments (GET /me/messages/{message-id}/attachments)
- [ ] get-mail-attachment (GET /me/messages/{message-id}/attachments/{attachment-id})
- [ ] delete-mail-attachment (DELETE /me/messages/{message-id}/attachments/{attachment-id})
- [ ] send-mail (POST /me/sendMail)

### Mail — shared mailboxes (requires org-mode)
- [ ] list-shared-mailbox-messages (GET /users/{user-id}/messages)
- [ ] list-shared-mailbox-folder-messages (GET /users/{user-id}/mailFolders/{mailFolder-id}/messages)
- [ ] get-shared-mailbox-message (GET /users/{user-id}/messages/{message-id})
- [ ] send-shared-mailbox-mail (POST /users/{user-id}/sendMail)

### Calendar — own
- [10] list-calendar-events (GET /me/events)
- [ ] get-calendar-event (GET /me/events/{event-id})
- [ ] create-calendar-event (POST /me/events)
- [ ] update-calendar-event (PATCH /me/events/{event-id})
- [ ] delete-calendar-event (DELETE /me/events/{event-id})
- [ ] list-specific-calendar-events (GET /me/calendars/{calendar-id}/events)
- [ ] get-specific-calendar-event (GET /me/calendars/{calendar-id}/events/{event-id})
- [ ] create-specific-calendar-event (POST /me/calendars/{calendar-id}/events)
- [ ] update-specific-calendar-event (PATCH /me/calendars/{calendar-id}/events/{event-id})
- [ ] delete-specific-calendar-event (DELETE /me/calendars/{calendar-id}/events/{event-id})
- [ ] get-calendar-view (GET /me/calendarView)
- [ ] list-calendars (GET /me/calendars)

### Calendar — helper (shared availability; requires org-mode)
- [ ] find-meeting-times (POST /me/findMeetingTimes)

### Cross-cutting
- [ ] search-query (POST /search/query) — verify it respects header-scoped user




- basic email interaction:
  - X find mails/read mails (Mostly doesnt work properly)
    - O over all accessible mailboxes
    - X in a specific mailbox
    - X by searchstring in subject or body or sender (usually malforms requests)
    - X by specific sender (usually malforms requests)
  - X summarize mails
  - O reply to mail (can work, if it already knows which mail is being referred to)
  - O forward mail (can work, if it already knows which mail is being referred to)
  - X generate draft (Bad Request)
  - O send mail (works but prone to error (your-email@example.com), works well when specifying direct recipient)
  - X delete mail (refuses)
  - X create mail folder (can list folders but not create)
  - X move mail to folder

- test calendar interaction:
  - O list events (today, this week, next week, by specific date, etc) (works relatively well)
  - X create event (incl. invitations, setting of reminders, teams meeting sync possible?, etc) (doesnt work, bad request)
  - X edit/move events (malformed request)
  - X delete events (malformed request)

- test contacts:
  - O listing contacts works
  - O sending mail to contact works (after listing them)


  phase 2:

  - basic email interaction:
  - 8 find mails/read mails (finds mails related to rob, finds my gehalts rechnungen)
    - O over all accessible mailboxes
    - 8 in a specific mailbox (finds mail in specific mail box, tested with archive)
    - 8 by searchstring in subject or body or sender (working pretty good now)
    - 8 by specific sender (seems to work)
  - 8 summarize mails (Works by summarizing each mail individually so far)
  - O reply to mail (can work, if it already knows which mail is being referred to)
  - O forward mail (can work, if it already knows which mail is being referred to)
  - 8 generate draft (works but doesnt set recipient, works with recipient now :D)
  - O send mail (works but prone to error (your-email@example.com), works well when specifying direct recipient)
  - 8 delete draft mail
  - 8 delete mail (works)
  - XX create mail folder (can list folders but not create) (Not supported)
  - -5/8 move mail to folder (works now, however before could move mail to non existant folders, thus "deleting" them)

- test calendar interaction:
  - O list events (today, this week, next week, by specific date, etc) (works relatively well)
  - O create event (incl. invitations, setting of reminders, teams meeting sync possible?, etc) (doesnt work, bad request)
  - O edit/move events (malformed request)
  - O delete events (malformed request)

- test contacts:
  - O listing contacts works
  - O sending mail to contact works (after listing them)