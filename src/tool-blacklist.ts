// Aliases we do not expose directly via MCP because they are wrapped by consolidated tools.
export const TOOL_BLACKLIST = [
  // Calendar (primary + specific variants)
  'get-calendar-event',
  'get-specific-calendar-event',
  'create-calendar-event',
  'create-specific-calendar-event',
  'update-calendar-event',
  'update-specific-calendar-event',
  'delete-calendar-event',
  'delete-specific-calendar-event',
  'list-calendar-events',
  'list-specific-calendar-events',

  // Mail (own mailbox, folders, shared mailboxes)
  'list-mail-messages',
  'list-mail-folder-messages',
  'list-shared-mailbox-messages',
  'list-shared-mailbox-folder-messages',
  'get-mail-message',
  'get-shared-mailbox-message',
  'send-mail',
  'send-shared-mailbox-mail',
] as const;

export type ToolBlacklistEntry = (typeof TOOL_BLACKLIST)[number];

