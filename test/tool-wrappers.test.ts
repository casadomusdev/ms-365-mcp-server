import { afterEach, beforeEach, describe, expect, it, vi, type MockInstance } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
  },
}));

const ENDPOINTS: any[] = [];

vi.mock('../src/generated/client.js', () => ({
  api: {
    get endpoints() {
      return ENDPOINTS;
    },
  },
}));

const pathParam = (name: string) => ({
  name,
  type: 'Path',
  schema: z.string(),
});

const bodyParam = () => ({
  name: 'body',
  type: 'Body',
  schema: z.object({ subject: z.string().optional() }),
});

const minimalResponse = {
  content: [{ text: JSON.stringify({ value: [] }) }],
};

const envBackup = {
  mail: process.env.MS365_MCP_ENABLE_MAIL,
  calendar: process.env.MS365_MCP_ENABLE_CALENDAR,
};

const buildEndpoints = () => {
  ENDPOINTS.length = 0;
  ENDPOINTS.push(
    { alias: 'list-mail-messages', method: 'GET', path: '/me/messages', description: 'list mail' },
    {
      alias: 'list-mail-folder-messages',
      method: 'GET',
      path: '/me/mailFolders/:mailFolderId/messages',
      description: 'folder mail',
      parameters: [pathParam('mailFolderId')],
    },
    {
      alias: 'list-shared-mailbox-messages',
      method: 'GET',
      path: '/users/:userId/messages',
      description: 'shared mail',
      parameters: [pathParam('userId')],
    },
    {
      alias: 'list-shared-mailbox-folder-messages',
      method: 'GET',
      path: '/users/:userId/mailFolders/:mailFolderId/messages',
      description: 'shared folder mail',
      parameters: [pathParam('userId'), pathParam('mailFolderId')],
    },
    {
      alias: 'get-mail-message',
      method: 'GET',
      path: '/me/messages/:messageId',
      description: 'get mail',
      parameters: [pathParam('messageId')],
    },
    {
      alias: 'get-shared-mailbox-message',
      method: 'GET',
      path: '/users/:userId/messages/:messageId',
      description: 'get shared mail',
      parameters: [pathParam('userId'), pathParam('messageId')],
    },
    {
      alias: 'send-mail',
      method: 'POST',
      path: '/me/sendMail',
      description: 'send mail',
      parameters: [bodyParam()],
    },
    {
      alias: 'send-shared-mailbox-mail',
      method: 'POST',
      path: '/users/:userId/sendMail',
      description: 'send shared mail',
      parameters: [pathParam('userId'), bodyParam()],
    },
    { alias: 'list-calendar-events', method: 'GET', path: '/me/events', description: 'list events' },
    {
      alias: 'list-specific-calendar-events',
      method: 'GET',
      path: '/me/calendars/:calendarId/events',
      description: 'list specific events',
      parameters: [pathParam('calendarId')],
    },
    {
      alias: 'get-calendar-event',
      method: 'GET',
      path: '/me/events/:eventId',
      description: 'get event',
      parameters: [pathParam('eventId')],
    },
    {
      alias: 'get-specific-calendar-event',
      method: 'GET',
      path: '/me/calendars/:calendarId/events/:eventId',
      description: 'get specific event',
      parameters: [pathParam('calendarId'), pathParam('eventId')],
    },
    {
      alias: 'create-calendar-event',
      method: 'POST',
      path: '/me/events',
      description: 'create event',
      parameters: [bodyParam()],
    },
    {
      alias: 'create-specific-calendar-event',
      method: 'POST',
      path: '/me/calendars/:calendarId/events',
      description: 'create specific event',
      parameters: [pathParam('calendarId'), bodyParam()],
    },
    {
      alias: 'update-calendar-event',
      method: 'PATCH',
      path: '/me/events/:eventId',
      description: 'update event',
      parameters: [pathParam('eventId'), bodyParam()],
    },
    {
      alias: 'update-specific-calendar-event',
      method: 'PATCH',
      path: '/me/calendars/:calendarId/events/:eventId',
      description: 'update specific event',
      parameters: [pathParam('calendarId'), pathParam('eventId'), bodyParam()],
    },
    {
      alias: 'delete-calendar-event',
      method: 'DELETE',
      path: '/me/events/:eventId',
      description: 'delete event',
      parameters: [pathParam('eventId')],
    },
    {
      alias: 'delete-specific-calendar-event',
      method: 'DELETE',
      path: '/me/calendars/:calendarId/events/:eventId',
      description: 'delete specific event',
      parameters: [pathParam('calendarId'), pathParam('eventId')],
    }
  );
};

describe('Outlook wrapper consolidation', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolSpy: MockInstance;

  beforeEach(() => {
    buildEndpoints();
    process.env.MS365_MCP_ENABLE_MAIL = 'true';
    process.env.MS365_MCP_ENABLE_CALENDAR = 'true';
    server = new McpServer({ name: 'test', version: '1.0.0' });
    graphClient = {
      graphRequest: vi.fn().mockResolvedValue(minimalResponse),
    } as unknown as GraphClient;
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => ({} as any));
  });

  afterEach(() => {
    process.env.MS365_MCP_ENABLE_MAIL = envBackup.mail;
    process.env.MS365_MCP_ENABLE_CALENDAR = envBackup.calendar;
  });

  it('registers wrappers instead of the raw aliases', () => {
    registerGraphTools(server, graphClient, false, undefined, true);

    const registeredNames = toolSpy.mock.calls.map((call) => call[0]);

    expect(registeredNames).toContain('list-mail-messages');
    expect(registeredNames).toContain('get-mail-message');
    expect(registeredNames).toContain('send-mail');
    expect(registeredNames).toContain('list-calendar-events');
    expect(registeredNames).toContain('get-calendar-event');
    expect(registeredNames).toContain('create-calendar-event');
    expect(registeredNames).toContain('update-calendar-event');
    expect(registeredNames).toContain('delete-calendar-event');

    expect(registeredNames).not.toContain('list-mail-folder-messages');
    expect(registeredNames).not.toContain('list-shared-mailbox-messages');
    expect(registeredNames).not.toContain('get-shared-mailbox-message');
    expect(registeredNames).not.toContain('send-shared-mailbox-mail');
    expect(registeredNames).not.toContain('list-specific-calendar-events');
    expect(registeredNames).not.toContain('get-specific-calendar-event');
  });

  it('routes list-mail-messages to the correct internal alias', async () => {
    registerGraphTools(server, graphClient, false, undefined, true);
    const listMailCall = toolSpy.mock.calls.find((call) => call[0] === 'list-mail-messages');
    expect(listMailCall).toBeDefined();
    const handler = listMailCall![4] as (params: any) => Promise<unknown>;

    const requestMock = graphClient.graphRequest as any;

    // Personal mailbox root
    await handler({});
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/messages');

    // Personal mailbox folder
    await handler({ folderId: 'A' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/mailFolders/A/messages');

    // Shared mailbox (by email)
    await handler({ sharedMailboxEmail: 'shared@example.com' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/users/shared%40example.com/messages');

    // Shared mailbox folder (by ID)
    await handler({ folderId: 'B', sharedMailboxId: 'guid-123' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/users/guid-123/mailFolders/B/messages');

    // Shared mailbox folder (by email, ID takes precedence)
    await handler({ folderId: 'C', sharedMailboxId: 'guid-456', sharedMailboxEmail: 'other@example.com' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/users/guid-456/mailFolders/C/messages');
  });

  it('routes get-mail-message to the correct internal alias', async () => {
    registerGraphTools(server, graphClient, false, undefined, true);
    const getMailCall = toolSpy.mock.calls.find((call) => call[0] === 'get-mail-message');
    expect(getMailCall).toBeDefined();
    const handler = getMailCall![4] as (params: any) => Promise<unknown>;

    const requestMock = graphClient.graphRequest as any;

    // Personal mailbox
    await handler({ messageId: 'msg-123' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/messages/msg-123');

    // Shared mailbox (by email)
    await handler({ messageId: 'msg-456', sharedMailboxEmail: 'shared@example.com' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/users/shared%40example.com/messages/msg-456');

    // Shared mailbox (by ID)
    await handler({ messageId: 'msg-789', sharedMailboxId: 'guid-123' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/users/guid-123/messages/msg-789');

    // Shared mailbox (ID takes precedence over email)
    await handler({ 
      messageId: 'msg-999', 
      sharedMailboxId: 'guid-456', 
      sharedMailboxEmail: 'other@example.com' 
    });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/users/guid-456/messages/msg-999');
  });

  it('routes send-mail to the correct internal alias', async () => {
    registerGraphTools(server, graphClient, false, undefined, true);
    const sendMailCall = toolSpy.mock.calls.find((call) => call[0] === 'send-mail');
    expect(sendMailCall).toBeDefined();
    const handler = sendMailCall![4] as (params: any) => Promise<unknown>;

    const requestMock = graphClient.graphRequest as any;

    // Personal mailbox
    await handler({ body: { subject: 'Test' } });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/sendMail');

    // Shared mailbox (by email)
    await handler({ body: { subject: 'Test' }, sharedMailboxEmail: 'shared@example.com' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/users/shared%40example.com/sendMail');

    // Shared mailbox (by ID)
    await handler({ body: { subject: 'Test' }, sharedMailboxId: 'guid-123' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/users/guid-123/sendMail');

    // Shared mailbox (ID takes precedence)
    await handler({ 
      body: { subject: 'Test' }, 
      sharedMailboxId: 'guid-456', 
      sharedMailboxEmail: 'other@example.com' 
    });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/users/guid-456/sendMail');
  });

  it('routes calendar wrappers to the correct internal alias', async () => {
    registerGraphTools(server, graphClient, false, undefined, true);
    const requestMock = graphClient.graphRequest as any;

    // Test list-calendar-events
    const listEventsCall = toolSpy.mock.calls.find((call) => call[0] === 'list-calendar-events');
    expect(listEventsCall).toBeDefined();
    const listHandler = listEventsCall![4] as (params: any) => Promise<unknown>;

    await listHandler({});
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/events');

    await listHandler({ calendarId: 'cal-123' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/calendars/cal-123/events');

    // Test get-calendar-event
    const getEventCall = toolSpy.mock.calls.find((call) => call[0] === 'get-calendar-event');
    expect(getEventCall).toBeDefined();
    const getHandler = getEventCall![4] as (params: any) => Promise<unknown>;

    await getHandler({ eventId: 'evt-123' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/events/evt-123');

    await getHandler({ eventId: 'evt-456', calendarId: 'cal-789' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/calendars/cal-789/events/evt-456');

    // Test create-calendar-event
    const createEventCall = toolSpy.mock.calls.find((call) => call[0] === 'create-calendar-event');
    expect(createEventCall).toBeDefined();
    const createHandler = createEventCall![4] as (params: any) => Promise<unknown>;

    await createHandler({ body: { subject: 'Test' } });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/events');

    await createHandler({ body: { subject: 'Test' }, calendarId: 'cal-123' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/calendars/cal-123/events');

    // Test update-calendar-event
    const updateEventCall = toolSpy.mock.calls.find((call) => call[0] === 'update-calendar-event');
    expect(updateEventCall).toBeDefined();
    const updateHandler = updateEventCall![4] as (params: any) => Promise<unknown>;

    await updateHandler({ eventId: 'evt-123', body: { subject: 'Updated' } });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/events/evt-123');

    await updateHandler({ eventId: 'evt-456', calendarId: 'cal-789', body: { subject: 'Updated' } });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/calendars/cal-789/events/evt-456');

    // Test delete-calendar-event
    const deleteEventCall = toolSpy.mock.calls.find((call) => call[0] === 'delete-calendar-event');
    expect(deleteEventCall).toBeDefined();
    const deleteHandler = deleteEventCall![4] as (params: any) => Promise<unknown>;

    await deleteHandler({ eventId: 'evt-123' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/events/evt-123');

    await deleteHandler({ eventId: 'evt-456', calendarId: 'cal-789' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/calendars/cal-789/events/evt-456');
  });

  it('wrapper schemas include optional routing parameters', () => {
    registerGraphTools(server, graphClient, false, undefined, true);

    // Check list-mail-messages schema (schema is a plain object with Zod schemas as values)
    const listMailCall = toolSpy.mock.calls.find((call) => call[0] === 'list-mail-messages');
    expect(listMailCall).toBeDefined();
    const listMailSchema = listMailCall![2] as Record<string, z.ZodTypeAny>;
    expect(listMailSchema.folderId).toBeDefined();
    expect(listMailSchema.sharedMailboxId).toBeDefined();
    expect(listMailSchema.sharedMailboxEmail).toBeDefined();
    // Check that folderId is optional and has description
    const folderIdSchema = listMailSchema.folderId as z.ZodOptional<z.ZodString>;
    expect(folderIdSchema._def.description).toContain('Optional');
    const sharedMailboxIdSchema = listMailSchema.sharedMailboxId as z.ZodOptional<z.ZodString>;
    expect(sharedMailboxIdSchema._def.description).toContain('Optional');

    // Check get-mail-message schema
    const getMailCall = toolSpy.mock.calls.find((call) => call[0] === 'get-mail-message');
    expect(getMailCall).toBeDefined();
    const getMailSchema = getMailCall![2] as Record<string, z.ZodTypeAny>;
    expect(getMailSchema.sharedMailboxId).toBeDefined();
    expect(getMailSchema.sharedMailboxEmail).toBeDefined();

    // Check send-mail schema
    const sendMailCall = toolSpy.mock.calls.find((call) => call[0] === 'send-mail');
    expect(sendMailCall).toBeDefined();
    const sendMailSchema = sendMailCall![2] as Record<string, z.ZodTypeAny>;
    expect(sendMailSchema.sharedMailboxId).toBeDefined();
    expect(sendMailSchema.sharedMailboxEmail).toBeDefined();

    // Check calendar wrapper schemas
    const getEventCall = toolSpy.mock.calls.find((call) => call[0] === 'get-calendar-event');
    expect(getEventCall).toBeDefined();
    const getEventSchema = getEventCall![2] as Record<string, z.ZodTypeAny>;
    expect(getEventSchema.calendarId).toBeDefined();
    const calendarIdSchema = getEventSchema.calendarId as z.ZodOptional<z.ZodString>;
    expect(calendarIdSchema._def.description).toContain('Optional');

    const listEventsCall = toolSpy.mock.calls.find((call) => call[0] === 'list-calendar-events');
    expect(listEventsCall).toBeDefined();
    const listEventsSchema = listEventsCall![2] as Record<string, z.ZodTypeAny>;
    expect(listEventsSchema.calendarId).toBeDefined();
  });

  it('blacklisted tools are not registered', () => {
    registerGraphTools(server, graphClient, false, undefined, true);
    const registeredNames = toolSpy.mock.calls.map((call) => call[0]);

    // Verify all blacklisted tools are NOT registered
    const blacklisted = [
      'list-mail-folder-messages',
      'list-shared-mailbox-messages',
      'list-shared-mailbox-folder-messages',
      'get-shared-mailbox-message',
      'send-shared-mailbox-mail',
      'list-specific-calendar-events',
      'get-specific-calendar-event',
      'create-specific-calendar-event',
      'update-specific-calendar-event',
      'delete-specific-calendar-event',
    ];

    for (const alias of blacklisted) {
      expect(registeredNames).not.toContain(alias);
    }
  });

  it('handles empty or whitespace-only optional params correctly', async () => {
    registerGraphTools(server, graphClient, false, undefined, true);
    const requestMock = graphClient.graphRequest as any;

    // Test calendarId with empty string
    const getEventCall = toolSpy.mock.calls.find((call) => call[0] === 'get-calendar-event');
    const getHandler = getEventCall![4] as (params: any) => Promise<unknown>;

    await getHandler({ eventId: 'evt-123', calendarId: '' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/events/evt-123');

    await getHandler({ eventId: 'evt-123', calendarId: '   ' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/events/evt-123');

    // Test sharedMailboxEmail with empty string
    const listMailCall = toolSpy.mock.calls.find((call) => call[0] === 'list-mail-messages');
    const listHandler = listMailCall![4] as (params: any) => Promise<unknown>;

    await listHandler({ sharedMailboxEmail: '' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/messages');

    await listHandler({ sharedMailboxId: '', sharedMailboxEmail: '   ' });
    expect(requestMock.mock.calls.at(-1)?.[0]).toBe('/me/messages');
  });
});

