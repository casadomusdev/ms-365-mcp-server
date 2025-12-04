import { afterEach, beforeEach, describe, expect, it, vi, type MockInstance } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { parseArgs } from '../src/cli.js';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/cli.js', () => {
  const parseArgsMock = vi.fn();
  return {
    parseArgs: parseArgsMock,
  };
});

vi.mock('../src/generated/client.js', () => {
  return {
    api: {
      endpoints: [
        {
          alias: 'list-mail-messages',
          method: 'get',
          path: '/me/messages',
          parameters: [],
        },
        {
          alias: 'list-mail-folder-messages',
          method: 'get',
          path: '/me/mailFolders/:mailFolderId/messages',
          parameters: [{ name: 'mailFolderId', type: 'Path', schema: {} }],
        },
        {
          alias: 'list-shared-mailbox-messages',
          method: 'get',
          path: '/users/:userId/messages',
          parameters: [{ name: 'userId', type: 'Path', schema: {} }],
        },
        {
          alias: 'list-shared-mailbox-folder-messages',
          method: 'get',
          path: '/users/:userId/mailFolders/:mailFolderId/messages',
          parameters: [
            { name: 'userId', type: 'Path', schema: {} },
            { name: 'mailFolderId', type: 'Path', schema: {} },
          ],
        },
        {
          alias: 'get-mail-message',
          method: 'get',
          path: '/me/messages/:messageId',
          parameters: [{ name: 'messageId', type: 'Path', schema: {} }],
        },
        {
          alias: 'get-shared-mailbox-message',
          method: 'get',
          path: '/users/:userId/messages/:messageId',
          parameters: [
            { name: 'userId', type: 'Path', schema: {} },
            { name: 'messageId', type: 'Path', schema: {} },
          ],
        },
        {
          alias: 'send-mail',
          method: 'post',
          path: '/me/sendMail',
          parameters: [{ name: 'body', type: 'Body', schema: {} }],
        },
        {
          alias: 'send-shared-mailbox-mail',
          method: 'post',
          path: '/users/:userId/sendMail',
          parameters: [
            { name: 'userId', type: 'Path', schema: {} },
            { name: 'body', type: 'Body', schema: {} },
          ],
        },
        {
          alias: 'delete-mail-message',
          method: 'delete',
          path: '/me/messages/{message-id}',
          parameters: [],
        },
      ],
    },
  };
});

vi.mock('../src/logger.js', () => {
  return {
    default: {
      info: vi.fn(),
      warn: vi.fn(),
      error: vi.fn(),
      debug: vi.fn(),
    },
  };
});

// Mock fs.readFileSync for endpoints.json
vi.mock('fs', async () => {
  const actualFs = await vi.importActual<typeof import('fs')>('fs');
  return {
    ...actualFs,
    readFileSync: vi.fn((filePath: string, ...args: any[]) => {
      // Return mock endpoints.json data when graph-tools tries to load it
      if (filePath.includes('endpoints.json')) {
        return JSON.stringify([
          {
            pathPattern: '/me/messages',
            method: 'get',
            toolName: 'list-mail-messages',
            scopes: ['Mail.Read'],
          },
          {
            pathPattern: '/me/mailFolders/{mailFolder-id}/messages',
            method: 'get',
            toolName: 'list-mail-folder-messages',
            scopes: ['Mail.Read'],
          },
          {
            pathPattern: '/users/{user-id}/messages',
            method: 'get',
            toolName: 'list-shared-mailbox-messages',
            scopes: ['Mail.Read.Shared'],
          },
          {
            pathPattern: '/users/{user-id}/mailFolders/{mailFolder-id}/messages',
            method: 'get',
            toolName: 'list-shared-mailbox-folder-messages',
            scopes: ['Mail.Read.Shared'],
          },
          {
            pathPattern: '/me/messages/{message-id}',
            method: 'get',
            toolName: 'get-mail-message',
            scopes: ['Mail.Read'],
          },
          {
            pathPattern: '/users/{user-id}/messages/{message-id}',
            method: 'get',
            toolName: 'get-shared-mailbox-message',
            scopes: ['Mail.Read.Shared'],
          },
          {
            pathPattern: '/me/sendMail',
            method: 'post',
            toolName: 'send-mail',
            scopes: ['Mail.Send'],
          },
          {
            pathPattern: '/users/{user-id}/sendMail',
            method: 'post',
            toolName: 'send-shared-mailbox-mail',
            scopes: ['Mail.Send.Shared'],
          },
          {
            pathPattern: '/me/messages/{message-id}',
            method: 'delete',
            toolName: 'delete-mail-message',
            scopes: ['Mail.ReadWrite'],
          },
        ]);
      }
      // For other files, use the actual fs
      return actualFs.readFileSync(filePath, ...args);
    }),
  };
});

describe('Read-Only Mode', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolSpy: MockInstance;

  beforeEach(() => {
    vi.clearAllMocks();

    delete process.env.READ_ONLY;
    process.env.MS365_MCP_ENABLE_MAIL = 'true';

    server = new McpServer({ name: 'test', version: '1.0.0' });
    graphClient = {
      graphRequest: vi.fn().mockResolvedValue({
        content: [{ text: JSON.stringify({ value: [] }) }],
      }),
    } as unknown as GraphClient;
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => ({} as any));
  });

  afterEach(() => {
    delete process.env.MS365_MCP_ENABLE_MAIL;
    vi.resetAllMocks();
  });


  it('should respect --read-only flag from CLI', () => {
    vi.mocked(parseArgs).mockReturnValue({ readOnly: true } as ReturnType<typeof parseArgs>);

    const options = parseArgs();
    expect(options.readOnly).toBe(true);

    registerGraphTools(server, graphClient, options.readOnly, undefined, false);

    // In read-only mode, only GET operations should be registered
    // Wrappers: list-mail-messages (GET), get-mail-message (GET)
    // Direct: none (delete-mail-message is DELETE, filtered out)
    expect(toolSpy).toHaveBeenCalledTimes(2);

    const toolCalls = toolSpy.mock.calls.map((call: unknown[]) => call[0]);
    expect(toolCalls).toContain('list-mail-messages');
    expect(toolCalls).toContain('get-mail-message');
    expect(toolCalls).not.toContain('send-mail');
    expect(toolCalls).not.toContain('delete-mail-message');
  });

  it('should register all endpoints when not in read-only mode', () => {
    vi.mocked(parseArgs).mockReturnValue({ readOnly: false } as ReturnType<typeof parseArgs>);

    const options = parseArgs();
    expect(options.readOnly).toBe(false);

    registerGraphTools(server, graphClient, options.readOnly, undefined, false);

    // In non-read-only mode, all operations should be registered
    // Wrappers: list-mail-messages (GET), get-mail-message (GET), send-mail (POST)
    // Direct: delete-mail-message (DELETE)
    expect(toolSpy).toHaveBeenCalledTimes(4);

    const toolCalls = toolSpy.mock.calls.map((call: unknown[]) => call[0]);
    expect(toolCalls).toContain('list-mail-messages');
    expect(toolCalls).toContain('get-mail-message');
    expect(toolCalls).toContain('send-mail');
    expect(toolCalls).toContain('delete-mail-message');
  });
});
