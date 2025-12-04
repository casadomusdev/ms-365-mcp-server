import { afterEach, beforeEach, describe, expect, it, vi, type MockInstance } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  },
}));

vi.mock('../src/tool-blacklist.js', () => ({
  TOOL_BLACKLIST: [],
}));

vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'list-mail-messages',
        method: 'GET',
        path: '/me/messages',
        description: 'List mail messages',
      },
      { alias: 'send-mail', method: 'POST', path: '/me/sendMail', description: 'Send mail' },
      {
        alias: 'list-calendar-events',
        method: 'GET',
        path: '/me/events',
        description: 'List calendar events',
      },
      {
        alias: 'list-excel-worksheets',
        method: 'GET',
        path: '/workbook/worksheets',
        description: 'List Excel worksheets',
      },
      { alias: 'get-current-user', method: 'GET', path: '/me', description: 'Get current user' },
    ],
  },
}));

const envBackup = {
  mail: process.env.MS365_MCP_ENABLE_MAIL,
  calendar: process.env.MS365_MCP_ENABLE_CALENDAR,
  files: process.env.MS365_MCP_ENABLE_FILES,
  excel: process.env.MS365_MCP_ENABLE_EXCEL_POWERPOINT,
  user: process.env.MS365_MCP_ENABLE_USER,
};

describe('Tool Filtering', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolSpy: MockInstance;

  beforeEach(() => {
    process.env.MS365_MCP_ENABLE_MAIL = 'true';
    process.env.MS365_MCP_ENABLE_CALENDAR = 'true';
    process.env.MS365_MCP_ENABLE_FILES = 'true';
    process.env.MS365_MCP_ENABLE_EXCEL_POWERPOINT = 'true';
    process.env.MS365_MCP_ENABLE_USER = 'true';

    server = new McpServer({ name: 'test', version: '1.0.0' });
    graphClient = {} as GraphClient;
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => ({} as any));
  });

  afterEach(() => {
    process.env.MS365_MCP_ENABLE_MAIL = envBackup.mail;
    process.env.MS365_MCP_ENABLE_CALENDAR = envBackup.calendar;
    process.env.MS365_MCP_ENABLE_FILES = envBackup.files;
    process.env.MS365_MCP_ENABLE_EXCEL_POWERPOINT = envBackup.excel;
    process.env.MS365_MCP_ENABLE_USER = envBackup.user;
  });

  it('should register all tools when no filter is provided', () => {
    registerGraphTools(server, graphClient, false);

    expect(toolSpy).toHaveBeenCalledTimes(5);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-mail-messages',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'send-mail',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-calendar-events',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-excel-worksheets',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'get-current-user',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
  });

  it('should filter tools by regex pattern - mail only', () => {
    registerGraphTools(server, graphClient, false, 'mail');

    expect(toolSpy).toHaveBeenCalledTimes(2);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-mail-messages',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'send-mail',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
  });

  it('should filter tools by regex pattern - calendar or excel', () => {
    registerGraphTools(server, graphClient, false, 'calendar|excel');

    expect(toolSpy).toHaveBeenCalledTimes(2);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-calendar-events',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-excel-worksheets',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
  });

  it('should handle invalid regex patterns gracefully', () => {
    registerGraphTools(server, graphClient, false, '[invalid regex');

    expect(toolSpy).toHaveBeenCalledTimes(5);
  });

  it('should combine read-only and filtering correctly', () => {
    registerGraphTools(server, graphClient, true, 'mail');

    expect(toolSpy).toHaveBeenCalledTimes(1);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-mail-messages',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
  });

  it('should register no tools when pattern matches nothing', () => {
    registerGraphTools(server, graphClient, false, 'nonexistent');

    expect(toolSpy).toHaveBeenCalledTimes(0);
  });
});
