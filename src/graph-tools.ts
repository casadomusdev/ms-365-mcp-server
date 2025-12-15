import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { CallToolRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import logger from './logger.js';
import GraphClient from './graph-client.js';
import { api } from './generated/client.js';
import { z } from 'zod';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import crypto from 'crypto';
import { getToolDescription } from './tool-descriptions.js';
import { ImpersonationContext, MailboxDiscoveryCache } from './impersonation/index.js';
import { TOOL_BLACKLIST } from './tool-blacklist.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function sanitizeAqsSearch(value: string): string {
  // AQS treats '-' as operator; unquoted tokens like CE-I6 error. Strategy:
  // - If it already looks like an advanced query (contains ':', boolean ops, parentheses, or quotes), leave as-is
  // - Else, if it contains any non-alphanumeric, wrap the whole term in double quotes
  // - Escape any embedded double quotes
  const raw = String(value ?? '').trim();
  if (raw.length === 0) return raw;

  const hasAdvancedSyntax = /[\(\)]/.test(raw) || /\bAND\b|\bOR\b|\bNOT\b/i.test(raw) || raw.includes(':') || raw.includes('"');
  if (hasAdvancedSyntax) {
    return raw;
  }

  const needsQuoting = /[^A-Za-z0-9]/.test(raw);
  if (!needsQuoting) {
    return raw;
  }

  const escaped = raw.replace(/"/g, '\\"');
  return `"${escaped}"`;
}

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
  returnDownloadUrl?: boolean;
}

const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

type TextContent = {
  type: 'text';
  text: string;
  [key: string]: unknown;
};

type ImageContent = {
  type: 'image';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type AudioContent = {
  type: 'audio';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type ResourceTextContent = {
  type: 'resource';
  resource: {
    text: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceBlobContent = {
  type: 'resource';
  resource: {
    blob: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceContent = ResourceTextContent | ResourceBlobContent;

type ContentItem = TextContent | ImageContent | AudioContent | ResourceContent;

interface CallToolResult {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

type ToolMetadata = {
  title: string;
  readOnlyHint: boolean;
};

interface GraphToolDefinition {
  alias: string;
  safeName: string;
  description: string;
  paramSchema: Record<string, unknown>;
  metadata: ToolMetadata;
  handler: (params: any) => Promise<CallToolResult>;
}

export function registerGraphTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  enabledToolsPattern?: string,
  orgMode: boolean = false
): void {
  const stripHtml = (html: string): string =>
    String(html ?? '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<[^>]+>/g, '')
      .replace(/\s+\n/g, '\n')
      .trim();

  const scrubBodies = (
    json: any,
    includeBody: boolean,
    bodyFormat?: 'html' | 'text',
    previewCap: number = 800
  ): void => {
    if (!json || typeof json !== 'object') return;

    const processOne = (item: any) => {
      if (!item || typeof item !== 'object') return;
      const hasBody = item.body && typeof item.body === 'object';
      if (!hasBody) return;

      // Ensure bodyPreview exists and is short/plain by default
      const original =
        typeof item.body.content === 'string' ? String(item.body.content) : undefined;
      if (!includeBody) {
        const preview =
          item.bodyPreview && typeof item.bodyPreview === 'string'
            ? String(item.bodyPreview)
            : original
            ? stripHtml(original)
            : '';
        item.bodyPreview = preview.slice(0, previewCap);
        if (item.body && 'content' in item.body) {
          delete item.body.content;
        }
      } else if (original && bodyFormat === 'text') {
        // Convert HTML to plain text when requested
        item.body = {
          contentType: 'Text',
          content: stripHtml(original),
        };
      }
    };

    if (Array.isArray(json.value)) {
      for (const it of json.value) processOne(it);
    } else {
      processOne(json);
    }
  };
  const parseBool = (val: string | undefined, dflt = true): boolean => {
    if (val == null) return dflt;
    const v = val.toLowerCase();
    return v === '1' || v === 'true' || v === 'yes' || v === 'on';
  };

  // Feature toggles (default true)
  // Disabled by default; opt-in via env
  const enableMail = parseBool(process.env.MS365_MCP_ENABLE_MAIL, false);
  const enableCalendar = parseBool(process.env.MS365_MCP_ENABLE_CALENDAR, false);
  const enableFiles = parseBool(process.env.MS365_MCP_ENABLE_FILES, false); // OneDrive + SharePoint
  const enableTeams = parseBool(process.env.MS365_MCP_ENABLE_TEAMS, false); // Chats + Channels
  const enableExcelPpt = parseBool(process.env.MS365_MCP_ENABLE_EXCEL_POWERPOINT, false);
  const enableOneNote = parseBool(process.env.MS365_MCP_ENABLE_ONENOTE, false);
  const enableTasks = parseBool(process.env.MS365_MCP_ENABLE_TASKS, false); // To Do + Planner
  const enableContacts = parseBool(process.env.MS365_MCP_ENABLE_CONTACTS, false);
  const enableUser = parseBool(process.env.MS365_MCP_ENABLE_USER, false);
  const enableSearch = parseBool(process.env.MS365_MCP_ENABLE_SEARCH, false);

  const getCategoryEnabled = (alias: string, path: string): boolean => {
    const p = path.toLowerCase();
    const a = alias.toLowerCase();

    // Mail
    if (
      p.includes('/me/messages') ||
      p.includes('/users/:userid/messages') || // generated client path style
      p.includes('/mailfolders') ||
      p.includes('/sendmail') ||
      p.includes('/attachments')
    ) {
      return enableMail;
    }

    // Calendar
    if (
      p.includes('/events') ||
      p.includes('/calendarview') ||
      p.includes('/calendars') ||
      p.includes('/findmeetingtimes')
    ) {
      return enableCalendar;
    }

    // Files: OneDrive + SharePoint
    if (p.includes('/drives') || p.startsWith('/sites')) {
      return enableFiles;
    }

    // Excel (workbook endpoints live under /drives/.../workbook)
    if (p.includes('/workbook/')) {
      return enableExcelPpt;
    }

    // OneNote
    if (p.includes('/onenote/')) {
      return enableOneNote;
    }

    // Tasks (To Do + Planner)
    if (p.includes('/todo/') || p.includes('/planner/')) {
      return enableTasks;
    }

    // Teams (Chats/Channels)
    if (
      p.includes('/chats') || // covers /me/chats and /chats
      p.includes('/joinedteams') ||
      p.startsWith('/teams')
    ) {
      return enableTeams;
    }

    // Contacts
    if (p.includes('/contacts')) {
      return enableContacts;
    }

    // User info (get-current-user)
    if (p === '/me' || p === '/users') {
      return enableUser;
    }

    // Cross-cutting search
    if (p === '/search/query') {
      return enableSearch;
    }

    // Default: disabled (explicit opt-in via env for known groups only)
    return false;
  };
  const createSafeToolName = (raw: string): string => {
    // Sanitize to allowed chars (letters, numbers, underscore, dash), lowercased
    const sanitized = raw
      .toLowerCase()
      .replace(/[^a-z0-9_-]/g, '_')
      .replace(/_+/g, '_')
      .replace(/^-+/, '')
      .replace(/_+$/g, '');

    // Leave ample headroom for clients that prepend GUID prefixes (e.g., 37 chars incl. underscore); cap at 24
    const MAX = 24;
    if (sanitized.length <= MAX) return sanitized;

    const hash = crypto.createHash('md5').update(sanitized).digest('hex').slice(0, 8);
    const baseMax = MAX - 1 - hash.length; // room for '-' + hash
    const base = sanitized.slice(0, Math.max(0, baseMax));
    return `${base}-${hash}`;
  };
  let enabledToolsRegex: RegExp | undefined;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
    }
  }

  const toolDefinitions = new Map<string, GraphToolDefinition>();
  const sanitizedTools: Array<{ original: string; sanitized: string }> = [];
  const rsvpEnabledAliases = new Set(['update-calendar-event', 'update-specific-calendar-event']);
  const recentEventCache = new Map<string, Set<string>>();
  const responseActionEnum = z
    .enum(['accept', 'decline', 'tentative'])
    .describe(
      'RSVP action to take. Preferred way to accept/decline/tentatively accept without providing attendees.'
    );
  const responseStatusEnum = z
    .enum(['accepted', 'declined', 'tentativelyAccepted'])
    .describe('Graph RSVP status values (accepted / declined / tentativelyAccepted).');
  const rememberEventsForUser = (userEmail: string | undefined, eventIds: string[]) => {
    if (!userEmail || eventIds.length === 0) return;
    const normalized = userEmail.toLowerCase();
    const existing = recentEventCache.get(normalized) ?? new Set<string>();
    for (const id of eventIds) {
      if (typeof id === 'string' && id.trim().length > 0) {
        existing.add(id);
      }
    }
    recentEventCache.set(normalized, existing);
  };

  const annotateMessageObject = (message: any, mailboxKey: string | undefined, ordinal?: number): void => {
    if (!message || typeof message !== 'object') {
      return;
    }
    if (typeof message.id === 'string' && message.id.trim().length > 0) {
      message.messageIdToUse = message.id;
      message._useThisMessageIdForActions = message.id;
    }
    if (typeof ordinal === 'number') {
      message._messageIndex = ordinal;
    }
    if (mailboxKey) {
      message._mailboxKeyUsed = mailboxKey;
    }
    if (typeof message.parentFolderId === 'string') {
      message._parentFolderIdHint = message.parentFolderId;
    }
    if (Object.prototype.hasOwnProperty.call(message, 'isDraft')) {
      message._isDraftMessage = message.isDraft === true;
    }
    if (typeof message.lastModifiedDateTime === 'string') {
      message._lastModifiedDateTime = message.lastModifiedDateTime;
    }
    message._selectionHint = buildMessageSelectionHint(message, ordinal);
  };

  const buildMessageSelectionHint = (message: any, ordinal?: number): string => {
    const parts: string[] = [];
    if (typeof ordinal === 'number') {
      parts.push(`#${ordinal}`);
    }

    const stateLabels: string[] = [];
    if (message.isDraft === true) {
      stateLabels.push('draft');
    }
    if (message.isRead === false) {
      stateLabels.push('unread');
    }
    if (stateLabels.length === 0) {
      stateLabels.push('message');
    }
    parts.push(stateLabels.join('+'));

    const subject =
      typeof message.subject === 'string' && message.subject.trim().length > 0
        ? `subject "${message.subject.slice(0, 60)}${message.subject.length > 60 ? '…' : ''}"`
        : 'subject <none>';
    parts.push(subject);

    const lastModified =
      message.lastModifiedDateTime || message.createdDateTime || message.receivedDateTime;
    if (typeof lastModified === 'string') {
      parts.push(`lastModified ${lastModified}`);
    }

    let attachmentInfo: string | undefined;
    if (Array.isArray(message.attachments)) {
      attachmentInfo = `attachments:${message.attachments.length}`;
    } else if (typeof message.hasAttachments === 'boolean') {
      attachmentInfo = `attachments:${message.hasAttachments ? 'yes' : 'no'}`;
    }
    if (attachmentInfo) {
      parts.push(attachmentInfo);
    }

    const truncatedId =
      typeof message.messageIdToUse === 'string'
        ? message.messageIdToUse.slice(0, 16)
        : typeof message.id === 'string'
        ? message.id.slice(0, 16)
        : undefined;
    if (truncatedId) {
      parts.push(`id ${truncatedId}${truncatedId.length === 16 ? '…' : ''}`);
    }

    return parts.join(' | ');
  };

  const resolveMailboxKey = (
    impersonated: string | undefined,
    params: Record<string, unknown>,
    path: string,
    resolvedPathParams: Record<string, string>
  ): string | undefined => {
    const paramObject = params as Record<string, unknown>;
    const userParam =
      (typeof paramObject.userId === 'string' && paramObject.userId) ||
      (typeof paramObject['user-id'] === 'string' && (paramObject['user-id'] as string)) ||
      resolvedPathParams.userId;
    if (typeof userParam === 'string' && userParam.trim().length > 0) {
      return userParam.toLowerCase();
    }
    const userMatch = path.match(/\/users\/([^/?]+)/i);
    if (userMatch && typeof userMatch[1] === 'string') {
      try {
        return decodeURIComponent(userMatch[1]).toLowerCase();
      } catch {
        return userMatch[1].toLowerCase();
      }
    }
    return impersonated?.toLowerCase();
  };

  const mailListAliases = new Set([
    'list-mail-messages',
    'list-mail-folder-messages',
    'list-shared-mailbox-messages',
    'list-shared-mailbox-folder-messages',
  ]);
  const mailSingleMessageAliases = new Set(['get-mail-message', 'get-shared-mailbox-message']);

  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      logger.info(`Skipping work account tool ${tool.alias} - not in org mode`);
      continue;
    }

    if (readOnly && tool.method.toUpperCase() !== 'GET') {
      logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
      continue;
    }

    if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
      logger.info(`Skipping tool ${tool.alias} - doesn't match filter pattern`);
      continue;
    }

    // Feature toggle filtering
    if (!getCategoryEnabled(tool.alias, tool.path)) {
      logger.info(`Skipping tool ${tool.alias} - disabled by feature toggles`);
      continue;
    }

    const paramSchema: Record<string, unknown> = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        paramSchema[param.name] = param.schema || z.any();
      }
    }

    if (tool.method.toUpperCase() === 'GET' && tool.path.includes('/')) {
      paramSchema['fetchAllPages'] = z
        .boolean()
        .describe('Automatically fetch all pages of results')
        .optional();
    }

    if (rsvpEnabledAliases.has(tool.alias)) {
      paramSchema['responseAction'] = responseActionEnum
        .describe(
          'Preferred RSVP indicator. Set to "accept", "decline", or "tentative" to respond without modifying attendees.'
        )
        .optional();
      paramSchema['responseStatus'] = responseStatusEnum
        .describe('Alternative RSVP indicator using Graph status values (accepted / declined / tentativelyAccepted).')
        .optional();
      paramSchema['sendResponse'] = z
        .boolean()
        .describe('Optional SendResponse flag for RSVP endpoints (maps to Graph SendResponse).')
        .optional();
      paramSchema['comment'] = z
        .string()
        .describe('Optional comment/notes to include in the RSVP response (maps to Graph Comment).')
        .optional();
      paramSchema['body'] = z
        .object({
          responseAction: responseActionEnum.optional(),
          responseStatus: responseStatusEnum.optional(),
          SendResponse: z.boolean().optional(),
          sendResponse: z.boolean().optional(),
          Comment: z.string().optional(),
          comment: z.string().optional(),
        })
        .passthrough()
        .describe(
          'Event patch object. For RSVP, set responseAction/responseStatus (preferred) and optional SendResponse/Comment. Do NOT include attendees to respond—the server automatically responds as the authenticated mailbox. Other Event fields may be patched as usual.'
        )
        .optional();
    }

    // Add includeHeaders parameter for all tools to capture ETags and other headers
    paramSchema['includeHeaders'] = z
      .boolean()
      .describe('Include response headers (including ETag) in the response metadata')
      .optional();

    // Add excludeResponse parameter to only return success/failure indication
    paramSchema['excludeResponse'] = z
      .boolean()
      .describe('Exclude the full response body and only return success or failure indication')
      .optional();
    // Add includeBody/bodyFormat controls for endpoints that may contain HTML bodies
    paramSchema['includeBody'] = z
      .boolean()
      .describe('Include full body fields (HTML or text) instead of safe preview')
      .optional();
    paramSchema['bodyFormat'] = z
      .enum(['html', 'text'])
      .describe('Format of returned body when includeBody=true')
      .optional();

    const safeName = createSafeToolName(tool.alias);
    const safeTitle = createSafeToolName(tool.alias);
    if (tool.alias !== safeName) {
      sanitizedTools.push({ original: tool.alias, sanitized: safeName });
      logger.warn(`Tool alias sanitized`, { alias: tool.alias, safeName });
    }

    const finalDescription = getToolDescription(
      tool.alias,
      tool.description ||
        `Execute ${tool.method.toUpperCase()} request to ${tool.path} (${tool.alias})`
    );

    const readOnlyHint = tool.method.toUpperCase() === 'GET';
    const handler = async (params: any) => {
        logger.info(`Tool ${safeName} called with params: ${JSON.stringify(params)}`);
        try {
          logger.info(`params: ${JSON.stringify(params)}`);

          // Impersonation: Extract from all possible sources and log which one is used
          const impersonateHeaderName = (process.env.MS365_MCP_IMPERSONATE_HEADER || 'X-Impersonate-User').toLowerCase();
          
          // First check for _meta in AsyncLocalStorage (set by request interceptor)
          const storedMeta = ImpersonationContext.getMeta();
          const fromMetaHeaders = storedMeta?.headers?.[impersonateHeaderName] as string | undefined;
          
          const fromContext = ImpersonationContext.getImpersonatedUser();
          const fromEnv = (process.env.MS365_MCP_IMPERSONATE_USER || '').trim() || undefined;
          
          // Priority: _meta headers > AsyncLocalStorage context > env var
          let impersonated: string | undefined;
          let impersonationSource: 'meta-header' | 'http-context' | 'env-var' | 'none' = 'none';
          
          if (fromMetaHeaders?.trim()) {
            impersonated = fromMetaHeaders.trim();
            impersonationSource = 'meta-header';
            ImpersonationContext.setImpersonatedUser(impersonated);
          } else if (fromContext) {
            impersonated = fromContext;
            impersonationSource = 'http-context';
          } else if (fromEnv) {
            impersonated = fromEnv;
            impersonationSource = 'env-var';
            ImpersonationContext.setImpersonatedUser(impersonated);
          }
          
          // Debug logging for impersonation mode
          logger.info(`[Impersonation Debug] Tool: ${safeName}`, {
            mode: impersonationSource,
            headerName: impersonateHeaderName,
            headerValue: fromMetaHeaders?.trim() || 'not set',
            contextValue: fromContext || 'not set',
            envVarValue: fromEnv || 'not set',
            finalValue: impersonated || 'none',
            source: impersonationSource,
            storedMetaPresent: !!storedMeta,
            storedMetaHeadersPresent: !!storedMeta?.headers,
            allStoredMetaHeaders: storedMeta?.headers ? Object.keys(storedMeta.headers) : []
          });

          const cache = new MailboxDiscoveryCache();
          let allowedEmails: string[] = [];
          if (impersonated) {
            const allowed = await cache.getMailboxes(impersonated);
            allowedEmails = allowed.map((m) => m.email.toLowerCase());
          }

          logger.info(`Allowed mailboxes after mailbox discovery: ${JSON.stringify(allowedEmails)}`);

          const parameterDefinitions = tool.parameters || [];

          let path = tool.path;
          
          const queryParams: Record<string, string> = {};
          const resolvedPathParams: Record<string, string> = {};
          const headers: Record<string, string> = {};
          let body: unknown = null;

          for (let [paramName, paramValue] of Object.entries(params)) {
            // Skip pagination control parameter - it's not part of the Microsoft Graph API - I think 🤷
            if (paramName === 'fetchAllPages') {
              continue;
            }

            // Skip headers control parameter - it's not part of the Microsoft Graph API
            if (paramName === 'includeHeaders') {
              continue;
            }

            // Skip excludeResponse control parameter - it's not part of the Microsoft Graph API
            if (paramName === 'excludeResponse') {
              continue;
            }

            // Skip response parameter - it's used to route to correct endpoint, not sent to API
            if (paramName === 'response' || paramName === 'tentative') {
              continue;
            }

            // Ok, so, MCP clients (such as claude code) doesn't support $ in parameter names,
            // and others might not support __, so we strip them in hack.ts and restore them here
            const odataParams = [
              'filter',
              'select',
              'expand',
              'orderby',
              'skip',
              'top',
              'count',
              'search',
              'format',
            ];
            const fixedParamName = odataParams.includes(paramName.toLowerCase())
              ? `$${paramName.toLowerCase()}`
              : paramName;
            const paramDef = parameterDefinitions.find((p) => p.name === paramName);

            if (paramDef) {
              switch (paramDef.type) {
                case 'Path':
                  // If this is userId and impersonation is active, enforce allowed set
                  if (
                    impersonated &&
                    (paramName.toLowerCase() === 'userid' || paramName.toLowerCase() === 'user-id')
                  ) {
                    const requested = String(paramValue || '').trim();
                    const normalized = requested.toLowerCase();
                    const defaultUser = impersonated.toLowerCase();
                    let finalUser = defaultUser;
                    if (requested && allowedEmails.includes(normalized)) {
                      finalUser = requested;
                    } else if (requested && !allowedEmails.includes(normalized)) {
                      // Ignore disallowed and force impersonated to avoid leaking
                      logger.info(
                        `[impersonation] Overriding disallowed userId=${requested} → ${impersonated}`
                      );
                    }
                    path = path
                      .replace(`{${paramName}}`, encodeURIComponent(finalUser))
                      .replace(`:${paramName}`, encodeURIComponent(finalUser));
                  } else {
                    path = path
                      .replace(`{${paramName}}`, encodeURIComponent(paramValue as string))
                      .replace(`:${paramName}`, encodeURIComponent(paramValue as string));
                  }
                  if (typeof paramValue === 'string' && paramValue.trim().length > 0) {
                    const lowered = paramName.toLowerCase();
                    if (lowered === 'messageid' || lowered === 'message-id') {
                      resolvedPathParams.messageId = paramValue;
                    } else if (lowered === 'userid' || lowered === 'user-id') {
                      resolvedPathParams.userId = paramValue;
                    }
                  }
                  break;

                case 'Query':
                  // Skip empty query parameters to avoid Graph API errors
                  if (paramValue === '' || paramValue === null || paramValue === undefined) {
                    break;
                  }
                  if (fixedParamName === '$search') {
                    const sanitized = sanitizeAqsSearch(String(paramValue));
                    // Only add if sanitized value is not empty
                    if (sanitized.trim().length > 0) {
                      queryParams[fixedParamName] = `${sanitized}`;
                    }
                  } else {
                    queryParams[fixedParamName] = `${paramValue}`;
                  }
                  break;

                case 'Body':
                  if (paramDef.schema) {
                    const parseResult = paramDef.schema.safeParse(paramValue);
                    if (!parseResult.success) {
                      const wrapped = { [paramName]: paramValue };
                      const wrappedResult = paramDef.schema.safeParse(wrapped);
                      if (wrappedResult.success) {
                        logger.info(
                          `Auto-corrected parameter '${paramName}': AI passed nested field directly, wrapped it as {${paramName}: ...}`
                        );
                        body = wrapped;
                      } else {
                        body = paramValue;
                      }
                    } else {
                      body = paramValue;
                    }
                  } else {
                    body = paramValue;
                  }
                  break;

                case 'Header':
                  headers[fixedParamName] = `${paramValue}`;
                  break;
              }
            } else if (paramName === 'body') {
              body = paramValue;
              logger.info(`Set body param: ${JSON.stringify(body)}`);
            }
          }

          // If impersonation active and tool targets /users/:userId but none was provided, force impersonated
          if (impersonated && path.includes('/users/:userId')) {
            path = path.replace(':userId', encodeURIComponent(impersonated));
          }

          // Sanitize search payloads for /search/query (Microsoft Search API)
          try {
            const isSearchQueryEndpoint = tool.path === '/search/query' || path === '/search/query';
            if (isSearchQueryEndpoint && body && typeof body === 'object') {
              const anyBody: any = body;
              if (Array.isArray(anyBody.requests)) {
                for (const req of anyBody.requests) {
                  if (req && req.query && typeof req.query.queryString === 'string') {
                    req.query.queryString = sanitizeAqsSearch(req.query.queryString);
                  }
                }
              } else if (anyBody.query && typeof anyBody.query.queryString === 'string') {
                anyBody.query.queryString = sanitizeAqsSearch(anyBody.query.queryString);
              }
            }
          } catch (e) {
            logger.warn(`Search query sanitization skipped due to unexpected body shape: ${e}`);
          }

          if (Object.keys(queryParams).length > 0) {
            const queryString = Object.entries(queryParams)
              .filter(([_, value]) => value !== '' && value !== null && value !== undefined)
              .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
              .join('&');
            if (queryString) {
              path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
            }
          }

          const mailboxKeyForRequest = resolveMailboxKey(impersonated, params, path, resolvedPathParams);

          if (tool.alias === 'create-draft-email' && params.body && typeof params.body === 'object') {
            const draftBody = params.body as Record<string, unknown>;
            if (Array.isArray(draftBody.attachments) && draftBody.attachments.length > 0) {
              logger.warn(
                `[create-draft-email] Inline attachments detected; Graph requires add-mail-attachment after draft creation`
              );
              draftBody._attachmentsInlineWarning =
                'Attachments passed inline to create-draft-email are not supported. Create the draft without attachments, then call add-mail-attachment with the draft\'s messageIdToUse.';
            }
          }

          if (tool.alias === 'add-mail-attachment' && params.body && typeof params.body === 'object') {
            const attachmentBody = params.body as Record<string, unknown>;
            const contentBytes =
              typeof attachmentBody.contentBytes === 'string'
                ? attachmentBody.contentBytes
                : typeof (attachmentBody as any).ContentBytes === 'string'
                ? ((attachmentBody as any).ContentBytes as string)
                : undefined;

            if (!contentBytes || contentBytes.trim().length === 0) {
              throw new Error(
                'add-mail-attachment requires a file payload with a base64-encoded "contentBytes" string. Example: { "@odata.type": "#microsoft.graph.fileAttachment", "name": "file.png", "contentType": "image/png", "contentBytes": "iVBORw0..." }.'
              );
            }

            if (!attachmentBody['@odata.type']) {
              attachmentBody['@odata.type'] = '#microsoft.graph.fileAttachment';
              logger.info(`[add-mail-attachment] Defaulted @odata.type to #microsoft.graph.fileAttachment`);
            }
          }

          const options: {
            method: string;
            headers: Record<string, string>;
            body?: string;
            rawResponse?: boolean;
            includeHeaders?: boolean;
            excludeResponse?: boolean;
            [key: string]: unknown;
          } = {
            method: tool.method.toUpperCase(),
            headers,
          };

          if (options.method !== 'GET' && body) {
            options.body = typeof body === 'string' ? body : JSON.stringify(body);
          }

          const isProbablyMediaContent =
            tool.errors?.some((error) => error.description === 'Retrieved media content') ||
            path.endsWith('/content');

          if (endpointConfig?.returnDownloadUrl && path.endsWith('/content')) {
            path = path.replace(/\/content$/, '');
            logger.info(
              `Auto-returning download URL for ${tool.alias} (returnDownloadUrl=true in endpoints.json)`
            );
          } else if (isProbablyMediaContent) {
            options.rawResponse = true;
          }

          // Set includeHeaders if requested
          if (params.includeHeaders === true) {
            options.includeHeaders = true;
          }

          // Set excludeResponse if requested
          if (params.excludeResponse === true) {
            options.excludeResponse = true;
          }

          // Auto-route: If update-calendar-event is used for RSVP, automatically route to respond endpoint
          // MUST happen BEFORE the request is made
          // Store eventId used for RSVP so we can verify it later
          let rsvpEventId: string | null = null;
          
          const isRsvpCapableTool = rsvpEnabledAliases.has(tool.alias);

          if (isRsvpCapableTool) {
            const rawBody =
              params.body && typeof params.body === 'object' ? ({ ...(params.body as any) } as Record<string, any>) : {};
            const body = rawBody;

            if (params.sendResponse !== undefined && body.SendResponse === undefined && body.sendResponse === undefined) {
              body.SendResponse = params.sendResponse;
            }
            if (params.comment !== undefined && body.Comment === undefined && body.comment === undefined) {
              body.Comment = params.comment;
            }
            if (body.sendResponse !== undefined && body.SendResponse === undefined) {
              body.SendResponse = body.sendResponse;
              delete body.sendResponse;
            }
            if (body.comment !== undefined && body.Comment === undefined) {
              body.Comment = body.comment;
              delete body.comment;
            }
            const userEmailForRSVP = impersonated?.toLowerCase() || '';

            const mapResponseAction = (value: unknown): 'accept' | 'tentative' | 'decline' | null => {
              if (value === undefined || value === null) return null;
              const normalized = String(value).trim().toLowerCase();
              if (!normalized) return null;
              if (['accept', 'accepted', 'acceptance', 'yes', 'confirm', 'confirmed'].includes(normalized)) {
                return 'accept';
              }
              if (['decline', 'declined', 'reject', 'rejected', 'no'].includes(normalized)) {
                return 'decline';
              }
              if (
                [
                  'tentative',
                  'tentativelyaccepted',
                  'tentatively accepted',
                  'maybe',
                  'tentativeaccept',
                  'tentative acceptance',
                ].includes(normalized)
              ) {
                return 'tentative';
              }
              return null;
            };

            const mapGraphStatusToAction = (status: unknown): 'accept' | 'tentative' | 'decline' | null => {
              const normalized = typeof status === 'string' ? status : undefined;
              if (!normalized) return null;
              if (normalized === 'accepted') return 'accept';
              if (normalized === 'tentativelyAccepted') return 'tentative';
              if (normalized === 'declined') return 'decline';
              return null;
            };

            let responseAction: 'accept' | 'tentative' | 'decline' | null = null;
            let responseSource: string | null = null;

            const trySetResponseAction = (value: unknown, source: string) => {
              if (responseAction !== null) return;
              const mapped = mapResponseAction(value);
              if (mapped) {
                responseAction = mapped;
                responseSource = source;
              }
            };

            trySetResponseAction(body.responseAction, 'body.responseAction');
            trySetResponseAction(body.responseStatus, 'body.responseStatus');
            trySetResponseAction(body.response, 'body.response');
            trySetResponseAction(body.rsvp, 'body.rsvp');
            trySetResponseAction(body.intent, 'body.intent');
            trySetResponseAction(params.responseAction, 'params.responseAction');
            trySetResponseAction(params.responseStatus, 'params.responseStatus');
            trySetResponseAction(params.response, 'params.response');
            trySetResponseAction((params as any)['response-action'], "params['response-action']");

            // Fallback: look at attendees only if no explicit response field was provided
            if (responseAction === null && body.attendees && Array.isArray(body.attendees)) {
              const userAttendee = body.attendees.find((a: any) => {
                const email = a.emailAddress?.address?.toLowerCase() || '';
                return email === userEmailForRSVP || email.includes(userEmailForRSVP.split('@')[0]);
              });

              if (userAttendee && userAttendee.status?.response) {
                const mapped = mapGraphStatusToAction(userAttendee.status.response);
                if (mapped) {
                  responseAction = mapped;
                  responseSource = 'body.attendees';
                }
              }
            }

            if (!responseAction) {
              logger.debug(
                `[update-calendar-event] No RSVP markers detected; proceeding with regular PATCH operation`
              );
            } else {
              logger.info(
                `[update-calendar-event] Detected RSVP intent (${responseAction}) from ${responseSource || 'unknown source'}`
              );

              // Extract eventId from params ONLY (path still has {event-id} placeholder at this point)
              const eventId = params.eventId || params['event-id'] || params['eventId'];

              logger.info(
                `[update-calendar-event] EventId extraction: params.eventId=${params.eventId || 'none'}, params['event-id']=${params['event-id'] || 'none'}, final=${eventId || 'NONE!'}`
              );

              if (!eventId) {
                logger.error(
                  `[update-calendar-event] CRITICAL: Could not extract eventId from params! Available params: ${JSON.stringify(
                    Object.keys(params)
                  )}, path=${path}`
                );
              }

              rsvpEventId = eventId || rsvpEventId;

              if (eventId) {
                const normalizedUser = userEmailForRSVP || impersonated?.toLowerCase() || '';
                const knownIds = normalizedUser ? recentEventCache.get(normalizedUser) : undefined;
                if (!knownIds || !knownIds.has(eventId)) {
                  throw new Error(
                    `EventId ${eventId.substring(
                      0,
                      50
                    )}... not found in recent list/get results. Always copy the helper field "eventIdToUse" (or "_useThisEventIdForUpdates") from the event you intend to modify.`
                  );
                }

                logger.info(
                  `[update-calendar-event] RSVP routing: eventId=${eventId.substring(
                    0,
                    50
                  )}..., action=${responseAction}, user=${userEmailForRSVP}`
                );

                // Check status BEFORE the call to verify it actually changes
                let statusBefore: string | null = null;
                try {
                  const checkPath = `/me/events/${encodeURIComponent(eventId)}`;
                  const checkResponse = await graphClient.graphRequest(checkPath, { method: 'GET' });
                  if (checkResponse.content && checkResponse.content.length > 0) {
                    const checkData = JSON.parse(checkResponse.content[0].text);
                    
                    // Log organizer info for debugging (but don't block - let API decide)
                    const organizerEmail = checkData.organizer?.emailAddress?.address || '';
                    const userEmailLower = userEmailForRSVP?.toLowerCase() || '';
                    logger.info(`[update-calendar-event] Event organizer: ${organizerEmail}, User: ${userEmailLower}, isOrganizer: ${checkData.isOrganizer}`);
                    
                    if (checkData.attendees && Array.isArray(checkData.attendees)) {
                      const checkAttendee = checkData.attendees.find((a: any) => {
                        const email = a.emailAddress?.address?.toLowerCase() || '';
                        return email === userEmailForRSVP;
                      });
                      if (checkAttendee) {
                        statusBefore = checkAttendee.status?.response || 'none';
                        logger.info(`[update-calendar-event] Status BEFORE respond call: ${statusBefore}`);
                      }
                    }
                  }
                } catch (e) {
                  logger.warn(`[update-calendar-event] Could not check status before call: ${e}`);
                }

                if (responseAction === 'tentative') {
                  path = `/me/events/${encodeURIComponent(eventId)}/tentativelyAccept`;
                } else if (responseAction === 'decline') {
                  path = `/me/events/${encodeURIComponent(eventId)}/decline`;
                } else {
                  path = `/me/events/${encodeURIComponent(eventId)}/accept`;
                }

                logger.info(`[update-calendar-event] RSVP call: path=${path}, method=POST`);

                // Change method to POST (respond endpoints are POST, not PATCH)
                options.method = 'POST';

                // Remove helper fields before sending body to Graph
                const {
                  attendees,
                  responseAction: _responseActionField,
                  responseStatus: _responseStatusField,
                  response: _responseField,
                  rsvp: _rsvpField,
                  intent: _intentField,
                  ...restBody
                } = body;

                if (Object.keys(restBody).length > 0) {
                  options.body = typeof restBody === 'string' ? restBody : JSON.stringify(restBody);
                  logger.info(
                    `[update-calendar-event] Keeping body fields for RSVP call: ${Object.keys(restBody).join(', ')}`
                  );
                } else {
                  options.body = '{}';
                  logger.info(`[update-calendar-event] Sending empty body for RSVP`);
                }

                logger.info(`[update-calendar-event] Routed to ${path} (POST) for RSVP response`);
              }
            }
          }

          logger.info(`Making graph request to ${path} with options: ${JSON.stringify(options)}`);
          let response = await graphClient.graphRequest(path, options);
          
          // Log response status for RSVP calls
          if (isRsvpCapableTool && (path.includes('/accept') || path.includes('/decline') || path.includes('/tentativelyAccept'))) {
            logger.info(`[update-calendar-event] RSVP call response: path=${path}, responseSize=${response.content?.length || 0}, isError=${response.isError || false}`);
          }

          const fetchAllPages = params.fetchAllPages === true;
          if (fetchAllPages && response && response.content && response.content.length > 0) {
            try {
              let combinedResponse = JSON.parse(response.content[0].text);
              let allItems = combinedResponse.value || [];
              let nextLink = combinedResponse['@odata.nextLink'];
              let pageCount = 1;

              while (nextLink) {
                logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);

                const url = new URL(nextLink);
                const nextPath = url.pathname.replace('/v1.0', '');
                const nextOptions = { ...options };

                const nextQueryParams: Record<string, string> = {};
                for (const [key, value] of url.searchParams.entries()) {
                  nextQueryParams[key] = value;
                }
                nextOptions.queryParams = nextQueryParams;

                const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
                if (nextResponse && nextResponse.content && nextResponse.content.length > 0) {
                  const nextJsonResponse = JSON.parse(nextResponse.content[0].text);
                  if (nextJsonResponse.value && Array.isArray(nextJsonResponse.value)) {
                    allItems = allItems.concat(nextJsonResponse.value);
                  }
                  nextLink = nextJsonResponse['@odata.nextLink'];
                  pageCount++;

                  if (pageCount > 100) {
                    logger.warn(`Reached maximum page limit (100) for pagination`);
                    break;
                  }
                } else {
                  break;
                }
              }

              combinedResponse.value = allItems;
              if (combinedResponse['@odata.count']) {
                combinedResponse['@odata.count'] = allItems.length;
              }
              delete combinedResponse['@odata.nextLink'];

              response.content[0].text = JSON.stringify(combinedResponse);

              logger.info(
                `Pagination complete: collected ${allItems.length} items across ${pageCount} pages`
              );
            } catch (e) {
              logger.error(`Error during pagination: ${e}`);
            }
          }

          if (response && response.content && response.content.length > 0) {
            const responseText = response.content[0].text;
            const responseSize = responseText.length;
            // Keep high-level metrics at info
            logger.info(`Response size: ${responseSize} characters`);
            // Enforce hard cap to avoid flooding agent context
            const MAX_TEXT = Number(process.env.MS365_MCP_TEXT_CAP || '200000');
            let textTruncatedMeta: { textTruncated: boolean; originalSize: number } | undefined;

            try {
              const jsonResponse = JSON.parse(responseText);
              
              // Check for error responses and enhance them with helpful messages
              if (response.isError || (jsonResponse && typeof jsonResponse === 'object' && jsonResponse.error)) {
                const errorMessage = typeof jsonResponse.error === 'string' ? jsonResponse.error : JSON.stringify(jsonResponse.error);
                const is404Error = errorMessage.includes('404') || errorMessage.includes('ErrorItemNotFound') || errorMessage.includes('not found');
                const isOrganizerError = errorMessage.includes('organizer') || errorMessage.includes('meeting organizer') || 
                                       (jsonResponse.error && typeof jsonResponse.error === 'object' && 
                                        (jsonResponse.error.message?.toLowerCase().includes('organizer') || 
                                         jsonResponse.error.code === 'ErrorInvalidRequest'));
                
                // Enhance organizer errors for RSVP operations
                if (isOrganizerError && isRsvpCapableTool && 
                    (path.includes('/accept') || path.includes('/decline') || path.includes('/tentativelyAccept'))) {
                  // Extract organizer info from error or try to get it from the event
                  let organizerInfo = '';
                  try {
                    const errorDetail = jsonResponse.error && typeof jsonResponse.error === 'object' ? jsonResponse.error : {};
                    const eventIdMatch = path.match(/\/events\/([^\/]+)/);
                    if (eventIdMatch) {
                      const eventId = decodeURIComponent(eventIdMatch[1]);
                      logger.info(`[update-calendar-event] Attempting to fetch event ${eventId.substring(0, 50)}... to check organizer`);
                      // Note: We could fetch the event here, but that would be another API call
                      // Instead, suggest the user check the event details
                    }
                  } catch (e) {
                    logger.warn(`[update-calendar-event] Could not extract organizer info: ${e}`);
                  }
                  
                  const currentUserEmail = impersonated || 'current user';
                  const enhancedError = `⚠️  CANNOT RSVP: Microsoft Graph API reports you cannot respond to this meeting because you're the organizer. ` +
                    `However, if you believe you are NOT the organizer, this may be an API issue. ` +
                    `To verify: Call get-calendar-event with the eventId and check the "organizer.emailAddress.address" field. ` +
                    `The actual organizer should be listed there. If the organizer email doesn't match your email (${currentUserEmail}), ` +
                    `this may be a Microsoft Graph API limitation or impersonation issue. ` +
                    `Organizers cannot accept/decline their own meetings—they are already considered accepted. ` +
                    `If you need to cancel the meeting, use delete-calendar-event instead.`;
                  logger.error(enhancedError);
                  jsonResponse._organizerError = true;
                  jsonResponse._error = enhancedError;
                  jsonResponse._warning = enhancedError;
                  jsonResponse._cannotRSVP = true;
                  jsonResponse._suggestedCheck = 'Call get-calendar-event to verify the organizer.emailAddress.address field';
                }
                
                // Enhance 404 errors for folder-related operations
                if (is404Error && (tool.alias === 'list-mail-folder-messages' || path.includes('/mailFolders/'))) {
                  const folderId = params.mailFolderId || params.folderId || resolvedPathParams.mailFolderId;
                  if (folderId) {
                    const enhancedError = `⚠️  FOLDER NOT FOUND: The folder ID "${folderId.substring(0, 50)}..." is invalid or doesn't exist. ` +
                      `This usually means: (1) The folder ID is incorrect (e.g., using parentFolderId instead of folder id), ` +
                      `(2) The folder was deleted, or (3) Case-sensitive folder name mismatch when looking up the folder. ` +
                      `SOLUTION: Call list-mail-folders with $select=id,displayName to get the correct folder ID. ` +
                      `Use the folder's "id" field (NOT parentFolderId). ` +
                      `Helper fields "folderIdToUse" or "_useThisIdForFolderQueries" from list-mail-folders point to the correct ID.`;
                    logger.error(enhancedError);
                    jsonResponse._folderNotFound = true;
                    jsonResponse._error = enhancedError;
                    jsonResponse._warning = enhancedError;
                    jsonResponse._invalidFolderId = true;
                  }
                }
              }
              
              // Special handling for update-calendar-event when it was routed to a respond endpoint
              const isRespondEndpoint = isRsvpCapableTool && 
                                       (path.includes('/accept') || path.includes('/decline') || path.includes('/tentativelyAccept'));
              
              // Accept/decline endpoints return 204 No Content, which becomes an empty response or {success: true}
              const isEmptyResponse = !responseText || responseText.trim() === '' || responseText === '{}' || 
                                     (jsonResponse && (jsonResponse.success === true || jsonResponse.message === 'OK!'));
              
              if (isRespondEndpoint && isEmptyResponse) {
                logger.info(`[update-calendar-event] RSVP endpoint called successfully (empty/204 response expected)`);
                // Determine response type from path
                let response: 'accept' | 'tentative' | 'decline' = 'accept';
                if (path.includes('/tentativelyAccept')) {
                  response = 'tentative';
                } else if (path.includes('/decline')) {
                  response = 'decline';
                } else {
                  response = 'accept';
                }
                
                // Verify the acceptance actually worked by fetching the event
                // Use the stored eventId from routing (should always be set if routing happened)
                const eventId = rsvpEventId || (params.eventId || params['event-id'] || 'unknown') as string;
                
                if (!rsvpEventId) {
                  logger.warn(`[update-calendar-event] WARNING: No stored eventId for verification! Using fallback: ${eventId.substring(0, 50)}...`);
                }
                
                logger.info(`[update-calendar-event] Verifying RSVP for eventId: ${eventId.substring(0, 50)}... (stored: ${rsvpEventId ? 'yes' : 'no'})`);
                try {
                  // Wait longer for the API to process the change - RSVP endpoints can take time to propagate
                  // Try multiple times with increasing delays to account for sync delays
                  let verificationAttempts = 0;
                  const maxAttempts = 3;
                  let eventData: any = null;
                  let verifyResponse: any = null;
                  
                  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
                    const delayMs = attempt * 1000; // 1s, 2s, 3s
                    logger.info(`[update-calendar-event] Verification attempt ${attempt}/${maxAttempts}, waiting ${delayMs}ms...`);
                    await new Promise(resolve => setTimeout(resolve, delayMs));
                    
                    const verifyPath = `/me/events/${encodeURIComponent(eventId)}`;
                    logger.info(`[update-calendar-event] Fetching event to verify (attempt ${attempt}): ${verifyPath}`);
                    verifyResponse = await graphClient.graphRequest(verifyPath, { method: 'GET' });
                    
                    if (verifyResponse.content && verifyResponse.content.length > 0) {
                      eventData = JSON.parse(verifyResponse.content[0].text);
                      const verifyUserEmail = impersonated?.toLowerCase() || '';
                      const userAttendee = eventData.attendees?.find((a: any) => {
                        const email = a.emailAddress?.address?.toLowerCase() || '';
                        return email === verifyUserEmail;
                      });
                      
                      if (userAttendee) {
                        const actualStatus = userAttendee.status?.response || 'none';
                        const expectedStatus = response === 'accept' ? 'accepted' :
                                             response === 'tentative' ? 'tentativelyAccepted' :
                                             'declined';
                        
                        logger.info(`[update-calendar-event] Attempt ${attempt}: User ${verifyUserEmail} status is ${actualStatus}, expected ${expectedStatus}`);
                        
                        if (actualStatus === expectedStatus) {
                          logger.info(`[update-calendar-event] Verification successful on attempt ${attempt}`);
                          break; // Success, exit loop
                        } else if (attempt < maxAttempts) {
                          logger.info(`[update-calendar-event] Status mismatch on attempt ${attempt}, will retry...`);
                        }
                      }
                    }
                  }
                  
                  if (!eventData && verifyResponse && verifyResponse.content && verifyResponse.content.length > 0) {
                    eventData = JSON.parse(verifyResponse.content[0].text);
                  }
                  
                  if (eventData) {
                    // Use the impersonated user already determined in the handler
                    // (from header, context, or env var - in that priority order)
                    const verifyUserEmail = impersonated?.toLowerCase() || '';
                    const eventSubject = eventData.subject || 'NO SUBJECT';
                    logger.info(`[update-calendar-event] Final verification: subject="${eventSubject}", eventId=${eventId.substring(0, 50)}..., user=${verifyUserEmail}, attendees=${eventData.attendees?.length || 0}`);
                    
                    if (eventData.attendees && Array.isArray(eventData.attendees)) {
                      // Log all attendees for debugging
                      logger.info(`[update-calendar-event] Event attendees: ${JSON.stringify(eventData.attendees.map((a: any) => ({ 
                        email: a.emailAddress?.address, 
                        status: a.status?.response,
                        type: a.type 
                      })))}`);
                      
                      // Find the exact matching attendee - be precise with email matching
                      const userAttendee = eventData.attendees.find((a: any) => {
                        const email = a.emailAddress?.address?.toLowerCase() || '';
                        // Exact match only - don't use loose matching
                        return email === verifyUserEmail;
                      });
                      
                      // Check organizer first - if user is organizer, they can't RSVP
                      const organizerEmail = eventData.organizer?.emailAddress?.address?.toLowerCase() || '';
                      const isUserOrganizer = organizerEmail === verifyUserEmail || eventData.isOrganizer === true;
                      
                      if (isUserOrganizer) {
                        jsonResponse.success = false;
                        jsonResponse._organizerError = true;
                        jsonResponse._error = `Cannot verify RSVP: You are the organizer of this meeting (${organizerEmail}). Organizers cannot accept/decline their own meetings.`;
                        jsonResponse.message = `⚠️ Cannot verify RSVP: You are the organizer. Organizers cannot RSVP to their own meetings.`;
                        logger.warn(`[update-calendar-event] Verification skipped: User ${verifyUserEmail} is the organizer`);
                      } else if (userAttendee) {
                        const actualStatus = userAttendee.status?.response || 'none';
                        const expectedStatus = response === 'accept' ? 'accepted' :
                                             response === 'tentative' ? 'tentativelyAccepted' :
                                             'declined';
                        
                        logger.info(`[update-calendar-event] Verification: checking attendee ${userAttendee.emailAddress?.address}, expected ${expectedStatus}, actual ${actualStatus}`);
                        
                        // Also log organizer info for debugging
                        logger.info(`[update-calendar-event] Event organizer: ${eventData.organizer?.emailAddress?.address || 'unknown'}, isOrganizer: ${eventData.isOrganizer}`);
                        
                        if (actualStatus === expectedStatus) {
                          const action = response === 'accept' ? 'accepted' :
                                        response === 'tentative' ? 'tentatively accepted' :
                                        'declined';
                          jsonResponse.success = true;
                          jsonResponse.action = action;
                          jsonResponse.message = `Meeting invitation ${action} successfully`;
                          jsonResponse._verified = true;
                          jsonResponse._organizerEmail = eventData.organizer?.emailAddress?.address || 'unknown';
                          jsonResponse._yourEmail = verifyUserEmail;
                          jsonResponse._yourStatus = actualStatus;
                          // Include all attendee statuses for transparency
                          jsonResponse._allAttendees = eventData.attendees.map((a: any) => ({
                            email: a.emailAddress?.address,
                            status: a.status?.response || 'none',
                            type: a.type
                          }));
                          logger.info(`✓ Calendar event ${action} verified: eventId ${eventId.substring(0, 50)}...`);
                        } else {
                          jsonResponse.success = false;
                          jsonResponse._acceptFailed = true;
                          jsonResponse._error = `Accept operation returned 204 success but event was not actually ${response === 'accept' ? 'accepted' : response === 'tentative' ? 'tentatively accepted' : 'declined'}. Current status: ${actualStatus}, Expected: ${expectedStatus}`;
                          jsonResponse.message = `⚠️ Accept operation may have failed. Expected status: ${expectedStatus}, Actual status: ${actualStatus}`;
                          jsonResponse._organizerEmail = eventData.organizer?.emailAddress?.address || 'unknown';
                          jsonResponse._yourEmail = verifyUserEmail;
                          jsonResponse._yourStatus = actualStatus;
                          logger.error(`⚠️ Calendar event accept verification failed: expected ${expectedStatus}, got ${actualStatus} for eventId ${eventId.substring(0, 50)}...`);
                          logger.error(`⚠️ Event attendees: ${JSON.stringify(eventData.attendees.map((a: any) => ({ email: a.emailAddress?.address, status: a.status?.response })))}`);
                        }
                      } else {
                        logger.warn(`Could not find user attendee (${verifyUserEmail}) in event for verification. Event attendees: ${eventData.attendees.map((a: any) => a.emailAddress?.address).join(', ')}`);
                        jsonResponse.success = false;
                        jsonResponse._attendeeNotFound = true;
                        jsonResponse._error = `Could not verify RSVP: User ${verifyUserEmail} not found in event attendees. Available attendees: ${eventData.attendees.map((a: any) => a.emailAddress?.address).join(', ')}`;
                        jsonResponse.message = `⚠️ Could not verify RSVP: Your email (${verifyUserEmail}) was not found in the event attendees.`;
                        // Still mark as success but note verification couldn't be done
                        const action = response === 'accept' ? 'accepted' :
                                      response === 'tentative' ? 'tentatively accepted' :
                                      'declined';
                        jsonResponse.success = true;
                        jsonResponse.action = action;
                        jsonResponse.message = `Meeting invitation ${action} successfully (verification incomplete - user not found in attendees)`;
                        jsonResponse._verificationIncomplete = true;
                        jsonResponse._warning = `Could not verify: user ${verifyUserEmail} not found in event attendees`;
                      }
                    } else {
                      logger.warn(`Event has no attendees array for verification`);
                      const action = response === 'accept' ? 'accepted' :
                                    response === 'tentative' ? 'tentatively accepted' :
                                    'declined';
                      jsonResponse.success = true;
                      jsonResponse.action = action;
                      jsonResponse.message = `Meeting invitation ${action} successfully (verification incomplete - no attendees)`;
                      jsonResponse._verificationIncomplete = true;
                    }
                  } else {
                    logger.warn(`Could not parse verification response`);
                    const action = response === 'accept' ? 'accepted' :
                                  response === 'tentative' ? 'tentatively accepted' :
                                  'declined';
                    jsonResponse.success = true;
                    jsonResponse.action = action;
                    jsonResponse.message = `Meeting invitation ${action} successfully (verification incomplete)`;
                    jsonResponse._verificationIncomplete = true;
                  }
                } catch (verifyError) {
                  logger.error(`Could not verify calendar event acceptance: ${verifyError}`);
                  // Still mark as success but note verification failed
                  const action = response === 'accept' ? 'accepted' :
                                response === 'tentative' ? 'tentatively accepted' :
                                'declined';
                  jsonResponse.success = true;
                  jsonResponse.action = action;
                  jsonResponse.message = `Meeting invitation ${action} successfully (verification failed: ${verifyError})`;
                  jsonResponse._verificationFailed = true;
                  jsonResponse._verificationError = String(verifyError);
                }
              }
              
              // Special handling for list-mail-folders: add helper field and optionally fetch actual message counts
              if (tool.alias === 'list-mail-folders' && jsonResponse.value && Array.isArray(jsonResponse.value)) {
                // Only fetch actual counts if explicitly requested (to avoid rate limits)
                const fetchActualCounts = params.includeActualCounts === true;
                
                let folderCounts: Array<{ id: string | null; actualCount: number | null; unreadCount: number | null }> = [];
                
                // Always fetch actual counts (user requested this), but with aggressive rate limiting
                // Process folders sequentially with delays to avoid 429 errors
                const BATCH_SIZE = 1; // Process 1 folder at a time to avoid rate limits
                const DELAY_BETWEEN_FOLDERS = 300; // 300ms delay between folders
                const SKIP_UNREAD_COUNTS = true; // Skip unread counts to reduce API calls by 50%
                
                for (let i = 0; i < jsonResponse.value.length; i += BATCH_SIZE) {
                  const batch = jsonResponse.value.slice(i, i + BATCH_SIZE);
                  
                  for (const folder of batch) {
                    if (!folder || typeof folder !== 'object' || !folder.id) {
                      folderCounts.push({ id: null, actualCount: null, unreadCount: null });
                      continue;
                    }
                    
                    try {
                      // Use $count=true to get count efficiently
                      const basePath = path.includes('/users/')
                        ? path.replace(/\/mailFolders.*$/, '')
                        : '/me';
                      const folderPath = `${basePath}/mailFolders/${encodeURIComponent(folder.id)}/messages`;
                      
                      // Get total count
                      const countResponse = await graphClient.graphRequest(
                        `${folderPath}?$count=true&$top=0&$select=id`,
                        { method: 'GET' }
                      );
                      
                      if (countResponse.content && countResponse.content.length > 0) {
                        const countData = JSON.parse(countResponse.content[0].text);
                        const totalCount = countData['@odata.count'] ?? null;
                        
                        // Skip unread count to reduce API calls (use metadata unreadItemCount instead)
                        const unreadCount = SKIP_UNREAD_COUNTS ? null : null;
                        
                        folderCounts.push({ id: folder.id, actualCount: totalCount, unreadCount });
                      } else {
                        folderCounts.push({ id: folder.id, actualCount: null, unreadCount: null });
                      }
                    } catch (error: any) {
                      // Handle rate limiting gracefully - if we get 429, stop fetching and use metadata
                      if (error.message && error.message.includes('429')) {
                        logger.warn(`Rate limited while fetching counts. Stopping count fetching and using metadata for remaining ${jsonResponse.value.length - i} folders.`);
                        // Fill remaining folders with null counts
                        for (let j = i; j < jsonResponse.value.length; j++) {
                          folderCounts.push({ id: jsonResponse.value[j]?.id || null, actualCount: null, unreadCount: null });
                        }
                        break; // Stop fetching counts
                      } else {
                        logger.debug(`Could not fetch actual count for folder ${folder.id}: ${error}`);
                        folderCounts.push({ id: folder.id, actualCount: null, unreadCount: null });
                      }
                    }
                    
                    // Delay between folders to avoid rate limits
                    if (i + 1 < jsonResponse.value.length) {
                      await new Promise(resolve => setTimeout(resolve, DELAY_BETWEEN_FOLDERS));
                    }
                  }
                }
                
                // Create a map of actual counts
                const countsMap = new Map(
                  folderCounts
                    .filter((c: any) => c.id && c.actualCount !== null)
                    .map((c: any) => [c.id, { actualCount: c.actualCount, unreadCount: c.unreadCount }])
                );
                
                for (const folder of jsonResponse.value) {
                  if (folder && typeof folder === 'object' && folder.id) {
                    // Add explicit helper field pointing to the correct ID to use
                    folder.folderIdToUse = folder.id;
                    folder._useThisIdForFolderQueries = folder.id;
                    
                    // Add actual counts if available
                    const actualCounts = countsMap.get(folder.id);
                    if (actualCounts) {
                      folder.actualMessageCount = actualCounts.actualCount;
                      folder.actualUnreadCount = actualCounts.unreadCount;
                      
                      // Show both metadata and actual counts, but emphasize actual
                      if (folder.totalItemCount !== undefined && folder.totalItemCount !== null) {
                        const metadataCount = folder.totalItemCount;
                        const actualCount = actualCounts.actualCount;
                        
                        if (Math.abs(metadataCount - actualCount) > 2) {
                          folder._countDiscrepancy = true;
                          folder._countNote = `Metadata shows ${metadataCount} items, but actual message count is ${actualCount}. The actual count (${actualCount}) is accurate.`;
                        }
                      }
                    } else {
                      // If we couldn't get actual count, keep the warning
                      if (folder.totalItemCount !== undefined && folder.totalItemCount !== null) {
                        folder._metadataNote = `⚠️  Note: totalItemCount (${folder.totalItemCount}) may be stale, cached, or include subfolders. ` +
                          `To get accurate message count, query the folder directly using list-mail-messages with folderId="${folder.id}".`;
                      }
                    }
                  }
                }
              }
              
              // Validation for list-mail-folder-messages: verify messages actually belong to the requested folder
              if (tool.alias === 'list-mail-folder-messages' && jsonResponse.value && Array.isArray(jsonResponse.value)) {
                const requestedFolderId = params.mailFolderId || resolvedPathParams.mailFolderId;
                if (requestedFolderId && typeof requestedFolderId === 'string') {
                  // Fetch folder metadata to verify we're querying the correct folder
                  try {
                    const folderPath = path.includes('/users/') 
                      ? path.replace(/\/messages.*$/, '')
                      : `/me/mailFolders/${encodeURIComponent(requestedFolderId)}`;
                    const folderResponse = await graphClient.graphRequest(folderPath, { method: 'GET' });
                    if (folderResponse.content && folderResponse.content.length > 0) {
                      const folderData = JSON.parse(folderResponse.content[0].text);
                      const folderDisplayName = folderData.displayName || 'Unknown';
                      const folderId = folderData.id || requestedFolderId;
                      
                      // Check for subfolders - they might explain the count discrepancy
                      let hasSubfolders = false;
                      let subfolderCount = 0;
                      try {
                        const subfoldersPath = folderPath.replace(/\/$/, '') + '/childFolders';
                        const subfoldersResponse = await graphClient.graphRequest(subfoldersPath, { method: 'GET' });
                        if (subfoldersResponse.content && subfoldersResponse.content.length > 0) {
                          const subfoldersData = JSON.parse(subfoldersResponse.content[0].text);
                          if (subfoldersData.value && Array.isArray(subfoldersData.value) && subfoldersData.value.length > 0) {
                            hasSubfolders = true;
                            subfolderCount = subfoldersData.value.length;
                          }
                        }
                      } catch (e) {
                        // Subfolder check failed, ignore
                      }
                      
                      // Add folder info to response for verification
                      jsonResponse._queriedFolder = {
                        id: folderId,
                        displayName: folderDisplayName,
                        totalItemCount: folderData.totalItemCount,
                        unreadItemCount: folderData.unreadItemCount,
                        hasSubfolders: hasSubfolders,
                        subfolderCount: subfolderCount,
                        parentFolderId: folderData.parentFolderId
                      };
                      
                      const folderTotalCount = folderData.totalItemCount || 0;
                      const returnedCount = jsonResponse.value.length;
                      
                      logger.info(`📁 Queried folder: "${folderDisplayName}" (ID: ${folderId.substring(0, 50)}..., Total: ${folderTotalCount}, Unread: ${folderData.unreadItemCount || 'unknown'}, Returned: ${returnedCount})`);
                      
                      // Warn if folder ID doesn't match (shouldn't happen, but good to check)
                      if (folderId !== requestedFolderId) {
                        const warning = `⚠️  Folder ID mismatch: Requested "${requestedFolderId.substring(0, 50)}..." but got "${folderId.substring(0, 50)}..." (displayName: "${folderDisplayName}")`;
                        logger.warn(warning);
                        jsonResponse._folderIdMismatch = true;
                        jsonResponse._warning = warning;
                      }
                      
                      // Warn if folder metadata count differs significantly from returned message count
                      // This can happen if: (1) folder metadata is stale/cached, (2) folder contains subfolders,
                      // (3) folder contains items other than messages, (4) there's a sync issue, or
                      // (5) folder was created from another folder and inherited the count
                      if (folderTotalCount > 0 && returnedCount > 0) {
                        const discrepancy = Math.abs(folderTotalCount - returnedCount);
                        const discrepancyPercent = (discrepancy / Math.max(folderTotalCount, returnedCount)) * 100;
                        
                        // If discrepancy is large (more than 20% or more than 10 items), warn
                        if (discrepancyPercent > 20 || discrepancy > 10) {
                          let causes = [];
                          if (hasSubfolders) {
                            causes.push(`Folder contains ${subfolderCount} subfolder(s) - totalItemCount may include items in subfolders`);
                          }
                          causes.push('Folder metadata is stale/cached (Microsoft Graph API limitation)');
                          causes.push('Folder may have been created from another folder and inherited the count');
                          causes.push('Sync delay between Outlook and Graph API');
                          
                          const warning = `⚠️  COUNT DISCREPANCY: Folder "${folderDisplayName}" metadata reports ${folderTotalCount} total items, but query returned ${returnedCount} messages (difference: ${discrepancy}). ` +
                            `Possible causes: ${causes.join('; ')}. ` +
                            `The returned ${returnedCount} messages are the actual messages currently in this folder. ` +
                            `${hasSubfolders ? `NOTE: This folder has ${subfolderCount} subfolder(s) - check if items are in subfolders. ` : ''}` +
                            `To verify, check the folder in Outlook and compare with the returned message count.`;
                          logger.warn(warning);
                          jsonResponse._countDiscrepancy = true;
                          jsonResponse._countWarning = warning;
                          jsonResponse._metadataCount = folderTotalCount;
                          jsonResponse._returnedCount = returnedCount;
                        }
                      } else if (folderTotalCount > 0 && returnedCount === 0) {
                        // Folder metadata says there are items, but query returned none
                        const warning = `⚠️  EMPTY RESULT: Folder "${folderDisplayName}" metadata reports ${folderTotalCount} total items, but query returned 0 messages. ` +
                          `This may indicate: (1) All items are in subfolders, (2) Items are not messages, (3) Filter/query issue, or (4) Sync delay.`;
                        logger.warn(warning);
                        jsonResponse._emptyResultWarning = warning;
                      }
                    }
                  } catch (folderError) {
                    logger.warn(`Could not fetch folder metadata for validation: ${folderError}`);
                  }
                  
                  // Validate messages belong to the requested folder
                  let mismatchedCount = 0;
                  const mismatchedMessages: string[] = [];
                  
                  for (const message of jsonResponse.value) {
                    if (message && typeof message === 'object' && message.parentFolderId) {
                      if (message.parentFolderId !== requestedFolderId) {
                        mismatchedCount++;
                        if (mismatchedMessages.length < 5) {
                          mismatchedMessages.push(`Message ${message.id?.substring(0, 30) || 'unknown'} is in folder ${message.parentFolderId.substring(0, 30)}...`);
                        }
                      }
                    }
                  }
                  
                  if (mismatchedCount > 0) {
                    const errorMsg = `⚠️  FOLDER MISMATCH: Requested folder "${requestedFolderId.substring(0, 50)}..." but ${mismatchedCount} of ${jsonResponse.value.length} messages belong to different folders. ` +
                      `This usually means the folder ID is incorrect (e.g., using parentFolderId instead of folder id, or wrong folder entirely). ` +
                      `SOLUTION: Call list-mail-folders with $select=id,displayName and find the exact folder by displayName (case-sensitive). ` +
                      `Use the folder's "id" field (NOT parentFolderId) as folderId. ` +
                      `Helper fields "folderIdToUse" or "_useThisIdForFolderQueries" from list-mail-folders point to the correct ID. ` +
                      (mismatchedMessages.length > 0 ? `Sample mismatches: ${mismatchedMessages.join('; ')}` : '');
                    logger.error(errorMsg);
                    jsonResponse._folderMismatch = true;
                    jsonResponse._error = errorMsg;
                    jsonResponse._warning = errorMsg;
                    jsonResponse._mismatchedCount = mismatchedCount;
                    jsonResponse._totalCount = jsonResponse.value.length;
                  } else if (jsonResponse.value.length > 0) {
                    // All messages match - log success for debugging
                    logger.info(`✓ Folder validation passed: All ${jsonResponse.value.length} messages belong to requested folder ${requestedFolderId.substring(0, 50)}...`);
                  }
                }
              }
              
              // Special handling for calendar event queries: add clear response status and helper fields
              if (
                (tool.alias === 'list-calendar-events' || tool.alias === 'get-calendar-event' || tool.alias === 'get-calendar-view') &&
                jsonResponse &&
                typeof jsonResponse === 'object'
              ) {
                const events = (tool.alias === 'list-calendar-events' || tool.alias === 'get-calendar-view')
                  ? (jsonResponse.value || [])
                  : [jsonResponse];
                const cachedEventIds: string[] = [];
                
                for (const event of events) {
                  // Always add helper fields for all events with an ID, regardless of attendees
                  if (event && typeof event === 'object' && event.id && typeof event.id === 'string') {
                    cachedEventIds.push(event.id);
                    // Always add helper fields for all events so they can be used for updates/RSVPs
                    event.eventIdToUse = event.id;
                    event._useThisEventIdForUpdates = event.id;
                    event._calendarIdUsed = event.calendarId || null;
                  }
                  
                  // Process attendee-specific information if attendees exist
                  if (event && typeof event === 'object' && event.attendees && Array.isArray(event.attendees)) {
                    // Find the authenticated user's attendee entry
                    // Use the impersonated user already determined in the handler
                    // (from header, context, or env var - in that priority order)
                    const eventUserEmail = impersonated?.toLowerCase() || '';
                    
                    const userAttendee = event.attendees.find((a: any) => {
                      const email = a.emailAddress?.address?.toLowerCase() || '';
                      // Use exact match for consistency
                      return email === eventUserEmail;
                    });

                    if (userAttendee) {
                      const responseStatus = userAttendee.status?.response || 'none';

                      if (responseStatus === 'none' || responseStatus === null || responseStatus === undefined) {
                        event._responseStatus = 'notAccepted';
                        event._responseStatusMessage =
                          '⚠️ This meeting invitation has NOT been accepted yet. Call update-calendar-event with body.responseAction set to "accept" (or body.responseStatus="accepted"). Do NOT include attendee lists or ask which attendee to use—the MCP server already responds as the authenticated mailbox.';
                      } else if (responseStatus === 'tentativelyAccepted') {
                        event._responseStatus = 'tentative';
                        event._responseStatusMessage =
                          'This meeting is tentatively accepted (maybe). Call update-calendar-event with body.responseAction="accept" (or body.responseStatus="accepted") to fully accept. Do NOT prompt the user for attendee details.';
                      } else if (responseStatus === 'accepted') {
                        event._responseStatus = 'accepted';
                        event._responseStatusMessage = 'This meeting has been accepted.';
                      } else if (responseStatus === 'declined') {
                        event._responseStatus = 'declined';
                        event._responseStatusMessage = 'This meeting has been declined.';
                      }
                      // Add explicit note about how to respond
                      const respondEmail = eventUserEmail || impersonated || '(authenticated user)';
                      event._respondInstructions =
                        `To respond: call update-calendar-event with body.responseAction set to "accept", "decline", or "tentative" ` +
                        `(or use body.responseStatus with Graph values "accepted"/"declined"/"tentativelyAccepted"). ` +
                        `Do NOT include attendee lists and NEVER ask the human which attendee to use; always respond as ${respondEmail}.`;
                    } else if (event.organizer?.emailAddress?.address?.toLowerCase() === eventUserEmail) {
                      // User is the organizer, so no response needed
                      event._responseStatus = 'organizer';
                      event._responseStatusMessage = 'You are the organizer of this meeting.';
                    }
                  }
                }
                rememberEventsForUser(impersonated, cachedEventIds);
              }

              if (mailListAliases.has(tool.alias) && jsonResponse && Array.isArray(jsonResponse.value)) {
                jsonResponse.value.forEach((message: any, index: number) => {
                  annotateMessageObject(message, mailboxKeyForRequest, index + 1);
                });
              } else if (mailSingleMessageAliases.has(tool.alias) && jsonResponse && typeof jsonResponse === 'object') {
                annotateMessageObject(jsonResponse, mailboxKeyForRequest);
              }
              
              // Special handling for create-draft-email: explicitly state that attachments must be added separately
              if (tool.alias === 'create-draft-email' && jsonResponse && typeof jsonResponse === 'object') {
                const hasAttachments = Array.isArray(jsonResponse.attachments) && jsonResponse.attachments.length > 0;
                if (!hasAttachments) {
                  jsonResponse._noAttachments = true;
                  jsonResponse._attachmentInstructions = 
                    'This draft has NO attachments. To attach files, you MUST call add-mail-attachment with the messageIdToUse from this response. ' +
                    'Do NOT claim attachments were added unless you actually called add-mail-attachment and received a success response.';
                }
              }
              
              // Special handling for add-mail-attachment: mark success explicitly
              if (tool.alias === 'add-mail-attachment' && jsonResponse && typeof jsonResponse === 'object') {
                jsonResponse._attachmentAdded = true;
                jsonResponse._attachmentSuccessMessage = 
                  'Attachment successfully added to the draft. You can now call add-mail-attachment again for additional files, or call send-mail to send the draft.';
              }
              
              // Validation for move-mail-message: verify the move actually succeeded
              if (tool.alias === 'move-mail-message' && jsonResponse && typeof jsonResponse === 'object') {
                // Check if this is an error response first
                const isErrorResponse = jsonResponse.error || response.isError;
                
                if (isErrorResponse) {
                  // Extract destinationId for error reporting
                  let destinationId: string | undefined;
                  try {
                    if (options.body && typeof options.body === 'string') {
                      const sentBody = JSON.parse(options.body);
                      destinationId = sentBody.DestinationId || sentBody.destinationId;
                    }
                  } catch (e) {
                    // If parsing fails, fall back to params.body
                  }
                  if (!destinationId) {
                    destinationId = params.body?.DestinationId || params.body?.destinationId;
                  }
                  
                  const errorMessage = typeof jsonResponse.error === 'string' ? jsonResponse.error : JSON.stringify(jsonResponse.error);
                  const is404Error = errorMessage.includes('404') || errorMessage.includes('ErrorItemNotFound') || errorMessage.includes('not found');
                  
                  if (is404Error && destinationId) {
                    const enhancedError = `⚠️  MOVE FAILED: The destination folder ID "${destinationId.substring(0, 50)}..." is invalid or not found. ` +
                      `This usually means: (1) The folder ID is incorrect (e.g., using parentFolderId instead of folder id), ` +
                      `(2) The folder doesn't exist, or (3) Case-sensitive folder name mismatch (e.g., "Wichtig" vs "wichtig"). ` +
                      `SOLUTION: Call list-mail-folders with $select=id,displayName and find the exact folder by displayName (case-sensitive). ` +
                      `Use the folder's "id" field (NOT parentFolderId) as destinationId. ` +
                      `Helper fields "folderIdToUse" or "_useThisIdForFolderQueries" from list-mail-folders point to the correct ID.`;
                    logger.error(enhancedError);
                    jsonResponse._moveFailed = true;
                    jsonResponse._error = enhancedError;
                    jsonResponse._warning = enhancedError;
                    jsonResponse._invalidFolderId = true;
                  } else {
                    // Other error types
                    jsonResponse._moveFailed = true;
                    jsonResponse._error = errorMessage;
                  }
                } else {
                  // Success response - validate the move
                  // Read destinationId from the actual body that was sent to the API
                  // Try to parse from options.body (the actual sent body) first, fallback to params.body
                  let destinationId: string | undefined;
                  try {
                    if (options.body && typeof options.body === 'string') {
                      const sentBody = JSON.parse(options.body);
                      destinationId = sentBody.DestinationId || sentBody.destinationId;
                    }
                  } catch (e) {
                    // If parsing fails, fall back to params.body
                  }
                  // Fallback to params.body if not found in sent body
                  if (!destinationId) {
                    destinationId = params.body?.DestinationId || params.body?.destinationId;
                  }
                  
                  const actualParentFolderId = jsonResponse.parentFolderId;
                  
                  if (destinationId && actualParentFolderId) {
                    if (destinationId !== actualParentFolderId) {
                      const errorMsg = `⚠️  MOVE FAILED: Requested destinationId "${destinationId.substring(0, 50)}..." but message is in folder "${actualParentFolderId.substring(0, 50)}...". ` +
                        `The move operation did not succeed. This usually means the destinationId is invalid (e.g., using parentFolderId instead of folder id). ` +
                        `Use the folder's "id" field from list-mail-folders (NOT "parentFolderId"). ` +
                        `Helper fields "folderIdToUse" or "_useThisIdForFolderQueries" from list-mail-folders point to the correct ID.`;
                      logger.error(errorMsg);
                      // Add error indicator to response
                      jsonResponse._moveFailed = true;
                      jsonResponse._error = `Move failed: destinationId does not match actual parentFolderId. Requested: ${destinationId.substring(0, 30)}..., Actual: ${actualParentFolderId.substring(0, 30)}...`;
                      jsonResponse._warning = errorMsg;
                    } else {
                      logger.info(`✓ Move successful: message moved to folder ${destinationId.substring(0, 50)}...`);
                      jsonResponse._moveSucceeded = true;
                    }
                  } else if (destinationId && !actualParentFolderId) {
                    logger.warn(`⚠️  Could not verify move: destinationId provided but response missing parentFolderId`);
                  }
                }
              }
              
              // Scrub or include message bodies depending on params
              try {
                const includeBody = params.includeBody === true;
                const bodyFormat: 'html' | 'text' | undefined =
                  params.bodyFormat === 'text' ? 'text' : params.bodyFormat === 'html' ? 'html' : undefined;
                scrubBodies(jsonResponse, includeBody, bodyFormat);
                response.content[0].text = JSON.stringify(jsonResponse);
                if (response.content[0].text.length > MAX_TEXT) {
                  textTruncatedMeta = {
                    textTruncated: true,
                    originalSize: response.content[0].text.length,
                  };
                  response.content[0].text = response.content[0].text.slice(0, MAX_TEXT);
                }
              } catch (e) {
                logger.warn(`Body scrub/include step skipped: ${e}`);
              }
              if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
                logger.info(`Response contains ${jsonResponse.value.length} items`);
              }
              if (jsonResponse['@odata.nextLink']) {
                logger.info(`Response has pagination nextLink: ${jsonResponse['@odata.nextLink']}`);
              }
              // Detailed previews only at debug and behind env flag
              if (process.env.MS365_MCP_DEBUG === 'true') {
                const preview = responseText.substring(0, 500);
                logger.debug(
                  `Response preview: ${preview}${responseText.length > 500 ? '...' : ''}`
                );
              }
            } catch {
              if (process.env.MS365_MCP_DEBUG === 'true') {
                const preview = responseText.substring(0, 500);
                logger.debug(
                  `Response preview (non-JSON): ${preview}${responseText.length > 500 ? '...' : ''}`
                );
              }
            }
            // Attach truncation meta if applied
            if (textTruncatedMeta) {
              response._meta = { ...(response._meta || {}), ...textTruncatedMeta };
            }
          }

          // Convert McpResponse to CallToolResult with the correct structure
          const content: ContentItem[] = response.content.map((item) => {
            // GraphClient only returns text content items, so create proper TextContent items
            const textContent: TextContent = {
              type: 'text',
              text: item.text,
            };
            return textContent;
          });

          const result: CallToolResult = {
            content,
            _meta: response._meta,
            isError: response.isError,
          };

          return result;
        } catch (error) {
          logger.error(`Error in tool ${tool.alias}: ${(error as Error).message}`);
          const errorContent: TextContent = {
            type: 'text',
            text: JSON.stringify({
              error: `Error in tool ${tool.alias}: ${(error as Error).message}`,
            }),
          };

          return {
            content: [errorContent],
            isError: true,
          };
        }
      };

    toolDefinitions.set(tool.alias, {
      alias: tool.alias,
      safeName,
      description: finalDescription,
      paramSchema,
      metadata: {
        title: safeTitle,
        readOnlyHint,
      },
      handler,
    });
  }

  const blacklist = new Set<string>(TOOL_BLACKLIST as readonly string[]);
  for (const def of toolDefinitions.values()) {
    if (blacklist.has(def.alias)) {
      logger.info(`Skipping tool ${def.alias} - managed by wrapper`);
      continue;
    }

    server.tool(def.safeName, def.description, def.paramSchema as any, def.metadata, def.handler);
  }

  const invokeInternalTool = async (
    alias: string,
    params: Record<string, unknown>
  ): Promise<CallToolResult> => {
    const def = toolDefinitions.get(alias);
    if (!def) {
      throw new Error(`Internal tool ${alias} is not available`);
    }
    return def.handler(params);
  };

  const cloneParamSchema = (
    schema: Record<string, unknown> | undefined
  ): Record<string, unknown> => ({ ...(schema || {}) });

  const getToolDefinitionOrWarn = (alias: string): GraphToolDefinition | undefined => {
    const def = toolDefinitions.get(alias);
    if (!def) {
      logger.warn(
        `Wrapper registration skipped because required internal alias '${alias}' was not generated (feature disabled or filtered)`
      );
    }
    return def;
  };

  const registerWrapperTool = (
    alias: string,
    description: string,
    paramSchema: Record<string, unknown>,
    metadata: ToolMetadata,
    handler: (params: any) => Promise<CallToolResult>
  ) => {
    const safeName = createSafeToolName(alias);
    const safeTitle = createSafeToolName(alias);
    if (alias !== safeName) {
      sanitizedTools.push({ original: alias, sanitized: safeName });
    }
    server.tool(safeName, description, paramSchema as any, { ...metadata, title: safeTitle }, handler);
  };

  // Calendar wrappers (optional calendarId selector)
  const registerCalendarWrapper = (alias: string, specificAlias: string) => {
    const base = getToolDefinitionOrWarn(alias);
    const specific = getToolDefinitionOrWarn(specificAlias);
    if (!base || !specific) {
      return;
    }

    const schema = {
      ...cloneParamSchema(base.paramSchema),
      calendarId: z
        .string()
        .describe('Optional calendarId to target a non-primary calendar; defaults to primary when omitted')
        .optional(),
    };

    registerWrapperTool(alias, base.description, schema, base.metadata, async (params) => {
      const { calendarId, ...rest } = params;
      if (typeof calendarId === 'string' && calendarId.trim().length > 0) {
        return invokeInternalTool(specificAlias, { ...rest, calendarId });
      }
      return invokeInternalTool(alias, rest);
    });
  };

  registerCalendarWrapper('list-calendar-events', 'list-specific-calendar-events');
  registerCalendarWrapper('get-calendar-event', 'get-specific-calendar-event');
  registerCalendarWrapper('create-calendar-event', 'create-specific-calendar-event');
  registerCalendarWrapper('update-calendar-event', 'update-specific-calendar-event');
  registerCalendarWrapper('delete-calendar-event', 'delete-specific-calendar-event');

  // Mail wrappers (folder/shared selectors)
  const registerListMailWrapper = () => {
    const base = getToolDefinitionOrWarn('list-mail-messages');
    const folder = getToolDefinitionOrWarn('list-mail-folder-messages');
    const shared = getToolDefinitionOrWarn('list-shared-mailbox-messages');
    const sharedFolder = getToolDefinitionOrWarn('list-shared-mailbox-folder-messages');
    if (!base || !folder || !shared || !sharedFolder) {
      return;
    }

    const schema = {
      ...cloneParamSchema(base.paramSchema),
      folderId: z
        .string()
        .describe(
          'Optional mail folder id to scope results. MUST be the folder\'s "id" field from list-mail-folders response (NOT parentFolderId). Use the exact "id" value of the target folder. Alias: mailFolderId (both accepted).'
        )
        .optional(),
      mailFolderId: z
        .string()
        .describe(
          'Optional mail folder id (alias for folderId). MUST be the folder\'s "id" field from list-mail-folders response (NOT parentFolderId). Use the exact "id" value of the target folder.'
        )
        .optional(),
      sharedMailboxId: z
        .string()
        .describe('Optional shared mailbox object id/UPN; takes precedence over sharedMailboxEmail')
        .optional(),
      sharedMailboxEmail: z
        .string()
        .describe('Optional shared mailbox SMTP address when sharedMailboxId is unknown')
        .optional(),
    };

    registerWrapperTool(base.alias, base.description, schema, base.metadata, async (params) => {
      // Accept both folderId and mailFolderId for compatibility
      const folderId = params.folderId || params.mailFolderId;
      const { sharedMailboxId, sharedMailboxEmail, ...rest } = params;
      // Remove mailFolderId from rest if it was used, to avoid passing it twice
      if (params.mailFolderId && !params.folderId) {
        delete (rest as any).mailFolderId;
      }
      
      const sharedTarget =
        typeof sharedMailboxId === 'string' && sharedMailboxId.trim().length > 0
          ? sharedMailboxId
          : typeof sharedMailboxEmail === 'string' && sharedMailboxEmail.trim().length > 0
          ? sharedMailboxEmail
          : undefined;

      if (sharedTarget && typeof folderId === 'string' && folderId.trim().length > 0) {
        return invokeInternalTool('list-shared-mailbox-folder-messages', {
          ...rest,
          userId: sharedTarget,
          mailFolderId: folderId,
        });
      }

      if (sharedTarget) {
        return invokeInternalTool('list-shared-mailbox-messages', { ...rest, userId: sharedTarget });
      }

      if (typeof folderId === 'string' && folderId.trim().length > 0) {
        return invokeInternalTool('list-mail-folder-messages', { ...rest, mailFolderId: folderId });
      }

      return invokeInternalTool('list-mail-messages', rest);
    });
  };

  registerListMailWrapper();

  const registerGetMailWrapper = () => {
    const base = getToolDefinitionOrWarn('get-mail-message');
    const shared = getToolDefinitionOrWarn('get-shared-mailbox-message');
    if (!base || !shared) {
      return;
    }

    const schema = {
      ...cloneParamSchema(base.paramSchema),
      sharedMailboxId: z
        .string()
        .describe('Optional shared mailbox object id/UPN; takes precedence over sharedMailboxEmail')
        .optional(),
      sharedMailboxEmail: z
        .string()
        .describe('Optional shared mailbox SMTP address when sharedMailboxId is unknown')
        .optional(),
    };

    registerWrapperTool(base.alias, base.description, schema, base.metadata, async (params) => {
      const { sharedMailboxId, sharedMailboxEmail, ...rest } = params;
      const sharedTarget =
        typeof sharedMailboxId === 'string' && sharedMailboxId.trim().length > 0
          ? sharedMailboxId
          : typeof sharedMailboxEmail === 'string' && sharedMailboxEmail.trim().length > 0
          ? sharedMailboxEmail
          : undefined;

      if (sharedTarget) {
        return invokeInternalTool('get-shared-mailbox-message', { ...rest, userId: sharedTarget });
      }

      return invokeInternalTool('get-mail-message', rest);
    });
  };

  registerGetMailWrapper();

  const registerSendMailWrapper = () => {
    const base = getToolDefinitionOrWarn('send-mail');
    const shared = getToolDefinitionOrWarn('send-shared-mailbox-mail');
    if (!base || !shared) {
      return;
    }

    const schema = {
      ...cloneParamSchema(base.paramSchema),
      sharedMailboxId: z
        .string()
        .describe('Optional shared mailbox object id/UPN; takes precedence over sharedMailboxEmail')
        .optional(),
      sharedMailboxEmail: z
        .string()
        .describe('Optional shared mailbox SMTP address when sharedMailboxId is unknown')
        .optional(),
    };

    const ensureSaveToSentItemsDefault = (
      body: unknown
    ): { payload: Record<string, unknown>; explicitlyDisabled: boolean } => {
      if (body && typeof body === 'object') {
        const cloned = { ...(body as Record<string, unknown>) };
        const explicitlyDisabled = cloned.SaveToSentItems === false;
        if (cloned.SaveToSentItems === undefined || cloned.SaveToSentItems === null) {
          cloned.SaveToSentItems = true;
        }
        return { payload: cloned, explicitlyDisabled };
      }
      return { payload: { SaveToSentItems: true }, explicitlyDisabled: false };
    };

    registerWrapperTool(base.alias, base.description, schema, base.metadata, async (params) => {
      const { sharedMailboxId, sharedMailboxEmail, ...rest } = params;
      const { payload: body, explicitlyDisabled } = ensureSaveToSentItemsDefault(
        (rest as Record<string, unknown>).body
      );

      if (explicitlyDisabled) {
        logger.warn(
          '[send-mail] SaveToSentItems=false requested; honoring request but this should be rare.'
        );
      }

      const payload = {
        ...rest,
        body,
      };
      const sharedTarget =
        typeof sharedMailboxId === 'string' && sharedMailboxId.trim().length > 0
          ? sharedMailboxId
          : typeof sharedMailboxEmail === 'string' && sharedMailboxEmail.trim().length > 0
          ? sharedMailboxEmail
          : undefined;

      if (sharedTarget) {
        return invokeInternalTool('send-shared-mailbox-mail', { ...payload, userId: sharedTarget });
      }

      return invokeInternalTool('send-mail', payload);
    });
  };

  registerSendMailWrapper();

  // Log big warning if any tools were sanitized
  if (sanitizedTools.length > 0) {
    logger.warn('═══════════════════════════════════════════════════════════════════════════════');
    logger.warn('⚠️  WARNING: TOOL NAME SANITIZATION DETECTED ⚠️');
    logger.warn('═══════════════════════════════════════════════════════════════════════════════');
    logger.warn(`${sanitizedTools.length} tool(s) had their names sanitized during registration because they were too long:`);
    for (const { original, sanitized } of sanitizedTools) {
      if (TOOL_BLACKLIST.includes(original as any)) {
        logger.info(`  •OK "${original}" → "${sanitized}" (used in wrapper tool)`);
      } else {
        logger.error(`  • "${original}" → "${sanitized}"`);
      }
    }
    logger.warn('This may cause issues on tools not marked as OK.');
    logger.warn('Consider updating tool names to be shorter.');
    logger.warn('═══════════════════════════════════════════════════════════════════════════════');
  }

  // NOTE: The MCP TypeScript SDK's server.tool() method only passes the 'arguments' 
  // field to tool handlers, not the full CallToolRequestParams which contains _meta.
  // This means _meta data sent by MCPO is currently inaccessible in tool handlers.
  // 
  // The SDK would need to be updated to either:
  // 1. Pass the full params object (including _meta) to tool handlers
  // 2. Provide a setRequestHandler() method to intercept requests before tool handlers
  // 3. Expose _meta through a different mechanism
  //
  // For now, impersonation will work via:
  // - HTTP middleware (for HTTP mode) -> ImpersonationContext  
  // - MS365_MCP_IMPERSONATE_USER env var (fallback for all modes)
}
