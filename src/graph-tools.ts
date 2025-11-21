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
      logger.warn(`Tool alias exceeds 64 chars, renaming`, { alias: tool.alias, safeName });
    }

    const finalDescription = getToolDescription(
      tool.alias,
      tool.description ||
        `Execute ${tool.method.toUpperCase()} request to ${tool.path} (${tool.alias})`
    );

    server.tool(
      safeName,
      finalDescription,
      paramSchema as any,
      {
        title: safeTitle,
        readOnlyHint: tool.method.toUpperCase() === 'GET',
      },
      async (params: any) => {
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
                  break;

                case 'Query':
                  if (fixedParamName === '$search') {
                    const sanitized = sanitizeAqsSearch(String(paramValue ?? ''));
                    queryParams[fixedParamName] = `${sanitized}`;
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
              .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
              .join('&');
            path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
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

          logger.info(`Making graph request to ${path} with options: ${JSON.stringify(options)}`);
          let response = await graphClient.graphRequest(path, options);

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
      }
    );
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
