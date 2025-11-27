# Performance Optimization Plan

## Overview

This document outlines performance optimization opportunities discovered during analysis of the MS-365 MCP server codebase. Each optimization includes a detailed implementation plan with clear phases that can be converted into actionable todo items.

---

## CRITICAL FIX: Cache Instance Sharing (COMPLETED ✓)

### Problem
The `MailboxDiscoveryCache` was instantiated per tool call, creating a new Map-based cache for each request. This completely negated the intended performance benefit of caching.

### Solution Implemented
- Created module-level shared cache instance in `graph-tools.ts`
- Cache is initialized once on first `registerGraphTools()` call
- All tool calls now share the same cache instance

### Impact
- **Before**: 5-10 Graph API calls per request (0% cache hits)
- **After**: 5-10 Graph API calls on first request, ~0ms on subsequent requests (near-100% cache hits within TTL)
- **Performance gain**: ~500-2000ms per request after first successful discovery

---

## OPTIMIZATION 1: Token Refresh Enhancement

### Priority: HIGH
### Estimated Impact: Medium-High (reduces external API calls)
### Complexity: Medium

### Current State
- Every 401 response triggers fresh token refresh API call to Microsoft
- No caching of refreshed tokens
- No mutex protection against concurrent refresh attempts
- No proactive refresh before expiry

### Problem Analysis
1. **Race conditions**: Multiple concurrent requests with expired token each attempt refresh
2. **Redundant API calls**: Successful refreshes not cached/shared across requests
3. **User experience**: Requests wait for 401 before attempting refresh (slower)

### Implementation Plan

#### Phase 1: Token Metadata Tracking
**File**: `src/graph-client.ts`

Add token expiry tracking to GraphClient:

```typescript
class GraphClient {
  private accessToken: string | null = null;
  private refreshToken: string | null = null;
  private tokenExpiresAt: number = 0; // Unix timestamp
  private tokenRefreshPromise: Promise<void> | null = null; // Mutex

  setOAuthTokens(accessToken: string, refreshToken?: string, expiresIn?: number): void {
    this.accessToken = accessToken;
    this.refreshToken = refreshToken || null;
    
    // Calculate expiry timestamp (default to 1 hour if not provided)
    const expiresInSeconds = expiresIn || 3600;
    this.tokenExpiresAt = Date.now() + (expiresInSeconds * 1000);
    
    logger.debug('OAuth tokens updated', {
      expiresAt: new Date(this.tokenExpiresAt).toISOString(),
      expiresIn: expiresInSeconds
    });
  }
}
```

**Changes needed**:
- Add `tokenExpiresAt` property
- Add `tokenRefreshPromise` property for mutex
- Update `setOAuthTokens()` signature to accept `expiresIn` parameter
- Calculate and store token expiry timestamp

#### Phase 2: Proactive Token Refresh
**File**: `src/graph-client.ts`

Implement proactive refresh logic:

```typescript
private async ensureValidToken(): Promise<string> {
  // If token expires in less than 5 minutes, proactively refresh
  const REFRESH_THRESHOLD_MS = 5 * 60 * 1000; // 5 minutes
  const now = Date.now();
  
  if (this.tokenExpiresAt - now < REFRESH_THRESHOLD_MS) {
    logger.info('Token expiring soon, proactively refreshing', {
      expiresAt: new Date(this.tokenExpiresAt).toISOString(),
      timeRemaining: Math.floor((this.tokenExpiresAt - now) / 1000) + 's'
    });
    
    await this.performTokenRefresh();
  }
  
  if (!this.accessToken) {
    throw new Error('No access token available');
  }
  
  return this.accessToken;
}
```

**Configuration**:
- Add `MS365_MCP_TOKEN_REFRESH_THRESHOLD` env var (default: 300 seconds)
- Configurable threshold for when to trigger proactive refresh

#### Phase 3: Refresh Mutex Implementation
**File**: `src/graph-client.ts`

Prevent concurrent refresh attempts:

```typescript
private async performTokenRefresh(): Promise<void> {
  // If refresh already in progress, wait for it
  if (this.tokenRefreshPromise) {
    logger.debug('Token refresh already in progress, waiting...');
    await this.tokenRefreshPromise;
    return;
  }
  
  // Start new refresh with mutex
  this.tokenRefreshPromise = this._doTokenRefresh();
  
  try {
    await this.tokenRefreshPromise;
  } finally {
    this.tokenRefreshPromise = null;
  }
}

private async _doTokenRefresh(): Promise<void> {
  if (!this.refreshToken) {
    throw new Error('No refresh token available');
  }
  
  logger.info('Refreshing access token');
  const tenantId = process.env.MS365_MCP_TENANT_ID || 'common';
  const clientId = process.env.MS365_MCP_CLIENT_ID || '084a3e9f-a9f4-43f7-89f9-d229cf97853e';
  const clientSecret = process.env.MS365_MCP_CLIENT_SECRET;
  
  if (!clientSecret) {
    throw new Error('MS365_MCP_CLIENT_SECRET not configured');
  }
  
  const response = await refreshAccessToken(
    this.refreshToken, 
    clientId, 
    clientSecret, 
    tenantId
  );
  
  // Update tokens with expiry info
  this.setOAuthTokens(
    response.access_token,
    response.refresh_token,
    response.expires_in
  );
  
  logger.info('Access token refreshed successfully', {
    expiresIn: response.expires_in
  });
}
```

#### Phase 4: Integration
**File**: `src/graph-client.ts`

Update `makeRequest()` to use proactive refresh:

```typescript
async makeRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<unknown> {
  // Ensure token is valid before making request
  let accessToken = options.accessToken || this.accessToken;
  
  if (!accessToken) {
    accessToken = await this.authManager.getToken();
  } else if (this.refreshToken) {
    // Use proactive refresh if we have refresh token
    await this.ensureValidToken();
    accessToken = this.accessToken!;
  }
  
  // ... rest of existing logic
}
```

#### Phase 5: Response Processing
**File**: `src/lib/microsoft-auth.ts`

Ensure `refreshAccessToken()` returns `expires_in`:

```typescript
interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number; // Add this field
  token_type: string;
  scope: string;
}
```

### Testing Checklist
- [ ] Verify token refresh triggered before expiry (not after 401)
- [ ] Confirm concurrent requests don't trigger multiple refreshes
- [ ] Test with varying `MS365_MCP_TOKEN_REFRESH_THRESHOLD` values
- [ ] Measure reduction in 401 responses
- [ ] Verify no regression in token handling

### Environment Variables
```bash
# Token refresh threshold (seconds before expiry to trigger refresh)
# Default: 300 (5 minutes)
MS365_MCP_TOKEN_REFRESH_THRESHOLD=300
```

---

## OPTIMIZATION 2: HTTP Connection Pooling

### Priority: HIGH  
### Estimated Impact: Medium (reduces TLS handshake overhead)
### Complexity: Low-Medium

### Current State
- Uses native `fetch()` without HTTP agent configuration
- Each request likely creates new TCP connection
- TLS handshake overhead on every request (~100-300ms)

### Problem Analysis
1. **Connection overhead**: New TCP + TLS handshake per request
2. **Resource usage**: Unnecessary socket creation/destruction
3. **Latency**: Handshake adds 100-300ms per request

### Implementation Plan

#### Phase 1: HTTP Agent Setup
**File**: `src/graph-client.ts`

Install and configure HTTP agent:

```bash
npm install undici
```

Add agent configuration:

```typescript
import { Agent, setGlobalDispatcher } from 'undici';

// Module-level HTTP agent for connection pooling
const httpAgent = new Agent({
  connections: 128, // Max concurrent sockets
  pipelining: 1,    // HTTP pipelining depth
  keepAliveTimeout: 60000, // 60 seconds keep-alive
  keepAliveMaxTimeout: 300000, // 5 minutes max
});

// Set as global dispatcher for all fetches
setGlobalDispatcher(httpAgent);
```

**Configuration via env vars**:
```bash
MS365_MCP_HTTP_POOL_SIZE=128
MS365_MCP_HTTP_KEEPALIVE_TIMEOUT=60000
MS365_MCP_HTTP_KEEPALIVE_MAX=300000
```

#### Phase 2: Connection Metrics
**File**: `src/graph-client.ts`

Add connection pool monitoring:

```typescript
private logConnectionStats(): void {
  if (process.env.MS365_MCP_DEBUG === 'true') {
    const stats = httpAgent.stats();
    logger.debug('HTTP connection pool stats', {
      connected: stats.connected,
      pending: stats.pending,
      running: stats.running,
      size: stats.size
    });
  }
}
```

#### Phase 3: Graceful Shutdown
**File**: `src/server.ts`

Add cleanup for HTTP connections:

```typescript
async shutdown(): Promise<void> {
  logger.info('Shutting down HTTP connection pool');
  await httpAgent.close();
  logger.info('Server shutdown complete');
}
```

### Testing Checklist
- [ ] Verify connection reuse with network inspection
- [ ] Measure latency improvement (expect ~100-200ms per request)
- [ ] Test connection limit enforcement
- [ ] Verify graceful shutdown doesn't leak connections
- [ ] Test with high concurrency (50+ concurrent requests)

### Expected Performance Gain
- **First request**: No change (~300ms TLS handshake)
- **Subsequent requests**: -100-200ms per request

---

## OPTIMIZATION 3: Response Processing Efficiency

### Priority: MEDIUM
### Estimated Impact: Low-Medium (reduces CPU usage)
### Complexity: Medium

### Current State
- Large responses parsed/stringified multiple times
- Body scrubbing iterates through all items
- Truncation happens after full parse
- Multiple serialization passes

### Problem Analysis
1. **CPU overhead**: Multiple parse/stringify cycles on large responses
2. **Memory pressure**: Entire response held in memory during processing
3. **Inefficiency**: Processing happens even when data is truncated

### Implementation Plan

#### Phase 1: Single-Pass Processing
**File**: `src/graph-tools.ts`

Optimize scrubBodies to work in-place:

```typescript
const scrubBodies = (
  json: any,
  includeBody: boolean,
  bodyFormat?: 'html' | 'text',
  previewCap: number = 800,
  maxItems?: number // New: early termination
): void => {
  if (!json || typeof json !== 'object') return;

  const processOne = (item: any, index: number) => {
    // Early termination if maxItems specified
    if (maxItems !== undefined && index >= maxItems) {
      return false; // Signal to stop processing
    }
    
    // ... existing scrubbing logic
    return true; // Continue processing
  };

  if (Array.isArray(json.value)) {
    // Process with early termination support
    for (let i = 0; i < json.value.length; i++) {
      if (!processOne(json.value[i], i)) {
        // Truncate array if we hit limit
        json.value = json.value.slice(0, i);
        json._truncated = true;
        break;
      }
    }
  } else {
    processOne(json, 0);
  }
};
```

#### Phase 2: Lazy Body Conversion
**File**: `src/graph-tools.ts`

Only convert HTML to text when explicitly requested:

```typescript
// Instead of always stripping HTML, keep original and convert on-demand
if (!includeBody) {
  // Store minimal preview, defer full conversion
  item.bodyPreview = (item.bodyPreview || '').slice(0, previewCap);
  item.body = { content: '[body omitted - use includeBody=true to retrieve]' };
} else if (bodyFormat === 'text' && item.body?.contentType === 'HTML') {
  // Only strip HTML when explicitly requested as text
  item.body.content = stripHtml(item.body.content);
  item.body.contentType = 'Text';
}
```

#### Phase 3: Stream Processing for Large Responses
**File**: `src/graph-client.ts`

Add streaming support for very large responses:

```typescript
async graphRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<McpResponse> {
  const MAX_MEMORY_SIZE = Number(process.env.MS365_MCP_MAX_RESPONSE_SIZE || '10485760'); // 10MB
  
  try {
    const result = await this.makeRequest(endpoint, options);
    const resultStr = JSON.stringify(result);
    
    // If response is huge, switch to streaming mode
    if (resultStr.length > MAX_MEMORY_SIZE) {
      logger.warn('Response exceeds memory limit, using streaming mode', {
        size: resultStr.length,
        limit: MAX_MEMORY_SIZE
      });
      
      return this.streamLargeResponse(result, options);
    }
    
    return this.formatJsonResponse(result, options.rawResponse, options.excludeResponse);
  } catch (error) {
    // ... error handling
  }
}

private streamLargeResponse(data: unknown, options: GraphRequestOptions): McpResponse {
  // Implement chunked processing for very large responses
  // This is a future enhancement for edge cases
  logger.warn('Stream processing not yet implemented, truncating response');
  const truncated = JSON.stringify(data).slice(0, 200000);
  return {
    content: [{ type: 'text', text: truncated }],
    _meta: { truncated: true, reason: 'exceeded_memory_limit' }
  };
}
```

### Testing Checklist
- [ ] Test with large email responses (100+ items)
- [ ] Verify HTML stripping only when needed
- [ ] Measure CPU usage reduction
- [ ] Test early termination with maxItems
- [ ] Benchmark parse/stringify cycles

### Environment Variables
```bash
# Maximum response size before streaming mode (bytes)
# Default: 10485760 (10MB)
MS365_MCP_MAX_RESPONSE_SIZE=10485760

# Maximum items to process in array responses
# Default: unlimited
MS365_MCP_MAX_ITEMS=1000
```

---

## OPTIMIZATION 4: Pagination Parallelization

### Priority: MEDIUM
### Estimated Impact: High (when fetchAllPages used)
### Complexity: Medium

### Current State
- Sequential page fetching (wait for page N before fetching N+1)
- No concurrency control
- Hard 100-page limit

### Problem Analysis
1. **Latency multiplication**: 100 pages × 500ms = 50 seconds
2. **No concurrency**: Only 1 request at a time
3. **Inefficient**: Could fetch 3-5 pages concurrently

### Implementation Plan

#### Phase 1: Concurrent Page Fetcher
**File**: `src/graph-tools.ts`

Implement parallel pagination:

```typescript
async function fetchPagesInParallel(
  graphClient: GraphClient,
  initialResponse: any,
  options: any,
  concurrency: number = 3
): Promise<any> {
  const allItems = initialResponse.value || [];
  let nextLinks: string[] = [];
  
  if (initialResponse['@odata.nextLink']) {
    nextLinks.push(initialResponse['@odata.nextLink']);
  }
  
  const maxPages = Number(process.env.MS365_MCP_MAX_PAGES || '100');
  let pageCount = 1;
  
  // Process pages in batches with concurrency control
  while (nextLinks.length > 0 && pageCount < maxPages) {
    const batch = nextLinks.splice(0, concurrency);
    
    logger.debug(`Fetching page batch`, {
      batchSize: batch.length,
      totalPages: pageCount,
      remainingLinks: nextLinks.length
    });
    
    // Fetch batch concurrently
    const responses = await Promise.all(
      batch.map(async (link) => {
        try {
          const url = new URL(link);
          const path = url.pathname.replace('/v1.0', '');
          const queryParams: Record<string, string> = {};
          
          for (const [key, value] of url.searchParams.entries()) {
            queryParams[key] = value;
          }
          
          return await graphClient.graphRequest(path, {
            ...options,
            queryParams
          });
        } catch (error) {
          logger.error(`Error fetching page: ${error}`);
          return null;
        }
      })
    );
    
    // Process responses
    for (const response of responses) {
      if (!response?.content?.[0]?.text) continue;
      
      try {
        const pageData = JSON.parse(response.content[0].text);
        
        if (pageData.value && Array.isArray(pageData.value)) {
          allItems.push(...pageData.value);
        }
        
        if (pageData['@odata.nextLink']) {
          nextLinks.push(pageData['@odata.nextLink']);
        }
        
        pageCount++;
      } catch (error) {
        logger.error(`Error parsing page response: ${error}`);
      }
    }
  }
  
  return {
    value: allItems,
    '@odata.count': allItems.length,
    _pagesFetched: pageCount
  };
}
```

#### Phase 2: Integration
**File**: `src/graph-tools.ts`

Replace sequential pagination with parallel:

```typescript
const fetchAllPages = params.fetchAllPages === true;
if (fetchAllPages && response?.content?.[0]?.text) {
  try {
    const initialResponse = JSON.parse(response.content[0].text);
    
    if (initialResponse['@odata.nextLink']) {
      const concurrency = Number(process.env.MS365_MCP_PAGE_CONCURRENCY || '3');
      
      logger.info('Starting parallel pagination', {
        concurrency,
        firstPageItems: initialResponse.value?.length || 0
      });
      
      const combinedResponse = await fetchPagesInParallel(
        graphClient,
        initialResponse,
        options,
        concurrency
      );
      
      response.content[0].text = JSON.stringify(combinedResponse);
      
      logger.info('Pagination complete', {
        totalItems: combinedResponse.value.length,
        pages: combinedResponse._pagesFetched
      });
    }
  } catch (error) {
    logger.error(`Error during parallel pagination: ${error}`);
  }
}
```

#### Phase 3: Smart Concurrency
**File**: `src/graph-tools.ts`

Adjust concurrency based on response times:

```typescript
class AdaptiveConcurrencyController {
  private concurrency: number;
  private maxConcurrency: number;
  private minConcurrency: number;
  private avgResponseTime: number = 0;
  private requestCount: number = 0;
  
  constructor(initial: number = 3) {
    this.concurrency = initial;
    this.maxConcurrency = Number(process.env.MS365_MCP_MAX_PAGE_CONCURRENCY || '5');
    this.minConcurrency = 1;
  }
  
  recordResponse(durationMs: number): void {
    this.requestCount++;
    this.avgResponseTime = 
      (this.avgResponseTime * (this.requestCount - 1) + durationMs) / this.requestCount;
    
    // Increase concurrency if responses are fast
    if (this.avgResponseTime < 300 && this.concurrency < this.maxConcurrency) {
      this.concurrency++;
      logger.debug('Increasing pagination concurrency', { 
        concurrency: this.concurrency,
        avgResponseTime: this.avgResponseTime
      });
    }
    
    // Decrease if responses are slow
    if (this.avgResponseTime > 1000 && this.concurrency > this.minConcurrency) {
      this.concurrency--;
      logger.debug('Decreasing pagination concurrency', {
        concurrency: this.concurrency,
        avgResponseTime: this.avgResponseTime
      });
    }
  }
  
  getConcurrency(): number {
    return this.concurrency;
  }
}
```

### Testing Checklist
- [ ] Test with large datasets (1000+ items across many pages)
- [ ] Verify concurrency limits are respected
- [ ] Test with slow network (ensure no timeout issues)
- [ ] Confirm all pages are fetched correctly
- [ ] Measure speedup vs sequential (expect 2-3x faster)

### Environment Variables
```bash
# Concurrent page fetches
# Default: 3
MS365_MCP_PAGE_CONCURRENCY=3

# Maximum concurrent page fetches
# Default: 5
MS365_MCP_MAX_PAGE_CONCURRENCY=5

# Maximum total pages to fetch
# Default: 100
MS365_MCP_MAX_PAGES=100
```

### Expected Performance Gain
- **100 pages sequential**: ~50 seconds
- **100 pages parallel (3x)**: ~17 seconds
- **Speedup**: 3x faster

---

## OPTIMIZATION 5: Request Deduplication

### Priority: LOW-MEDIUM
### Estimated Impact: Medium (in high-concurrency scenarios)
### Complexity: Medium-High

### Current State
- No deduplication of in-flight requests
- Multiple identical requests all hit the API
- Especially wasteful for mailbox discovery

### Problem Analysis
1. **Duplicate work**: Same endpoint + params = duplicate API calls
2. **Resource waste**: Unnecessary network/compute overhead
3. **Rate limiting**: Duplicate requests count against quotas

### Implementation Plan

#### Phase 1: Request Key Generation
**File**: `src/utils/request-cache.ts` (NEW)

Create request deduplication utility:

```typescript
import crypto from 'crypto';

export interface PendingRequest<T> {
  promise: Promise<T>;
  createdAt: number;
}

export class RequestDeduplicator<T = any> {
  private pending = new Map<string, PendingRequest<T>>();
  private ttl: number;
  
  constructor(ttlMs: number = 5000) {
    this.ttl = ttlMs;
  }
  
  generateKey(method: string, path: string, params?: Record<string, any>): string {
    const data = JSON.stringify({ method, path, params });
    return crypto.createHash('md5').update(data).digest('hex');
  }
  
  async deduplicate<R extends T>(
    key: string,
    executor: () => Promise<R>
  ): Promise<R> {
    // Check for existing in-flight request
    const existing = this.pending.get(key);
    
    if (existing) {
      const age = Date.now() - existing.createdAt;
      
      if (age < this.ttl) {
        logger.debug('Request deduplication hit', { key, age });
        return existing.promise as Promise<R>;
      } else {
        // Expired, remove it
        this.pending.delete(key);
      }
    }
    
    // Execute new request
    const promise = executor();
    
    this.pending.set(key, {
      promise,
      createdAt: Date.now()
    });
    
    // Clean up when complete
    promise
      .finally(() => {
        this.pending.delete(key);
      })
      .catch(() => {}); // Prevent unhandled rejection
    
    return promise;
  }
  
  clear(): void {
    this.pending.clear();
  }
  
  size(): number {
    return this.pending.size;
  }
}
```

#### Phase 2: Integration with GraphClient
**File**: `src/graph-client.ts`

Add deduplication to API requests:

```typescript
import { RequestDeduplicator } from './utils/request-cache.js';

class GraphClient {
  // ... existing properties
  private requestDeduplicator: RequestDeduplicator;
  
  constructor(authManager: TokenProvider) {
    this.authManager = authManager;
    
    const dedupTTL = Number(process.env.MS365_MCP_REQUEST_DEDUP_TTL || '5000');
    this.requestDeduplicator = new RequestDeduplicator(dedupTTL);
  }
  
  async makeRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<unknown> {
    const deduplicationEnabled = process.env.MS365_MCP_REQUEST_DEDUPLICATION !== 'false';
    
    if (!deduplicationEnabled || options.method !== 'GET') {
      // Only deduplicate GET requests
      return this._makeRequestInternal(endpoint, options);
    }
    
    const key = this.requestDeduplicator.generateKey(
      options.method || 'GET',
      endpoint,
      options
    );
    
    return this.requestDeduplicator.deduplicate(key, () => 
      this._makeRequestInternal(endpoint, options)
    );
  }
  
  private async _makeRequestInternal(
    endpoint: string, 
    options: GraphRequestOptions
  ): Promise<unknown> {
    // Existing makeRequest logic moved here
    // ...
  }
}
```

#### Phase 3: Metrics & Monitoring
**File**: `src/graph-client.ts`

Add deduplication metrics:

```typescript
private deduplicationStats = {
  hits: 0,
  misses: 0,
  savedRequests: 0
};

logDeduplicationStats(): void {
  if (this.deduplicationStats.hits > 0 || this.deduplicationStats.misses > 0) {
    const total = this.deduplicationStats.hits + this.deduplicationStats.misses;
    const hitRate = (this.deduplicationStats.hits / total * 100).toFixed(1);
    
    logger.info('Request deduplication stats', {
      hits: this.deduplicationStats.hits,
      misses: this.deduplicationStats.misses,
      hitRate: `${hitRate}%`,
      savedRequests: this.deduplicationStats.savedRequests,
      pendingRequests: this.requestDeduplicator.size()
    });
  }
}
```

### Testing Checklist
- [ ] Test concurrent identical requests (verify only 1 API call)
- [ ] Test slightly different requests (verify separate calls)
- [ ] Test TTL expiration
- [ ] Verify no memory leaks (pending requests cleared)
- [ ] Measure hit rate in production workloads

### Environment Variables
```bash
# Enable/disable request deduplication
# Default: true
MS365_MCP_REQUEST_DEDUPLICATION=true

# TTL for in-flight request tracking (ms)
# Default: 5000 (5 seconds)
MS365_MCP_REQUEST_DEDUP_TTL=5000
```

### Expected Performance Gain
- **High concurrency**: Up to 50% reduction in duplicate requests
- **Typical usage**: 10-20% reduction in total API calls

---

## OPTIMIZATION 6: Logging Optimization

### Priority: LOW
### Estimated Impact: Low (reduces CPU in debug mode)
### Complexity: Low

### Current State
- Extensive `JSON.stringify()` calls for all log levels
- Stringification happens even when logs aren't written
- No lazy evaluation

### Problem Analysis
1. **Wasted CPU**: Stringify large objects even when debug logging disabled
2. **Memory pressure**: Large strings created unnecessarily
3. **Code clarity**: Harder to read with inline stringification

### Implementation Plan

#### Phase 1: Lazy Log Message Evaluation
**File**: `src/logger.ts`

Add lazy evaluation support:

```typescript
type LogMessage = string | (() => string);

class Logger {
  // ... existing code
  
  debug(message: string, data?: any | (() => any)): void {
    if (!this.shouldLog('debug')) {
      return; // Skip evaluation entirely
    }
    
    const resolvedData = typeof data === 'function' ? data() : data;
    
    // ... existing log logic
  }
  
  info(message: string, data?: any | (() => any)): void {
    if (!this.shouldLog('info')) {
      return;
    }
    
    const resolvedData = typeof data === 'function' ? data() : data;
    
    // ... existing log logic
  }
  
  private shouldLog(level: string): boolean {
    const levels = ['error', 'warn', 'info', 'debug'];
    const currentLevel = process.env.MS365_MCP_LOG_LEVEL || 'info';
    
    return levels.indexOf(level) <= levels.indexOf(currentLevel);
  }
}
```

#### Phase 2: Update Call Sites
**File**: `src/graph-tools.ts` and others

Use lazy evaluation for expensive stringification:

```typescript
// Before: Always stringifies
logger.debug(`Response preview: ${JSON.stringify(largeObject)}`);

// After: Only stringifies if debug logging enabled
logger.debug('Response preview', () => JSON.stringify(largeObject));

// Before: Complex computation always runs
logger.debug(`Computed value: ${expensiveComputation(data)}`);

// After: Computation only runs if needed
logger.debug('Computed value', () => expensiveComputation(data));
```

#### Phase 3: Structured Logging
**File**: `src/logger.ts`

Support structured logging without stringification:

```typescript
info(message: string, meta?: Record<string, any>): void {
  if (!this.shouldLog('info')) {
    return;
  }
  
  // For JSON format, include meta directly (no stringify needed)
  if (this.format === 'json') {
    console.log(JSON.stringify({
      timestamp: new Date().toISOString(),
      level: 'info',
      message,
      ...meta // Meta included directly, single stringify
    }));
  } else {
    // For console format, pretty print
    console.log(`[INFO] ${message}`, meta || '');
  }
}
```

### Testing Checklist
- [ ] Verify debug logs don't impact performance when disabled
- [ ] Test lazy evaluation actually skips computation
- [ ] Measure CPU reduction in production mode
- [ ] Ensure structured logging formats correctly

### Expected Performance Gain
- **Debug mode disabled**: 5-10% CPU reduction (avoids stringify)
- **Debug mode enabled**: No change (same work done)

---

## OPTIMIZATION 7: Rate Limiting Efficiency

### Priority: LOW
### Estimated Impact: Low (prevents memory leaks)
### Complexity: Low

### Current State  
- Map-based rate limiting without cleanup
- Map grows unbounded in long-running servers
- No LRU eviction

### Problem Analysis
1. **Memory leak**: Rate limit buckets never removed
2. **Growing overhead**: Map operations slow as size increases
3. **Resource waste**: Tracking expired buckets indefinitely

### Implementation Plan

#### Phase 1: Periodic Cleanup
**File**: `src/server.ts`

Add cleanup interval:

```typescript
// In HTTP server setup
const rateLimitCleanupInterval = setInterval(() => {
  const now = Date.now();
  const windowMs = Number(process.env.MS365_MCP_RATE_LIMIT_WINDOW_MS || '60000');
  
  let cleaned = 0;
  for (const [ip, bucket] of buckets.entries()) {
    if (now - bucket.windowStart >= windowMs * 2) {
      buckets.delete(ip);
      cleaned++;
    }
  }
  
  if (cleaned > 0) {
    logger.debug('Cleaned up expired rate limit buckets', {
      cleaned,
      remaining: buckets.size
    });
  }
}, 60000); // Run every minute

// Clean up on shutdown
process.on('SIGTERM', () => {
  clearInterval(rateLimitCleanupInterval);
});
```

#### Phase 2: LRU Implementation (Optional)
**File**: `src/utils/lru-cache.ts` (NEW)

For very high-traffic scenarios:

```typescript
export class LRUCache<K, V> {
  private cache = new Map<K, V>();
  private maxSize: number;
  
  constructor(maxSize: number = 10000) {
    this.maxSize = maxSize;
  }
  
  get(key: K): V | undefined {
    const value = this.cache.get(key);
    if (value !== undefined) {
      // Move to end (most recently used)
      this.cache.delete(key);
      this.cache.set(key, value);
    }
    return value;
  }
  
  set(key: K, value: V): void {
    // Remove if exists (will re-add at end)
    this.cache.delete(key);
    
    // Evict oldest if at capacity
    if (this.cache.size >= this.maxSize) {
      const firstKey = this.cache.keys().next().value;
      this.cache.delete(firstKey);
    }
    
    this.cache.set(key, value);
  }
  
  size(): number {
    return this.cache.size;
  }
}
```

### Testing Checklist
- [ ] Monitor bucket count in long-running instance
- [ ] Verify cleanup runs and removes expired buckets
- [ ] Test with high traffic (thousands of IPs)
- [ ] Ensure no memory leaks after cleanup

### Environment Variables
```bash
# Rate limit bucket cleanup interval (ms)
# Default: 60000 (1 minute)
MS365_MCP_RATE_LIMIT_CLEANUP_INTERVAL=60000

# Maximum rate limit buckets to track (LRU eviction)
# Default: 10000
MS365_MCP_RATE_LIMIT_MAX_BUCKETS=10000
```

---

## Summary & Recommendations

### Implementation Priority

1. **✅ COMPLETED**: Cache Instance Sharing (immediate ~500-2000ms improvement)
2. **HIGH**: Token Refresh Enhancement (prevents redundant auth calls)
3. **HIGH**: HTTP Connection Pooling (~100-200ms per request)
4. **MEDIUM**: Pagination Parallelization (3x speedup for multi-page fetches)
5. **MEDIUM**: Response Processing Efficiency (CPU reduction)
6. **LOW-MEDIUM**: Request Deduplication (benefits high-concurrency scenarios)
7. **LOW**: Logging Optimization (cleanup)
8. **LOW**: Rate Limiting Efficiency (prevents memory leaks)

### Quick Wins

Start with these three for maximum impact with minimal effort:

1. **Cache fix** (✅ done) - Immediate dramatic improvement
2. **HTTP connection pooling** - Low complexity, high impact
3. **Token refresh** - Medium complexity, prevents cascading failures

### Performance Testing Strategy

After implementing each optimization:

1. **Baseline measurement**: Record current performance metrics
2. **A/B testing**: Enable optimization for subset of requests
3. **Monitoring**: Track latency, throughput, error rates
4. **Gradual rollout**: Increase traffic percentage if metrics improve
5. **Documentation**: Record actual performance gains

### Metrics to Track

```typescript
interface PerformanceMetrics {
  requestLatency: {
    p50: number;
    p95: number;
    p99: number;
  };
  cacheHitRate: number;
  tokenRefreshCount: number;
  duplicateRequestsPrevented: number;
  apiCallsSaved: number;
  connectionReuseRate: number;
}
```

---

## Converting to TODO List

Each optimization phase can be converted into actionable todo items. For example, for **Optimization 1: Token Refresh Enhancement**:

```markdown
## Token Refresh Enhancement
- [ ] Phase 1: Add token expiry tracking
  - [ ] Add `tokenExpiresAt` property to GraphClient
  - [ ] Add `tokenRefreshPromise` mutex property
  - [ ] Update `setOAuthTokens()` signature
  - [ ] Calculate and store expiry timestamp
- [ ] Phase 2: Implement proactive refresh
  - [ ] Create `ensureValidToken()` method
  - [ ] Add `MS365_MCP_TOKEN_REFRESH_THRESHOLD` env var
  - [ ] Implement threshold checking logic
- [ ] Phase 3: Add refresh mutex
  - [ ] Create `performTokenRefresh()` method
  - [ ] Create `_doTokenRefresh()` internal method
  - [ ] Prevent concurrent refreshes
- [ ] Phase 4: Integration
  - [ ] Update `makeRequest()` to call `ensureValidToken()`
  - [ ] Test token refresh before 401
- [ ] Phase 5: Update response processing
  - [ ] Ensure `expires_in` field in token response
  - [ ] Update TypeScript interfaces
- [ ] Testing
  - [ ] Verify proactive refresh triggers
  - [ ] Test concurrent request handling
  - [ ] Measure 401 reduction
```

Each optimization in this document follows the same pattern and can be converted to a similar checklist structure.

---

## Notes

- All environment variables maintain backward compatibility (sensible defaults)
- Optimizations are designed to be independent (can be implemented separately)
- Each optimization includes rollback strategy (disable via env var)
- Performance gains are estimates based on typical Microsoft Graph API latency
- Actual improvements may vary based on network conditions and workload patterns
