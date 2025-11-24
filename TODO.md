# Mailbox Discovery & Impersonation - Implementation Tasks

## Phase 1: Core Discovery Service
- [ ] Create `src/impersonation/MailboxDiscoveryService.ts`
  - [ ] Extract Strategy 1 (calendar permissions) from auth.ts
  - [ ] Extract Strategy 2 (shared mailbox detection) from auth.ts
  - [ ] Extract Strategy 3 (SendAs permissions) from auth.ts
  - [ ] Extract Strategy 4 (mailbox access verification) from auth.ts
  - [ ] Add constructor with GraphClient dependency
  - [ ] Implement `discoverMailboxes(userEmail)` main method
  - [ ] Add JSDoc documentation for each method
  - [ ] Include timeout handling (5s default)
  - [ ] Include concurrent request limiting (5 max)
  - [ ] Add comprehensive error handling

## Phase 2: Optional User Validation
- [ ] Create `src/impersonation/UserValidationCache.ts`
  - [ ] Implement email format validation (regex)
  - [ ] Implement user existence check via Graph API
  - [ ] Add TTL-based caching logic
  - [ ] Add constructor with GraphClient and TTL parameters
  - [ ] Implement `validateUser(email)` method
  - [ ] Add JSDoc documentation
  - [ ] Handle cache expiration
  - [ ] Add error handling for Graph API failures

## Phase 3: Refactor MailboxDiscoveryCache
- [ ] Update `src/impersonation/MailboxDiscoveryCache.ts`
  - [ ] Remove all references to `MS365_MCP_IMPERSONATE_ALLOWED_MAILBOXES`
  - [ ] Add MailboxDiscoveryService instantiation in constructor
  - [ ] Update `getMailboxes()` to use service
  - [ ] Keep existing cache validation logic
  - [ ] Update JSDoc documentation
  - [ ] Test cache hit/miss scenarios

## Phase 4: Impersonation Resolver
- [ ] Create `src/impersonation/ImpersonationResolver.ts`
  - [ ] Implement source precedence logic (header > context > env)
  - [ ] Add email format validation
  - [ ] Add optional user existence validation
  - [ ] Implement fallback behavior for empty values
  - [ ] Create clear error messages with source info
  - [ ] Add constructor with optional UserValidationCache
  - [ ] Implement `resolveImpersonatedUser()` method
  - [ ] Add comprehensive JSDoc documentation

## Phase 5: Integration Updates
- [ ] Update `src/graph-tools.ts`
  - [ ] Add MailboxDiscoveryCache initialization
  - [ ] Add optional UserValidationCache initialization
  - [ ] Add ImpersonationResolver initialization
  - [ ] Replace verbose impersonation detection with resolver
  - [ ] Update logging to show resolution source
  - [ ] Test integration with tool handlers
- [ ] Update `src/auth.ts`
  - [ ] Remove duplicate discovery strategies (~250 lines)
  - [ ] Add MailboxDiscoveryService usage
  - [ ] Update `listImpersonatedMailboxes()` method
  - [ ] Keep error handling and formatting
  - [ ] Test CLI auth commands

## Phase 6: Documentation
- [ ] Create `IMPERSONATION.md`
  - [ ] Write overview section
  - [ ] Document configuration variables
  - [ ] Create source precedence table
  - [ ] Explain validation modes
  - [ ] Document mailbox discovery process
  - [ ] Explain caching strategy
  - [ ] Add troubleshooting section
  - [ ] Provide example scenarios
- [ ] Update `.env.example`
  - [ ] Add impersonation configuration section
  - [ ] Document all new environment variables
  - [ ] Provide clear comments and defaults

## Phase 7: Cleanup
- [ ] Remove obsolete code
  - [ ] Search for `MS365_MCP_IMPERSONATE_ALLOWED_MAILBOXES` references
  - [ ] Remove env var from all files
  - [ ] Remove duplicate discovery logic from auth.ts
  - [ ] Clean up verbose debug logging in graph-tools.ts
- [ ] Update exports
  - [ ] Update `src/impersonation/index.ts` with new exports
  - [ ] Export MailboxDiscoveryService
  - [ ] Export UserValidationCache
  - [ ] Export ImpersonationResolver
  - [ ] Export MailboxInfo type

## Testing & Verification
- [ ] Manual testing scenarios
  - [ ] Test basic impersonation (env var)
  - [ ] Test header override
  - [ ] Test empty header fallback
  - [ ] Test validation enabled with invalid user
  - [ ] Test user with no shared mailboxes
  - [ ] Test user with shared mailbox access
  - [ ] Test cache performance (hit/miss/expiration)
- [ ] Code review
  - [ ] Review all new files for quality
  - [ ] Check JSDoc documentation completeness
  - [ ] Verify error messages are clear
  - [ ] Confirm no duplicate code remains
  - [ ] Validate performance optimizations

## Future Improvements
- [ ] Add unit tests for MailboxDiscoveryService
- [ ] Add unit tests for UserValidationCache
- [ ] Add unit tests for ImpersonationResolver
- [ ] Add integration tests for full discovery flow
- [ ] Consider proactive cache warming
- [ ] Consider metrics/monitoring implementation
