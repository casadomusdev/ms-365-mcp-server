# PowerShell-Based Shared Mailbox Discovery - Implementation Tasks

## Phase 1: Environment Setup & Prerequisites
- [ ] Install PowerShell Core 7.x in development environment
  - [ ] Test PowerShell installation: `pwsh --version`
  - [ ] Verify pwsh is available in PATH
- [ ] Install Exchange Online PowerShell module
  - [ ] Run: `pwsh -Command "Install-Module -Name ExchangeOnlineManagement -Force"`
  - [ ] Test module: `pwsh -Command "Get-Module -ListAvailable ExchangeOnlineManagement"`
- [ ] Certificate setup for authentication
  - [ ] Generate self-signed certificate for development
  - [ ] Add certificate to Azure app registration
  - [ ] Store certificate securely (decide: file system or env variable)
  - [ ] Document certificate generation process
- [ ] Update .env.example with PowerShell environment variables
  - [ ] Add MS365_CERTIFICATE_PATH
  - [ ] Add MS365_CERTIFICATE_PASSWORD (optional)
  - [ ] Add MS365_POWERSHELL_ENABLED (feature flag)
  - [ ] Add MS365_POWERSHELL_TIMEOUT (default 30000ms)

## Phase 2: PowerShell Integration Layer
- [ ] Create src/lib/PowerShellService.ts
  - [ ] Implement PowerShell process spawning (child_process)
  - [ ] Add stdin/stdout communication handling
  - [ ] Implement JSON output parsing
  - [ ] Add timeout handling
  - [ ] Add error handling and logging
  - [ ] Create method: `execute(script: string, args: Record<string, any>)`
  - [ ] Create method: `checkPermissions(userEmail: string)`
  - [ ] Add JSDoc documentation
- [ ] Add PowerShell service tests
  - [ ] Test process spawning
  - [ ] Test JSON parsing
  - [ ] Test timeout handling
  - [ ] Test error scenarios

## Phase 3: Exchange PowerShell Scripts
- [ ] Create scripts/check-mailbox-permissions.ps1
  - [ ] Implement Exchange Online connection (certificate auth)
  - [ ] Add parameter handling from Node.js
  - [ ] Query shared mailboxes (Get-Mailbox -RecipientTypeDetails SharedMailbox)
  - [ ] Check Full Access permissions (Get-MailboxPermission)
  - [ ] Check SendAs permissions (Get-RecipientPermission)
  - [ ] Format output as JSON
  - [ ] Add error handling
  - [ ] Add logging/verbose output
  - [ ] Test script manually with pwsh
- [ ] Create helper PowerShell utilities
  - [ ] Connection validation script
  - [ ] Permission verification script
  - [ ] Troubleshooting/diagnostic script

## Phase 4: Refactor MailboxDiscoveryService
- [ ] Update src/impersonation/MailboxDiscoveryService.ts
  - [ ] Add PowerShellService dependency to constructor
  - [ ] Remove checkCalendarDelegation() method
  - [ ] Remove checkCalendarPermissions() method
  - [ ] Remove calendar delegation code from discoverMailboxes()
  - [ ] Add new discoverViaExchangePowerShell() method
  - [ ] Update discoverMailboxes() to use PowerShell
  - [ ] Keep personal mailbox detection via Graph API
  - [ ] Combine personal + shared mailbox results
  - [ ] Add feature flag check (MS365_POWERSHELL_ENABLED)
  - [ ] Add fallback error handling if PowerShell unavailable
  - [ ] Update JSDoc comments
  - [ ] Update logging messages
- [ ] Verify MailboxDiscoveryCache still works (no changes needed)
- [ ] Update type definitions if needed
  - [ ] Review MailboxInfo type
  - [ ] Add permission types if needed

## Phase 5: Docker Integration
- [ ] Update Dockerfile
  - [ ] Add PowerShell Core installation
  - [ ] Add Exchange Online PowerShell module installation
  - [ ] Add certificate handling (COPY or VOLUME)
  - [ ] Test Docker build
- [ ] Update docker-compose.yaml
  - [ ] Add volume mounts for certificates (if needed)
  - [ ] Add PowerShell environment variables
  - [ ] Test docker-compose up
- [ ] Update .dockerignore if needed
  - [ ] Exclude development certificates

## Phase 6: Documentation Updates
- [ ] Create POWERSHELL_SETUP.md
  - [ ] PowerShell Core installation guide
  - [ ] Exchange Online module installation
  - [ ] Certificate generation step-by-step
  - [ ] Azure app registration configuration
  - [ ] Troubleshooting common issues
- [ ] Update README.md
  - [ ] Add PowerShell requirements section
  - [ ] Link to POWERSHELL_SETUP.md
  - [ ] Update feature list
- [ ] Update SERVER_SETUP.md
  - [ ] Add PowerShell installation instructions
  - [ ] Add certificate configuration steps
- [ ] Update USER_IMPERSONATE.md
  - [ ] Update shared mailbox discovery explanation
  - [ ] Document PowerShell vs Graph API approaches
  - [ ] Update limitations section
  - [ ] Add performance notes
- [ ] Update .env.example
  - [ ] Add all new PowerShell environment variables
  - [ ] Add comments explaining each variable

## Phase 7: Code Cleanup
- [ ] Remove obsolete calendar delegation code
  - [ ] Review all files for calendar permission references
  - [ ] Remove unused imports
  - [ ] Remove unused type definitions
- [ ] Update error messages
  - [ ] Replace calendar delegation errors
  - [ ] Add PowerShell-specific error messages
- [ ] Clean up comments and documentation
  - [ ] Remove references to calendar delegation approach
  - [ ] Add PowerShell implementation notes

## Phase 8: Testing
- [ ] Unit tests
  - [ ] PowerShellService.execute() tests
  - [ ] PowerShellService.checkPermissions() tests
  - [ ] JSON parsing edge cases
  - [ ] Timeout scenarios
  - [ ] Error handling
- [ ] Integration tests
  - [ ] End-to-end mailbox discovery
  - [ ] Cache integration
  - [ ] PowerShell unavailable scenario
  - [ ] Certificate auth failures
- [ ] Manual testing scenarios
  - [ ] User with no shared mailbox access
  - [ ] User with Full Access only
  - [ ] User with SendAs only
  - [ ] User with both Full Access and SendAs
  - [ ] Admin user with elevated permissions
  - [ ] PowerShell disabled (feature flag off)
  - [ ] PowerShell timeout
  - [ ] Invalid certificate
  - [ ] Missing Exchange module
- [ ] Performance testing
  - [ ] Measure PowerShell execution time
  - [ ] Verify cache effectiveness
  - [ ] Test with multiple concurrent requests

## Phase 9: Deployment Preparation
- [ ] Update deployment documentation
  - [ ] Add PowerShell prerequisites
  - [ ] Certificate deployment guide
  - [ ] Environment variable configuration
- [ ] Create migration guide
  - [ ] Explain breaking changes
  - [ ] Provide rollback procedure
  - [ ] Document new requirements
- [ ] Security review
  - [ ] Certificate storage security
  - [ ] PowerShell script injection prevention
  - [ ] Audit logging considerations
- [ ] Performance optimization
  - [ ] Profile PowerShell execution
  - [ ] Optimize permission queries
  - [ ] Consider parallel processing

## Phase 10: Final Verification
- [ ] Code review
  - [ ] Review all new files
  - [ ] Check JSDoc completeness
  - [ ] Verify error messages clarity
  - [ ] Confirm no duplicate code
  - [ ] Validate type safety
- [ ] Documentation review
  - [ ] Verify all docs updated
  - [ ] Check for broken links
  - [ ] Ensure examples are accurate
  - [ ] Confirm troubleshooting steps
- [ ] Build and deploy
  - [ ] Run `npm run build`
  - [ ] Test Docker image build
  - [ ] Deploy to staging environment
  - [ ] Run smoke tests
- [ ] Git housekeeping
  - [ ] Commit with descriptive message
  - [ ] Tag release if applicable
  - [ ] Update CHANGELOG if exists

## Future Improvements (Post-Implementation)
- [ ] Performance enhancements
  - [ ] Background cache warming for active users
  - [ ] Parallel PowerShell execution
  - [ ] Incremental permission checking
  - [ ] Connection pooling for PowerShell sessions
- [ ] Monitoring and observability
  - [ ] PowerShell execution time metrics
  - [ ] Permission query success/failure rates
  - [ ] Cache hit/miss analytics
  - [ ] Alert on excessive failures
- [ ] Advanced features
  - [ ] Real-time permission change detection
  - [ ] Permission audit logging
  - [ ] Detailed permission type filtering
  - [ ] Support for user-to-user delegation (calendar)
- [ ] Alternative authentication
  - [ ] Azure Managed Identity support
  - [ ] Azure Key Vault integration
  - [ ] Automatic certificate rotation
- [ ] Testing enhancements
  - [ ] Add more unit test coverage
  - [ ] E2E test automation
  - [ ] Performance regression tests
  - [ ] Security scanning

## Notes
- PowerShell execution will be slow (~2-5 seconds) compared to Graph API
- Cache (TTL-based, 1 hour default) is critical for performance
- Feature flag (MS365_POWERSHELL_ENABLED) allows gradual rollout
- Certificate-based auth is more secure than client secret for PowerShell
- Keep Graph API for personal mailbox detection (faster, works well)
- PowerShell is ONLY used for shared mailbox permission detection
