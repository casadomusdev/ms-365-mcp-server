# Future Enhancements

This document outlines potential improvements and enhancements for the MS-365 MCP Server. The core PowerShell-based shared mailbox discovery implementation is complete and tested.

## Performance Enhancements

- [ ] Background cache warming for active users
- [ ] Parallel PowerShell execution for faster discovery
- [ ] Incremental permission checking
- [ ] Connection pooling for PowerShell sessions

## Monitoring and Observability

- [ ] PowerShell execution time metrics
- [ ] Permission query success/failure rates
- [ ] Cache hit/miss analytics
- [ ] Alert on excessive failures

## Advanced Features

- [ ] Real-time permission change detection
- [ ] Permission audit logging
- [ ] Detailed permission type filtering
- [ ] Support for user-to-user delegation (calendar)

## Alternative Authentication

- [ ] Azure Managed Identity support
- [ ] Azure Key Vault integration
- [ ] Automatic certificate rotation

## Testing Enhancements

- [ ] Add more unit test coverage
- [ ] E2E test automation
- [ ] Performance regression tests
- [ ] Security scanning
