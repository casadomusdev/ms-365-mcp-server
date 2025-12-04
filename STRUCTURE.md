# Project Structure

## Comprehensive Project Overview

The **ms-365-mcp-server** is a Model Context Protocol (MCP) server implementation that provides comprehensive access to Microsoft 365 and Office services through the Microsoft Graph API. It enables AI assistants and MCP clients to interact with Outlook mail, Calendar, OneDrive, Teams, SharePoint, and other Microsoft 365 services.

### Key Features
- **MCP Protocol Implementation**: Full MCP server with stdio and HTTP transport support
- **Microsoft Graph API Integration**: Comprehensive coverage of Microsoft 365 services
- **Authentication Flexibility**: Device code flow, OAuth authorization code flow, and BYOT (Bring Your Own Token)
- **Organization Mode**: Support for both personal and work/school accounts
- **Security Controls**: Read-only mode, tool filtering, bearer token authentication
- **Docker Support**: Containerized deployment with STDIO wrapper for isolated execution
- **Token Management**: Secure credential storage via OS credential store with file fallback

### Architecture Principles
- **Stateless Design**: HTTP mode supports stateless bearer token authentication
- **Smart Tool Routing**: Consolidated tools with optional parameters (e.g., calendar tools accept optional calendarId)
- **Flexible Deployment**: Local development, Docker containers, or NPM package usage
- **Security First**: Optional read-only mode, granular tool filtering, secure token storage

## Complete Directory Structure

```
mcp_outlook/
├── README.md                           # Main documentation and getting started guide
├── STRUCTURE.md                        # This file - comprehensive project structure
├── QUICK_START.md                      # Quick start guide with setup scripts
├── SERVER_SETUP.md                     # Production deployment guide
├── AUTH.md                             # Authentication documentation
├── HEALTH_CHECK.md                     # Health check documentation
├── DRYRUN.md                           # Dry-run mode documentation
├── TOOLS.md                            # Available tools documentation
├── USER_IMPERSONATE.md                 # User impersonation documentation
├── testing.md                          # Testing documentation and strategy
├── functionality_testing.md            # Functionality testing scenarios
├── function_documentation.md           # Function documentation
├── howto.md                            # How-to guides
├── recent.md                           # Recent changes and updates
├── wikicontent.md                      # Wiki content
├── openwebuiSystemprompt.md           # OpenWebUI system prompt
├── LICENSE                             # MIT license
├── package.json                        # NPM package configuration
├── package-lock.json                   # NPM dependency lock file
├── tsconfig.json                       # TypeScript configuration
├── tsup.config.ts                      # Build configuration (tsup bundler)
├── eslint.config.js                    # ESLint configuration
├── vitest.config.js                    # Vitest test configuration
├── .prettierrc                         # Prettier code formatting configuration
├── .releaserc.json                     # Semantic release configuration
├── .gitignore                          # Git ignore patterns
├── .npmignore                          # NPM publish ignore patterns
├── .dockerignore                       # Docker build ignore patterns
├── .env.example                        # Environment variables template
├── .scripts-lib.sh                     # Shell script library functions
│
├── Dockerfile                          # Docker container definition
├── docker-compose.yaml                 # Docker Compose configuration
├── docker-mcp-wrapper.sh              # STDIO wrapper for Docker integration
│
├── start.sh                            # Start server script
├── stop.sh                             # Stop server script
├── health-check.sh                     # Health check script
│
├── auth-login.sh                       # Authentication: Login
├── auth-logout.sh                      # Authentication: Logout
├── auth-verify.sh                      # Authentication: Verify
├── auth-list-accounts.sh              # Authentication: List accounts
├── auth-list-mailboxes.sh             # Authentication: List mailboxes
├── auth-list-impersonated-mailboxes.sh # Authentication: List impersonated mailboxes
├── auth-export-tokens.sh              # Authentication: Export tokens
├── auth-import-tokens.sh              # Authentication: Import tokens
│
├── util-list-tools.sh                  # Utility: List available tools
│
├── mocks.json                          # Mock data for testing
├── remove-recursive-refs.js           # Utility to remove recursive references
├── test-calendar-fix.js               # Calendar fix test script
├── test-real-calendar.js              # Real calendar test script
│
├── bin/                                # Build and generation scripts
│   ├── generate-graph-client.mjs      # Generate Graph API client from OpenAPI spec
│   └── modules/                        # Modular generation components
│       ├── download-openapi.mjs       # Download Microsoft Graph OpenAPI specification
│       ├── extract-descriptions.mjs   # Extract tool descriptions from OpenAPI
│       ├── generate-mcp-tools.mjs     # Generate MCP tool definitions
│       └── simplified-openapi.mjs     # Generate simplified OpenAPI client
│
├── scripts/                            # Utility scripts
│   └── keychain-helper.js             # macOS Keychain integration helper
│
├── src/                                # Source code
│   ├── index.ts                       # Main entry point
│   ├── server.ts                      # MCP server implementation
│   ├── cli.ts                         # CLI argument parsing and handler
│   ├── version.ts                     # Version information
│   │
│   ├── auth.ts                        # Authentication core logic
│   ├── auth-tools.ts                  # Authentication MCP tools (login, logout, verify)
│   ├── oauth-provider.ts              # OAuth provider for HTTP mode
│   │
│   ├── graph-client.ts                # Microsoft Graph API client wrapper
│   ├── graph-tools.ts                 # Microsoft Graph MCP tools implementation
│   ├── endpoints.json                 # Graph API endpoint definitions
│   ├── tool-descriptions.ts           # Tool descriptions and metadata
│   ├── tool-blacklist.ts              # Tool filtering/blacklist logic
│   ├── toolDescriptions.ts            # Legacy tool descriptions
│   ├── toolDescriptions.js            # Legacy tool descriptions (JS)
│   │
│   ├── logger.ts                      # Centralized logging service
│   │
│   ├── lib/                           # Shared libraries
│   │   └── microsoft-auth.ts          # Microsoft authentication library integration
│   │
│   ├── impersonation/                 # User/mailbox impersonation
│   │   ├── index.ts                   # Impersonation exports
│   │   ├── ImpersonationContext.ts    # Impersonation context management
│   │   └── MailboxDiscoveryCache.ts   # Mailbox discovery caching
│   │
│   ├── mock/                          # Mock mode for testing
│   │   ├── index.ts                   # Mock exports
│   │   ├── defaults.ts                # Default mock responses
│   │   ├── loader.ts                  # Mock data loader
│   │   ├── MockResponse.ts            # Mock response handler
│   │   ├── pathMatcher.ts             # Path matching for mocks
│   │   └── registry.ts                # Mock registry
│   │
│   └── generated/                     # Auto-generated code
│       ├── README.md                  # Generated code documentation
│       ├── client.ts                  # Generated Graph API client (not in repo)
│       ├── endpoint-types.ts          # Generated endpoint type definitions
│       └── hack.ts                    # Temporary hacks/workarounds
│
└── test/                              # Test suite
    ├── auth-tools.test.ts             # Authentication tools tests
    ├── calendar-fix.test.js           # Calendar fix tests
    ├── cli.test.ts                    # CLI tests
    ├── graph-api.test.ts              # Graph API integration tests
    ├── read-only.test.ts              # Read-only mode tests
    ├── tool-filtering.test.ts         # Tool filtering tests
    ├── tool-wrappers.test.ts          # Tool wrapper tests
    ├── dryrun-mockmode.test.ts        # Dry-run mock mode tests
    ├── dryrun-partial.test.ts         # Dry-run partial tests
    ├── test-hack.ts                   # Test hack utilities
    └── fixtures/                      # Test fixtures
        └── dryrun-mocks.json          # Dry-run mock data
```

## Core Components

### 1. Server Layer (`src/server.ts`)
The MCP server implementation that handles:
- **Protocol Management**: JSON-RPC 2.0 over stdio or HTTP
- **Transport Layer**: Stdio transport (default) or HTTP transport with Express.js
- **Tool Registration**: Dynamic tool registration based on mode and filters
- **Resource Management**: MCP resource endpoints
- **OAuth Integration**: OAuth capabilities advertisement and endpoint handling (HTTP mode)

### 2. Authentication (`src/auth.ts`, `src/auth-tools.ts`, `src/oauth-provider.ts`)
Multi-method authentication system:
- **Device Code Flow**: Interactive browser-based authentication
- **OAuth Authorization Code Flow**: Standard OAuth 2.0 flow (HTTP mode)
- **Bearer Token Authentication**: Stateless token authentication with auto-refresh
- **BYOT**: Bring Your Own Token for external token management
- **Token Storage**: Secure OS credential store with encrypted file fallback
- **Smart Auto-Detection**: Automatically selects best available auth method

### 3. Graph API Client (`src/graph-client.ts`)
Microsoft Graph API integration:
- **Request Handling**: HTTP request construction and execution
- **Token Management**: Automatic token refresh and retry logic
- **Error Handling**: Graph API error parsing and reporting
- **Mock Mode**: Development/testing with mock responses
- **Dry-Run Mode**: Test mode without actual API calls

### 4. Tool System (`src/graph-tools.ts`, `src/tool-descriptions.ts`)
MCP tool implementations:
- **Dynamic Tool Generation**: Tools generated from OpenAPI specifications
- **Smart Routing**: Consolidated tools with optional parameters
- **Tool Filtering**: Regex-based tool enabling/disabling
- **Read-Only Mode**: Automatic filtering of write operations
- **Organization Mode**: Additional tools for work/school accounts

### 5. Impersonation System (`src/impersonation/`)
User and mailbox impersonation:
- **Context Management**: Track impersonation context across requests
- **Mailbox Discovery**: Automatic discovery and caching of accessible mailboxes
- **Shared Mailbox Access**: Tools for accessing shared mailboxes
- **Permission Validation**: Verify delegated permissions

### 6. Mock System (`src/mock/`)
Testing and development support:
- **Mock Registry**: Pre-defined mock responses for common endpoints
- **Path Matching**: Flexible pattern matching for mock responses
- **Custom Mocks**: Load custom mock data from JSON files
- **Dry-Run Mode**: Non-destructive testing mode

### 7. Logging System (`src/logger.ts`)
Centralized logging service:
- **Log Levels**: debug, info, warn, error
- **Output Formats**: Console (development) and JSON (production)
- **File Output**: Rotating log files with configurable retention
- **Correlation IDs**: Request tracking across components
- **Configuration**: Environment variable based configuration

### 8. Code Generation (`bin/`)
Automated client generation:
- **OpenAPI Download**: Fetch latest Microsoft Graph OpenAPI spec
- **Client Generation**: Generate TypeScript client from OpenAPI
- **Tool Generation**: Create MCP tool definitions from endpoints
- **Type Generation**: Generate TypeScript types for endpoints

## Data Flow

### MCP Request Flow (STDIO Mode)
```
MCP Client (e.g., Claude Desktop)
    ↓ (stdio JSON-RPC)
Server (src/server.ts)
    ↓
Tool Handler (src/graph-tools.ts)
    ↓
Graph Client (src/graph-client.ts)
    ↓ (HTTPS with Bearer token)
Microsoft Graph API
    ↓
Response parsing & formatting
    ↓
MCP Client receives result
```

### MCP Request Flow (HTTP Mode)
```
MCP Client (e.g., MCP Inspector)
    ↓ (HTTP POST /mcp with Bearer token)
Express Server (src/server.ts)
    ↓ (Bearer token extraction & validation)
Authentication Layer
    ↓
Tool Handler (src/graph-tools.ts)
    ↓
Graph Client (src/graph-client.ts)
    ↓ (HTTPS with Bearer token)
Microsoft Graph API
    ↓
Response parsing & formatting
    ↓
HTTP Response to MCP Client
```

### Authentication Flow (Device Code)
```
User initiates login
    ↓
Device code request to Azure AD
    ↓
Display URL + code to user
    ↓
User completes browser authentication
    ↓
Token polling succeeds
    ↓
Store tokens in OS credential store
    ↓
Ready for Graph API calls
```

### Authentication Flow (OAuth - HTTP Mode)
```
MCP Client discovers OAuth capabilities
    ↓
Client redirects user to /auth/authorize
    ↓
User authenticates with Microsoft
    ↓
Callback to /auth/token with code
    ↓
Exchange code for access token
    ↓
Client receives token
    ↓
Client includes token in MCP requests
```

### Bearer Token Authentication Flow
```
Client sends request with Authorization header
    ↓
Server extracts bearer token
    ↓
Graph Client uses bearer token for API calls
    ↓
If token expired & refresh token provided:
    ↓ (automatic refresh)
Microsoft Token Endpoint
    ↓
New access token returned
    ↓
API request retried with new token
```

## Configuration

### Environment Variables

**Authentication**
- `MS365_MCP_CLIENT_ID`: Azure app client ID (default: built-in app)
- `MS365_MCP_CLIENT_SECRET`: Azure app client secret (enables client credentials flow)
- `MS365_MCP_TENANT_ID`: Azure tenant ID (default: 'common')
- `MS365_MCP_OAUTH_TOKEN`: Pre-existing OAuth token (BYOT mode)

**Server Behavior**
- `MS365_MCP_ORG_MODE`: Enable organization/work mode (true/1)
- `MS365_MCP_READ_ONLY`: Enable read-only mode (true/1)
- `MS365_MCP_ENABLED_TOOLS`: Regex pattern for tool filtering

**Logging**
- `MS365_MCP_LOG_LEVEL`: Logging level (debug/info/warn/error, default: info)
- `MS365_MCP_LOG_FORMAT`: Log format (console/json, default: console)
- `MS365_MCP_LOG_DIR`: Log file directory (default: logs/)
- `MS365_MCP_LOG_RETENTION`: Log retention days (default: 14)
- `MS365_MCP_SILENT`: Disable console output (true/1)

### CLI Flags

**Authentication Commands**
- `--login`: Login using device code flow
- `--logout`: Logout and clear credentials
- `--verify-login`: Verify current login status

**Server Options**
- `--org-mode`, `--work-mode`: Enable organization/work features
- `--read-only`: Enable read-only mode
- `--http [port]`: Use HTTP transport (default port: 3000)
- `--enable-auth-tools`: Enable login/logout tools in HTTP mode
- `--enabled-tools <pattern>`: Filter tools by regex pattern
- `-v`: Enable verbose logging

## Testing Strategy

### Unit Tests
Located in `test/` directory, covering:
- **Authentication Tools**: Login/logout/verify functionality
- **Graph API Integration**: API call handling and error management
- **Tool Filtering**: Tool enabling/disabling logic
- **Read-Only Mode**: Write operation blocking
- **CLI**: Command-line argument parsing
- **Dry-Run Mode**: Mock mode testing

### Test Execution
```bash
# Run all tests
npm test

# Run specific test file
npm test test/auth-tools.test.ts

# Run with coverage
npm run test:coverage

# Verify all quality checks
npm run verify
```

### Manual Testing
See `testing.md` and `functionality_testing.md` for:
- Integration testing scenarios
- Manual test procedures
- OAuth flow testing
- Shared mailbox access testing
- Docker deployment testing

### Test Scripts
- `test-calendar-fix.js`: Calendar functionality testing
- `test-real-calendar.js`: Real calendar API testing
- `util-list-tools.sh`: List available tools for verification

## Development Workflow

### Setup
```bash
# Clone repository
git clone https://github.com/casadomusdev/ms-365-mcp-server.git
cd ms-365-mcp-server

# Install dependencies
npm install

# Generate client code
npm run generate

# Build project
npm run build
```

### Development Cycle
1. Make code changes in `src/`
2. Run `npm run build` to compile TypeScript
3. Test changes with `npm test`
4. Run verification: `npm run verify`
5. Commit changes following conventional commits

### Code Generation
When Microsoft Graph API changes:
```bash
npm run generate
```
This will:
1. Download latest OpenAPI spec
2. Generate TypeScript client
3. Generate MCP tool definitions
4. Update type definitions

### Docker Development
```bash
# Build and start container
./start.sh

# Verify authentication
./auth-login.sh
./auth-verify.sh

# Test functionality
./util-list-tools.sh

# Stop container
./stop.sh
```

### Release Process
The project uses semantic-release for automated versioning:
1. Commit following conventional commits format
2. Push to main branch
3. GitHub Actions automatically:
   - Runs tests
   - Builds project
   - Determines version bump
   - Creates GitHub release
   - Publishes to NPM

## Extension Points

### Adding New Tools
1. Update `src/endpoints.json` with new Graph API endpoints
2. Add tool descriptions in `src/tool-descriptions.ts`
3. Run `npm run generate` to regenerate client
4. Add tests in `test/`

### Custom Authentication
Implement custom auth provider:
1. Create new provider in `src/lib/`
2. Update `src/auth.ts` to support new method
3. Add configuration options
4. Document in `AUTH.md`

### Mock Responses
Add custom mocks:
1. Add entries to `mocks.json`
2. Update `src/mock/registry.ts`
3. Test with dry-run mode

### Tool Filtering
Create custom filters:
1. Update `src/tool-blacklist.ts`
2. Add environment variable configuration
3. Document in README.md

## Dependencies

### Runtime Dependencies
- `@modelcontextprotocol/sdk`: MCP protocol implementation
- `@microsoft/microsoft-graph-client`: Graph API client
- `@azure/msal-node`: Microsoft Authentication Library
- `express`: HTTP server (for HTTP mode)
- `keytar`: Secure credential storage
- `winston`: Logging framework

### Development Dependencies
- `typescript`: TypeScript compiler
- `tsup`: Build bundler
- `vitest`: Test framework
- `eslint`: Code linting
- `prettier`: Code formatting

## Security Considerations

### Token Storage
- Primary: OS credential store (macOS Keychain, Windows Credential Manager, Linux Secret Service)
- Fallback: Encrypted file storage in `~/.ms365-mcp/`

### Read-Only Mode
When enabled:
- Filters all write operations (create, update, delete, send)
- Allows only GET requests to Graph API
- Prevents accidental data modification

### Tool Filtering
- Regex-based tool enabling/disabling
- Reduces attack surface
- Allows granular permission control

### Bearer Token Authentication
- Stateless authentication support
- No server-side session storage
- Client-controlled token lifecycle
- Automatic token refresh with refresh token

## Architecture Decisions

### Why Smart Tool Routing?
Consolidating related operations (e.g., all calendar operations) into single tools with optional parameters reduces the MCP surface area while maintaining full functionality.

### Why Multiple Auth Methods?
Different deployment scenarios require different auth methods:
- **Device Code**: Best for CLI and desktop apps
- **OAuth**: Standard for web applications
- **Bearer Token**: Stateless API integration
- **BYOT**: Integration with existing auth systems

### Why OS Credential Store?
Provides secure, encrypted token storage without requiring users to manage encryption keys or passwords.

### Why Docker STDIO Wrapper?
Enables complete container isolation for MCP communication without exposing network ports, providing security and portability.

## Future Improvements

### Planned Features
- Additional Microsoft 365 services (Power Automate, Forms, etc.)
- Enhanced caching for frequently accessed data
- Webhook support for real-time notifications
- Batch operations for improved performance
- Enhanced error recovery and retry logic
- Multi-account support
- Custom scope configuration

### Performance Optimizations
- Request batching
- Response caching
- Connection pooling
- Lazy loading for tool definitions

### Security Enhancements
- Token rotation policies
- Enhanced audit logging
- Fine-grained permission models
- Rate limiting support

---

**Last Updated**: 2025-12-04
**Project Version**: Based on package.json version
**Maintainer**: Casadomus (fork of Softeria/ms-365-mcp-server)
