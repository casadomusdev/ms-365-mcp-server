# Health Check Documentation

## Overview

The `health-check.sh` script provides a simple way to verify that the MS-365 MCP server is properly configured, authenticated, and can communicate with Microsoft Graph API.

## Requirements

- `jq` - JSON processor (used for parsing API responses)
- Authenticated MS-365 account (via `--login` command)
- Built server application (`npm run build` must have been run)

## Usage

### Local Testing (Host Machine)

If you have the project built locally:

```bash
cd projects/ms-365-mcp-server
./health-check.sh
```

### Docker Container Testing

To run the health check inside the Docker container:

```bash
# Using docker-compose exec (container must be running)
docker-compose exec ms365-mcp /app/health-check.sh

# Or using docker-compose run (one-off container)
docker-compose run --rm ms365-mcp /app/health-check.sh
```

### As a Docker Healthcheck

You can integrate this into your docker-compose.yaml for automatic health monitoring:

```yaml
services:
  ms365-mcp:
    # ... other configuration ...
    
    healthcheck:
      test: ["CMD", "/app/health-check.sh"]
      interval: 5m
      timeout: 30s
      retries: 3
      start_period: 30s
```

## Exit Codes

The script returns the following exit codes:

- `0` - Health check **passed** - Authentication verified and Graph API connection successful
- `1` - Health check **failed** - Authentication or connection issues
- `2` - Container not running (only when using docker-compose)

## Output Examples

### Successful Health Check

```
MS-365 MCP Server Health Check
===============================

ℹ Verifying authentication and Graph API connection...

✓ Authentication verified
✓ Microsoft Graph API connection successful

Authenticated as:
  Name:  John Doe
  Email: john.doe@company.com
```

### Failed Health Check

```
MS-365 MCP Server Health Check
===============================

ℹ Verifying authentication and Graph API connection...

✗ Health check failed

Error details:
[Error message details here]
```

## Authentication Setup

Before running the health check, you need to authenticate with Microsoft 365. Use the authentication helper scripts:

```bash
# Initial authentication (one-time setup)
./auth-login.sh

# Verify authentication worked
./auth-verify.sh
```

See [AUTH.md](AUTH.md) for complete authentication management documentation.

## Troubleshooting

### "Authentication failed"

The health check requires valid authentication. If you see authentication errors:

```bash
# Verify authentication status
./auth-verify.sh

# If not authenticated, run the login script
./auth-login.sh
```

### "command not found: jq"

Install jq:

```bash
# macOS
brew install jq

# Debian/Ubuntu
apt-get install jq

# Alpine Linux
apk add jq
```

The Docker image already includes `jq`, so this only affects local testing.

### "Authentication failed"

Run the authentication flow:

```bash
# Using docker-compose
docker-compose run --rm ms365-mcp node dist/index.js --login

# Or locally
node dist/index.js --login
```

### "dist/index.js not found"

Build the application first:

```bash
npm run build
```

## Integration with Monitoring Tools

### Prometheus/Grafana

You can wrap this script in a monitoring exporter:

```bash
#!/bin/bash
if ./health-check.sh > /dev/null 2>&1; then
  echo "ms365_mcp_health 1"
else
  echo "ms365_mcp_health 0"
fi
```

### Cron-based Monitoring

Schedule periodic health checks:

```cron
*/15 * * * * cd /path/to/ms-365-mcp-server && ./health-check.sh || echo "Health check failed" | mail -s "MS-365 MCP Alert" admin@company.com
```

### Kubernetes Liveness Probe

```yaml
livenessProbe:
  exec:
    command:
      - /app/health-check.sh
  initialDelaySeconds: 30
  periodSeconds: 300
  timeoutSeconds: 30
  failureThreshold: 3
```

## What It Checks

The health check script:

1. ✅ Verifies the server application is built and accessible
2. ✅ Confirms authentication tokens are valid
3. ✅ Tests connectivity to Microsoft Graph API
4. ✅ Retrieves current user information as proof of working authentication

## Security Note

The health check only displays the authenticated user's name and email. It does not expose sensitive information like tokens or credentials.
