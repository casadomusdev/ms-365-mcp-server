#!/bin/bash
#
# MS-365 MCP Server Health Check Script
#
# This script verifies that the MCP server is properly configured and can
# connect to Microsoft Graph API.
#
# Dual-mode operation:
#   - Detects if running inside Docker container or on host
#   - Automatically uses correct execution method
#
# Usage:
#   ./health-check.sh                    # Auto-detects environment
#
# Exit codes:
#   0 - Health check passed
#   1 - Health check failed
#   2 - Container not running (host mode only)

set -e

# Detect if running inside Docker container
if [ -f /.dockerenv ] || grep -q docker /proc/1/cgroup 2>/dev/null; then
    INSIDE_DOCKER=true
else
    INSIDE_DOCKER=false
fi

# If running on host, delegate to container
if [ "$INSIDE_DOCKER" = false ]; then
    # Load COMPOSE_PROJECT_NAME from .env
    if [ -f .env ]; then
        export $(grep -v '^#' .env | grep COMPOSE_PROJECT_NAME | xargs)
    fi
    
    # Use compose project name or default
    COMPOSE_PROJECT=${COMPOSE_PROJECT_NAME:-ms365-mcp}
    
    # Execute health check inside container
    exec docker compose exec ms365-mcp /app/health-check.sh
    exit $?
fi

# Color codes for output (only used when running inside container)
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print colored output
print_status() {
    local status=$1
    local message=$2
    
    case $status in
        "ok")
            echo -e "${GREEN}✓${NC} $message"
            ;;
        "error")
            echo -e "${RED}✗${NC} $message"
            ;;
        "info")
            echo -e "${YELLOW}ℹ${NC} $message"
            ;;
    esac
}

# Main health check
echo "MS-365 MCP Server Health Check"
echo "==============================="
echo ""

# Run the verify-login command
print_status "info" "Verifying authentication and Graph API connection..."
echo ""

if OUTPUT=$(node dist/index.js --verify-login 2>&1); then
    # Parse the JSON output
    if echo "$OUTPUT" | jq -e '.displayName' > /dev/null 2>&1; then
        DISPLAY_NAME=$(echo "$OUTPUT" | jq -r '.displayName')
        USER_EMAIL=$(echo "$OUTPUT" | jq -r '.userPrincipalName // .mail // "N/A"')
        
        echo ""
        print_status "ok" "Authentication verified"
        print_status "ok" "Microsoft Graph API connection successful"
        echo ""
        echo "Authenticated as:"
        echo "  Name:  $DISPLAY_NAME"
        echo "  Email: $USER_EMAIL"
        echo ""
        exit 0
    else
        print_status "error" "Authentication failed - invalid response"
        echo ""
        echo "Response: $OUTPUT"
        exit 1
    fi
else
    print_status "error" "Health check failed"
    echo ""
    echo "Error details:"
    echo "$OUTPUT"
    echo ""
    exit 1
fi
