#!/bin/bash
#
# MS-365 MCP Server Health Check Script
#
# This script verifies that the MCP server is properly configured and can
# connect to Microsoft Graph API.
#
# Usage:
#   ./health-check.sh                    # Run health check locally
#   docker-compose exec ms365-mcp ./health-check.sh   # Run in container
#
# Exit codes:
#   0 - Health check passed
#   1 - Health check failed
#   2 - Container not running (docker-compose only)

set -e

# Color codes for output
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
