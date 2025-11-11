#!/bin/bash
#
# MS-365 MCP Server - Verify Authentication Script
#
# Verifies that authentication is working and displays the authenticated user.
# This is useful to confirm tokens are valid without starting the full server.
#
# Usage:
#   ./auth-verify.sh
#
# Exit codes:
#   0 - Authentication verified
#   1 - Authentication failed

set -e

# Color codes
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${YELLOW}MS-365 MCP Server - Verify Authentication${NC}"
echo "==========================================="
echo ""

# Run the verify command
if OUTPUT=$(docker compose run --rm ms365-mcp node dist/index.js --verify-login 2>&1); then
    # Parse the JSON output
    if echo "$OUTPUT" | jq -e '.displayName' > /dev/null 2>&1; then
        DISPLAY_NAME=$(echo "$OUTPUT" | jq -r '.displayName')
        USER_EMAIL=$(echo "$OUTPUT" | jq -r '.userPrincipalName // .mail // "N/A"')
        
        echo -e "${GREEN}✓ Authentication verified${NC}"
        echo ""
        echo "Authenticated as:"
        echo "  Name:  $DISPLAY_NAME"
        echo "  Email: $USER_EMAIL"
        echo ""
        exit 0
    else
        echo -e "${RED}✗ Authentication failed - invalid response${NC}"
        echo ""
        echo "Response: $OUTPUT"
        exit 1
    fi
else
    echo -e "${RED}✗ Authentication failed${NC}"
    echo ""
    echo "You need to run ./auth-login.sh first to authenticate."
    echo ""
    exit 1
fi
