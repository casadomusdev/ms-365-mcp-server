#!/bin/bash
#
# MS-365 MCP Server - Verify Authentication Script
#
# Verifies that authentication is working and displays the authenticated user.
# Supports dual-mode operation: Docker or local Node.js
#
# Usage:
#   ./auth-verify.sh
#
# Exit codes:
#   0 - Authentication verified
#   1 - Authentication failed

set -e

# Get script directory and source shared library
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/.scripts-lib.sh"

echo -e "${YELLOW}MS-365 MCP Server - Verify Authentication${NC}"
echo "==========================================="
echo ""

# Detect execution mode and run appropriately
detect_execution_mode

if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    # Use docker compose run for one-off verification
    OUTPUT=$(docker compose run --rm ms365-mcp node dist/index.js --verify-login 2>&1)
else
    # Run directly
    OUTPUT=$(node dist/index.js --verify-login 2>&1)
fi

# Check if output contains valid JSON with success=true
if echo "$OUTPUT" | jq -e '.success == true and (.userData.displayName != null and .userData.displayName != "")' > /dev/null 2>&1; then
    
    # Extract user data
    DISPLAY_NAME=$(echo "$OUTPUT" | jq -r '.userData.displayName')
    USER_EMAIL=$(echo "$OUTPUT" | jq -r '.userData.userPrincipalName // .userData.mail // "N/A"')
    
    echo -e "${GREEN}✓ Authentication verified${NC}"
    echo ""
    echo "Authenticated as:"
    echo "  Name:  $DISPLAY_NAME"
    echo "  Email: $USER_EMAIL"
    echo ""
    exit 0
else
    echo -e "${RED}✗ Authentication failed${NC}"
    echo ""
    echo "You need to run ./auth-login.sh first to authenticate."
    echo ""
    echo "Debug output: $OUTPUT"
    exit 1
fi
