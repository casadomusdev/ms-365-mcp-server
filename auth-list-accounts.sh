#!/bin/bash
#
# MS-365 MCP Server - List Accounts Script
#
# Lists all authenticated accounts cached in the token store.
# Useful for multi-account setups to see which accounts are available.
#
# Usage:
#   ./auth-list-accounts.sh
#
# Exit codes:
#   0 - Success
#   1 - Failed

set -e

# Color codes
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

echo -e "${YELLOW}MS-365 MCP Server - List Accounts${NC}"
echo "===================================="
echo ""

# Run the list-accounts command
if OUTPUT=$(docker compose run --rm ms365-mcp node dist/index.js --list-accounts 2>&1); then
    # Parse the JSON output
    if echo "$OUTPUT" | jq -e '.accounts' > /dev/null 2>&1; then
        ACCOUNT_COUNT=$(echo "$OUTPUT" | jq -r '.accounts | length')
        
        if [ "$ACCOUNT_COUNT" -eq 0 ]; then
            echo "No accounts found."
            echo ""
            echo "Run ./auth-login.sh to authenticate."
            echo ""
            exit 0
        fi
        
        echo "Found $ACCOUNT_COUNT account(s):"
        echo ""
        
        # List each account
        echo "$OUTPUT" | jq -r '.accounts[] | 
            "  " + 
            (if .selected then "→ " else "  " end) +
            .username + 
            (if .name then " (" + .name + ")" else "" end) +
            (if .selected then " [SELECTED]" else "" end)'
        
        echo ""
        echo "Legend:"
        echo "  → = Currently selected account"
        echo ""
        echo "To select a different account:"
        echo "  docker compose run --rm ms365-mcp node dist/index.js --select-account <account-id>"
        echo ""
        exit 0
    else
        echo "Failed to parse account list"
        echo ""
        echo "Response: $OUTPUT"
        exit 1
    fi
else
    echo "Failed to list accounts"
    echo ""
    exit 1
fi
