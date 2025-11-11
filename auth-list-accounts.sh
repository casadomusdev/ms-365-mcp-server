#!/bin/bash
#
# MS-365 MCP Server - List Accounts Script
#
# Lists all authenticated accounts cached in the token store.
# Supports dual-mode operation: Docker or local Node.js
#
# Usage:
#   ./auth-list-accounts.sh
#
# Exit codes:
#   0 - Success
#   1 - Failed

set -e

# Get script directory and source shared library
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/.scripts-lib.sh"

echo -e "${YELLOW}MS-365 MCP Server - List Accounts${NC}"
echo "===================================="
echo ""

# Detect execution mode and run appropriately
detect_execution_mode

if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    OUTPUT=$(docker compose run --rm ms365-mcp node dist/index.js --list-accounts 2>&1)
else
    OUTPUT=$(node dist/index.js --list-accounts 2>&1)
fi

# Process the output
if [ $? -eq 0 ]; then
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
