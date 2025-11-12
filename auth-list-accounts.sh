#!/bin/bash
#
# MS-365 MCP Server - List & Select Accounts Script
#
# Lists all authenticated accounts and allows interactive selection.
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

echo -e "${YELLOW}MS-365 MCP Server - List & Select Accounts${NC}"
echo "=============================================="
echo ""

# Detect execution mode and run appropriately
detect_execution_mode

if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    OUTPUT=$(docker compose run --rm ms365-mcp node dist/index.js --list-accounts 2>&1)
else
    OUTPUT=$(node dist/index.js --list-accounts 2>&1)
fi

# Parse the JSON output using here-string to avoid quote issues
if jq -e '.accounts' <<< "$OUTPUT" > /dev/null 2>&1; then
    ACCOUNT_COUNT=$(jq -r '.accounts | length' <<< "$OUTPUT")
    
    if [ "$ACCOUNT_COUNT" -eq 0 ]; then
        echo "No accounts found."
        echo ""
        echo "Run ./auth-login.sh to authenticate."
        echo ""
        exit 0
    fi
    
    echo "Found $ACCOUNT_COUNT account(s):"
    echo ""
    
    # Create arrays to store account data
    declare -a ACCOUNT_IDS
    declare -a ACCOUNT_NAMES
    declare -a ACCOUNT_EMAILS
    declare -a IS_SELECTED
    
    # Parse accounts into arrays
    INDEX=0
    while IFS= read -r account; do
        ACCOUNT_IDS[$INDEX]=$(jq -r '.username' <<< "$account")
        ACCOUNT_NAMES[$INDEX]=$(jq -r '.name // "N/A"' <<< "$account")
        ACCOUNT_EMAILS[$INDEX]=$(jq -r '.username' <<< "$account")
        IS_SELECTED[$INDEX]=$(jq -r '.selected // false' <<< "$account")
        ((INDEX++))
    done < <(jq -c '.accounts[]' <<< "$OUTPUT")
    
    # Display numbered list
    for i in "${!ACCOUNT_IDS[@]}"; do
        NUM=$((i + 1))
        MARKER="  "
        SUFFIX=""
        
        if [ "${IS_SELECTED[$i]}" = "true" ]; then
            MARKER="→ "
            SUFFIX=" ${GREEN}[SELECTED]${NC}"
        fi
        
        echo -e "${NUM}. ${MARKER}${ACCOUNT_NAMES[$i]} (${ACCOUNT_EMAILS[$i]})${SUFFIX}"
    done
    
    echo ""
    
    # If only one account, no need to select
    if [ "$ACCOUNT_COUNT" -eq 1 ]; then
        echo -e "${GREEN}✓ Only one account available${NC}"
        exit 0
    fi
    
    # Prompt for selection
    echo -e "${CYAN}Select an account to make it active:${NC}"
    read -p "Enter number (1-$ACCOUNT_COUNT) or press Enter to keep current: " SELECTION
    
    # If user pressed Enter without input, exit
    if [ -z "$SELECTION" ]; then
        echo ""
        echo "No changes made."
        exit 0
    fi
    
    # Validate selection
    if ! [[ "$SELECTION" =~ ^[0-9]+$ ]] || [ "$SELECTION" -lt 1 ] || [ "$SELECTION" -gt "$ACCOUNT_COUNT" ]; then
        echo -e "${RED}✗ Invalid selection${NC}"
        echo "Please enter a number between 1 and $ACCOUNT_COUNT"
        exit 1
    fi
    
    # Get the account ID (arrays are 0-indexed)
    ARRAY_INDEX=$((SELECTION - 1))
    SELECTED_ID="${ACCOUNT_IDS[$ARRAY_INDEX]}"
    SELECTED_NAME="${ACCOUNT_NAMES[$ARRAY_INDEX]}"
    
    echo ""
    echo "Selecting account: $SELECTED_NAME ($SELECTED_ID)"
    echo ""
    
    # Execute select-account command
    if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
        if docker compose run --rm ms365-mcp node dist/index.js --select-account "$SELECTED_ID" > /dev/null 2>&1; then
            echo -e "${GREEN}✓ Account selected successfully${NC}"
            echo ""
            echo "Active account is now: $SELECTED_NAME"
            echo ""
            exit 0
        else
            echo -e "${RED}✗ Failed to select account${NC}"
            exit 1
        fi
    else
        if node dist/index.js --select-account "$SELECTED_ID" > /dev/null 2>&1; then
            echo -e "${GREEN}✓ Account selected successfully${NC}"
            echo ""
            echo "Active account is now: $SELECTED_NAME"
            echo ""
            exit 0
        else
            echo -e "${RED}✗ Failed to select account${NC}"
            exit 1
        fi
    fi
else
    echo "Failed to parse account list"
    echo ""
    echo "Response: $OUTPUT"
    exit 1
fi
