#!/bin/bash
#
# MS-365 MCP Server - List Mailboxes Script
#
# Lists all accessible mailboxes (personal, delegated, shared).
# Supports dual-mode operation: Docker or local Node.js
#
# Usage:
#   ./auth-list-mailboxes.sh
#
# Exit codes:
#   0 - Success
#   1 - Failed

set -e

# Get script directory and source shared library
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/.scripts-lib.sh"

echo -e "${YELLOW}MS-365 MCP Server - List Accessible Mailboxes${NC}"
echo "=============================================="
echo ""

# Detect execution mode and run appropriately
detect_execution_mode

if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    OUTPUT=$(docker compose run --rm ms365-mcp node dist/index.js --list-mailboxes 2>&1)
else
    OUTPUT=$(node dist/index.js --list-mailboxes 2>&1)
fi

# Parse the JSON output
if jq -e '.success == true' <<< "$OUTPUT" > /dev/null 2>&1; then
    MAILBOX_COUNT=$(jq -r '.mailboxes | length' <<< "$OUTPUT")
    
    if [ "$MAILBOX_COUNT" -eq 0 ]; then
        echo "No mailboxes found."
        echo ""
        echo "Run ./auth-login.sh to authenticate."
        echo ""
        exit 0
    fi
    
    echo "Found $MAILBOX_COUNT mailbox(es):"
    echo ""
    
    # Create arrays to store mailbox data
    declare -a MAILBOX_TYPES
    declare -a MAILBOX_NAMES
    declare -a MAILBOX_EMAILS
    declare -a IS_PRIMARY
    
    # Parse mailboxes into arrays
    INDEX=0
    while IFS= read -r mailbox; do
        MAILBOX_TYPES[$INDEX]=$(jq -r '.type' <<< "$mailbox")
        MAILBOX_NAMES[$INDEX]=$(jq -r '.displayName // "N/A"' <<< "$mailbox")
        MAILBOX_EMAILS[$INDEX]=$(jq -r '.email // "N/A"' <<< "$mailbox")
        IS_PRIMARY[$INDEX]=$(jq -r '.isPrimary // false' <<< "$mailbox")
        INDEX=$((INDEX + 1))
    done < <(jq -c '.mailboxes[]' <<< "$OUTPUT")
    
    # Display numbered list with type badges
    for i in "${!MAILBOX_TYPES[@]}"; do
        NUM=$((i + 1))
        TYPE="${MAILBOX_TYPES[$i]}"
        
        # Format type with color
        case "$TYPE" in
            personal)
                TYPE_BADGE="${GREEN}[PERSONAL]${NC}"
                ;;
            delegated)
                TYPE_BADGE="${CYAN}[DELEGATED]${NC}"
                ;;
            shared)
                TYPE_BADGE="${YELLOW}[SHARED]${NC}"
                ;;
            *)
                TYPE_BADGE="[${TYPE^^}]"
                ;;
        esac
        
        PRIMARY_MARKER=""
        if [ "${IS_PRIMARY[$i]}" = "true" ]; then
            PRIMARY_MARKER=" ${GREEN}★${NC}"
        fi
        
        echo -e "${NUM}. ${TYPE_BADGE} ${MAILBOX_NAMES[$i]} (${MAILBOX_EMAILS[$i]})${PRIMARY_MARKER}"
    done
    
    echo ""
    echo -e "${GREEN}✓ Mailbox listing complete${NC}"
    echo ""
    exit 0
else
    echo -e "${RED}✗ Failed to list mailboxes${NC}"
    echo ""
    
    # Check if there's an error message in the JSON
    if jq -e '.error' <<< "$OUTPUT" > /dev/null 2>&1; then
        ERROR_MSG=$(jq -r '.error' <<< "$OUTPUT")
        echo "Error: $ERROR_MSG"
    else
        echo "Response: $OUTPUT"
    fi
    
    echo ""
    echo "Make sure you are authenticated with ./auth-login.sh"
    echo ""
    exit 1
fi
