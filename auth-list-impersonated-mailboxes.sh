#!/bin/bash
#
# MS-365 MCP Server - List Impersonated User Mailboxes Script
#
# Lists all accessible mailboxes (personal, delegated, shared) for the user
# specified in MS365_MCP_IMPERSONATE_USER environment variable.
# Supports dual-mode operation: Docker or local Node.js
#
# Usage:
#   ./auth-list-impersonated-mailboxes.sh
#
# Exit codes:
#   0 - Success
#   1 - Failed

set -e

# Get script directory and source shared library
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/.scripts-lib.sh"

echo -e "${YELLOW}MS-365 MCP Server - List Impersonated User Mailboxes${NC}"
echo "=========================================================="
echo ""

# Detect execution mode and run appropriately
detect_execution_mode

if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    OUTPUT=$(docker compose run --rm ms365-mcp node dist/index.js --list-impersonated-mailboxes 2>&1)
else
    OUTPUT=$(node dist/index.js --list-impersonated-mailboxes 2>&1)
fi

# Parse the JSON output
if jq -e '.success == true' <<< "$OUTPUT" > /dev/null 2>&1; then
    USER_EMAIL=$(jq -r '.userEmail' <<< "$OUTPUT")
    MAILBOX_COUNT=$(jq -r '.mailboxes | length' <<< "$OUTPUT")
    
    echo "Impersonated User: $USER_EMAIL"
    echo ""
    
    if [ "$MAILBOX_COUNT" -eq 0 ]; then
        echo "No mailboxes found."
        echo ""
        echo "Make sure MS365_MCP_IMPERSONATE_USER is correctly configured."
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
        ((INDEX++))
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
    
    # Check if there's a note
    if jq -e '.note' <<< "$OUTPUT" > /dev/null 2>&1; then
        NOTE=$(jq -r '.note' <<< "$OUTPUT")
        echo ""
        echo -e "${YELLOW}Note:${NC} $NOTE"
    fi
    
    echo ""
    echo -e "${GREEN}✓ Mailbox listing complete${NC}"
    echo ""
    exit 0
else
    echo -e "${RED}✗ Failed to list impersonated mailboxes${NC}"
    echo ""
    
    # Check if there's an error message in the JSON
    if jq -e '.error' <<< "$OUTPUT" > /dev/null 2>&1; then
        ERROR_MSG=$(jq -r '.error' <<< "$OUTPUT")
        echo "Error: $ERROR_MSG"
    else
        echo "Response: $OUTPUT"
    fi
    
    echo ""
    echo "Make sure:"
    echo "  1. You are authenticated with ./auth-login.sh"
    echo "  2. MS365_MCP_IMPERSONATE_USER is set in your .env file"
    echo "  3. You have the necessary permissions (User.ReadBasic.All)"
    echo ""
    exit 1
fi

