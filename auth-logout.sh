#!/bin/bash
#
# MS-365 MCP Server - Logout Script
#
# Logs out and clears all cached credentials from the token cache.
# Supports dual-mode operation: Docker or local Node.js
#
# Usage:
#   ./auth-logout.sh
#
# Exit codes:
#   0 - Logout successful
#   1 - Logout failed

set -e

# Get script directory and source shared library
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/.scripts-lib.sh"

echo -e "${YELLOW}MS-365 MCP Server - Logout${NC}"
echo "============================"
echo ""
echo -e "${YELLOW}WARNING: This will remove all cached credentials.${NC}"
echo "You will need to re-authenticate with ./auth-login.sh"
echo ""
read -p "Are you sure you want to logout? (y/N): " -n 1 -r
echo ""

if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "Logout cancelled"
    exit 0
fi

echo ""
echo "Logging out..."
echo ""

# Detect execution mode
detect_execution_mode

# Build and run command based on mode
if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    # Docker mode - use docker compose run
    DOCKER_CMD="docker compose run --rm ms365-mcp node dist/index.js --logout"
    
    if eval "$DOCKER_CMD"; then
        LOGOUT_SUCCESS=true
    else
        LOGOUT_SUCCESS=false
    fi
else
    # Direct mode - run Node.js locally
    if node dist/index.js --logout; then
        LOGOUT_SUCCESS=true
    else
        LOGOUT_SUCCESS=false
    fi
fi

# Check result
if [ "$LOGOUT_SUCCESS" = true ]; then
    echo ""
    echo -e "${GREEN}✓ Logged out successfully${NC}"
    echo ""
    echo "All credentials have been cleared."
    echo "Run ./auth-login.sh to authenticate again."
    echo ""
    exit 0
else
    echo ""
    echo -e "${RED}✗ Logout failed${NC}"
    echo ""
    exit 1
fi
