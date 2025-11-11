#!/bin/bash
#
# MS-365 MCP Server - Logout Script
#
# Logs out and clears all cached credentials from the token cache.
# Use this to remove authentication completely.
#
# Usage:
#   ./auth-logout.sh
#
# Exit codes:
#   0 - Logout successful
#   1 - Logout failed

set -e

# Color codes
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

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

# Run the logout command
if docker compose run --rm ms365-mcp node dist/index.js --logout; then
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
