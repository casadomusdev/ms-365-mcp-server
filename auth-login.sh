#!/bin/bash
#
# MS-365 MCP Server - Login Script
#
# Initiates the device code authentication flow.
# This is a one-time process that caches tokens for automatic refresh.
#
# Usage:
#   ./auth-login.sh
#
# What to expect:
#   1. A URL will be displayed (e.g., https://microsoft.com/devicelogin)
#   2. A device code will be shown (e.g., AB12-CD34)
#   3. Visit the URL on any device and enter the code
#   4. Sign in with your Microsoft 365 account
#   5. Grant the requested permissions
#   6. Tokens will be cached in the Docker volume
#
# Exit codes:
#   0 - Login successful
#   1 - Login failed

set -e

# Color codes
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${YELLOW}╔════════════════════════════════════════════════════════════════╗${NC}"
echo -e "${YELLOW}║         MS-365 MCP Server - Device Code Login                  ║${NC}"
echo -e "${YELLOW}╔════════════════════════════════════════════════════════════════╗${NC}"
echo ""
echo "Starting device code authentication flow..."
echo ""
echo "Instructions:"
echo "  1. A URL and code will appear below"
echo "  2. Visit the URL on ANY device (laptop, phone, etc.)"
echo "  3. Enter the code when prompted"
echo "  4. Sign in with your Microsoft 365 account"
echo "  5. Review and approve the permissions"
echo ""
echo -e "${YELLOW}════════════════════════════════════════════════════════════════${NC}"
echo ""

# Run the login command
if docker compose run --rm ms365-mcp node dist/index.js --login; then
    echo ""
    echo -e "${GREEN}✓ Login successful!${NC}"
    echo ""
    echo "Tokens have been cached. You can now:"
    echo "  - Run ./auth-verify.sh to verify authentication"
    echo "  - Start the server with: docker compose up -d"
    echo ""
    exit 0
else
    echo ""
    echo "✗ Login failed"
    echo ""
    exit 1
fi
