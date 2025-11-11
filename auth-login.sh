#!/bin/bash
#
# MS-365 MCP Server - Login Script
#
# Initiates the device code authentication flow.
# This is a one-time process that caches tokens for automatic refresh.
#
# Usage:
#   ./auth-login.sh [--force-file-cache]
#
# Options:
#   --force-file-cache    Force tokens to be saved to files instead of system keychain
#                         (Useful when you need to export tokens for transfer)
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

# Parse arguments
FORCE_FILE_CACHE=false
INTERACTIVE=true

for arg in "$@"; do
    case $arg in
        --force-file-cache)
            FORCE_FILE_CACHE=true
            INTERACTIVE=false
            shift
            ;;
        *)
            echo "Unknown option: $arg"
            echo "Usage: ./auth-login.sh [--force-file-cache]"
            exit 1
            ;;
    esac
done

# If no arguments provided, ask user interactively
if [ "$INTERACTIVE" = true ]; then
    echo "Token Storage Options:"
    echo "  1. System Keychain (default, secure)"
    echo "  2. File-based cache (allows token export for server transfer)"
    echo ""
    read -p "Choose storage method (1 or 2) [1]: " choice
    choice=${choice:-1}
    
    if [ "$choice" = "2" ]; then
        FORCE_FILE_CACHE=true
        echo ""
        echo "✓ File-based cache selected"
        echo "  You'll be able to export tokens with ./auth-export-tokens.sh"
        echo ""
    elif [ "$choice" = "1" ]; then
        echo ""
        echo "✓ System keychain selected (default)"
        echo ""
    else
        echo "Invalid choice. Using default (system keychain)."
        echo ""
    fi
fi

# Color codes
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

echo -e "${YELLOW}╔════════════════════════════════════════════════════════════════╗${NC}"
echo -e "${YELLOW}║         MS-365 MCP Server - Device Code Login                  ║${NC}"
echo -e "${YELLOW}╔════════════════════════════════════════════════════════════════╗${NC}"
echo ""

if [ "$FORCE_FILE_CACHE" = true ]; then
    echo -e "${CYAN}ℹ File-based cache mode enabled${NC}"
    echo "  Tokens will be saved to files instead of system keychain"
    echo "  This allows you to export tokens for transfer to other machines"
    echo ""
fi

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

# Build docker compose command with optional environment variable
DOCKER_CMD="docker compose run --rm"
if [ "$FORCE_FILE_CACHE" = true ]; then
    DOCKER_CMD="$DOCKER_CMD -e FORCE_FILE_CACHE=true"
fi
DOCKER_CMD="$DOCKER_CMD ms365-mcp node dist/index.js --login"

# Run the login command
if eval "$DOCKER_CMD"; then
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
