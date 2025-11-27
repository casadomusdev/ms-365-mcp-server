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

# Get script directory and source shared library
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/.scripts-lib.sh"

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

echo -e "${YELLOW}╔════════════════════════════════════════════════════════════════╗${NC}"
echo -e "${YELLOW}║         MS-365 MCP Server - Authentication Setup               ║${NC}"
echo -e "${YELLOW}╔════════════════════════════════════════════════════════════════╗${NC}"
echo ""

# Check if client secret is configured (client credentials mode)
if [ -f .env ]; then
    CLIENT_SECRET=$(grep "^MS365_MCP_CLIENT_SECRET=" .env | cut -d '=' -f2)
fi

if [ -n "$CLIENT_SECRET" ]; then
    echo -e "${CYAN}ℹ Client Credentials Mode Detected${NC}"
    echo "  MS365_MCP_CLIENT_SECRET is configured"
    echo "  No interactive login required - using app permissions"
    echo ""
    
    # Check if PowerShell certificate is configured
    if [ -f .env ]; then
        CERT_PATH=$(grep "^MS365_CERT_PATH=" .env | cut -d '=' -f2)
    fi
    
    # Check if certificate exists
    if [ -z "$CERT_PATH" ] || [ ! -f "$CERT_PATH" ]; then
        echo -e "${YELLOW}⚠ PowerShell Certificate Not Found${NC}"
        echo ""
        echo "For shared mailbox discovery, you need a PowerShell certificate."
        echo "This is required for Exchange Online app-only authentication."
        echo ""
        read -p "Generate PowerShell certificate now? (Y/n): " -n 1 -r
        echo
        
        if [[ ! $REPLY =~ ^[Nn]$ ]]; then
            echo ""
            if "$SCRIPT_DIR/auth-generate-cert.sh"; then
                echo ""
                echo -e "${GREEN}✓ Certificate generated successfully${NC}"
                echo ""
            else
                echo ""
                echo -e "${YELLOW}⚠ Certificate generation failed or was skipped${NC}"
                echo "  You can generate it later with: ./auth-generate-cert.sh"
                echo ""
            fi
        else
            echo ""
            echo "Skipping certificate generation."
            echo "Run './auth-generate-cert.sh' later to enable PowerShell features."
            echo ""
        fi
    else
        echo -e "${GREEN}✓ PowerShell certificate found${NC}"
        echo ""
    fi
    
    echo "Verifying connectivity with Microsoft Graph API..."
    
    # In client credentials mode, just verify that the setup works
    exec "$SCRIPT_DIR/auth-verify.sh"
fi

# Device code flow mode
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

# Detect execution mode
detect_execution_mode

# Build and run command based on mode
if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    # Docker mode - use docker compose run
    DOCKER_CMD="docker compose run --rm"
    if [ "$FORCE_FILE_CACHE" = true ]; then
        DOCKER_CMD="$DOCKER_CMD -e MS365_MCP_FORCE_FILE_CACHE=true"
    fi
    DOCKER_CMD="$DOCKER_CMD ms365-mcp node dist/index.js --login"
    
    if eval "$DOCKER_CMD"; then
        LOGIN_SUCCESS=true
    else
        LOGIN_SUCCESS=false
    fi
else
    # Direct mode - run Node.js locally
    if [ "$FORCE_FILE_CACHE" = true ]; then
        export MS365_MCP_FORCE_FILE_CACHE=true
    fi
    
    if node dist/index.js --login; then
        LOGIN_SUCCESS=true
    else
        LOGIN_SUCCESS=false
    fi
fi

# Check result
if [ "$LOGIN_SUCCESS" = true ]; then
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
