#!/bin/bash
#
# MS-365 MCP Server - Import Tokens Script
#
# Imports authentication tokens from a compressed archive.
# Supports dual-mode operation: Docker or local Node.js
# Supports importing to both keychain and file-based storage
#
# Usage:
#   ./auth-import-tokens.sh <archive-file> [options]
#
# Arguments:
#   archive-file  Path to tokens-*.tar.gz file to import
#
# Options:
#   --to-keychain  Import tokens to system keychain (default for local mode)
#   --to-file      Import tokens to file-based storage
#
# Exit codes:
#   0 - Import successful
#   1 - Import failed
#   2 - Invalid archive file

set -e

# Get script directory and source shared library
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/.scripts-lib.sh"

echo -e "${CYAN}MS-365 MCP Server - Import Tokens${NC}"
echo "===================================="
echo ""

# Parse arguments
ARCHIVE_FILE=""
IMPORT_MODE=""

while [[ $# -gt 0 ]]; do
    case $1 in
        --to-keychain)
            IMPORT_MODE="keychain"
            shift
            ;;
        --to-file)
            IMPORT_MODE="file"
            shift
            ;;
        *)
            if [ -z "$ARCHIVE_FILE" ]; then
                ARCHIVE_FILE="$1"
            else
                echo -e "${RED}✗ Unknown option: $1${NC}"
                echo ""
                echo "Usage: ./auth-import-tokens.sh <archive-file> [--to-keychain|--to-file]"
                exit 2
            fi
            shift
            ;;
    esac
done

# Check if archive file was provided
if [ -z "$ARCHIVE_FILE" ]; then
    echo -e "${RED}✗ No archive file specified${NC}"
    echo ""
    echo "Usage: ./auth-import-tokens.sh <archive-file> [options]"
    echo ""
    echo "Options:"
    echo "  --to-keychain  Import tokens to system keychain (default for local)"
    echo "  --to-file      Import tokens to file-based storage"
    echo ""
    echo "Example:"
    echo "  ./auth-import-tokens.sh tokens-20231201-143022.tar.gz --to-keychain"
    echo ""
    exit 2
fi

# Check if archive file exists
if [ ! -f "$ARCHIVE_FILE" ]; then
    echo -e "${RED}✗ Archive file not found: $ARCHIVE_FILE${NC}"
    echo ""
    exit 2
fi

# Check file extension
if [[ ! "$ARCHIVE_FILE" =~ \.tar\.gz$ ]]; then
    echo -e "${YELLOW}⚠  Warning: File doesn't have .tar.gz extension${NC}"
    echo "  Attempting to extract anyway..."
    echo ""
fi

echo "Importing tokens from archive..."
echo "  Archive: $ARCHIVE_FILE"
echo ""

# Create temporary directory for extraction
EXTRACT_DIR=$(mktemp -d)
trap "rm -rf $EXTRACT_DIR" EXIT

# Extract archive
echo "Extracting archive..."
if tar -xzf "$ARCHIVE_FILE" -C "$EXTRACT_DIR" 2>/dev/null; then
    echo -e "${GREEN}✓ Archive extracted${NC}"
else
    echo -e "${RED}✗ Failed to extract archive${NC}"
    echo ""
    echo "The file may be corrupted or not a valid tar.gz archive."
    exit 1
fi

# Verify token file exists in archive
TOKEN_FILE="$EXTRACT_DIR/tokens/.token-cache.json"
if [ ! -f "$TOKEN_FILE" ]; then
    echo -e "${RED}✗ No token file found in archive${NC}"
    echo ""
    echo "The archive doesn't contain a valid token file."
    exit 2
fi

# Show export info if available
INFO_FILE="$EXTRACT_DIR/tokens/export-info.txt"
if [ -f "$INFO_FILE" ]; then
    echo ""
    echo "Archive Information:"
    echo "──────────────────────────────────────"
    cat "$INFO_FILE"
    echo "──────────────────────────────────────"
    echo ""
fi

# Detect execution mode
detect_execution_mode

# Determine import mode if not specified
if [ -z "$IMPORT_MODE" ]; then
    if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
        # Docker always uses file-based storage
        IMPORT_MODE="file"
        echo -e "${CYAN}ℹ Auto-selected file-based import (Docker mode)${NC}"
    else
        # Local mode defaults to keychain
        IMPORT_MODE="keychain"
        echo -e "${CYAN}ℹ Auto-selected keychain import (local mode)${NC}"
        echo "  Use --to-file to import to file-based storage instead"
    fi
    echo ""
fi

# Import tokens based on mode
IMPORT_SUCCESS=false

if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    # Docker mode - always import to container volume (file-based)
    if [ "$IMPORT_MODE" = "keychain" ]; then
        echo -e "${YELLOW}⚠  Warning: Cannot use keychain in Docker mode${NC}"
        echo "  Switching to file-based import for Docker container"
        echo ""
        IMPORT_MODE="file"
    fi
    
    echo "Importing to Docker volume..."
    
    # Ensure container data directory exists
    docker compose exec ms365-mcp mkdir -p /app/data 2>/dev/null || true
    
    # Copy token cache to container
    if docker compose cp "$TOKEN_FILE" ms365-mcp:/app/data/.token-cache.json 2>/dev/null; then
        IMPORT_LOCATION="docker-volume"
        IMPORT_SUCCESS=true
        
        # Also copy selected account file if it exists
        if [ -f "$EXTRACT_DIR/tokens/.selected-account.json" ]; then
            docker compose cp "$EXTRACT_DIR/tokens/.selected-account.json" ms365-mcp:/app/data/.selected-account.json 2>/dev/null || true
        fi
    else
        echo -e "${RED}✗ Failed to copy tokens to container${NC}"
        IMPORT_SUCCESS=false
    fi
    
elif [ "$IMPORT_MODE" = "keychain" ]; then
    # Import to keychain using helper script
    echo "Importing to system keychain..."
    
    if node "$SCRIPT_DIR/scripts/keychain-helper.js" import-to-keychain "$EXTRACT_DIR/tokens" 2>&1; then
        IMPORT_LOCATION="keychain"
        IMPORT_SUCCESS=true
    else
        echo -e "${RED}✗ Failed to import to keychain${NC}"
        echo ""
        echo "You may need to:"
        echo "  - Grant keychain access when prompted"
        echo "  - Or use --to-file to import to file-based storage"
        IMPORT_SUCCESS=false
    fi
    
else
    # Import to file-based storage
    echo "Importing to local file-based storage..."
    
    # Determine local token location
    if [ -d /app/data ]; then
        # Inside container
        TARGET_DIR="/app/data"
    else
        # Local machine
        TARGET_DIR="$SCRIPT_DIR"
    fi
    
    # Copy token cache file
    if cp "$TOKEN_FILE" "$TARGET_DIR/.token-cache.json"; then
        IMPORT_LOCATION="file-based ($TARGET_DIR)"
        IMPORT_SUCCESS=true
        
        # Also copy selected account file if it exists
        if [ -f "$EXTRACT_DIR/tokens/.selected-account.json" ]; then
            cp "$EXTRACT_DIR/tokens/.selected-account.json" "$TARGET_DIR/.selected-account.json" || true
        fi
    else
        echo -e "${RED}✗ Failed to copy tokens${NC}"
        IMPORT_SUCCESS=false
    fi
fi

# Report result
if [ "$IMPORT_SUCCESS" = true ]; then
    echo ""
    echo -e "${GREEN}✓ Tokens imported successfully${NC}"
    echo ""
    echo "Import Details:"
    echo "  Location: $IMPORT_LOCATION"
    echo "  Mode: $IMPORT_MODE"
    echo ""
    
    if [ "$IMPORT_MODE" = "file" ]; then
        echo -e "${YELLOW}Note:${NC} File-based token storage is less secure than keychain."
        echo "      Consider using keychain storage for better security."
        echo ""
    fi
    
    echo "Next steps:"
    echo "  1. Verify authentication: ./auth-verify.sh"
    if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
        echo "  2. Start the server: ./start.sh or docker compose up -d"
    else
        echo "  2. Start the server: node dist/index.js"
    fi
    echo ""
    exit 0
else
    echo ""
    echo "Import failed. Please check the error messages above."
    echo ""
    exit 1
fi
