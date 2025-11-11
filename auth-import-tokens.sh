#!/bin/bash
#
# MS-365 MCP Server - Import Tokens Script
#
# Imports authentication tokens from a compressed archive.
# Supports dual-mode operation: Docker or local Node.js
#
# Usage:
#   ./auth-import-tokens.sh <archive-file>
#
# Arguments:
#   archive-file  Path to tokens-*.tar.gz file to import
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

# Check if archive file was provided
if [ -z "$1" ]; then
    echo -e "${RED}✗ No archive file specified${NC}"
    echo ""
    echo "Usage: ./auth-import-tokens.sh <archive-file>"
    echo ""
    echo "Example:"
    echo "  ./auth-import-tokens.sh tokens-20231201-143022.tar.gz"
    echo ""
    exit 2
fi

ARCHIVE_FILE="$1"

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

# Extract archive
echo "Extracting archive..."
if tar -xzf "$ARCHIVE_FILE" -C "$EXTRACT_DIR" 2>/dev/null; then
    echo -e "${GREEN}✓ Archive extracted${NC}"
else
    echo -e "${RED}✗ Failed to extract archive${NC}"
    echo ""
    echo "The file may be corrupted or not a valid tar.gz archive."
    rm -rf "$EXTRACT_DIR"
    exit 1
fi

# Verify token file exists in archive
TOKEN_FILE="$EXTRACT_DIR/tokens/.token-cache.json"
if [ ! -f "$TOKEN_FILE" ]; then
    echo -e "${RED}✗ No token file found in archive${NC}"
    echo ""
    echo "The archive doesn't contain a valid token file."
    rm -rf "$EXTRACT_DIR"
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

# Import tokens based on mode
if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    echo "Importing to Docker volume..."
    
    # Ensure container data directory exists
    docker compose exec ms365-mcp mkdir -p /app/data 2>/dev/null || true
    
    # Copy token file to container
    if docker compose cp "$TOKEN_FILE" ms365-mcp:/app/data/.token-cache.json 2>/dev/null; then
        IMPORT_LOCATION="docker-volume"
        IMPORT_SUCCESS=true
    else
        echo -e "${RED}✗ Failed to copy tokens to container${NC}"
        IMPORT_SUCCESS=false
    fi
else
    echo "Importing to local directory..."
    
    # Determine local token location
    if [ -d /app/data ]; then
        # Inside container
        TARGET_DIR="/app/data"
    else
        # Local machine
        TARGET_DIR="$SCRIPT_DIR"
    fi
    
    # Copy token file
    if cp "$TOKEN_FILE" "$TARGET_DIR/.token-cache.json"; then
        IMPORT_LOCATION="local ($TARGET_DIR)"
        IMPORT_SUCCESS=true
    else
        echo -e "${RED}✗ Failed to copy tokens${NC}"
        IMPORT_SUCCESS=false
    fi
fi

# Cleanup
rm -rf "$EXTRACT_DIR"

# Report result
if [ "$IMPORT_SUCCESS" = true ]; then
    echo -e "${GREEN}✓ Tokens imported successfully${NC}"
    echo ""
    echo "Import Details:"
    echo "  Location: $IMPORT_LOCATION"
    echo ""
    echo "Next steps:"
    echo "  1. Verify authentication: ./auth-verify.sh"
    echo "  2. Start the server: ./start.sh"
    echo ""
    exit 0
else
    echo ""
    echo "Import failed. Please check the error messages above."
    echo ""
    exit 1
fi
