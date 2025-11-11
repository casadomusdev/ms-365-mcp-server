#!/bin/bash
#
# MS-365 MCP Server - Export Tokens Script
#
# Exports authentication tokens to a timestamped compressed archive.
# Supports dual-mode operation: Docker or local Node.js
#
# Usage:
#   ./auth-export-tokens.sh [output-file]
#
# Arguments:
#   output-file  Optional: Custom output filename (default: tokens-YYYYMMDD-HHMMSS.tar.gz)
#
# Exit codes:
#   0 - Export successful
#   1 - Export failed
#   2 - No tokens found to export

set -e

# Get script directory and source shared library
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/.scripts-lib.sh"

echo -e "${CYAN}MS-365 MCP Server - Export Tokens${NC}"
echo "===================================="
echo ""

# Detect execution mode
detect_execution_mode

# Determine token file location based on mode
if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    # Docker mode - tokens are in container volume
    TOKEN_LOCATION="docker-volume"
    TEMP_DIR=$(mktemp -d)
    
    # Copy token file from container to temp location
    echo "Checking for tokens in Docker volume..."
    if docker compose exec ms365-mcp test -f /app/data/.token-cache.json 2>/dev/null; then
        docker compose cp ms365-mcp:/app/data/.token-cache.json "$TEMP_DIR/.token-cache.json" 2>/dev/null || true
    fi
    
    TOKEN_FILE="$TEMP_DIR/.token-cache.json"
else
    # Direct mode - tokens are in local directory
    TOKEN_LOCATION="local"
    if [ -f .token-cache.json ]; then
        TOKEN_FILE=".token-cache.json"
    elif [ -f /app/data/.token-cache.json ]; then
        TOKEN_FILE="/app/data/.token-cache.json"
    else
        TOKEN_FILE=".token-cache.json"
    fi
fi

# Check if tokens exist
if [ ! -f "$TOKEN_FILE" ]; then
    echo -e "${RED}✗ No authentication tokens found${NC}"
    echo ""
    echo "You need to authenticate first:"
    echo "  ./auth-login.sh --force-file-cache"
    echo ""
    echo "Note: Token export only works with file-based cache."
    echo "      System keychain tokens cannot be exported."
    echo ""
    
    # Cleanup temp dir if created
    [ -d "$TEMP_DIR" ] && rm -rf "$TEMP_DIR"
    
    exit 2
fi

# Generate output filename with timestamp
TIMESTAMP=$(date +"%Y%m%d-%H%M%S")
if [ -n "$1" ]; then
    OUTPUT_FILE="$1"
else
    OUTPUT_FILE="tokens-${TIMESTAMP}.tar.gz"
fi

echo "Exporting tokens..."
echo "  Source: $TOKEN_LOCATION"
echo "  Output: $OUTPUT_FILE"
echo ""

# Create temporary directory for archive contents
ARCHIVE_DIR=$(mktemp -d)
mkdir -p "$ARCHIVE_DIR/tokens"

# Copy token file
cp "$TOKEN_FILE" "$ARCHIVE_DIR/tokens/.token-cache.json"

# Create metadata file
cat > "$ARCHIVE_DIR/tokens/export-info.txt" << EOF
MS-365 MCP Server Token Export
===============================

Export Date: $(date '+%Y-%m-%d %H:%M:%S %Z')
Hostname: $(hostname)
Export Mode: $TOKEN_LOCATION

Important:
- Keep this file secure and encrypted
- Tokens provide access to your Microsoft 365 account
- Import on target machine with: ./auth-import-tokens.sh $OUTPUT_FILE
EOF

# Create compressed archive
cd "$ARCHIVE_DIR"
tar -czf "$SCRIPT_DIR/$OUTPUT_FILE" tokens/

# Cleanup
cd "$SCRIPT_DIR"
rm -rf "$ARCHIVE_DIR"
[ -d "$TEMP_DIR" ] && rm -rf "$TEMP_DIR"

# Verify archive was created
if [ -f "$OUTPUT_FILE" ]; then
    FILE_SIZE=$(du -h "$OUTPUT_FILE" | cut -f1)
    echo -e "${GREEN}✓ Tokens exported successfully${NC}"
    echo ""
    echo "Archive Details:"
    echo "  File: $OUTPUT_FILE"
    echo "  Size: $FILE_SIZE"
    echo ""
    echo -e "${YELLOW}⚠  SECURITY WARNING:${NC}"
    echo "  This file contains sensitive authentication tokens."
    echo "  Store it securely and transport it safely."
    echo ""
    echo "To import on another machine:"
    echo "  1. Transfer the file securely (scp, encrypted transfer, etc.)"
    echo "  2. Run: ./auth-import-tokens.sh $OUTPUT_FILE"
    echo ""
    exit 0
else
    echo -e "${RED}✗ Failed to create archive${NC}"
    echo ""
    exit 1
fi
