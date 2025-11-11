#!/bin/bash
#
# MS-365 MCP Server - Export Tokens Script
#
# Exports authentication tokens from the Docker container to local files.
# This allows you to transfer tokens to another machine or create backups.
#
# Usage:
#   ./auth-export-tokens.sh [output-directory]
#
# Arguments:
#   output-directory  Optional. Directory to save tokens (default: ./tokens-backup)
#
# Exit codes:
#   0 - Export successful
#   1 - Export failed

set -e

# Color codes
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Default output directory
OUTPUT_DIR="${1:-./tokens-backup}"
TIMESTAMP=$(date +%Y%m%d-%H%M%S)

echo -e "${YELLOW}MS-365 MCP Server - Export Tokens${NC}"
echo "====================================="
echo ""

# Create output directory
mkdir -p "$OUTPUT_DIR"

echo "Exporting tokens from Docker container..."
echo "Output directory: $OUTPUT_DIR"
echo ""

# Check if container exists
if ! docker compose ps | grep -q ms365-mcp; then
    echo -e "${RED}✗ Container 'ms365-mcp' not found${NC}"
    echo ""
    echo "Make sure the container has been created with:"
    echo "  docker compose up -d"
    echo ""
    exit 1
fi

# Export token files from Docker volume
echo "Copying .token-cache.json..."
if docker compose exec -T ms365-mcp cat /app/data/.token-cache.json > "$OUTPUT_DIR/.token-cache.json" 2>/dev/null; then
    echo -e "${GREEN}✓${NC} Exported .token-cache.json"
else
    echo -e "${YELLOW}⚠${NC} .token-cache.json not found (container may not be authenticated yet)"
fi

echo "Copying .selected-account.json..."
if docker compose exec -T ms365-mcp cat /app/data/.selected-account.json > "$OUTPUT_DIR/.selected-account.json" 2>/dev/null; then
    echo -e "${GREEN}✓${NC} Exported .selected-account.json"
else
    echo -e "${YELLOW}⚠${NC} .selected-account.json not found"
fi

# Create a timestamp file for reference
echo "$TIMESTAMP" > "$OUTPUT_DIR/.export-timestamp"

echo ""

# Check if any files were actually exported
if [ -f "$OUTPUT_DIR/.token-cache.json" ]; then
    echo -e "${GREEN}✓ Export completed successfully${NC}"
    echo ""
    echo "Token files saved to: $OUTPUT_DIR"
    echo ""
    echo "⚠️  SECURITY WARNING:"
    echo "These files contain sensitive authentication tokens!"
    echo ""
    echo "To secure the tokens:"
    echo "  1. Encrypt: tar czf - \"$OUTPUT_DIR\" | gpg -c > tokens-$TIMESTAMP.tar.gz.gpg"
    echo "  2. Transfer securely (scp, encrypted USB, etc.)"
    echo "  3. Delete unencrypted files: rm -rf \"$OUTPUT_DIR\""
    echo ""
    echo "To import on another machine:"
    echo "  ./auth-import-tokens.sh $OUTPUT_DIR"
    echo ""
    exit 0
else
    echo -e "${RED}✗ Export failed - no token files found${NC}"
    echo ""
    echo "This usually means authentication has not been performed yet."
    echo "Run ./auth-login.sh first to authenticate."
    echo ""
    exit 1
fi
