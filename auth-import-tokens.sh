#!/bin/bash
#
# MS-365 MCP Server - Import Tokens Script
#
# Imports authentication tokens from local files into the Docker container.
# This allows you to transfer tokens from another machine or restore from backup.
#
# Usage:
#   ./auth-import-tokens.sh [input-directory]
#
# Arguments:
#   input-directory  Optional. Directory containing tokens (default: ./tokens-backup)
#
# Exit codes:
#   0 - Import successful
#   1 - Import failed

set -e

# Color codes
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Default input directory
INPUT_DIR="${1:-./tokens-backup}"

echo -e "${YELLOW}MS-365 MCP Server - Import Tokens${NC}"
echo "====================================="
echo ""

# Check if input directory exists
if [ ! -d "$INPUT_DIR" ]; then
    echo -e "${RED}✗ Input directory not found: $INPUT_DIR${NC}"
    echo ""
    echo "Usage: ./auth-import-tokens.sh [input-directory]"
    echo ""
    exit 1
fi

# Check if token files exist
if [ ! -f "$INPUT_DIR/.token-cache.json" ]; then
    echo -e "${RED}✗ Token cache file not found: $INPUT_DIR/.token-cache.json${NC}"
    echo ""
    echo "This directory doesn't appear to contain exported tokens."
    echo "Run ./auth-export-tokens.sh first on the source machine."
    echo ""
    exit 1
fi

echo "Importing tokens from: $INPUT_DIR"
echo ""

# Check if container exists and create it if needed
if ! docker compose ps | grep -q ms365-mcp; then
    echo "Container not running. Creating it..."
    docker compose up -d
    echo "Waiting for container to be ready..."
    sleep 2
fi

# Use docker cp to copy files into the volume
CONTAINER_NAME=$(docker compose ps -q ms365-mcp)

if [ -z "$CONTAINER_NAME" ]; then
    echo -e "${RED}✗ Failed to get container name${NC}"
    exit 1
fi

echo "Copying .token-cache.json..."
docker cp "$INPUT_DIR/.token-cache.json" "$CONTAINER_NAME:/app/data/.token-cache.json"
echo -e "${GREEN}✓${NC} Imported .token-cache.json"

if [ -f "$INPUT_DIR/.selected-account.json" ]; then
    echo "Copying .selected-account.json..."
    docker cp "$INPUT_DIR/.selected-account.json" "$CONTAINER_NAME:/app/data/.selected-account.json"
    echo -e "${GREEN}✓${NC} Imported .selected-account.json"
fi

# Fix permissions
echo "Setting correct permissions..."
docker compose exec -T ms365-mcp chown node:node /app/data/.token-cache.json
if [ -f "$INPUT_DIR/.selected-account.json" ]; then
    docker compose exec -T ms365-mcp chown node:node /app/data/.selected-account.json
fi

echo ""
echo -e "${GREEN}✓ Import completed successfully${NC}"
echo ""
echo "Tokens have been imported. Verify with:"
echo "  ./auth-verify.sh"
echo ""
echo "Or check inside the container:"
echo "  docker compose exec ms365-mcp /app/health-check.sh"
echo ""

# Optionally show export timestamp if available
if [ -f "$INPUT_DIR/.export-timestamp" ]; then
    EXPORT_TIME=$(cat "$INPUT_DIR/.export-timestamp")
    echo "Note: These tokens were exported at: $EXPORT_TIME"
    echo ""
fi
