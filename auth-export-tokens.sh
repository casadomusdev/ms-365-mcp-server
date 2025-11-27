#!/bin/bash
#
# MS-365 MCP Server - Export Tokens Script
#
# Exports authentication tokens to a timestamped compressed archive.
# Supports dual-mode operation: Docker or local Node.js
# Supports both keychain and file-based token storage
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

# Create temporary directory for token extraction
TEMP_DIR=$(mktemp -d)
trap "rm -rf $TEMP_DIR" EXIT

# Determine token file location based on mode
if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
    # Docker mode - tokens are in container volume (always file-based in Docker)
    TOKEN_LOCATION="docker-volume"
    
    echo "Checking for tokens in Docker volume..."
    if docker compose exec ms365-mcp test -f /app/data/.token-cache.json 2>/dev/null; then
        # Copy token files from container
        docker compose cp ms365-mcp:/app/data/.token-cache.json "$TEMP_DIR/.token-cache.json" 2>/dev/null || true
        docker compose cp ms365-mcp:/app/data/.selected-account.json "$TEMP_DIR/.selected-account.json" 2>/dev/null || true
        TOKENS_FOUND=true
    else
        TOKENS_FOUND=false
    fi
else
    # Local mode - tokens might be in keychain or files
    TOKEN_LOCATION="local"
    TOKENS_FOUND=false
    
    # First, try to export from keychain using helper script
    echo "Checking for tokens in system keychain..."
    if node "$SCRIPT_DIR/scripts/keychain-helper.js" export-from-keychain "$TEMP_DIR" 2>/dev/null; then
        echo -e "${GREEN}✓ Tokens exported from keychain${NC}"
        TOKEN_LOCATION="keychain"
        TOKENS_FOUND=true
    else
        # Keychain export failed, try file-based tokens
        echo "Keychain export failed, checking for file-based tokens..."
        
        if [ -f "$SCRIPT_DIR/.token-cache.json" ]; then
            cp "$SCRIPT_DIR/.token-cache.json" "$TEMP_DIR/.token-cache.json"
            TOKENS_FOUND=true
        elif [ -f /app/data/.token-cache.json ]; then
            cp /app/data/.token-cache.json" "$TEMP_DIR/.token-cache.json"
            TOKENS_FOUND=true
        fi
        
        # Also copy selected account if it exists
        if [ -f "$SCRIPT_DIR/.selected-account.json" ]; then
            cp "$SCRIPT_DIR/.selected-account.json" "$TEMP_DIR/.selected-account.json"
        elif [ -f /app/data/.selected-account.json ]; then
            cp /app/data/.selected-account.json" "$TEMP_DIR/.selected-account.json"
        fi
        
        if [ "$TOKENS_FOUND" = true ]; then
            TOKEN_LOCATION="file-based"
        fi
    fi
fi

# Check if tokens were found
if [ "$TOKENS_FOUND" = false ] || [ ! -f "$TEMP_DIR/.token-cache.json" ]; then
    echo -e "${RED}✗ No authentication tokens found${NC}"
    echo ""
    echo "You need to authenticate first:"
    echo "  For keychain storage: ./auth-login.sh"
    echo "  For file-based storage: ./auth-login.sh --force-file-cache"
    echo ""
    exit 2
fi

# Generate output filename with timestamp
TIMESTAMP=$(date +"%Y%m%d-%H%M%S")
if [ -n "$1" ]; then
    OUTPUT_FILE="$1"
else
    OUTPUT_FILE="tokens-${TIMESTAMP}.tar.gz"
fi

echo ""
echo "Exporting tokens..."
echo "  Source: $TOKEN_LOCATION"
echo "  Output: $OUTPUT_FILE"
echo ""

# Create temporary directory for archive contents
ARCHIVE_DIR=$(mktemp -d)
trap "rm -rf $ARCHIVE_DIR $TEMP_DIR" EXIT
mkdir -p "$ARCHIVE_DIR/tokens"

# Copy token files
cp "$TEMP_DIR/.token-cache.json" "$ARCHIVE_DIR/tokens/.token-cache.json"
if [ -f "$TEMP_DIR/.selected-account.json" ]; then
    cp "$TEMP_DIR/.selected-account.json" "$ARCHIVE_DIR/tokens/.selected-account.json"
fi

# Include PowerShell certificate if it exists
CERT_INCLUDED=false
if [ -f "$SCRIPT_DIR/certs/ms365-powershell.pfx" ]; then
    echo "Including PowerShell certificate..."
    mkdir -p "$ARCHIVE_DIR/tokens/certs"
    cp "$SCRIPT_DIR/certs/ms365-powershell.pfx" "$ARCHIVE_DIR/tokens/certs/ms365-powershell.pfx"
    cp "$SCRIPT_DIR/certs/ms365-powershell.cer" "$ARCHIVE_DIR/tokens/certs/ms365-powershell.cer" 2>/dev/null || true
    CERT_INCLUDED=true
    echo -e "${GREEN}✓ PowerShell certificate included${NC}"
fi

# Include certificate password from backup file if available
if [ -f "$SCRIPT_DIR/certs/.cert-password.txt" ] && [ "$CERT_INCLUDED" = true ]; then
    cp "$SCRIPT_DIR/certs/.cert-password.txt" "$ARCHIVE_DIR/tokens/certs/.cert-password.txt"
    echo -e "${GREEN}✓ Certificate password included${NC}"
fi

# Create metadata file
cat > "$ARCHIVE_DIR/tokens/export-info.txt" << EOF
MS-365 MCP Server Token Export
===============================

Export Date: $(date '+%Y-%m-%d %H:%M:%S %Z')
Hostname: $(hostname)
Export Mode: $TOKEN_LOCATION
Source: $EXECUTION_MODE
$([ "$CERT_INCLUDED" = true ] && echo "Certificate: Included (PFX + password)")

Files Included:
$(ls -1 "$ARCHIVE_DIR/tokens/" | grep -v "export-info.txt" | sed 's/^/  - /')

Important:
- Keep this file secure and encrypted
- Tokens provide access to your Microsoft 365 account
$([ "$CERT_INCLUDED" = true ] && echo "- Certificate is portable - works on any machine")
- Import on target machine with: ./auth-import-tokens.sh $OUTPUT_FILE
- Use --to-keychain flag to import into keychain
- Use --to-file flag to import into file-based storage
$([ "$CERT_INCLUDED" = true ] && echo "- Certificate auto-imported to certs/ directory")
EOF

# Create compressed archive
cd "$ARCHIVE_DIR"
tar -czf "$SCRIPT_DIR/$OUTPUT_FILE" tokens/

# Cleanup
cd "$SCRIPT_DIR"

# Verify archive was created
if [ -f "$OUTPUT_FILE" ]; then
    FILE_SIZE=$(du -h "$OUTPUT_FILE" | cut -f1)
    echo -e "${GREEN}✓ Tokens exported successfully${NC}"
    echo ""
    echo "Archive Details:"
    echo "  File: $OUTPUT_FILE"
    echo "  Size: $FILE_SIZE"
    echo "  Source: $TOKEN_LOCATION"
    if [ "$CERT_INCLUDED" = true ]; then
        echo "  Certificate: Included"
    fi
    echo ""
    echo -e "${YELLOW}⚠  SECURITY WARNING:${NC}"
    echo "  This file contains sensitive authentication tokens."
    if [ "$CERT_INCLUDED" = true ]; then
        echo "  This file also contains the PowerShell certificate + password."
    fi
    echo "  Store it securely and transport it safely."
    echo ""
    echo "To import on another machine:"
    echo "  1. Transfer the file securely (scp, encrypted transfer, etc.)"
    echo "  2. Import to keychain: ./auth-import-tokens.sh $OUTPUT_FILE --to-keychain"
    echo "  3. Or import to files: ./auth-import-tokens.sh $OUTPUT_FILE --to-file"
    if [ "$CERT_INCLUDED" = true ]; then
        echo "  4. Certificate will be auto-imported to certs/ directory"
    fi
    echo ""
    exit 0
else
    echo -e "${RED}✗ Failed to create archive${NC}"
    echo ""
    exit 1
fi
