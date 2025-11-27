#!/bin/bash
#
# PowerShell Debug Script
# Helps diagnose PowerShell execution failures
#

set -e

# Get script directory
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

echo -e "${CYAN}PowerShell Debugging Tool${NC}"
echo "=============================="
echo ""

# Step 1: Check PowerShell availability
echo -e "${YELLOW}[1/5] Checking PowerShell Core availability...${NC}"
if command -v pwsh &> /dev/null; then
    PWSH_VERSION=$(pwsh -Version 2>&1)
    echo -e "${GREEN}✓ PowerShell Core found: $PWSH_VERSION${NC}"
else
    echo -e "${RED}✗ PowerShell Core (pwsh) not found${NC}"
    echo "Install with: brew install --cask powershell"
    exit 1
fi
echo ""

# Step 2: Check Exchange Online module
echo -e "${YELLOW}[2/5] Checking Exchange Online module...${NC}"
MODULE_CHECK=$(pwsh -Command "Get-Module -ListAvailable ExchangeOnlineManagement" 2>&1)
if echo "$MODULE_CHECK" | grep -q "ExchangeOnlineManagement"; then
    echo -e "${GREEN}✓ Exchange Online module found${NC}"
    pwsh -Command "Get-InstalledModule ExchangeOnlineManagement | Select-Object Name, Version | Format-List"
else
    echo -e "${RED}✗ Exchange Online module not found${NC}"
    echo "Install with: pwsh -Command \"Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber\""
    exit 1
fi
echo ""

# Step 3: Check certificate configuration
echo -e "${YELLOW}[3/5] Checking certificate configuration...${NC}"

# Load .env if it exists
if [ -f ".env" ]; then
    source .env
fi

# Use standard certificate paths (convention over configuration)
CERT_PATH="$SCRIPT_DIR/certs/ms365-powershell.pfx"
CERT_PASSWORD_FILE="$SCRIPT_DIR/certs/.cert-password.txt"

# Check if certificate file exists
if [ ! -f "$CERT_PATH" ]; then
    echo -e "${RED}✗ Certificate file not found: $CERT_PATH${NC}"
    echo "Run './auth-generate-cert.sh' to generate a certificate"
    exit 1
fi

echo -e "${GREEN}✓ Certificate file found: $CERT_PATH${NC}"

# Check certificate password file
if [ ! -f "$CERT_PASSWORD_FILE" ]; then
    echo -e "${RED}✗ Certificate password file not found: $CERT_PASSWORD_FILE${NC}"
    echo "Run './auth-generate-cert.sh' to generate a certificate"
    exit 1
fi

# Read password from file
CERT_PASSWORD=$(cat "$CERT_PASSWORD_FILE")
PASSWORD_LENGTH=${#CERT_PASSWORD}
echo -e "${GREEN}✓ Certificate password file found (length: $PASSWORD_LENGTH chars)${NC}"
echo ""

# Step 4: Get required parameters
echo -e "${YELLOW}[4/5] Gathering parameters...${NC}"

# Get client ID
CLIENT_ID="${MS365_MCP_CLIENT_ID:-}"
if [ -z "$CLIENT_ID" ]; then
    echo -e "${RED}✗ MS365_MCP_CLIENT_ID not set in .env${NC}"
    exit 1
fi
echo "Client ID: $CLIENT_ID"

# Get organization domain
ORGANIZATION="${MS365_MCP_ORGANIZATION:-}"
if [ -z "$ORGANIZATION" ]; then
    echo -e "${RED}✗ MS365_MCP_ORGANIZATION not set in .env${NC}"
    echo "Set your tenant domain (e.g., contoso.onmicrosoft.com) in .env"
    exit 1
fi
echo "Organization: $ORGANIZATION"

# Get user email
USER_EMAIL="${MS365_MCP_IMPERSONATE_USER:-}"
if [ -z "$USER_EMAIL" ]; then
    echo -e "${YELLOW}⚠ MS365_MCP_IMPERSONATE_USER not set${NC}"
    echo "Enter user email to test:"
    read USER_EMAIL
fi
echo "User email: $USER_EMAIL"
echo ""

# Step 5: Run PowerShell script with detailed output
echo -e "${YELLOW}[5/5] Testing PowerShell script...${NC}"
echo "Running: pwsh scripts/check-mailbox-permissions.ps1"
echo ""

# Run PowerShell with verbose output
pwsh -NoProfile -NonInteractive -Command "
    \$ErrorActionPreference = 'Continue'
    \$VerbosePreference = 'Continue'
    
    Write-Host '--- PowerShell Script Execution ---' -ForegroundColor Cyan
    Write-Host 'Script: scripts/check-mailbox-permissions.ps1' -ForegroundColor Cyan
    Write-Host 'User Email: $USER_EMAIL' -ForegroundColor Cyan
    Write-Host 'App ID: $CLIENT_ID' -ForegroundColor Cyan
    Write-Host 'Organization: $ORGANIZATION' -ForegroundColor Cyan
    Write-Host 'Certificate: $CERT_PATH' -ForegroundColor Cyan
    Write-Host ''
    
    try {
        & './scripts/check-mailbox-permissions.ps1' \
            -UserEmail '$USER_EMAIL' \
            -CertificatePath '$CERT_PATH' \
            -CertificatePassword '$CERT_PASSWORD' \
            -AppId '$CLIENT_ID' \
            -Organization '$ORGANIZATION' \
            -Verbose
        
        Write-Host ''
        Write-Host '--- Script completed successfully ---' -ForegroundColor Green
    } catch {
        Write-Host ''
        Write-Host '--- Script failed ---' -ForegroundColor Red
        Write-Host \"Error: \$_\" -ForegroundColor Red
        Write-Host \"Error Type: \$(\$_.Exception.GetType().FullName)\" -ForegroundColor Red
        Write-Host \"Stack Trace:\" -ForegroundColor Red
        Write-Host \$_.ScriptStackTrace -ForegroundColor Red
        exit 1
    }
"

PWSH_EXIT_CODE=$?

echo ""
if [ $PWSH_EXIT_CODE -eq 0 ]; then
    echo -e "${GREEN}✓ PowerShell script executed successfully${NC}"
else
    echo -e "${RED}✗ PowerShell script failed with exit code: $PWSH_EXIT_CODE${NC}"
    echo ""
    echo "Common issues:"
    echo "1. Certificate not uploaded to Azure AD - upload the .cer file"
    echo "2. Insufficient permissions - check Azure AD app permissions"
    echo "3. Invalid organization domain - verify MS365_MCP_ORGANIZATION in .env"
    echo "4. Network issues - check connection to Exchange Online"
    echo "5. Certificate expired - generate a new one with ./auth-generate-cert.sh"
fi

echo ""
echo "Debug complete."
