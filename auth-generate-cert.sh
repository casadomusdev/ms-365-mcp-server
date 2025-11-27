#!/bin/bash
#
# PowerShell Certificate Generator for Exchange Online
# Generates a self-signed certificate for app-only authentication
#

set -e

# Get script directory
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Load common functions
source .scripts-lib.sh

# Certificate configuration
CERT_DIR="$SCRIPT_DIR/certs"
CERT_NAME="ms365-powershell"
CERT_PFX="$CERT_DIR/${CERT_NAME}.pfx"
CERT_CER="$CERT_DIR/${CERT_NAME}.cer"
DEFAULT_VALIDITY_YEARS=2

echo_section "PowerShell Certificate Generator"

# Check if OpenSSL is available (cross-platform)
if ! command -v openssl &> /dev/null; then
    echo_error "OpenSSL is required but not found"
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "Install with: brew install openssl"
    else
        echo "Install with: sudo apt-get install openssl"
    fi
    exit 1
fi

# Check if certificate already exists
if [ -f "$CERT_PFX" ]; then
    echo_warning "Certificate already exists: $CERT_PFX"
    echo ""
    read -p "Do you want to regenerate the certificate? (y/N): " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        echo_info "Keeping existing certificate"
        echo_success "Certificate and password file already exist in ./certs/"
        exit 0
    fi
    echo_info "Regenerating certificate..."
fi

# Create certs directory if it doesn't exist
mkdir -p "$CERT_DIR"

# Ask user for certificate validity period
echo ""
echo "Certificate validity period:"
echo "  - Default: $DEFAULT_VALIDITY_YEARS years (recommended)"
echo "  - Azure AD max: 3 years"
echo "  - Matches client secret standard: 2 years"
echo ""
read -p "Enter validity in years [$DEFAULT_VALIDITY_YEARS]: " VALIDITY_YEARS
VALIDITY_YEARS=${VALIDITY_YEARS:-$DEFAULT_VALIDITY_YEARS}

# Validate input
if ! [[ "$VALIDITY_YEARS" =~ ^[0-9]+$ ]] || [ "$VALIDITY_YEARS" -lt 1 ] || [ "$VALIDITY_YEARS" -gt 3 ]; then
    echo_error "Invalid validity period. Must be 1-3 years."
    exit 1
fi

# Calculate expiry date
EXPIRY_DATE=$(date -v +${VALIDITY_YEARS}y '+%Y-%m-%d' 2>/dev/null || date -d "+${VALIDITY_YEARS} years" '+%Y-%m-%d')

echo_info "Generating certificate valid for $VALIDITY_YEARS year(s) (until $EXPIRY_DATE)..."

# Load .env to get tenant info
if [ -f ".env" ]; then
    source .env
fi

TENANT_ID="${MS365_MCP_TENANT_ID:-your-tenant-id}"
CLIENT_ID="${MS365_MCP_CLIENT_ID:-your-client-id}"

# Generate certificate using OpenSSL (cross-platform)
echo_info "Creating self-signed certificate with OpenSSL..."

# Generate a secure random password (32 alphanumeric characters)
CERT_PASSWORD=$(openssl rand -base64 32 | tr -dc 'a-zA-Z0-9' | head -c 32)

# Certificate subject
CERT_SUBJECT="/CN=MS365-MCP-PowerShell-${CLIENT_ID}/O=MS365-MCP-Server"

# Calculate validity in days
VALIDITY_DAYS=$((VALIDITY_YEARS * 365))

# Temporary files
CERT_KEY="$CERT_DIR/.temp_private.key"
CERT_CSR="$CERT_DIR/.temp_cert.csr"
CERT_CRT="$CERT_DIR/.temp_cert.crt"

# Step 1: Generate private key
echo_info "Generating RSA private key (2048 bits)..."
openssl genrsa -out "$CERT_KEY" 2048 2>/dev/null

if [ $? -ne 0 ]; then
    echo_error "Failed to generate private key"
    rm -f "$CERT_KEY"
    exit 1
fi

# Step 2: Generate certificate signing request (CSR)
echo_info "Creating certificate signing request..."
openssl req -new -key "$CERT_KEY" -out "$CERT_CSR" -subj "$CERT_SUBJECT" 2>/dev/null

if [ $? -ne 0 ]; then
    echo_error "Failed to create CSR"
    rm -f "$CERT_KEY" "$CERT_CSR"
    exit 1
fi

# Step 3: Generate self-signed certificate
echo_info "Generating self-signed certificate (valid for $VALIDITY_YEARS year(s))..."
openssl x509 -req -days "$VALIDITY_DAYS" -in "$CERT_CSR" -signkey "$CERT_KEY" -out "$CERT_CRT" -sha256 2>/dev/null

if [ $? -ne 0 ]; then
    echo_error "Failed to generate certificate"
    rm -f "$CERT_KEY" "$CERT_CSR" "$CERT_CRT"
    exit 1
fi

# Step 4: Export as PFX (PKCS#12) with password
echo_info "Exporting PFX (private key + certificate)..."
openssl pkcs12 -export -out "$CERT_PFX" -inkey "$CERT_KEY" -in "$CERT_CRT" -password "pass:$CERT_PASSWORD" 2>/dev/null

if [ $? -ne 0 ]; then
    echo_error "Failed to export PFX"
    rm -f "$CERT_KEY" "$CERT_CSR" "$CERT_CRT"
    exit 1
fi

# Step 5: Export public key as CER (DER format for Azure AD)
echo_info "Exporting CER (public key for Azure AD)..."
openssl x509 -outform der -in "$CERT_CRT" -out "$CERT_CER" 2>/dev/null

if [ $? -ne 0 ]; then
    echo_error "Failed to export CER"
    rm -f "$CERT_KEY" "$CERT_CSR" "$CERT_CRT"
    exit 1
fi

# Clean up temporary files
rm -f "$CERT_KEY" "$CERT_CSR" "$CERT_CRT"

echo_success "Certificate generated successfully!"
echo ""
echo "Certificate files created:"
echo "  Private key (PFX): $CERT_PFX"
echo "  Public key (CER):  $CERT_CER"
echo ""

# Save password to backup file (for automatic use by PowerShellService)
echo_info "Saving certificate password..."
CERT_PASSWORD_FILE="$CERT_DIR/.cert-password.txt"
echo "$CERT_PASSWORD" > "$CERT_PASSWORD_FILE"
chmod 600 "$CERT_PASSWORD_FILE"  # Secure the password file

echo_success "Password saved to $CERT_PASSWORD_FILE"
echo_info "PowerShellService will automatically use certificate from: $CERT_DIR"

# Display Azure AD upload instructions
echo ""
echo_section "NEXT STEP: Upload Certificate to Azure AD"
echo ""
echo "1. Go to Azure Portal: https://portal.azure.com"
echo "2. Navigate to: Azure Active Directory → App registrations"
echo "3. Select your app: $CLIENT_ID"
echo "4. Go to: Certificates & secrets → Certificates tab"
echo "5. Click: [Upload certificate]"
echo "6. Select file: $CERT_CER"
echo "7. Add a description: 'PowerShell Exchange Online (expires $EXPIRY_DATE)'"
echo "8. Click: [Add]"
echo ""
echo_warning "⚠ Certificate will only work AFTER you upload it to Azure AD!"
echo ""
echo "Once uploaded, run './auth-verify.sh' to test the connection"
echo ""
