#!/bin/bash
#
# MS-365 MCP Server - Start Script
#
# Sets up and starts the MS-365 MCP server with Docker Compose.
# Validates environment configuration before starting.
#
# Usage:
#   ./start.sh [--build]
#
# Options:
#   --build    Force rebuild of Docker image
#
# Exit codes:
#   0 - Server started successfully
#   1 - Configuration error or startup failed

set -e

# Color codes
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# Parse arguments
FORCE_BUILD=false
for arg in "$@"; do
    case $arg in
        --build)
            FORCE_BUILD=true
            shift
            ;;
        *)
            echo "Unknown option: $arg"
            echo "Usage: ./start.sh [--build]"
            exit 1
            ;;
    esac
done

echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
echo -e "${CYAN}  MS-365 MCP Server - Setup & Start${NC}"
echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
echo ""

# Check if .env file exists
if [ ! -f .env ]; then
    echo -e "${RED}✗ Error: .env file not found${NC}"
    echo ""
    echo "To create your .env file:"
    echo "  1. Copy the example: cp .env.example .env"
    echo "  2. Edit .env and fill in your Azure AD credentials"
    echo ""
    echo "Required values:"
    echo "  - MS365_MCP_CLIENT_ID     (from Azure AD App Registration)"
    echo "  - MS365_MCP_TENANT_ID     (use 'common' or your tenant ID)"
    echo ""
    echo "Optional but recommended:"
    echo "  - MS365_MCP_CLIENT_SECRET (if using confidential client)"
    echo ""
    echo "For detailed setup instructions, see:"
    echo "  https://github.com/your-repo/ms-365-mcp-server#setup"
    echo ""
    exit 1
fi

echo -e "${GREEN}✓ Found .env file${NC}"

# Load .env file
set -a
source .env
set +a

# Validate required environment variables
MISSING_VARS=()

if [ -z "$MS365_MCP_CLIENT_ID" ]; then
    MISSING_VARS+=("MS365_MCP_CLIENT_ID")
fi

if [ -z "$MS365_MCP_TENANT_ID" ]; then
    MISSING_VARS+=("MS365_MCP_TENANT_ID")
fi

# Check if any required variables are missing
if [ ${#MISSING_VARS[@]} -ne 0 ]; then
    echo -e "${RED}✗ Missing required environment variables:${NC}"
    for var in "${MISSING_VARS[@]}"; do
        echo "  - $var"
    done
    echo ""
    echo "Please edit your .env file and set these values."
    echo "See .env.example for details."
    echo ""
    exit 1
fi

echo -e "${GREEN}✓ Environment variables validated${NC}"

# Show current configuration
echo ""
echo -e "${CYAN}Current Configuration:${NC}"
echo "  Client ID: ${MS365_MCP_CLIENT_ID:0:8}..."
echo "  Tenant ID: $MS365_MCP_TENANT_ID"
echo "  Org Mode:  ${MS365_MCP_ORG_MODE:-false}"
echo "  Log Level: ${LOG_LEVEL:-info}"
echo ""

# Check if Docker is running
if ! docker info > /dev/null 2>&1; then
    echo -e "${RED}✗ Error: Docker is not running${NC}"
    echo ""
    echo "Please start Docker Desktop and try again."
    echo ""
    exit 1
fi

echo -e "${GREEN}✓ Docker is running${NC}"

# Check if image needs to be built
IMAGE_EXISTS=$(docker images -q ms365-mcp-server 2> /dev/null)

if [ -z "$IMAGE_EXISTS" ] || [ "$FORCE_BUILD" = true ]; then
    if [ "$FORCE_BUILD" = true ]; then
        echo ""
        echo -e "${YELLOW}Building Docker image (forced rebuild)...${NC}"
    else
        echo ""
        echo -e "${YELLOW}Building Docker image (first time)...${NC}"
    fi
    echo "This may take a few minutes..."
    echo ""
    
    if docker compose build; then
        echo ""
        echo -e "${GREEN}✓ Docker image built successfully${NC}"
    else
        echo ""
        echo -e "${RED}✗ Docker build failed${NC}"
        echo ""
        exit 1
    fi
else
    echo -e "${GREEN}✓ Docker image already built${NC}"
fi

# Start the containers
echo ""
echo -e "${YELLOW}Starting containers...${NC}"
echo ""

if docker compose up -d; then
    echo ""
    echo -e "${GREEN}✓ Containers started successfully${NC}"
else
    echo ""
    echo -e "${RED}✗ Failed to start containers${NC}"
    echo ""
    exit 1
fi

# Wait a moment for containers to initialize
sleep 2

# Check container status
CONTAINER_STATUS=$(docker compose ps --format json | jq -r '.[0].State' 2>/dev/null || echo "unknown")

echo ""
echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
echo -e "${GREEN}✓ MS-365 MCP Server is running!${NC}"
echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
echo ""

# Check if authenticated
echo "Next Steps:"
echo ""
echo "1. Authenticate with Microsoft 365:"
echo -e "   ${CYAN}./auth-login.sh${NC}"
echo ""
echo "2. Verify authentication:"
echo -e "   ${CYAN}./auth-verify.sh${NC}"
echo ""
echo "3. Check server health:"
echo -e "   ${CYAN}./health-check.sh${NC}"
echo ""
echo "Useful Commands:"
echo "  View logs:         docker compose logs -f"
echo "  Stop server:       docker compose down"
echo "  Restart server:    docker compose restart"
echo "  Rebuild & restart: ./start.sh --build"
echo ""

exit 0
