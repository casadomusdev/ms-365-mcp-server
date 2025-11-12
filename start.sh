#!/bin/bash
#
# MS-365 MCP Server - Start Script
#
# Sets up and starts the MS-365 MCP server in Docker or local mode.
# Validates environment configuration before starting.
#
# Usage:
#   ./start.sh [--build] [--docker|--local]
#
# Options:
#   --build   Force rebuild of Docker image
#   --docker  Run in Docker mode
#   --local   Run locally without Docker
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
RUN_MODE=""

for arg in "$@"; do
    case $arg in
        --build)
            FORCE_BUILD=true
            shift
            ;;
        --docker)
            RUN_MODE="docker"
            shift
            ;;
        --local)
            RUN_MODE="local"
            shift
            ;;
        *)
            echo "Unknown option: $arg"
            echo "Usage: ./start.sh [--build] [--docker|--local]"
            echo ""
            echo "Options:"
            echo "  --build   Force rebuild of Docker image"
            echo "  --docker  Run in Docker mode"
            echo "  --local   Run locally without Docker"
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

if [ ${#MISSING_VARS[@]} -ne 0 ]; then
    echo -e "${RED}✗ Missing required environment variables:${NC}"
    for var in "${MISSING_VARS[@]}"; do
        echo "  - $var"
    done
    echo ""
    echo "Please edit your .env file and set these values."
    exit 1
fi

echo -e "${GREEN}✓ Environment variables validated${NC}"

# Show configuration
echo ""
echo -e "${CYAN}Current Configuration:${NC}"
echo "  Client ID: ${MS365_MCP_CLIENT_ID:0:8}..."
echo "  Tenant ID: $MS365_MCP_TENANT_ID"
echo ""

# Check Docker availability
DOCKER_AVAILABLE=false
if command -v docker >/dev/null 2>&1 && docker info >/dev/null 2>&1; then
    DOCKER_AVAILABLE=true
fi

# Determine run mode if not specified
if [ -z "$RUN_MODE" ]; then
    if [ "$DOCKER_AVAILABLE" = true ]; then
        echo -e "${CYAN}Select run mode:${NC}"
        echo "  1. Docker (recommended, isolated environment)"
        echo "  2. Local (direct Node.js, no Docker needed)"
        echo ""
        read -p "Choose mode (1 or 2) [1]: " MODE_CHOICE
        MODE_CHOICE=${MODE_CHOICE:-1}
        
        if [ "$MODE_CHOICE" = "2" ]; then
            RUN_MODE="local"
        else
            RUN_MODE="docker"
        fi
    else
        echo -e "${YELLOW}⚠  Docker not available - using local mode${NC}"
        RUN_MODE="local"
    fi
fi

# Validate mode selection
if [ "$RUN_MODE" = "docker" ] && [ "$DOCKER_AVAILABLE" = false ]; then
    echo -e "${RED}✗ Docker mode selected but Docker is not available${NC}"
    echo ""
    echo "Please either:"
    echo "  - Start Docker Desktop"
    echo "  - Run in local mode: ./start.sh --local"
    echo ""
    exit 1
fi

echo ""
echo -e "${GREEN}✓ Run mode: $RUN_MODE${NC}"
echo ""

# Docker mode
if [ "$RUN_MODE" = "docker" ]; then
    # Get compose project name
    COMPOSE_PROJECT=${COMPOSE_PROJECT_NAME:-ms365-mcp}
    IMAGE_NAME="${COMPOSE_PROJECT}-ms365-mcp"
    
    # Check if image exists
    IMAGE_EXISTS=$(docker images -q "$IMAGE_NAME" 2>/dev/null)
    
    if [ -z "$IMAGE_EXISTS" ] || [ "$FORCE_BUILD" = true ]; then
        if [ "$FORCE_BUILD" = true ]; then
            echo -e "${YELLOW}Building Docker image (forced rebuild)...${NC}"
        else
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
            exit 1
        fi
    else
        echo -e "${GREEN}✓ Docker image exists ($IMAGE_NAME)${NC}"
    fi
    
    # Start containers
    echo ""
    echo -e "${YELLOW}Starting containers...${NC}"
    echo ""
    
    if docker compose up -d; then
        echo ""
        echo -e "${GREEN}✓ Containers started successfully${NC}"
    else
        echo ""
        echo -e "${RED}✗ Failed to start containers${NC}"
        exit 1
    fi
    
    # Wait for initialization
    sleep 2
    
    echo ""
    echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
    echo -e "${GREEN}✓ MS-365 MCP Server is running (Docker mode)${NC}"
    echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
    echo ""
    echo "Next Steps:"
    echo "  1. Authenticate:     ./auth-login.sh"
    echo "  2. Verify:           ./auth-verify.sh"
    echo "  3. Health check:     ./health-check.sh"
    echo ""
    echo "Useful Commands:"
    echo "  View logs:           docker compose logs -f"
    echo "  Stop server:         docker compose down"
    echo "  Restart:             docker compose restart"
    echo "  Rebuild:             ./start.sh --build"
    echo ""

# Local mode
else
    echo -e "${YELLOW}Setting up local environment...${NC}"
    echo ""
    
    # Check for Node.js
    if ! command -v node >/dev/null 2>&1; then
        echo -e "${RED}✗ Node.js not found${NC}"
        echo "Please install Node.js to run in local mode"
        exit 1
    fi
    
    echo -e "${GREEN}✓ Node.js found: $(node --version)${NC}"
    
    # Install dependencies if needed
    if [ ! -d "node_modules" ]; then
        echo ""
        echo "Installing dependencies..."
        npm install
        echo ""
    fi
    
    # Build if needed
    if [ ! -d "dist" ] || [ "$FORCE_BUILD" = true ]; then
        echo "Building project..."
        npm run build
        echo ""
    fi
    
    echo -e "${GREEN}✓ Project built${NC}"
    echo ""
    echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
    echo -e "${GREEN}✓ Starting MS-365 MCP Server (Local mode)${NC}"
    echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
    echo ""
    echo "Next Steps:"
    echo "  1. Authenticate:     ./auth-login.sh"
    echo "  2. Verify:           ./auth-verify.sh"
    echo "  3. Health check:     ./health-check.sh"
    echo ""
    echo "The server will start now..."
    echo "Press Ctrl+C to stop"
    echo ""
    
    # Start the server
    node dist/index.js
fi

exit 0
