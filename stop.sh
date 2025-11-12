#!/bin/bash
#
# MS-365 MCP Server - Stop Script
#
# Stops the MCP server (Docker or local process).
# Attempts Docker stop first, then checks for local Node.js process.
#
# Usage:
#   ./stop.sh [--force]
#
# Options:
#   --force    Force stop (kill -9 for local, docker compose down for Docker)
#
# Exit codes:
#   0 - Server stopped successfully
#   1 - Failed to stop or nothing running

set -e

# Color codes
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# Parse arguments
FORCE_STOP=false
for arg in "$@"; do
    case $arg in
        --force)
            FORCE_STOP=true
            shift
            ;;
        *)
            echo "Unknown option: $arg"
            echo "Usage: ./stop.sh [--force]"
            exit 1
            ;;
    esac
done

echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
echo -e "${CYAN}  MS-365 MCP Server - Stop${NC}"
echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
echo ""

# Load .env if exists for COMPOSE_PROJECT_NAME
if [ -f .env ]; then
    set -a
    source .env
    set +a
fi

STOPPED_SOMETHING=false

# Try Docker first
if command -v docker >/dev/null 2>&1 && docker info >/dev/null 2>&1; then
    COMPOSE_PROJECT=${COMPOSE_PROJECT_NAME:-ms365-mcp}
    
    # Check if docker compose has any containers (running or stopped)
    CONTAINER_COUNT=$(docker compose ps -a --format json 2>/dev/null | jq -s 'length' 2>/dev/null || echo "0")
    
    if [ "$CONTAINER_COUNT" -gt 0 ]; then
        echo "Found Docker containers ($CONTAINER_COUNT)..."
        echo ""
        
        if [ "$FORCE_STOP" = true ]; then
            echo -e "${YELLOW}Force stopping containers (docker compose down)...${NC}"
            if docker compose down; then
                echo ""
                echo -e "${GREEN}✓ Docker containers stopped and removed${NC}"
                STOPPED_SOMETHING=true
            else
                echo -e "${RED}✗ Failed to stop Docker containers${NC}"
                exit 1
            fi
        else
            echo "Stopping Docker containers..."
            if docker compose stop; then
                echo ""
                echo -e "${GREEN}✓ Docker containers stopped${NC}"
                echo ""
                echo "Containers are stopped but not removed."
                echo "To remove: docker compose down"
                echo "To restart: docker compose start"
                STOPPED_SOMETHING=true
            else
                echo -e "${RED}✗ Failed to stop Docker containers${NC}"
                exit 1
            fi
        fi
    fi
fi

# Check for local Node.js process
if [ "$STOPPED_SOMETHING" = false ]; then
    echo "Checking for local Node.js process..."
    echo ""
    
    # Find node process running dist/index.js
    PID=$(ps aux | grep "node.*dist/index.js" | grep -v grep | awk '{print $2}' | head -1)
    
    if [ -n "$PID" ]; then
        echo "Found Node.js process (PID: $PID)"
        echo ""
        
        if [ "$FORCE_STOP" = true ]; then
            echo -e "${YELLOW}Force stopping Node.js process...${NC}"
            if kill -9 $PID 2>/dev/null; then
                echo ""
                echo -e "${GREEN}✓ Node.js process killed${NC}"
                STOPPED_SOMETHING=true
            else
                echo -e "${RED}✗ Failed to kill process${NC}"
                exit 1
            fi
        else
            echo "Stopping Node.js process gracefully..."
            if kill -TERM $PID 2>/dev/null; then
                # Wait a moment for graceful shutdown
                sleep 2
                
                # Check if still running
                if ps -p $PID >/dev/null 2>&1; then
                    echo -e "${YELLOW}Process still running, sending SIGKILL...${NC}"
                    kill -9 $PID 2>/dev/null || true
                fi
                
                echo ""
                echo -e "${GREEN}✓ Node.js process stopped${NC}"
                STOPPED_SOMETHING=true
            else
                echo -e "${RED}✗ Failed to stop process${NC}"
                exit 1
            fi
        fi
    fi
fi

# Final status
echo ""
if [ "$STOPPED_SOMETHING" = true ]; then
    echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
    echo -e "${GREEN}✓ MS-365 MCP Server stopped${NC}"
    echo -e "${CYAN}════════════════════════════════════════════════════════════════${NC}"
    echo ""
    exit 0
else
    echo -e "${YELLOW}⚠  No running server found${NC}"
    echo ""
    echo "Neither Docker containers nor local Node.js process detected."
    echo ""
    exit 1
fi
