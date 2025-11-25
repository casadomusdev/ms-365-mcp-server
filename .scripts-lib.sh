#!/bin/bash
#
# Shared Library for MS-365 MCP Server Scripts
#
# Provides common functions for environment detection and Docker operations
#
# Usage:
#   source .scripts-lib.sh
#   detect_execution_mode
#   if [ "$EXECUTION_MODE" = "docker-delegate" ]; then ...

# Color codes
export RED='\033[0;31m'
export GREEN='\033[0;32m'
export YELLOW='\033[1;33m'
export CYAN='\033[0;36m'
export NC='\033[0m' # No Color

# Detect execution mode
# Sets EXECUTION_MODE to either:
#   - "direct" = run Node.js directly (inside container or local machine)
#   - "docker-delegate" = delegate to Docker container (only if container is running)
detect_execution_mode() {
    # Check if we're inside a Docker container
    if [ -f /.dockerenv ] || grep -q docker /proc/1/cgroup 2>/dev/null; then
        EXECUTION_MODE="direct"
        return
    fi
    
    # Check if Docker container is RUNNING (not just if Docker is available)
    if command -v docker >/dev/null 2>&1; then
        if docker compose ps ms365-mcp 2>/dev/null | grep -q "Up"; then
            EXECUTION_MODE="docker-delegate"
            return
        fi
    fi
    
    # Default: run directly with local Node.js
    EXECUTION_MODE="direct"
}

# Get container name based on COMPOSE_PROJECT_NAME
get_container_name() {
    if [ -f .env ]; then
        export $(grep -v '^#' .env | grep COMPOSE_PROJECT_NAME | xargs)
    fi
    echo "${COMPOSE_PROJECT_NAME:-ms365-mcp}"
}

# Execute in appropriate mode
# Usage: exec_node_command "node dist/index.js --login"
exec_node_command() {
    local cmd="$1"
    
    detect_execution_mode
    
    if [ "$EXECUTION_MODE" = "docker-delegate" ]; then
        # Delegate to Docker container
        exec docker compose exec ms365-mcp bash -c "cd /app && $cmd"
    else
        # Run directly
        eval "$cmd"
    fi
}

# Check if tokens exist (for validation)
check_tokens_exist() {
    if [ -f .token-cache.json ] || [ -f /app/data/.token-cache.json ]; then
        return 0
    else
        return 1
    fi
}
