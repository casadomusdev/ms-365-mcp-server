#!/bin/bash
# docker-mcp-wrapper.sh
# MCP STDIO wrapper for Docker-based ms-365-mcp-server
# This script bridges Claude Desktop's STDIO communication to a Docker container

set -e

# Configuration
CONTAINER_NAME="${COMPOSE_PROJECT_NAME:-ms365-mcp}-server"
COMPOSE_FILE="$(dirname "$0")/docker-compose.yaml"
MAX_WAIT_TIME=30

# Parse arguments to pass to the MCP server
MCP_ARGS="$@"

# Function to check if container is running
is_container_running() {
    docker ps --filter "name=${CONTAINER_NAME}" --format '{{.Names}}' | grep -q "^${CONTAINER_NAME}$"
}

# Function to start the container if not running
ensure_container_running() {
    if ! is_container_running; then
        echo "Starting ms-365-mcp-server container..." >&2
        
        # Check if we're in the project directory
        if [ ! -f "$COMPOSE_FILE" ]; then
            echo "Error: docker-compose.yaml not found at $COMPOSE_FILE" >&2
            echo "Please run this script from the ms-365-mcp-server directory or set COMPOSE_FILE" >&2
            exit 1
        fi
        
        # Start container in detached mode
        cd "$(dirname "$COMPOSE_FILE")"
        docker compose up -d
        
        # Wait for container to be ready
        WAIT_TIME=0
        while [ $WAIT_TIME -lt $MAX_WAIT_TIME ]; do
            if is_container_running; then
                echo "Container started successfully" >&2
                sleep 2  # Give it a moment to fully initialize
                return 0
            fi
            sleep 1
            WAIT_TIME=$((WAIT_TIME + 1))
        done
        
        echo "Error: Container failed to start within ${MAX_WAIT_TIME} seconds" >&2
        exit 1
    fi
}

# Ensure container is running
ensure_container_running

# Execute the MCP server in the container with STDIO passthrough
# This command:
# - Uses docker exec -i for interactive STDIN
# - Runs the node process with MCP args
# - Passes through STDIO for MCP protocol communication
exec docker exec -i "${CONTAINER_NAME}" node /app/dist/index.js ${MCP_ARGS}
