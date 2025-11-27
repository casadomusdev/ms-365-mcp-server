# Build stage
FROM node:22-bookworm-slim AS builder

WORKDIR /app

# Copy package files
COPY package.json package-lock.json ./

# Install dependencies
RUN npm ci

# Copy source code
COPY . .

# Generate API client code
RUN npm run generate

# Build the project
RUN npm run build

# Production stage
FROM node:22-bookworm-slim

WORKDIR /app

# Install dependencies including PowerShell Core
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    dumb-init \
    curl \
    iputils-ping \
    libsecret-1-0 \
    jq \
    wget \
    apt-transport-https \
    software-properties-common && \
    # Install PowerShell Core
    wget -q https://packages.microsoft.com/config/debian/12/packages-microsoft-prod.deb && \
    dpkg -i packages-microsoft-prod.deb && \
    rm packages-microsoft-prod.deb && \
    apt-get update && \
    apt-get install -y powershell && \
    # Install Exchange Online PowerShell module
    pwsh -Command "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted" && \
    pwsh -Command "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope AllUsers" && \
    # Cleanup
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Copy package files
COPY package.json package-lock.json ./

# Install production dependencies only
RUN npm ci --production

# Copy built application from builder
COPY --from=builder /app/dist ./dist

# Copy health check script
COPY --chmod=755 *.sh ./

# Create directory for token cache with proper permissions
RUN mkdir -p /app/data && \
    chown -R node:node /app

# Switch to non-root user
USER node

# Set environment variable for token cache location
ENV TOKEN_CACHE_DIR=/app/data

# Use dumb-init to handle signals properly
ENTRYPOINT ["/usr/bin/dumb-init", "--"]

# Run the MCP server
CMD ["node", "dist/index.js"]
