#!/usr/bin/env node
/**
 * JWT Token Decoder for Microsoft 365 MCP Server
 * Decodes Graph and Exchange tokens to inspect permissions
 */

import 'dotenv/config';
import AuthManager, { buildScopesFromEndpoints } from './dist/auth.js';

function decodeJWT(token) {
  try {
    // JWT has 3 parts separated by dots: header.payload.signature
    const parts = token.split('.');
    if (parts.length !== 3) {
      throw new Error('Invalid JWT format');
    }

    // Decode the payload (second part)
    const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString());
    return payload;
  } catch (error) {
    throw new Error(`Failed to decode JWT: ${error.message}`);
  }
}

function formatTimestamp(timestamp) {
  if (!timestamp) return 'N/A';
  const date = new Date(timestamp * 1000);
  return date.toLocaleString('en-US', { 
    timeZone: 'Europe/Berlin',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  });
}

function analyzeToken(tokenType, token) {
  console.log(`\n${'='.repeat(70)}`);
  console.log(`${tokenType} TOKEN ANALYSIS`);
  console.log('='.repeat(70));

  if (!token) {
    console.log('❌ No token available\n');
    return false;
  }

  try {
    const decoded = decodeJWT(token);
    
    // Basic token info
    console.log('\n📋 Basic Information:');
    console.log(`  Audience (aud): ${decoded.aud || 'N/A'}`);
    console.log(`  Issuer (iss): ${decoded.iss || 'N/A'}`);
    console.log(`  App ID (appid): ${decoded.appid || 'N/A'}`);
    console.log(`  Tenant ID (tid): ${decoded.tid || 'N/A'}`);
    
    // Timestamps
    console.log('\n⏰ Token Validity:');
    console.log(`  Issued At (iat): ${formatTimestamp(decoded.iat)}`);
    console.log(`  Not Before (nbf): ${formatTimestamp(decoded.nbf)}`);
    console.log(`  Expires (exp): ${formatTimestamp(decoded.exp)}`);
    
    const now = Math.floor(Date.now() / 1000);
    const timeLeft = decoded.exp - now;
    const minutesLeft = Math.floor(timeLeft / 60);
    
    if (timeLeft > 0) {
      console.log(`  ✅ Valid for ${minutesLeft} more minutes`);
    } else {
      console.log(`  ❌ EXPIRED ${Math.abs(minutesLeft)} minutes ago`);
    }
    
    // Permissions/Roles
    console.log('\n🔐 Permissions (roles claim):');
    if (decoded.roles && Array.isArray(decoded.roles)) {
      if (decoded.roles.length === 0) {
        console.log('  ⚠️  NO ROLES FOUND - Token has no application permissions!');
      } else {
        decoded.roles.forEach(role => {
          console.log(`  ✓ ${role}`);
        });
      }
    } else {
      console.log('  ⚠️  No roles claim found (might be delegated permissions)');
    }
    
    // Scopes (delegated permissions)
    if (decoded.scp) {
      console.log('\n📝 Scopes (delegated permissions):');
      const scopes = decoded.scp.split(' ');
      scopes.forEach(scope => {
        console.log(`  - ${scope}`);
      });
    }
    
    // Check for Exchange-specific permissions
    console.log('\n🔍 Exchange Permission Check:');
    if (decoded.aud === 'https://outlook.office365.com' || decoded.aud === 'https://outlook.office.com') {
      console.log('  ✅ Token audience is correct for Exchange Online');
      
      if (decoded.roles && decoded.roles.some(r => r.includes('Exchange'))) {
        console.log('  ✅ Has Exchange permissions');
      } else {
        console.log('  ❌ NO Exchange permissions found in roles!');
        console.log('  💡 Add Exchange.ManageAsApp in Azure AD app permissions');
      }
    } else {
      console.log(`  ❌ Wrong audience for Exchange: ${decoded.aud}`);
      console.log('  💡 This token is for a different API');
    }
    
    return true;
  } catch (error) {
    console.log(`\n❌ Error decoding token: ${error.message}\n`);
    return false;
  }
}

async function main() {
  try {
    console.log('Microsoft 365 MCP Server - Token Decoder');
    console.log('=========================================\n');
    
    const scopes = buildScopesFromEndpoints(true);
    const authManager = new AuthManager(undefined, scopes);
    await authManager.loadTokenCache();
    
    // Get both tokens
    console.log('Extracting tokens...\n');
    
    const graphToken = await authManager.getToken();
    const exchangeToken = await authManager.getExchangeToken();
    
    // Analyze Graph token
    analyzeToken('GRAPH API', graphToken);
    
    // Analyze Exchange token
    analyzeToken('EXCHANGE ONLINE', exchangeToken);
    
    console.log('\n' + '='.repeat(70));
    console.log('TOKEN ANALYSIS COMPLETE');
    console.log('='.repeat(70) + '\n');
    
  } catch (error) {
    console.error(`\n❌ Error: ${error.message}\n`);
    process.exit(1);
  }
}

main();
