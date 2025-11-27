import 'dotenv/config';
import AuthManager, { buildScopesFromEndpoints } from '../dist/auth.js';

async function extractToken() {
  try {
    const scopes = buildScopesFromEndpoints(true);
    const authManager = new AuthManager(undefined, scopes);
    await authManager.loadTokenCache();
    
    const token = await authManager.getToken();
    
    if (token) {
      console.log(token);
    } else {
      console.error('ERROR: No token available');
      process.exit(1);
    }
  } catch (error) {
    console.error('ERROR:', error.message);
    process.exit(1);
  }
}

extractToken();
