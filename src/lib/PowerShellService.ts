import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import path from 'path';
import { execSync } from 'child_process';
import type AuthManager from '../auth.js';
import logger from '../logger.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * Represents a mailbox permission result from PowerShell
 */
export interface MailboxPermission {
  id: string;
  email: string;
  displayName: string;
  permissions: string[];
}

/**
 * Service for executing PowerShell scripts to query Exchange Online
 * 
 * This service provides integration with Exchange Online PowerShell to query
 * mailbox delegation permissions (Full Access, SendAs) that are not available
 * through Microsoft Graph API.
 * 
 * Authentication is handled by reusing the access token from AuthManager,
 * eliminating the need for certificate-based authentication.
 */
export class PowerShellService {
  private readonly timeout: number;
  private readonly enabled: boolean;
  private readonly available: boolean;

  /**
   * Creates a new PowerShellService instance
   * 
   * @param authManager - The AuthManager instance to get access tokens from
   */
  constructor(private readonly authManager: AuthManager) {
    this.timeout = parseInt(process.env.MS365_POWERSHELL_TIMEOUT || '30000', 10);
    
    // Default to enabled unless explicitly set to false
    const envValue = process.env.MS365_POWERSHELL_ENABLED;
    this.enabled = envValue !== 'false' && envValue !== '0';
    
    // Auto-detect pwsh availability
    this.available = this.checkPowerShellAvailability();
    
    if (this.enabled) {
      if (this.available) {
        logger.info('PowerShell integration enabled and available');
        logger.debug(`PowerShell timeout: ${this.timeout}ms`);
      } else {
        logger.warn('PowerShell integration enabled but pwsh is not available on this system');
        logger.warn('Shared mailbox discovery will be disabled - only personal mailboxes will be accessible');
        logger.info('To enable shared mailbox discovery, install PowerShell Core 7.x and Exchange Online PowerShell module');
      }
    } else {
      logger.info('PowerShell integration explicitly disabled via MS365_POWERSHELL_ENABLED=false');
    }
  }

  /**
   * Checks if PowerShell (pwsh) is available on the system
   * 
   * @returns true if pwsh command is available, false otherwise
   */
  private checkPowerShellAvailability(): boolean {
    try {
      execSync('pwsh -Version', { stdio: 'pipe' });
      return true;
    } catch (error) {
      return false;
    }
  }

  /**
   * Checks if PowerShell integration is both enabled and available
   * 
   * @returns true if PowerShell is enabled in config and pwsh is available on system
   */
  isEnabled(): boolean {
    return this.enabled && this.available;
  }

  /**
   * Executes a PowerShell script with the provided arguments
   * 
   * @param scriptPath - Absolute path to the PowerShell script
   * @param args - Arguments to pass to the script as key-value pairs
   * @returns The parsed JSON output from the PowerShell script
   * @throws Error if PowerShell execution fails or times out
   */
  async execute(scriptPath: string, args: Record<string, any>): Promise<any> {
    if (!this.enabled) {
      throw new Error('PowerShell integration is disabled. Remove MS365_POWERSHELL_ENABLED=false to enable.');
    }

    if (!this.available) {
      throw new Error('PowerShell (pwsh) is not available on this system. Install PowerShell Core 7.x to enable shared mailbox discovery.');
    }

    logger.debug(`Executing PowerShell script: ${scriptPath}`);
    logger.debug(`Script arguments: ${JSON.stringify(Object.keys(args))}`);

    // Build PowerShell command arguments
    const psArgs: string[] = ['-NoProfile', '-NonInteractive', '-File', scriptPath];
    
    // Add script parameters
    for (const [key, value] of Object.entries(args)) {
      psArgs.push(`-${key}`, String(value));
    }

    return new Promise((resolve, reject) => {
      const process = spawn('pwsh', psArgs);
      
      let stdout = '';
      let stderr = '';
      let timedOut = false;

      // Set timeout
      const timeoutId = setTimeout(() => {
        timedOut = true;
        process.kill('SIGTERM');
        reject(new Error(`PowerShell script execution timed out after ${this.timeout}ms`));
      }, this.timeout);

      // Collect stdout
      process.stdout.on('data', (data) => {
        stdout += data.toString();
      });

      // Collect stderr
      process.stderr.on('data', (data) => {
        stderr += data.toString();
      });

      // Handle process completion
      process.on('close', (code) => {
        clearTimeout(timeoutId);

        if (timedOut) {
          return; // Already rejected
        }

        if (code !== 0) {
          logger.error(`PowerShell script failed with exit code ${code}`);
          logger.error(`STDERR: ${stderr}`);
          logger.debug(`STDOUT: ${stdout}`);
          reject(new Error(`PowerShell script failed with exit code ${code}: ${stderr || 'No error message'}`));
          return;
        }

        // Log stderr as warnings (PowerShell may output non-error info to stderr)
        if (stderr) {
          logger.warn(`PowerShell stderr output: ${stderr}`);
        }

        // Parse JSON output
        try {
          logger.debug(`PowerShell stdout length: ${stdout.length} characters`);
          const result = JSON.parse(stdout);
          logger.debug(`PowerShell script completed successfully`);
          resolve(result);
        } catch (parseError) {
          logger.error(`Failed to parse PowerShell output as JSON`);
          logger.error(`Output: ${stdout.substring(0, 500)}...`);
          reject(new Error(`Failed to parse PowerShell output: ${(parseError as Error).message}`));
        }
      });

      // Handle process errors
      process.on('error', (error) => {
        clearTimeout(timeoutId);
        logger.error(`PowerShell process error: ${error.message}`);
        reject(new Error(`Failed to spawn PowerShell process: ${error.message}`));
      });
    });
  }

  /**
   * Checks mailbox permissions for a specific user using Exchange Online PowerShell
   * 
   * This method uses certificate-based authentication for app-only access.
   * Certificate must be configured via MS365_CERT_PATH and MS365_CERT_PASSWORD.
   * 
   * @param userEmail - The email address of the user to check permissions for
   * @returns Array of mailboxes the user has access to with their permissions
   * @throws Error if PowerShell is not enabled, certificate is missing, or execution fails
   */
  async checkPermissions(userEmail: string): Promise<MailboxPermission[]> {
    if (!this.enabled) {
      throw new Error('PowerShell integration is disabled. Remove MS365_POWERSHELL_ENABLED=false to enable.');
    }

    if (!this.available) {
      throw new Error('PowerShell (pwsh) is not available on this system. Install PowerShell Core 7.x to enable shared mailbox discovery.');
    }

    logger.info(`Checking mailbox permissions for user: ${userEmail}`);

    try {
      // Get required configuration
      const clientId = process.env.MS365_MCP_CLIENT_ID;
      const organization = process.env.MS365_MCP_ORGANIZATION;

      if (!clientId || !organization) {
        throw new Error(
          'MS365_MCP_CLIENT_ID and MS365_MCP_ORGANIZATION must be configured. ' +
          'MS365_MCP_ORGANIZATION should be your tenant domain (e.g., contoso.onmicrosoft.com), not the tenant ID GUID.'
        );
      }

      // Use standard certificate paths (convention over configuration)
      const certPath = path.join(__dirname, '..', '..', 'certs', 'ms365-powershell.pfx');
      const certPasswordFile = path.join(__dirname, '..', '..', 'certs', '.cert-password.txt');

      // Check if certificate exists
      const fs = await import('fs/promises');
      try {
        await fs.access(certPath);
      } catch (error) {
        throw new Error(
          `PowerShell certificate not found at: ${certPath}\n` +
          'Run ./auth-generate-cert.sh to create a certificate, ' +
          'then upload the .cer file to Azure AD.'
        );
      }

      // Read certificate password from backup file
      let certPassword: string;
      try {
        certPassword = (await fs.readFile(certPasswordFile, 'utf-8')).trim();
      } catch (error) {
        throw new Error(
          `Certificate password file not found at: ${certPasswordFile}\n` +
          'This file should have been created by ./auth-generate-cert.sh.\n' +
          'You may need to regenerate the certificate.'
        );
      }

      logger.debug(`Using certificate: ${certPath}`);
      logger.debug(`App ID: ${clientId}`);
      logger.debug(`Organization: ${organization}`);

      // Construct path to PowerShell script
      const scriptPath = path.join(__dirname, '..', '..', 'scripts', 'check-mailbox-permissions.ps1');
      logger.debug(`Script path: ${scriptPath}`);

      // Execute PowerShell script with certificate authentication
      const result = await this.execute(scriptPath, {
        UserEmail: userEmail,
        CertificatePath: certPath,
        CertificatePassword: certPassword,
        AppId: clientId,
        Organization: organization
      });

      // Validate and return result
      if (!Array.isArray(result)) {
        throw new Error('PowerShell script did not return an array');
      }

      logger.info(`Found ${result.length} shared mailbox(es) for user ${userEmail}`);
      return result as MailboxPermission[];

    } catch (error) {
      logger.error(`Failed to check mailbox permissions: ${(error as Error).message}`);
      throw error;
    }
  }
}

export default PowerShellService;
