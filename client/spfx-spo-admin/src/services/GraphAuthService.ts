import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Client } from '@microsoft/microsoft-graph-client';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class GraphAuthService {
  /**
   * Get Microsoft Graph client using user mode (SPFx context)
   * @param context - SPFx WebPart context
   * @returns Promise<Client>
   */
  public static async getUserModeClient(context: WebPartContext): Promise<Client> {
    try {
      const msGraphClient: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
      
      // Return the Graph client directly
      return msGraphClient as unknown as Client;
    } catch (error) {
      console.error('Error creating user mode Graph client:', error);
      throw new Error('Failed to initialize Graph client in user mode');
    }
  }

  /**
   * Get Microsoft Graph client using app mode (client credentials)
   * Note: This requires server-side implementation for security reasons
   * Client credentials should never be exposed in client-side code
   * @param appId - Application (client) ID
   * @param secret - Client secret
   * @param tenantId - Tenant ID
   * @returns Promise<Client>
   */
  public static async getAppModeClient(appId: string, secret: string, tenantId: string): Promise<Client> {
    // WARNING: This is a placeholder implementation
    // In production, client credentials flow should be handled server-side
    // This implementation is for demonstration purposes only
    
    throw new Error(
      'App mode authentication must be implemented server-side. ' +
      'Client credentials should never be exposed in client-side code. ' +
      'Consider using Azure Functions or a secure backend service.'
    );

    // For reference, server-side implementation would use:
    // const credential = new ClientSecretCredential(tenantId, appId, secret);
    // const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    //   scopes: ['https://graph.microsoft.com/.default']
    // });
    // return Client.initWithMiddleware({ authProvider });
  }
}
