import { Client } from '@microsoft/microsoft-graph-client';

export class SiteUrlHelper {
  /**
   * Get SharePoint site ID from site URL using Microsoft Graph API
   * @param graphClient - Initialized Microsoft Graph client
   * @param siteUrl - Full SharePoint site URL (e.g., https://contoso.sharepoint.com/sites/sitename)
   * @returns Promise<string> - Site ID in format: hostname,siteId,webId
   */
  public static async getSiteIdFromUrl(graphClient: Client, siteUrl: string): Promise<string> {
    try {
      // Validate URL format
      if (!siteUrl || !siteUrl.startsWith('http')) {
        throw new Error('Invalid site URL format. Must be a valid https:// URL');
      }

      const url = new URL(siteUrl);
      const hostname = url.hostname;
      const serverRelativePath = url.pathname;

      // Remove trailing slash if present
      const path = serverRelativePath.endsWith('/') 
        ? serverRelativePath.slice(0, -1) 
        : serverRelativePath;

      // Call Graph API to get site ID
      // Format: GET /sites/{hostname}:/{serverRelativePath}
      const endpoint = `/sites/${hostname}:${path}`;
      
      const response = await graphClient
        .api(endpoint)
        .get();

      if (!response || !response.id) {
        throw new Error('Unable to retrieve site ID from the provided URL');
      }

      return response.id;
    } catch (error) {
      console.error('Error resolving site ID from URL:', error);
      
      if (error instanceof Error) {
        if (error.message.includes('404')) {
          throw new Error('Site not found. Please verify the URL is correct and you have access to the site.');
        } else if (error.message.includes('401') || error.message.includes('403')) {
          throw new Error('Access denied. You do not have permission to access this site.');
        }
        throw error;
      }
      
      throw new Error('Failed to resolve site ID from URL');
    }
  }

  /**
   * Validate SharePoint URL format
   * @param url - URL to validate
   * @returns boolean
   */
  public static isValidSharePointUrl(url: string): boolean {
    try {
      const parsedUrl = new URL(url);
      return parsedUrl.hostname.includes('sharepoint.com') && 
             (parsedUrl.protocol === 'https:');
    } catch {
      return false;
    }
  }
}
