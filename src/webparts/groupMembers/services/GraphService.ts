import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUser, IGroup } from '../types/interfaces';

// Type for Microsoft Graph Client
interface IMSGraphClient {
  api(path: string): IMSGraphClientRequest;
}

interface IMSGraphClientRequest {
  select(properties: string): IMSGraphClientRequest;
  get(): Promise<{ value?: unknown[] } & Record<string, unknown>>;
}

export interface IGraphService {
  getGroupMembers(groupId: string, role: 'admin' | 'member'): Promise<IUser[]>;
  getUserGroups(): Promise<IGroup[]>;
  getUserPresence(userId: string): Promise<unknown>;
  getUserPhoto(userId: string): Promise<string | undefined>;
}

export class GraphService implements IGraphService {
  private context: WebPartContext;
  private graphClient: IMSGraphClient | undefined;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  private async getGraphClient(): Promise<IMSGraphClient> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.graphClient;
  }

  public async getUserGroups(): Promise<IGroup[]> {
    try {
      const client = await this.getGraphClient();
      const response = await client.api('/me/memberOf').get();
      
      return (response.value || []).filter(
        (group: unknown) => (group as Record<string, unknown>)['@odata.type'] === '#microsoft.graph.group'
      ).map((group: unknown) => {
        const g = group as Record<string, unknown>;
        return {
          id: g.id as string,
          displayName: g.displayName as string,
          '@odata.type': g['@odata.type'] as string,
          description: g.description as string
        };
      });
    } catch (error) {
      console.error('Error fetching user groups:', error);
      throw new Error('Failed to fetch user groups');
    }
  }

  public async getGroupMembers(groupId: string, role: 'admin' | 'member'): Promise<IUser[]> {
    const cacheKey = `groupUsers_${role}_${groupId}`;
    const cachedData = sessionStorage.getItem(cacheKey);
    
    if (cachedData) {
      const parsed = JSON.parse(cachedData);
      // Check if cache is still valid (30 minutes)
      if (Date.now() - parsed.timestamp < 30 * 60 * 1000) {
        return parsed.data;
      }
      sessionStorage.removeItem(cacheKey);
    }

    try {
      const client = await this.getGraphClient();
      const endpoint = role === 'admin' ? `/groups/${groupId}/owners` : `/groups/${groupId}/members`;
      
      const response = await client
        .api(endpoint)
        .select('id,displayName,jobTitle,mail,userPrincipalName,department,officeLocation')
        .get();

      const users = response?.value || [];
      
      // Cache the result with timestamp
      sessionStorage.setItem(cacheKey, JSON.stringify({
        data: users,
        timestamp: Date.now()
      }));

      return users.map((user: unknown) => {
        const u = user as Record<string, unknown>;
        return {
          id: u.id as string,
          displayName: u.displayName as string,
          jobTitle: u.jobTitle as string,
          mail: u.mail as string,
          userPrincipalName: u.userPrincipalName as string,
          department: u.department as string,
          officeLocation: u.officeLocation as string
        };
      });
    } catch (error) {
      console.error(`Error fetching ${role}s for group ${groupId}:`, error);
      return [];
    }
  }

  public async getUserPresence(userId: string): Promise<unknown> {
    try {
      const client = await this.getGraphClient();
      const response = await client.api(`/users/${userId}/presence`).get();
      return response;
    } catch (error) {
      console.warn(`Could not fetch presence for user ${userId}:`, error);
      return undefined;
    }
  }

  public async getUserPhoto(userId: string): Promise<string | undefined> {
    // Check cache first
    const cacheKey = `profilePhoto_${userId}`;
    const cachedPhoto = sessionStorage.getItem(cacheKey);
    if (cachedPhoto) {
      return cachedPhoto;
    }

    // Check if we already know this user has no photo
    if (sessionStorage.getItem(`${cacheKey}_noPhoto`)) {
      return undefined;
    }

    try {
      const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      const token = await tokenProvider.getToken("https://graph.microsoft.com");
      
      const url = `https://graph.microsoft.com/v1.0/users/${userId}/photo/$value`;
      
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 8000);

      const response = await fetch(url, {
        headers: { 
          Authorization: `Bearer ${token}`,
          'Accept': 'image/*'
        },
        signal: controller.signal
      });

      clearTimeout(timeoutId);

      if (!response.ok) {
        if (response.status === 404) {
          // User has no profile photo, cache this result
          sessionStorage.setItem(`${cacheKey}_noPhoto`, 'true');
        }
        return undefined;
      }

      const buffer = await response.arrayBuffer();
      const contentType = response.headers.get("content-type") || "image/jpeg";
      const blob = new Blob([buffer], { type: contentType });
      const objectUrl = URL.createObjectURL(blob);

      // Cache the photo URL
      sessionStorage.setItem(cacheKey, objectUrl);
      
      // Set cleanup timer for object URL
      setTimeout(() => {
        URL.revokeObjectURL(objectUrl);
        sessionStorage.removeItem(cacheKey);
      }, 10 * 60 * 1000); // 10 minutes

      return objectUrl;
    } catch (error) {
      console.warn(`Failed to load profile photo for user ${userId}:`, error);
      return undefined;
    }
  }
}