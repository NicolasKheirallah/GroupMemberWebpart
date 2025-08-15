import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserPresence } from '../types/interfaces';
import { CacheService } from './CacheService';

// Type for Microsoft Graph Client
interface IMSGraphClient {
  api(path: string): IMSGraphClientRequest;
}

interface IMSGraphClientRequest {
  select(properties: string): IMSGraphClientRequest;
  get(): Promise<{ value?: unknown[] } & Record<string, unknown>>;
}

export interface IProfileService {
  getUserPhoto(userId: string): Promise<string | undefined>;
  getUserPresence(userId: string): Promise<IUserPresence | undefined>;
  getBatchUserPresence(userIds: string[]): Promise<Record<string, IUserPresence>>;
}

export class ProfileService implements IProfileService {
  private context: WebPartContext;
  private graphClient: IMSGraphClient | undefined;
  private readonly RATE_LIMIT_DELAY = 100; // ms between requests
  private cacheService: CacheService;

  constructor(context: WebPartContext) {
    this.context = context;
    this.cacheService = CacheService.getInstance();
  }

  private async getGraphClient(): Promise<IMSGraphClient> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.graphClient;
  }

  public async getUserPhoto(userId: string): Promise<string | undefined> {
    // Check LRU cache first
    const cachedPhoto = this.cacheService.getUserPhoto(userId);
    if (cachedPhoto) {
      return cachedPhoto;
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
          // User has no profile photo, cache this negative result
          this.cacheService.setUserPhoto(userId, 'NO_PHOTO');
        }
        return undefined;
      }

      const buffer = await response.arrayBuffer();
      const contentType = response.headers.get("content-type") || "image/jpeg";
      const blob = new Blob([buffer], { type: contentType });
      const objectUrl = URL.createObjectURL(blob);

      // Cache the photo URL in LRU cache
      this.cacheService.setUserPhoto(userId, objectUrl);
      
      // Set cleanup timer for object URL
      setTimeout(() => {
        URL.revokeObjectURL(objectUrl);
      }, 60 * 60 * 1000); // 60 minutes

      return objectUrl;
    } catch (error) {
      console.warn(`Failed to load profile photo for user ${userId}:`, error);
      return undefined;
    }
  }

  public async getUserPresence(userId: string): Promise<IUserPresence | undefined> {
    // Check cache first (5-minute TTL for presence)
    const cachedPresence = this.cacheService.getUserPresence(userId);
    if (cachedPresence) {
      return cachedPresence as IUserPresence;
    }

    try {
      const client = await this.getGraphClient();
      const response = await client.api(`/users/${userId}/presence`).get();
      const r = response as Record<string, unknown>;
      
      const presence: IUserPresence = {
        availability: r.availability as string || 'Unknown',
        activity: r.activity as string || 'Unknown',
        lastSeenDateTime: r.lastSeenDateTime as string
      };

      // Cache with 5-minute TTL
      this.cacheService.setUserPresence(userId, presence);
      
      return presence;
    } catch (error) {
      console.warn(`Could not fetch presence for user ${userId}:`, error);
      return undefined;
    }
  }

  public async getBatchUserPresence(userIds: string[]): Promise<Record<string, IUserPresence>> {
    const results: Record<string, IUserPresence> = {};
    const uncachedUserIds: string[] = [];

    // Check cache first
    for (const userId of userIds) {
      const cachedPresence = this.cacheService.getUserPresence(userId);
      if (cachedPresence) {
        results[userId] = cachedPresence as IUserPresence;
      } else {
        uncachedUserIds.push(userId);
      }
    }

    if (uncachedUserIds.length === 0) {
      return results;
    }

    try {
      // Batch presence requests (max 20 at a time)
      const batchSize = 20;
      for (let i = 0; i < uncachedUserIds.length; i += batchSize) {
        const batch = uncachedUserIds.slice(i, i + batchSize);
        
        const batchPromises = batch.map(async (userId) => {
          try {
            const presence = await this.getUserPresence(userId);
            if (presence) {
              results[userId] = presence;
            }
          } catch (error) {
            console.warn(`Failed to get presence for user ${userId}:`, error);
          }
        });

        // Manual Promise.allSettled equivalent for older TypeScript targets
        await Promise.all(batchPromises.map(async (promise) => {
          try {
            await promise;
          } catch {
            // Errors are already handled within individual promises
          }
        }));
        
        // Add small delay between batches to respect rate limits
        if (i + batchSize < uncachedUserIds.length) {
          await new Promise(resolve => setTimeout(resolve, this.RATE_LIMIT_DELAY));
        }
      }
    } catch (error) {
      console.error('Batch presence request failed:', error);
    }

    return results;
  }
}