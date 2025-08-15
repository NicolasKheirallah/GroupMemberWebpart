import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUser, ISite, IGroup } from '../types/interfaces';
import { CacheService } from './CacheService';
import { GroupMemberService } from './GroupMemberService';

// Type for Microsoft Graph Client
interface IMSGraphClient {
  api(path: string): IMSGraphClientRequest;
}

interface IMSGraphClientRequest {
  select(properties: string): IMSGraphClientRequest;
  get(): Promise<{ value?: unknown[] } & Record<string, unknown>>;
}

export interface ISitePermissionService {
  getCurrentSite(): Promise<ISite | undefined>;
  getSiteMembers(siteId: string): Promise<IUser[]>;
  getAllSiteMembers(): Promise<IUser[]>;
}

export class SitePermissionService implements ISitePermissionService {
  private context: WebPartContext;
  private graphClient: IMSGraphClient | undefined;
  private cacheService: CacheService;
  private groupMemberService: GroupMemberService;

  constructor(context: WebPartContext) {
    this.context = context;
    this.cacheService = CacheService.getInstance();
    this.groupMemberService = new GroupMemberService(context);
  }

  private async getGraphClient(): Promise<IMSGraphClient> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.graphClient;
  }

  public async getCurrentSite(): Promise<ISite | undefined> {
    try {
      const client = await this.getGraphClient();
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      
      // Get site information using the current site URL
      const hostname = new URL(siteUrl).hostname;
      const sitePath = new URL(siteUrl).pathname;
      
      const response = await client.api(`/sites/${hostname}:${sitePath}`).get();
      
      const r = response as Record<string, unknown>;
      const sharepointIds = r.sharepointIds as Record<string, unknown> || {};
      return {
        id: r.id as string,
        displayName: r.displayName as string,
        webUrl: r.webUrl as string,
        siteCollectionId: sharepointIds.siteId as string,
        webId: sharepointIds.webId as string
      };
    } catch (error) {
      console.warn('Could not get current site info:', error);
      return undefined;
    }
  }

  public async getSiteMembers(siteId: string): Promise<IUser[]> {
    const cacheKey = `siteMembers_${siteId}`;
    
    // Check LRU cache first
    const cachedData = this.cacheService.getUserData(cacheKey);
    if (cachedData) {
      return cachedData as IUser[];
    }

    try {
      const client = await this.getGraphClient();
      const allUsers: IUser[] = [];

      // Method 1: Get site permissions (includes inherited permissions)
      try {
        const permissionsResponse = await client
          .api(`/sites/${siteId}/permissions`)
          .select('id,roles,grantedToIdentitiesV2,grantedTo,inheritedFrom')
          .get();

        for (const permission of permissionsResponse.value || []) {
          const p = permission as Record<string, unknown>;
          const roles = p.roles as string[] || [];
          const grantedTo = p.grantedToIdentitiesV2 || p.grantedTo;
          
          if (grantedTo) {
            for (const identity of Array.isArray(grantedTo) ? grantedTo : [grantedTo]) {
              const id = identity as Record<string, unknown>;
              const accessLevel = this.mapSiteRolesToAccessLevel(roles);
              
              if (id.user) {
                // Direct user permission
                const user = id.user as Record<string, unknown>;
                allUsers.push({
                  id: user.id as string,
                  displayName: user.displayName as string,
                  mail: user.mail as string,
                  userPrincipalName: user.userPrincipalName as string,
                  jobTitle: user.jobTitle as string,
                  department: user.department as string,
                  officeLocation: user.officeLocation as string,
                  accessLevel,
                  source: 'site'
                });
              } else if (id.group) {
                // Security group or M365 group permission - resolve members
                try {
                  const group = id.group as Record<string, unknown>;
                  const groupMembers = await this.groupMemberService.resolveGroupMembers(group.id as string, accessLevel);
                  allUsers.push(...groupMembers);
                } catch (error) {
                  console.warn(`Failed to resolve group members for group ${id.group}:`, error);
                }
              }
            }
          }
        }
      } catch (error) {
        console.warn('Failed to get site permissions via /permissions endpoint:', error);
      }

      // Method 2: Try SharePoint REST API for additional members (fallback)
      try {
        const alternativeMembers = await this.getSharePointSiteMembers(siteId);
        allUsers.push(...alternativeMembers);
      } catch (error) {
        console.warn('Failed to get members via SharePoint REST API:', error);
      }

      // Method 3: Try to get site administrators specifically
      try {
        const siteAdmins = await this.getSiteAdministrators(siteId);
        allUsers.push(...siteAdmins);
      } catch (error) {
        console.warn('Failed to get site administrators:', error);
      }

      // Deduplicate users by ID, giving priority to higher access levels
      const userMap = new Map<string, IUser>();
      
      for (const user of allUsers) {
        const existingUser = userMap.get(user.id);
        if (!existingUser || this.getAccessLevelPriority(user.accessLevel) > this.getAccessLevelPriority(existingUser.accessLevel)) {
          userMap.set(user.id, user);
        }
      }

      const uniqueUsers = Array.from(userMap.values());

      // Cache the result
      this.cacheService.setUserData(cacheKey, uniqueUsers);
      
      return uniqueUsers;
    } catch (error) {
      console.error(`Error fetching site members for ${siteId}:`, error);
      return [];
    }
  }

  public async getAllSiteMembers(): Promise<IUser[]> {
    try {
      const currentSite = await this.getCurrentSite();
      if (!currentSite) {
        console.warn('Could not determine current site');
        return [];
      }

      const allUsers: IUser[] = [];
      const errors: string[] = [];

      // Try multiple approaches to find associated M365 group
      const associatedGroup = await this.findAssociatedGroup(currentSite);

      // Get M365 group members if there's an associated group
      if (associatedGroup) {
        try {
          // Get group owners (who are effectively both owners and admins in M365 groups)
          try {
            const groupOwners = await this.groupMemberService.getGroupMembers(associatedGroup.id, 'admin');
            const ownersWithLevel = groupOwners.map(user => ({ 
              ...user, 
              accessLevel: 'owner' as const, 
              source: 'group' as const 
            }));
            allUsers.push(...ownersWithLevel);
          } catch (error) {
            errors.push(`Failed to get owners from group: ${error}`);
          }

          // Get group members
          try {
            const groupMembers = await this.groupMemberService.getGroupMembers(associatedGroup.id, 'member');
            const membersWithLevel = groupMembers.map(user => ({ 
              ...user, 
              accessLevel: 'member' as const, 
              source: 'group' as const 
            }));
            allUsers.push(...membersWithLevel);
          } catch (error) {
            errors.push(`Failed to get members from group: ${error}`);
          }
        } catch (error) {
          errors.push(`Failed to get group members: ${error}`);
        }
      }

      // Always try to get direct site members (critical for Communication sites)
      try {
        const siteMembers = await this.getSiteMembersWithRetry(currentSite.id);
        allUsers.push(...siteMembers);
      } catch (error) {
        errors.push(`Failed to get site members: ${error}`);
        
        // If both group and site member retrieval fail, try fallback approach
        if (!associatedGroup) {
          try {
            const fallbackMembers = await this.getFallbackSiteMembers(currentSite.id);
            allUsers.push(...fallbackMembers);
          } catch (fallbackError) {
            errors.push(`Fallback method also failed: ${fallbackError}`);
          }
        }
      }

      // For Communication sites, also try to get visitors specifically
      try {
        const visitors = await this.getSiteVisitors(currentSite.id);
        allUsers.push(...visitors);
      } catch (error) {
        errors.push(`Failed to get site visitors: ${error}`);
      }

      // Log warnings if there were errors but we got some data
      if (errors.length > 0) {
        console.warn('Some member retrieval methods failed:', errors);
      }

      // If we have no users at all, try one more fallback approach
      if (allUsers.length === 0) {
        console.warn('No site members found through standard methods. Trying alternative approaches...');
        
        // For Communication sites, try getting site admins at least
        try {
          const currentUser = this.context.pageContext.user;
          if (currentUser) {
            // Add current user as a baseline
            allUsers.push({
              id: currentUser.loginName || currentUser.email || 'current',
              displayName: currentUser.displayName,
              mail: currentUser.email,
              userPrincipalName: currentUser.loginName || currentUser.email,
              accessLevel: 'admin' as const,
              source: 'site' as const
            });
            
            console.log('Added current user as baseline member');
          }
        } catch (error) {
          console.warn('Could not add current user as fallback:', error);
        }
        
        // If still no users, log detailed error information
        if (allUsers.length === 0) {
          console.error('No site members found. This might indicate:');
          console.error('1. Insufficient Microsoft Graph permissions');
          console.error('2. Site permissions API access issues');
          console.error('3. User does not have access to view site members');
          console.error('Required permissions: Sites.Read.All, User.Read.All, Group.Read.All');
          return [];
        }
      }

      // Deduplicate users by ID, giving priority to higher access levels
      const userMap = new Map<string, IUser>();
      
      for (const user of allUsers) {
        const existingUser = userMap.get(user.id);
        if (!existingUser || this.getAccessLevelPriority(user.accessLevel) > this.getAccessLevelPriority(existingUser.accessLevel)) {
          userMap.set(user.id, user);
        }
      }

      const uniqueUsers = Array.from(userMap.values());
      console.log(`Found ${uniqueUsers.length} unique site members from ${allUsers.length} total entries`);
      
      return uniqueUsers;
    } catch (error) {
      console.error('Critical error in getAllSiteMembers:', error);
      return [];
    }
  }

  private mapSiteRolesToAccessLevel(roles: string[]): 'owner' | 'admin' | 'member' | 'visitor' {
    // Convert roles to lowercase for case-insensitive comparison
    const lowerRoles = roles.map(role => role.toLowerCase());
    
    // Site Owner or Full Control
    if (lowerRoles.some(role => 
      role.includes('owner') || 
      role.includes('fullcontrol') || 
      role === 'full control' ||
      role.includes('siteadmin') ||
      role.includes('site admin')
    )) {
      return 'owner';
    }
    
    // Site Administrator, Design, or Manage permissions
    if (lowerRoles.some(role => 
      role.includes('admin') || 
      role.includes('manage') || 
      role.includes('design') ||
      role === 'manage hierarchy' ||
      role === 'approve' ||
      role.includes('moderate') ||
      role.includes('restrict')
    )) {
      return 'admin';
    }
    
    // Contributors, Edit, Write permissions
    if (lowerRoles.some(role => 
      role.includes('edit') || 
      role.includes('contribute') || 
      role.includes('write') ||
      role === 'add and customize pages' ||
      role === 'add items' ||
      role === 'edit items' ||
      role.includes('create') ||
      role.includes('modify')
    )) {
      return 'member';
    }
    
    // Visitors, Read-only permissions (be more explicit about visitor roles)
    if (lowerRoles.some(role => 
      role.includes('read') || 
      role.includes('view') ||
      role.includes('visitor') ||
      role === 'view only' ||
      role === 'limited access' ||
      role.includes('browse')
    )) {
      return 'visitor';
    }
    
    // Default to visitor for any other permissions
    return 'visitor';
  }

  // Additional method to get SharePoint site members via alternative API
  private async getSharePointSiteMembers(siteId: string): Promise<IUser[]> {
    const client = await this.getGraphClient();
    const allUsers: IUser[] = [];

    try {
      // Try to get site users via alternative Graph endpoints
      const siteDrive = await client.api(`/sites/${siteId}/drive`).get();
      
      if (siteDrive) {
        // Get users who have access to the site drive
        const drivePermissions = await client
          .api(`/sites/${siteId}/drive/root/permissions`)
          .select('id,roles,grantedToIdentitiesV2,grantedTo')
          .get();

        for (const permission of drivePermissions.value || []) {
          const p = permission as Record<string, unknown>;
          const roles = p.roles as string[] || [];
          const grantedTo = p.grantedToIdentitiesV2 || p.grantedTo;
          
          if (grantedTo) {
            for (const identity of Array.isArray(grantedTo) ? grantedTo : [grantedTo]) {
              const id = identity as Record<string, unknown>;
              const accessLevel = this.mapSiteRolesToAccessLevel(roles);
              
              if (id.user) {
                const user = id.user as Record<string, unknown>;
                allUsers.push({
                  id: user.id as string,
                  displayName: user.displayName as string,
                  mail: user.mail as string,
                  userPrincipalName: user.userPrincipalName as string,
                  jobTitle: user.jobTitle as string,
                  department: user.department as string,
                  officeLocation: user.officeLocation as string,
                  accessLevel,
                  source: 'site'
                });
              }
            }
          }
        }
      }
    } catch (error) {
      console.debug('Alternative SharePoint API method failed:', error);
    }

    return allUsers;
  }

  // Method to specifically get site administrators
  private async getSiteAdministrators(siteId: string): Promise<IUser[]> {
    const client = await this.getGraphClient();
    const admins: IUser[] = [];

    try {
      // Method 1: Try to get site information with owner details
      const siteInfo = await client
        .api(`/sites/${siteId}`)
        .select('id,displayName,createdBy,siteCollection')
        .get();

      const siteData = siteInfo as Record<string, unknown>;
      
      // Add site creator as owner if available
      if (siteData.createdBy) {
        const createdBy = siteData.createdBy as Record<string, unknown>;
        if (createdBy.user) {
          const user = createdBy.user as Record<string, unknown>;
          admins.push({
            id: user.id as string || 'creator',
            displayName: user.displayName as string || 'Site Creator',
            mail: user.email as string,
            userPrincipalName: user.userPrincipalName as string || user.email as string,
            accessLevel: 'owner',
            source: 'site'
          });
        }
      }

      // Method 2: Try to get site collection administrators
      try {
        const siteCollection = siteData.siteCollection as Record<string, unknown>;
        if (siteCollection && siteCollection.hostname) {
          // Additional logic to get tenant admins if needed
          console.debug('Site collection info available for admin discovery');
        }
      } catch (error) {
        console.debug('Could not get site collection admins:', error);
      }

    } catch (error) {
      console.debug('Could not get site administrators:', error);
    }

    return admins;
  }

  // Enhanced method to get visitors specifically
  private async getSiteVisitors(siteId: string): Promise<IUser[]> {
    const client = await this.getGraphClient();
    const visitors: IUser[] = [];

    try {
      // Look for "Everyone" or "Everyone except external users" permissions
      const permissionsResponse = await client
        .api(`/sites/${siteId}/permissions`)
        .select('id,roles,grantedToIdentitiesV2,grantedTo')
        .get();

      for (const permission of permissionsResponse.value || []) {
        const p = permission as Record<string, unknown>;
        const roles = p.roles as string[] || [];
        const grantedTo = p.grantedToIdentitiesV2 || p.grantedTo;
        
        // Only process if this is clearly a visitor-level permission
        if (this.mapSiteRolesToAccessLevel(roles) === 'visitor' && grantedTo) {
          for (const identity of Array.isArray(grantedTo) ? grantedTo : [grantedTo]) {
            const id = identity as Record<string, unknown>;
            
            if (id.user) {
              const user = id.user as Record<string, unknown>;
              visitors.push({
                id: user.id as string,
                displayName: user.displayName as string,
                mail: user.mail as string,
                userPrincipalName: user.userPrincipalName as string,
                jobTitle: user.jobTitle as string,
                department: user.department as string,
                officeLocation: user.officeLocation as string,
                accessLevel: 'visitor',
                source: 'site'
              });
            }
          }
        }
      }
    } catch (error) {
      console.debug('Could not get site visitors:', error);
    }

    return visitors;
  }

  private getAccessLevelPriority(level?: string): number {
    switch (level) {
      case 'owner': return 4;
      case 'admin': return 3;
      case 'member': return 2;
      case 'visitor': return 1;
      default: return 0;
    }
  }

  private async findAssociatedGroup(site: ISite): Promise<IGroup | undefined> {
    try {
      // Method 1: Try to get the site's group directly via Graph API
      try {
        const client = await this.getGraphClient();
        const response = await client.api(`/sites/${site.id}/drive`).get();
        const driveResponse = response as Record<string, unknown>;
        
        if (driveResponse.quota && (driveResponse.quota as Record<string, unknown>).deleted === undefined) {
          // This is likely a group-connected site, try to get the group
          const groupResponse = await client.api(`/sites/${site.id}/group`).get();
          const g = groupResponse as Record<string, unknown>;
          
          return {
            id: g.id as string,
            displayName: g.displayName as string,
            '@odata.type': '#microsoft.graph.group',
            description: g.description as string
          };
        }
      } catch {
        // Site might not have an associated group, continue with other methods
        console.log('No direct group association found, trying other methods');
      }

      // Method 2: Match by site URL patterns (more reliable than display name)
      const groups = await this.groupMemberService.getUserGroups();
      const siteUrl = site.webUrl.toLowerCase();
      
      // Try to find group by examining site URL structure
      const potentialGroup = groups.find(group => {
        const groupName = group.displayName.toLowerCase().replace(/\s+/g, '');
        const siteName = this.extractSiteNameFromUrl(siteUrl);
        
        return groupName === siteName || 
               siteName.includes(groupName) || 
               groupName.includes(siteName);
      });

      if (potentialGroup) {
        return potentialGroup;
      }

      // Method 3: Check if current user is owner of any groups that might match
      // This is useful for scenarios where the user has limited group visibility
      for (const group of groups) {
        try {
          const groupOwners = await this.groupMemberService.getGroupMembers(group.id, 'admin');
          const currentUser = this.context.pageContext.user;
          
          if (groupOwners.some(owner => 
            owner.userPrincipalName === currentUser.loginName ||
            owner.mail === currentUser.email
          )) {
            // User is owner of this group, might be the associated group
            const siteName = this.extractSiteNameFromUrl(siteUrl);
            const groupName = group.displayName.toLowerCase();
            
            if (groupName.includes(siteName) || siteName.includes(groupName)) {
              return group;
            }
          }
        } catch {
          // Continue to next group if this one fails
          continue;
        }
      }

      console.log('No associated M365 group found for this site');
      return undefined;
    } catch (error) {
      console.warn('Error finding associated group:', error);
      return undefined;
    }
  }

  private extractSiteNameFromUrl(siteUrl: string): string {
    try {
      const url = new URL(siteUrl);
      const pathParts = url.pathname.split('/').filter(part => part.length > 0);
      
      // For sites like /sites/sitename or /teams/teamname
      if (pathParts.length >= 2 && (pathParts[0] === 'sites' || pathParts[0] === 'teams')) {
        return pathParts[1].toLowerCase().replace(/[^a-z0-9]/g, '');
      }
      
      // For other patterns, try to extract meaningful name
      const lastPart = pathParts[pathParts.length - 1];
      return lastPart.toLowerCase().replace(/[^a-z0-9]/g, '');
    } catch {
      return '';
    }
  }

  private async getSiteMembersWithRetry(siteId: string, maxRetries: number = 3): Promise<IUser[]> {
    let lastError: unknown;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return await this.getSiteMembers(siteId);
      } catch (error) {
        lastError = error;
        console.warn(`Attempt ${attempt}/${maxRetries} failed for getSiteMembers:`, error);
        
        if (attempt < maxRetries) {
          // Exponential backoff: 1s, 2s, 4s
          await new Promise(resolve => setTimeout(resolve, Math.pow(2, attempt - 1) * 1000));
        }
      }
    }
    
    throw lastError;
  }

  private async getFallbackSiteMembers(siteId: string): Promise<IUser[]> {
    const allUsers: IUser[] = [];
    
    // Try multiple fallback approaches
    const fallbackMethods = [
      // Method 1: Try getting site information with creator
      async () => {
        try {
          const client = await this.getGraphClient();
          const response = await client
            .api(`/sites/${siteId}`)
            .select('createdBy')
            .get();
          
          const r = response as Record<string, unknown>;
          const createdBy = r.createdBy as Record<string, unknown>;
          
          if (createdBy && createdBy.user) {
            const user = createdBy.user as Record<string, unknown>;
            return [{
              id: user.id as string || 'creator',
              displayName: user.displayName as string || 'Site Creator',
              mail: user.email as string,
              userPrincipalName: user.email as string,
              accessLevel: 'owner' as const,
              source: 'site' as const
            }];
          }
          return [];
        } catch (error) {
          console.log('Site creator method failed:', error);
          return [];
        }
      },
      
      // Method 2: Try getting basic site information
      async () => {
        try {
          const client = await this.getGraphClient();
          await client
            .api(`/sites/${siteId}`)
            .get();
          
          // At minimum, we can show that the site exists and has the current user
          console.log('Site basic info retrieved, but no specific member information available');
          return [];
        } catch (error) {
          console.log('Site basic info method failed:', error);
          return [];
        }
      }
    ];
    
    // Try each fallback method
    for (const method of fallbackMethods) {
      try {
        const users = await method();
        if (users.length > 0) {
          allUsers.push(...users);
          console.log(`Fallback method found ${users.length} users`);
        }
      } catch (error) {
        console.log('Fallback method failed:', error);
        continue;
      }
    }
    
    if (allUsers.length === 0) {
      console.error('All fallback site member methods failed');
      throw new Error('No fallback methods could retrieve site members');
    }
    
    return allUsers;
  }
}