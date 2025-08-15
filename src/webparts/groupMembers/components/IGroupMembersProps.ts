import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IGroupMembersProps {
  context: WebPartContext;
  roles: string[];
  itemsPerPage?: number;
  sortField?: string;
  showPresenceIndicator?: boolean;
  showSearchBox?: boolean;
  ownerLabel?: string;
  adminLabel?: string;
  memberLabel?: string;
  visitorLabel?: string;
}
