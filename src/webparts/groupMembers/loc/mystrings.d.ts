declare interface IGroupMembersWebPartStrings {
  // Property Pane
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  
  // General Settings
  GeneralSettingsGroupName: string;
  WebPartTitleLabel: string;
  WebPartTitleDescription: string;
  WebPartTitlePlaceholder: string;
  
  // Display Options
  DisplayOptionsGroupName: string;
  ShowOwnersLabel: string;
  ShowAdminsLabel: string;
  ShowMembersLabel: string;
  ShowVisitorsLabel: string;
  YesText: string;
  NoText: string;
  
  // Features
  FeaturesGroupName: string;
  ShowSearchBoxLabel: string;
  ShowPresenceLabel: string;
  ItemsPerPageLabel: string;
  DefaultSortFieldLabel: string;
  EnabledText: string;
  DisabledText: string;
  NameSortOption: string;
  JobTitleSortOption: string;
  
  // Role Labels
  RoleLabelsGroupName: string;
  RoleLabelsDescription: string;
  OwnersLabelField: string;
  AdminsLabelField: string;
  MembersLabelField: string;
  VisitorsLabelField: string;
  OwnersPlaceholder: string;
  AdminsPlaceholder: string;
  MembersPlaceholder: string;
  VisitorsPlaceholder: string;
  
  // UI Labels
  OwnersDefaultLabel: string;
  AdminsDefaultLabel: string;
  MembersDefaultLabel: string;
  VisitorsDefaultLabel: string;
  MemberDefaultTitle: string;
  
  // Search
  SearchPlaceholder: string;
  ShowingText: string;
  OfText: string;
  UsersText: string;
  ClearText: string;
  
  // Loading and Messages
  LoadingText: string;
  LoadingDescription: string;
  FailedToLoadTitle: string;
  RetryText: string;
  CheckPermissionsText: string;
  NoUsersFoundText: string;
  AdjustSearchText: string;
  
  // Actions
  StartChatText: string;
  SendEmailText: string;
  LoadMoreText: string;
  PreviousText: string;
  NextText: string;
  PageText: string;
  
  // Presence
  AvailableText: string;
  BusyText: string;
  AwayText: string;
  OfflineText: string;
  UnknownPresenceText: string;
  
  // Error Messages
  NetworkErrorText: string;
  PermissionErrorText: string;
  UnexpectedErrorText: string;
  
  // Environment Messages
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'GroupMembersWebPartStrings' {
  const strings: IGroupMembersWebPartStrings;
  export = strings;
}
