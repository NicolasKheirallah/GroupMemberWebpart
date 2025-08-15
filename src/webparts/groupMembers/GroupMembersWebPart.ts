import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import EnhancedGroupMembers from './components/EnhancedGroupMembers';
import ErrorBoundary from './components/ErrorBoundary';
import { IGroupMembersWebPartProps } from './types/webPartProps';
import {
  WebPartTitle
} from '@pnp/spfx-controls-react/lib/WebPartTitle';

// Import the new architecture
import { 
  createServiceContainer, 
  setGlobalContainer, 
  disposeGlobalContainer,
  SERVICE_TOKENS
} from './services/ServiceContainer';
import { ILoggingService } from './services/LoggingService';
import { StateManager } from './state/StateManager';

export default class GroupMembersWebPart extends BaseClientSideWebPart<IGroupMembersWebPartProps> {
  private serviceContainer: ReturnType<typeof createServiceContainer>;
  private stateManager: StateManager | undefined;

  protected onInit(): Promise<void> {
    // Initialize the service container
    this.serviceContainer = createServiceContainer(this.context);
    setGlobalContainer(this.serviceContainer);

    // Initialize services
    const loggingService = this.serviceContainer.resolve<ILoggingService>(SERVICE_TOKENS.LOGGING_SERVICE);

    // Initialize state manager
    this.stateManager = StateManager.getInstance();

    // Log web part initialization
    loggingService.info('WebPart', 'Enhanced GroupMembers WebPart initialized', {
      instanceId: this.context.instanceId,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      userLoginName: this.context.pageContext.user.loginName
    });

    // Set default values for new properties if they're undefined
    if (this.properties.title === undefined) {
      this.properties.title = 'Site Members';
    }
    if (this.properties.showPresenceIndicator === undefined) {
      this.properties.showPresenceIndicator = false;
    }
    if (this.properties.showSearchBox === undefined) {
      this.properties.showSearchBox = true;
    }

    return super.onInit();
  }

  public render(): void {
    const loggingService = this.serviceContainer.resolve<ILoggingService>(SERVICE_TOKENS.LOGGING_SERVICE);
    
    // Clean up previous render first
    ReactDom.unmountComponentAtNode(this.domElement);
    
    try {
      // Convert toggle states to array of roles
      const roles = [
        this.properties.showOwners && 'owner',
        this.properties.showAdmins && 'admin',
        this.properties.showMembers && 'member',
        this.properties.showVisitors && 'visitor'
      ].filter(Boolean) as string[];

      if (roles.length === 0) {
        loggingService.warn('WebPart', 'No roles selected for display');
      }

      const webPartTitleElement = React.createElement(
        WebPartTitle,
        {
          displayMode: this.displayMode,
          title: this.properties.title,
          updateProperty: (value: string) => {
            this.properties.title = value;
          }
        }
      );

      const groupMembersElement = React.createElement(
        EnhancedGroupMembers,
        {
          context: this.context,
          roles,
          itemsPerPage: this.properties.itemsPerPage,
          sortField: this.properties.sortField,
          showSearchBox: this.properties.showSearchBox,
          showPresenceIndicator: this.properties.showPresenceIndicator,
          ownerLabel: this.properties.ownerLabel,
          adminLabel: this.properties.adminLabel,
          memberLabel: this.properties.memberLabel,
          visitorLabel: this.properties.visitorLabel
        }
      );

      const contentElement = React.createElement(
        'div',
        {},
        webPartTitleElement,
        groupMembersElement
      );

      const wrappedElement = React.createElement(
        ErrorBoundary,
        {
          level: 'critical' as const,
          context: 'GroupMembersWebPart',
          onError: (error, errorInfo) => {
            loggingService.critical('WebPart', 'Critical error in web part', error, {
              errorInfo,
              properties: this.properties
            });
          }
        },
        contentElement
      );

      ReactDom.render(wrappedElement, this.domElement);

      loggingService.debug('WebPart', 'Web part rendered successfully', {
        roles,
        itemsPerPage: this.properties.itemsPerPage,
        showSearchBox: this.properties.showSearchBox
      });

    } catch (error) {
      loggingService.critical('WebPart', 'Failed to render web part', error as Error);
      
      // Render error fallback
      const errorContent = React.createElement('div', { style: { padding: '20px' } }, 
        'An unexpected error occurred while loading the Group Members web part.'
      );

      const errorElement = React.createElement(
        ErrorBoundary,
        {
          level: 'critical' as const,
          context: 'WebPartRenderError'
        },
        errorContent
      );
      
      ReactDom.render(errorElement, this.domElement);
    }
  }

  protected onDispose(): void {
    try {
      const loggingService = this.serviceContainer?.resolve<ILoggingService>(SERVICE_TOKENS.LOGGING_SERVICE);
      
      if (loggingService) {
        loggingService.info('WebPart', 'Disposing GroupMembers WebPart');
      }

      // Clean up React
      ReactDom.unmountComponentAtNode(this.domElement);

      // Reset state manager
      if (this.stateManager) {
        this.stateManager.reset();
      }

      // Dispose services
      if (this.serviceContainer) {
        this.serviceContainer.dispose();
      }

      // Clean up global container
      disposeGlobalContainer();

    } catch (error) {
      console.error('Error during web part disposal:', error);
    }

    super.onDispose();
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    const loggingService = this.serviceContainer?.resolve<ILoggingService>(SERVICE_TOKENS.LOGGING_SERVICE);
    
    if (loggingService) {
      loggingService.debug('WebPart', 'Property pane field changed', {
        propertyPath,
        oldValue,
        newValue
      });
    }

    // Property changes are handled by the parent class
    // Configuration updates could be added here if needed

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configure which user roles to display and customize the appearance of your Group Members web part."
          },
          groups: [
            {
              groupName: "General Settings",
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title',
                  description: 'The title displayed at the top of the web part',
                  placeholder: 'e.g., Site Members, Team Directory'
                })
              ]
            },
            {
              groupName: "Display Options",
              groupFields: [
                PropertyPaneToggle('showOwners', {
                  label: 'Show Owners',
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneToggle('showAdmins', {
                  label: 'Show Administrators', 
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneToggle('showMembers', {
                  label: 'Show Members',
                  onText: 'Yes', 
                  offText: 'No'
                }),
                PropertyPaneToggle('showVisitors', {
                  label: 'Show Visitors',
                  onText: 'Yes',
                  offText: 'No'
                })
              ]
            },
            {
              groupName: "Features",
              groupFields: [
                PropertyPaneToggle('showSearchBox', {
                  label: 'Show Search Box',
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('showPresenceIndicator', {
                  label: 'Show Teams Presence',
                  onText: 'Enabled', 
                  offText: 'Disabled'
                }),
                PropertyPaneSlider('itemsPerPage', {
                  label: 'Items per page',
                  min: 5,
                  max: 50,
                  step: 5,
                  showValue: true
                }),
                PropertyPaneChoiceGroup('sortField', {
                  label: 'Default sort field',
                  options: [
                    { key: 'name', text: 'Name' },
                    { key: 'jobTitle', text: 'Job Title' }
                  ]
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Customize the labels for different user roles."
          },
          groups: [
            {
              groupName: "Role Labels",
              groupFields: [
                PropertyPaneTextField('ownerLabel', {
                  label: 'Owners Label',
                  placeholder: 'Owners'
                }),
                PropertyPaneTextField('adminLabel', {
                  label: 'Administrators Label', 
                  placeholder: 'Administrators'
                }),
                PropertyPaneTextField('memberLabel', {
                  label: 'Members Label',
                  placeholder: 'Members'
                }),
                PropertyPaneTextField('visitorLabel', {
                  label: 'Visitors Label',
                  placeholder: 'Visitors'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}