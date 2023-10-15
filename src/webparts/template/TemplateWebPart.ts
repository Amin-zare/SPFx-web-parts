import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneLink,
  PropertyPaneSlider,

} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TemplateWebPartStrings';
import Template from './components/Template';
import { ITemplateProps } from './components/ITemplateProps';

import { SPHttpClient } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

export interface ITemplateWebPartProps {
  description: string;
  checkbox: boolean;
  toggle: boolean;
  multiLineText: string;
  Rating: number;
  ListTitle: string;
}

export default class TemplateWebPart extends BaseClientSideWebPart<ITemplateWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITemplateProps> = React.createElement(
      Template,
      {
        description: this.properties.description,
        checkbox: this.properties.checkbox,
        toggle: this.properties.toggle,
        context: this.context,
        multiLineText: this.properties.multiLineText,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        Rating: this.properties.Rating
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  private validateDescription(value: string): string {
    if (value === null ||
      value.trim().length === 0) {
      return 'Provide a description';
    }

    if (value.length > 40) {
      return 'Description should not be longer than 40 characters';
    }

    return '';
  }


  private async validateListTitle(value: string): Promise<string> {
    if (value === null || value.length === 0) {
      return "Provide the list name";
    }

    try {
      const response = await this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl +
        `/_api/web/lists/getByTitle('${escape(value)}')?$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        return "";
      } else if (response.status === 404) {
        return `List '${escape(value)}' doesn't exist in the current site`;
      } else {
        return `Error: ${response.statusText}. Please try again`;
      }
    } catch (error) {
      return error.message;
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  onGetErrorMessage: this.validateDescription.bind(this)
                }),
                PropertyPaneCheckbox('checkbox', {
                  text: 'Checkbox'
                }),
                PropertyPaneToggle('toggle', {
                  label: strings.ToggleFieldLabel,
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneTextField('multiLineText', {
                  label: strings.MultiLineFieldLabel,
                  multiline: true
                }),
                PropertyPaneLink('linkProperty', {
                  href: 'https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/integrate-with-property-pane',
                  text: 'Click to view usage and configuration details',
                  target: '_blank'
                }),
                PropertyPaneSlider('Rating', {
                  label: strings.RatingFieldLabel,
                  min: 1,
                  max: 10,
                  step: 1,
                  showValue: true,
                  value: 1
                }),
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleFieldLabel,
                  onGetErrorMessage: this.validateListTitle.bind(this),
                  deferredValidationTime: 500 // This property specifies the number of milliseconds that the SharePoint Framework waits before starting the validation process
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
