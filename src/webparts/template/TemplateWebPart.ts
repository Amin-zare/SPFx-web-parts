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
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneChoiceGroup
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
  listName: string;
  itemName: string;
  preconfiguredListName: string;
  order: string;
  numberOfItems: number;
  style: string;
}

export default class TemplateWebPart extends BaseClientSideWebPart<ITemplateWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;
  private loadingIndicator: boolean = true;
  private items: IPropertyPaneDropdownOption[];
  private itemsDropdownDisabled: boolean = true;

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
        Rating: this.properties.Rating,
        listName: this.properties.listName,
        itemName: this.properties.itemName,
        preconfiguredListName: this.properties.preconfiguredListName,
        order: this.properties.order,
        numberOfItems: this.properties.numberOfItems,
        style: this.properties.style
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private async loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return await new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, _reject: (error: any) => void) => {
      setTimeout((): void => {
        resolve([{
          key: 'sharedDocuments',
          text: 'Shared Documents'
        },
        {
          key: 'myDocuments',
          text: 'My Documents'
        }]);
      }, 2000);
    });
  }

  private async loadItems(): Promise<IPropertyPaneDropdownOption[]> {
    if (!this.properties.listName) {
      // return empty options since no list has been selected
      return [];
    }

    // This is where you'd replace the mock data with the actual data from SharePoint
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return await new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
      // timeout to simulate async call
      setTimeout(() => {
        const items: { [key: string]: { key: string; text: string }[] } = {
          sharedDocuments: [
            {
              key: 'spfx_presentation.pptx',
              text: 'SPFx for the masses'
            },
            {
              key: 'hello-world.spapp',
              text: 'hello-world.spapp'
            }
          ],
          myDocuments: [
            {
              key: 'isaiah_cv.docx',
              text: 'Isaiah CV'
            },
            {
              key: 'isaiah_expenses.xlsx',
              text: 'Isaiah Expenses'
            }
          ]
        };
        resolve(items[this.properties.listName]);
      }, 2000);
    });
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    // disable the item selector until lists have been loaded
    this.listsDropdownDisabled = !this.lists;

    // disable the item selector until items have been loaded or if the list has not been selected
    this.itemsDropdownDisabled = !this.properties.listName || !this.items;

    // nothing to do until someone selects a list
    if (this.lists) {
      return;
    }

    // show a loading indicator in the property pane while loading lists and items
    this.loadingIndicator = true;
    this.context.propertyPane.refresh();

    // load the lists from SharePoint
    const listOptions: IPropertyPaneDropdownOption[] = await this.loadLists();
    this.lists = listOptions;
    this.listsDropdownDisabled = false;

    // load the items from SharePoint
    const itemOptions: IPropertyPaneDropdownOption[] = await this.loadItems();
    this.items = itemOptions;
    this.itemsDropdownDisabled = !this.properties.listName;

    // remove the loading indicator
    this.loadingIndicator = false;
    this.context.propertyPane.refresh();
  }


  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    if (propertyPath === 'listName' && newValue) {
      // communicate loading items
      this.loadingIndicator = true;

      // push new list value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

      // reset selected item
      this.properties.itemName = ''; // use empty string to force property pane to reset the selected item. undefined will not trigger the reset

      // disable item selector until new items are loaded
      this.itemsDropdownDisabled = true;

      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();

      // get new items
      const itemOptions: IPropertyPaneDropdownOption[] = await this.loadItems();

      // store items
      this.items = itemOptions;

      // enable item selector
      this.itemsDropdownDisabled = false;

      // clear status indicator
      this.loadingIndicator = false;

      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
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
      showLoadingIndicator: this.loadingIndicator,
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
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
                }),
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                  // onGetErrorMessage: this.validateListName.bind(this), 
                  // deferredValidationTime: 500 // This property specifies the number of milliseconds that the SharePoint Framework waits before starting the validation process
                }),
                PropertyPaneDropdown('itemName', {
                  label: strings.ItemNameFieldLabel,
                  options: this.items,
                  disabled: this.itemsDropdownDisabled,
                  selectedKey: this.properties.itemName // don't forget to bind this property so it is refreshed when the parent property changes
                }),
                PropertyPaneDropdown('preconfiguredListName', {
                  label: strings.ListNameFieldLabel,
                  options: [{
                    key: 'Documents',
                    text: 'Documents'
                  },
                  {
                    key: 'Images',
                    text: 'Images'
                  }]
                }),
                PropertyPaneChoiceGroup('order', {
                  label: strings.OrderFieldLabel,
                  options: [{
                    key: 'chronological',
                    text: strings.OrderFieldChronologicalOptionLabel
                  },
                  {
                    key: 'reversed',
                    text: strings.OrderFieldReversedOptionLabel
                  }]
                }),
                PropertyPaneSlider('numberOfItems', {
                  label: strings.NumberOfItemsFieldLabel,
                  min: 1,
                  max: 10,
                  step: 1
                }),
                PropertyPaneChoiceGroup('style', {
                  label: strings.StyleFieldLabel,
                  options: [{
                    key: 'thumbnails',
                    text: strings.StyleFieldThumbnailsOptionLabel
                  },
                  {
                    key: 'list',
                    text: strings.StyleFieldListOptionLabel
                  }]
                })]
            }
          ]
        }
      ]
    };
  }
}
