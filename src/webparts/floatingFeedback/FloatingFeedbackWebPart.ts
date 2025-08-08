import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as strings from 'FloatingFeedbackWebPartStrings';
import FloatingFeedback from './components/FloatingFeedback';
import { IFloatingFeedbackProps } from './components/IFloatingFeedbackProps';

export interface IFloatingFeedbackWebPartProps {
  description: string;
  listId: string;
  position: 'Top' | 'Bottom';
  titleColumn: string;
  descriptionColumn: string;
  ratingColumn: string;
  categoryColumn: string;
}

export default class FloatingFeedbackWebPart extends BaseClientSideWebPart<IFloatingFeedbackWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _lists: IPropertyPaneDropdownOption[] = [];
  private _listsDropdownDisabled: boolean = true;
  private _fields: IPropertyPaneDropdownOption[] = [];
  private _fieldsDropdownDisabled: boolean = true;

  public render(): void {
    const element: React.ReactElement<IFloatingFeedbackProps> = React.createElement(
      FloatingFeedback,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        pageName: document.title,
        listName: 'Feedback',
        position: this.properties.position,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    const message = await this._getEnvironmentMessage();
    this._environmentMessage = message;

    await this._getLists();

    // If we already have a list selected, try to load fields
    if (this.properties.listId) {
      await this._getFields(this.properties.listId);
    }
  }

  private _getLists(): Promise<void> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1) // eslint-disable-line @typescript-eslint/no-explicit-any
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((json) => {
        this._lists = json.value.map((list: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
          return {
            key: list.Id,
            text: list.Title
          };
        });
        this._listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
      });
  }

  private _getFields(listId: string): Promise<void> {
    if (!listId) return Promise.resolve();

    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists(guid'${listId}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((json) => {
        this._fields = json.value.map((field: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
          return {
            key: field.InternalName,
            text: `${field.Title} (${field.InternalName})`
          };
        });
        this._fieldsDropdownDisabled = false;
        this.context.propertyPane.refresh();
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void { // eslint-disable-line @typescript-eslint/no-explicit-any
    if (propertyPath === 'listId' && newValue) {
      this._fieldsDropdownDisabled = true;
      this._fields = [];
      this.properties.titleColumn = '';
      this.properties.descriptionColumn = '';
      this.properties.ratingColumn = '';
      this.properties.categoryColumn = '';

      // Fire and forget - _getFields calls refresh() when done
      this._getFields(newValue).catch((err) => {
        console.error('Error loading fields', err);
      });

      this.context.propertyPane.refresh();
    }
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
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
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('listId', {
                  label: 'Select Feedback List',
                  options: this._lists,
                  disabled: this._listsDropdownDisabled
                }),
                PropertyPaneDropdown('titleColumn', {
                  label: 'Title Column',
                  options: this._fields,
                  disabled: this._fieldsDropdownDisabled || !this.properties.listId
                }),
                PropertyPaneDropdown('descriptionColumn', {
                  label: 'Description Column',
                  options: this._fields,
                  disabled: this._fieldsDropdownDisabled || !this.properties.listId
                }),
                PropertyPaneDropdown('ratingColumn', {
                  label: 'Rating Column',
                  options: this._fields,
                  disabled: this._fieldsDropdownDisabled || !this.properties.listId
                }),
                PropertyPaneDropdown('categoryColumn', {
                  label: 'Category Column',
                  options: this._fields,
                  disabled: this._fieldsDropdownDisabled || !this.properties.listId
                }),
                PropertyPaneChoiceGroup('position', {
                  label: 'Position',
                  options: [
                    { key: 'Top', text: 'Top' },
                    { key: 'Bottom', text: 'Bottom' }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
