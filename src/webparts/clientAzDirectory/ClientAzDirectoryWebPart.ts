import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ClientAzDirectoryWebPartStrings';
import ClientAzDirectory from './components/ClientAzDirectory';
import { IClientAzDirectoryProps } from './components/IClientAzDirectoryProps';

export interface IClientAzDirectoryWebPartProps {
  listTitle: string;
  nameField: string;
  secondaryField: string;
  linkField: string;
}

export default class ClientAzDirectoryWebPart extends BaseClientSideWebPart<IClientAzDirectoryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IClientAzDirectoryProps> = React.createElement(
      ClientAzDirectory,
      {
        listTitle: this.properties.listTitle || 'Clients',
        nameField: this.properties.nameField || 'Title',
        secondaryField: this.properties.secondaryField || '',
        linkField: this.properties.linkField || '',
        spHttpClient: this.context.spHttpClient,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

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
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleFieldLabel,
                  description: strings.ListTitleFieldDescription
                }),
                PropertyPaneTextField('nameField', {
                  label: strings.NameFieldLabel,
                  description: strings.NameFieldDescription
                }),
                PropertyPaneTextField('secondaryField', {
                  label: strings.SecondaryFieldLabel,
                  description: strings.SecondaryFieldDescription
                }),
                PropertyPaneTextField('linkField', {
                  label: strings.LinkFieldLabel,
                  description: strings.LinkFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
