import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import * as strings from 'DetailsListCustomColumnsExampleWebPartStrings';
import DetailsListCustomColumnsExample from './components/DetailsListCustomColumnsExample';
import { IDetailsListCustomColumnsExampleProps } from './components/IDetailsListCustomColumnsExampleProps';
import { ClientMode } from './components/ClientMode';

export default class DetailsListCustomColumnsExampleWebPart extends BaseClientSideWebPart<IDetailsListCustomColumnsExampleProps> {

  public render(): void {
    const element: React.ReactElement<IDetailsListCustomColumnsExampleProps> = React.createElement(
      DetailsListCustomColumnsExample,
      {
        clientMode: this.properties.clientMode,
        context: this.context,     
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneChoiceGroup('clientMode', {
                  label: strings.ClientModeLabel,
                  options: [
                    { key: ClientMode.aad, text: "AadHttpClient"},
                    { key: ClientMode.graph, text: "MSGraphClient"},
                  ]
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
