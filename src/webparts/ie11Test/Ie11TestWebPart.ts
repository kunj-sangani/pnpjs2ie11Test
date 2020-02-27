import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'Ie11TestWebPartStrings';
import Ie11Test from './components/Ie11Test';
import { IIe11TestProps } from './components/IIe11TestProps';
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp/presets/all";

export interface IIe11TestWebPartProps {
  description: string;
}

export default class Ie11TestWebPart extends BaseClientSideWebPart<IIe11TestWebPartProps> {

  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      sp.setup({
        ie11:true,
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IIe11TestProps > = React.createElement(
      Ie11Test,
      {
        description: this.properties.description
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
