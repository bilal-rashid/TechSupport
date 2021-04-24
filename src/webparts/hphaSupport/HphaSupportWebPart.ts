import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HphaSupportWebPartStrings';
import HphaSupport from './components/HphaSupport';
import { IHphaSupportProps } from './components/IHphaSupportProps';
import { default as pnp } from "sp-pnp-js";

export interface IHphaSupportWebPartProps {
  equipment: string;
  issues: string;
  first: string;
  second: string;
  tips: string;
}

export default class HphaSupportWebPart extends BaseClientSideWebPart<IHphaSupportWebPartProps> {

  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
      // other init code may be present
      pnp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IHphaSupportProps> = React.createElement(
      HphaSupport,
      {
        equipment: (this.properties.equipment) ? this.properties.equipment : 'Choose Your Equipment',
        issues: (this.properties.issues) ? this.properties.issues : 'What issue are you having?',
        first: (this.properties.first) ? this.properties.first : 'Who to Call 1st ?',
        second: (this.properties.second) ? this.properties.second : 'Who to Call 2nd ?',
        tips: (this.properties.tips) ? this.properties.tips : 'Troubleshooting Tips',
        context: this.context
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
              groupName: 'Configurable Labels',
              groupFields: [
                PropertyPaneTextField('equipment', {
                  label: 'Equipment Label'
                }),
                PropertyPaneTextField('issues', {
                  label: 'Issues Label'
                }),
                PropertyPaneTextField('first', {
                  label: 'First Tier Label'
                }),
                PropertyPaneTextField('second', {
                  label: '2nd Tier Label'
                }),
                PropertyPaneTextField('tips', {
                  label: 'Tips Label'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
