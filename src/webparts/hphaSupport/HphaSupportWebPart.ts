import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, PropertyPaneCheckbox,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HphaSupportWebPartStrings';
import HphaSupport from './components/HphaSupport';
import { IHphaSupportProps } from './components/IHphaSupportProps';
import { default as pnp } from "sp-pnp-js";

export interface IHphaSupportWebPartProps {
  firstCategory: string;
  secondCategory: string;
  thirdCategory: string;
  issues: string;
  firstSupport: string;
  secondSupport: string;
  tips: string;
  link: string;
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
        firstCategory: (this.properties.firstCategory) ? this.properties.firstCategory : 'Primary Category',
        secondCategory: (this.properties.secondCategory) ? this.properties.secondCategory : 'Secondary Category',
        thirdCategory: (this.properties.thirdCategory) ? this.properties.thirdCategory : 'Third Category',
        issues: (this.properties.issues) ? this.properties.issues : 'What issue are you having?',
        firstSupport: (this.properties.firstSupport) ? this.properties.firstSupport : 'How to get help ?',
        secondSupport: (this.properties.secondSupport) ? this.properties.secondSupport : 'Backup Department',
        tips: (this.properties.tips) ? this.properties.tips : 'Troubleshooting Tips',
        link: this.properties.link,
        context: this.context}
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
                PropertyPaneTextField('firstCategory', {
                  label: 'First Category Label'
                }),
                PropertyPaneTextField('secondCategory', {
                  label: 'Second Category Label'
                }),
                PropertyPaneTextField('thirdCategory', {
                  label: 'Third Category Label'
                }),
                PropertyPaneTextField('issues', {
                  label: 'Specific Issues Label'
                }),
                PropertyPaneTextField('firstSupport', {
                  label: 'First Tier Support Label'
                }),
                PropertyPaneTextField('secondSupport', {
                  label: '2nd Tier Support Label'
                }),
                PropertyPaneTextField('tips', {
                  label: 'TroubleShooting Tips Label'
                }),
                PropertyPaneTextField('link', {
                  label: 'Help Link Label'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
