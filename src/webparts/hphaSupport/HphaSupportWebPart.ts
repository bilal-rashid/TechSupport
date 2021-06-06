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
  colorBackground: string;
  colorHeader: string;
  colorLightBackground: string;
}
var colorPickerBackground = null;
var colorPickerLightBackground = null;
var colorPickerHeader = null;
// "https://hpeits.sharepoint.com/sites/IMPACT/SiteAssets/hphasupport/"
export default class HphaSupportWebPart extends BaseClientSideWebPart<IHphaSupportWebPartProps> {

  protected onInit(): Promise<void> {

    const dynamicLibImport = import(
      /* webpackChunkName: 'CommonLibraryLibrary' */
      '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker'
      ).then(lib => {
      const isChrome = !!window['chrome'] && (!!window['chrome'].webstore || !!window['chrome'].runtime);
      if (isChrome) {
        colorPickerHeader = lib.PropertyFieldColorPicker('colorHeader',{
          label: 'Header Color',
          selectedColor: this.properties.colorHeader,
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          disabled: false,
          isHidden: false,
          alphaSliderHidden: false,
          style: 1,
          iconName: 'Precipitation',
          key: 'colorFieldId'
        });
        colorPickerBackground = lib.PropertyFieldColorPicker('colorBackground',{
          label: 'Background Color',
          selectedColor: this.properties.colorBackground,
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          disabled: false,
          isHidden: false,
          alphaSliderHidden: false,
          style: 1,
          iconName: 'Precipitation',
          key: 'colorFieldId'
        });
        colorPickerLightBackground = lib.PropertyFieldColorPicker('colorLightBackground',{
          label: 'Light Background Color',
          selectedColor: this.properties.colorLightBackground,
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          disabled: false,
          isHidden: false,
          alphaSliderHidden: false,
          style: 1,
          iconName: 'Precipitation',
          key: 'colorFieldId'
        });
      } else {
        colorPickerHeader = PropertyPaneTextField('colorHeader', {
          label: 'Header Color'
        });
        colorPickerLightBackground = PropertyPaneTextField('colorBackground', {
          label: 'Background Color'
        });
        colorPickerBackground = PropertyPaneTextField('colorLightBackground', {
          label: 'Light Background Color'
        });
      }
    });
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
        link: (this.properties.link) ? this.properties.link : 'Link to Support material',
        context: this.context,
        colorHeader:(this.properties.colorHeader) ?  this.properties.colorHeader: '#5B9BD5',
        colorBackground: (this.properties.colorBackground) ? this.properties.colorBackground : '#DEEAF6',
        colorLightBackground: (this.properties.colorLightBackground) ? this.properties.colorLightBackground : '#ffffff'
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
                }),
                colorPickerHeader,
                colorPickerBackground,
                colorPickerLightBackground
              ]
            }
          ]
        }
      ]
    };
  }
}
