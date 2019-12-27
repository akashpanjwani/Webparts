import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'BusinessAreaBlockWebPartStrings';
import BusinessAreaBlock from './components/BusinessAreaBlock';
import { IBusinessAreaBlockProps } from './components/IBusinessAreaBlockProps';
import { SPComponentLoader } from '@microsoft/sp-loader';


export default class BusinessAreaBlockWebPart extends BaseClientSideWebPart<IBusinessAreaBlockProps> {

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css');
    return super.onInit();
  }
  
  public render(): void {
    const element: React.ReactElement<IBusinessAreaBlockProps > = React.createElement(
      BusinessAreaBlock,
      {
        title: this.properties.title || "",
        site: this.context.pageContext.site.absoluteUrl,
        currentUser: this.context.pageContext.user.displayName,
        context: this.context,
        serviceScope: this.context.serviceScope,
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
