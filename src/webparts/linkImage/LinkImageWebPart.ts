import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'LinkImageWebPartStrings';
import LinkImage from './components/LinkImage';
import { ILinkImageProps } from './components/ILinkImageProps';


export default class LinkImageWebPart extends BaseClientSideWebPart<ILinkImageProps> {

  public render(): void {
    const element: React.ReactElement<ILinkImageProps > = React.createElement(
      LinkImage,
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
