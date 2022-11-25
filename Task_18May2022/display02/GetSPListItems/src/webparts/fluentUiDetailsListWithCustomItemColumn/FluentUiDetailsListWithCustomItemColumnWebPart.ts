import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FluentUiDetailsListWithCustomItemColumnWebPartStrings';
import FluentUiDetailsListWithCustomItemColumn from './components/FluentUiDetailsListWithCustomItemColumn';
import { IFluentUiDetailsListWithCustomItemColumnProps } from './components/IFluentUiDetailsListWithCustomItemColumnProps';

export interface IFluentUiDetailsListWithCustomItemColumnWebPartProps {
  description: string;
}

export default class FluentUiDetailsListWithCustomItemColumnWebPart extends BaseClientSideWebPart<IFluentUiDetailsListWithCustomItemColumnWebPartProps> {

  

  protected onInit(): Promise<void> {

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IFluentUiDetailsListWithCustomItemColumnProps> = React.createElement(
      FluentUiDetailsListWithCustomItemColumn,
      {
        description: this.properties.description,
        webURL:this.context.pageContext.web.absoluteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
