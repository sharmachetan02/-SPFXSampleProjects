import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'NavTilesWebPartStrings';
import NavTiles from './components/NavTiles';
import { INavTilesProps } from './components/INavTilesProps.types';
import { PNPService } from '../../common/services/PNPService';

export interface INavTilesWebPartProps {
  description: string;
}

export default class NavTilesWebPart extends BaseClientSideWebPart<INavTilesWebPartProps> {
  private pnpService: PNPService;

  protected onInit(): Promise<void> {
    this.pnpService = new PNPService(this.context);
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<INavTilesProps> = React.createElement(
      NavTiles,
      {
        pnpService: this.pnpService
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
