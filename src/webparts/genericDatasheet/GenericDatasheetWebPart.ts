import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import GenericView from './components/GenericDatasheet';
import { PNPService } from '../../common/services/PNPService';
import { IGenericDatasheetProps } from './components/Datasheet.types';
import { UserService } from '../../common/services/UserContextService';
// @typescript-eslint/no-empty-interface
export interface IGenericDatasheetWebPartProps {
  description: string;
}

export default class GenericDatasheetWebPart extends BaseClientSideWebPart<IGenericDatasheetWebPartProps> {
  private pnpService: PNPService;
  private userService: UserService;
  protected async onInit(): Promise<void> {
    this.pnpService = new PNPService(this.context);
    this.userService = new UserService(this.pnpService);
    await this.userService.setup();
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IGenericDatasheetProps> = React.createElement(
      GenericView,
      {
        userService: this.userService,
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
}
