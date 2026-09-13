import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import GenericForm from './components/Generic Form/GenericForm';
import { PNPService } from '../../common/services/PNPService';
import { IGenericFormProps } from './components/Generic Form/GenericForm.types';
import { UserService } from '../../common/services/UserContextService';
// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface IGenericFormWebPartProps {
}

export default class GenericFormWebPart extends BaseClientSideWebPart<IGenericFormWebPartProps> {
  private pnpService: PNPService;
  private userService: UserService;

  protected async onInit(): Promise<void> {
    this.pnpService = new PNPService(this.context);
    this.userService = new UserService(this.pnpService);
    await this.userService.setup();
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IGenericFormProps> = React.createElement(
      GenericForm,
      {
        pnpService: this.pnpService,
        userService: this.userService
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
