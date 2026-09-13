import {

  BaseApplicationCustomizer,

  PlaceholderContent,

  PlaceholderName

} from '@microsoft/sp-application-base';



import { override } from '@microsoft/decorators';
import "./AppCustomizer.module.scss";
import ReactDOM from 'react-dom';
import SiteHeader from './components/siteHeader/AppHeader';
import { ISiteHeaderProps } from './components/siteHeader/AppHeader.types';
import React from 'react';
import { PNPService } from '../../common/services/PNPService';
import { loadStyles } from '@microsoft/load-themed-styles';
import { UserService } from '../../common/services/UserContextService';
import { Navigation } from 'spfx-navigation';

loadStyles(`
  .top {
    background: url("/_layouts/15/images/body-map-bg.gif") repeat-x 0 0 !important;
    background-color: #f4f4f4;
    border-bottom: 1px solid rgb(210, 210, 210);
    width: 100%;
    line-height: 2.5;
    font-weight: bold;
    border: 1px solid red;
  }

  .top .ms-CommandBar {
    height: 35px;
    background-color: transparent;
    width: 100%;
    border: 1px solid red !important;
  }

  .top .ms-CommandBar-primaryCommands {
    margin: 0 !important;
    padding-left: 10px;
  }

  .top .ms-CommandBarItem {
    border: none;
    height: 35px;
    padding: 0;
    border: 1px solid red !important;
  }

  .top .ms-CommandBarItem-iconColor {
    display: none;
  }

  .top .ms-CommandBarItem-commandText,
  .top .ms-Icon {
    color: #007dba !important;
    border: 1px solid red !important;
  }

  .top .searchButtonContainer .ms-Icon {
    color: #f4f4f4 !important;
  }

  .top .ms-CommandBarItem:hover .ms-CommandBarItem-commandText,
  .top .ms-CommandBarItem:hover .ms-Icon {
    color: #007dba !important;
  }

  .top .ms-CommandBarItem-link {
    padding: 0 10px !important;
    background-color: inherit !important;
  }

  .top .ms-CommandBarItem-text {
    font-size: 9pt;
    font-weight: bold;
    height: 35px;
    line-height: 35px;
    text-align: center;
    border: 1px solid red !important;
  }
`);

export interface IAppCustomizerProperties {
  message: string;
}

export default class AppCustomizer
  extends BaseApplicationCustomizer<IAppCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  private _pnpService: PNPService;
  private _userService: UserService;

  @override
  public async onInit(): Promise<void> {
    this._pnpService = new PNPService(this.context);
    this._userService = new UserService(this._pnpService);
    if (!window['isNavigatedEventSubscribed']) {
      window['isNavigatedEventSubscribed'] = true;
      this.context.application.navigatedEvent.add(this, this.render);
    }
    await this._userService.setup();
    return super.onInit();
    //this._userService = new UserService(this._pnpService);
    // await this._userService.setup();
    //return super.onInit();
  }

  // private handleNavigate = (url: string): void => {
  //   // Use window.location.href for navigation
  //   // This triggers navigatedEvent and allows menu to stay visible
  //   //window.location.href = url;

  //   const targetUrl = new URL(url, window.location.origin);

  //   this.context.application.navigateTo(
  //     targetUrl,
  //     false // false = NO FULL PAGE RELOAD
  //   );
  // };

  private handleNavigate = (url: string): void => {

    //window.location.href = url;
    Navigation.navigate(url, false);

  };

  private renderSiteHeader = (): void => {
    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      const element: React.ReactElement<ISiteHeaderProps> = React.createElement(
        SiteHeader, { pnpService: this._pnpService, userService: this._userService, onNavigate: this.handleNavigate }
      );
      ReactDOM.render(element, this._topPlaceholder.domElement);
    } else {
      this.render();
    }
  };

  private render(): void {
    if (this.context.placeholderProvider.placeholderNames.indexOf(PlaceholderName.Top) !== -1) {
      if (!this._topPlaceholder || !this._topPlaceholder.domElement) {
        this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
          onDispose: this.onDispose
        });
      }
      this.renderSiteHeader();
    }
  }
}
