/* eslint-disable no-var */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-use-before-define */

import { PnPClientStorage, dateAdd } from '@pnp/common';
import DomHelper from './DomHelper';
import { UtilHelper } from './Util';

export enum storageKeys {
  EUT_MERGE_PORTAL_IS_EDIT_MODE_ACTIVE,
  EUT_MERGE_PORTAL_UI_CULTURE_NAME,
  EUT_MERGE_PORTAL_QUICK_APPS_DATA,
  EUT_MERGE_PORTAL_MY_APPS_DATA,
  EUT_MERGE_PORTAL_ESSENTIAL_INFO,
  EUT_MERGE_PORTAL_ONE_TEAM_DATA,
  EUT_MERGE_PORTAL_GROUP_EVENTS_DATA,
  EUT_MERGE_PORTAL_GROUP_VIDEOS_DATA,
  EUT_MERGE_PORTAL_GROUP_NEWS,
  EUT_MERGE_PORTAL_LOCAL_NEWS,
  EUT_MERGE_PORTAL_ALL_NEWS_PARSED_SEARCH_RESULTS,
  EUT_MERGE_PORTAL_ALL_NEWS_RENDITIONS,
  EUT_MERGE_PORTAL_M365_APPS,
  EUT_MERGE_PORTAL_ORG_TREE_DATA,
  EUT_MERGE_ONE_TEAM_DATA,
  EUT_MERGE_PORTAL_MY_APPS_WIDGET
}
declare var _ENVIRONMENT_: string;

enum storageType {
  Local,
  Session,
}

export class ClientStorageHelper {
  private static instance: ClientStorageHelper;
  private  isReady = false;
  private  store: PnPClientStorage;

  public storageType = storageType;

  private constructor() {
    this.store = new PnPClientStorage();
    this.store.session.deleteExpired();
    this.store.local.deleteExpired();
  }

  public static async getInstance(): Promise<ClientStorageHelper> {
    if (!ClientStorageHelper.instance) {
      ClientStorageHelper.instance = new ClientStorageHelper();
      ClientStorageHelper.instance.isReady = true;
      return ClientStorageHelper.instance;
    } else  if (ClientStorageHelper.instance.isReady) {
      return ClientStorageHelper.instance;
    } else {
      await DomHelper.delay(250);
      return await ClientStorageHelper.getInstance();
    }
  }

  public getValue(storage: storageType, key:storageKeys): any {

    let value = undefined;
    const storageKey = `${storageKeys[key]}_${_ENVIRONMENT_.toUpperCase()}`;
    if (storage === storageType.Local) {
      value = this.store.local.get(storageKey);
    }
    if (storage === storageType.Session) {
      value = this.store.session.get(storageKey);
    }
    return value;
  }

  public putValue(storage: storageType, key:storageKeys, item:any, expiresInMinutes:number) {
    let expirationDate = dateAdd(new Date(), 'minute', expiresInMinutes);
    const storageKey = `${storageKeys[key]}_${_ENVIRONMENT_.toUpperCase()}`;
    if (!UtilHelper.isToday(expirationDate)) {
      expirationDate = UtilHelper.getTomorrowDate();
    }
    if (storage === storageType.Local) {
      this.store.local.put(storageKey, item, expirationDate);
    }
    if (storage === storageType.Session) {
      this.store.session.put(storageKey, item, expirationDate);
    }
  }
  public deleteValue(storage: storageType, key:storageKeys): any {
    const storageKey = `${storageKeys[key]}_${_ENVIRONMENT_.toUpperCase()}`;
    let value = undefined;
    if (storage === storageType.Local) {
      value = this.store.local.delete(storageKey);
    }
    if (storage === storageType.Session) {
      value = this.store.session.delete(storageKey);
    }
    return value;
  }
}
