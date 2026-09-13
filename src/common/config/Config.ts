// eslint-disable-next-line no-var
declare var _ENVIRONMENT_: string;
import data from './Config.json';

 interface ITAXONOMY{
  GROUPING_TERMSETID:string;
  NAVIGATION_TERMSETID:string;
  NAVIGATION_CONTENT_TERMID:string;
}
 interface IADMINISTRATION{
  WEB_URL: string;
}
interface ISITE{
  ROOT_WEB_URL: string;
  SITE_TITLE: string;
}

export class Config {
  public static get TAXONOMY(): ITAXONOMY {
    return data.ENVIRONMENTS[_ENVIRONMENT_.toUpperCase()].TAXONOMY;
  }
  public static get ADMINISTRATION_SITE(): IADMINISTRATION {
    return data.ENVIRONMENTS[_ENVIRONMENT_.toUpperCase()].ADMINISTRATION_SITE;
  }
  public static get SITE(): ISITE {
    return data.ENVIRONMENTS[_ENVIRONMENT_.toUpperCase()].SITE;
  }
  public static get Environment(): string {
    return _ENVIRONMENT_;
  }
}
