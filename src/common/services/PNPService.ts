import '@pnp/sp/taxonomy';
import { IOrderedTermInfo, ITerms, ITermSet } from '@pnp/sp/taxonomy';
import { SPBrowser, spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/graph/users';
import '@pnp/graph/photos';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { IItems } from '@pnp/sp/items';
import { IList } from '@pnp/sp/lists';
import '@pnp/sp/site-users/web';
import { ISiteUser, IWebEnsureUserResult } from '@pnp/sp/site-users/types';
import { IWeb } from '@pnp/sp/webs';
import { ISearchQuery, SearchResults } from '@pnp/sp/search';
import '@pnp/sp/search';
import '@pnp/sp/fields';
import { IFolder } from '@pnp/sp/folders';
import { IField } from '@pnp/sp/fields';
export class PNPService {
  private spContext: BaseComponentContext;
  private sp: SPFI = null;
  private url = null;
  constructor(context: BaseComponentContext, webUrl?: string) {
    this.spContext = context;
    if (!webUrl) webUrl = context.pageContext.web.absoluteUrl;
    this.url = webUrl;
    this.sp = spfi(webUrl)
      .using(SPBrowser({ baseUrl: webUrl }))
      .using(SPFx({ pageContext: context.pageContext }));
  }

  public getContext(): BaseComponentContext {
    return this.spContext;
  }
  public getUrl(): string {
    return this.url;
  }

  public getTermsFromTermSet(termSetId: string): ITerms {
    return this.sp.termStore.sets.getById(termSetId).terms;
  }
  public getTermsTreeFromTermSet(termSetId: string): Promise<IOrderedTermInfo[]> {
    return this.sp.termStore.sets.getById(termSetId).getAllChildrenAsOrderedTree({ retrieveProperties: true });
  }
  public getTermSet(termSetId: string): ITermSet {
    return this.sp.termStore.sets.getById(termSetId);
  }

  public getListItems(listRelativeUrl: string, topNumber = 5000): IItems {
    return this.sp.web.getList(listRelativeUrl).items.top(topNumber);
  }

  public getFolder(folderRelativeUrl: string): IFolder {
    return this.sp.web.getFolderByServerRelativePath(folderRelativeUrl);
  }

  public getListByUrl(listRelativeUrl: string): IList {
    return this.sp.web.getList(listRelativeUrl);
  }

  public getListField(listRelativeUrl: string, fieldName: string): IField {
    return this.sp.web.getList(listRelativeUrl).fields
      .getByInternalNameOrTitle(fieldName);
  }

  public getListById(listId: string): IList {
    return this.sp.web.lists.getById(listId);
  }

  public getFileByUrl(fileUrl: string) {
    return this.sp.web.getFileByUrl(fileUrl);
  }

  public getCurrentUser(): ISiteUser {
    return this.sp.web.currentUser;
  }

  public getWeb(): IWeb {
    return this.sp.web;
  }
  public ensureUser(loginName: string): Promise<IWebEnsureUserResult> {
    return this.sp.web.ensureUser(loginName);
  }

  public async search(searchQuery: ISearchQuery): Promise<SearchResults> {
    const searchResults: Promise<SearchResults> = this.sp.search(searchQuery);
    return searchResults;
  }
}
