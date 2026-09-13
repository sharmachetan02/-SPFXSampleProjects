/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable no-use-before-define */
/* eslint-disable max-len */
/* eslint-disable no-loop-func */
/* eslint-disable no-inner-declarations */
import { ISearchConfig, ISearchMapping, ISearchResult, ISearchResults } from '../models/ISearch';
import { ISearchQuery, SearchResults } from '@pnp/sp/search';
import { PNPService } from '../services/PNPService';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { ILookupValue, IUserValue, IManagedMetadataValue, IUrlValue } from '../models/SPEntities';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
class SearchHelper {

  public static async search(pnpService: PNPService, searchConfig: ISearchConfig): Promise<ISearchResults> {

    const searchQuery: ISearchQuery = {};
    let searchResult: SearchResults = null;
    searchQuery.ClientType = 'ContentSearchRegular';
    searchQuery.Querytext = searchConfig.query;
    //searchQuery.EnableQueryRules = searchConfig.enableQueryRules ? searchConfig.enableQueryRules : false;

    if (searchConfig.resultSourceId) {
      searchQuery.SourceId = searchConfig.resultSourceId;
    } else {
      searchQuery.QueryTemplate = searchConfig.queryTemplate;
    }

    searchQuery.RowLimit = searchConfig.resultsCount ? searchConfig.resultsCount : 50;
    searchQuery.SelectProperties = searchConfig.selectedProperties;
    searchQuery.TrimDuplicates = false;
    if (searchConfig.sort) searchQuery.SortList = searchConfig.sort;
    //searchQuery.RowsPerPage = searchConfig.resultsCount;
    searchQuery.StartRow = searchConfig.startRow ? searchConfig.startRow : 0;
    searchResult =  await pnpService.search(searchQuery);
    const relevantResults: ISearchResult[] = SearchHelper.getRelevantResults(searchResult);
    return { RelevantResults: relevantResults, TotalRows: searchResult.TotalRows };
  }
  public static getRelevantResults(searchResult: SearchResults): ISearchResult[] {
    const relevantResults: ISearchResult[] = [];
    if (searchResult.RawSearchResults.PrimaryQueryResult) {
      const resultRows = searchResult.RawSearchResults.PrimaryQueryResult.RelevantResults.Table.Rows;
      resultRows.map((elt) => {
        const result: ISearchResult = {};
        elt.Cells.map((item) => {
          result[item.Key] = item.Value;
        });
        relevantResults.push(result);
      });
    }
    return relevantResults;
  }
  public static async searchGet(pnpService: PNPService, searchConfig: ISearchConfig) : Promise<ISearchResults> {
    const relevantResults: ISearchResult[] = [];
    let url: string = pnpService.getContext().pageContext.web.absoluteUrl + '/_api/search/query?querytext=';
    //const sort = searchConfig.sort.map((s) => `${s.Property}:${s.Direction}`).join(',');
    url += `'${searchConfig.query}'`;
    url += '&selectproperties=';
    url += `'${searchConfig.selectedProperties.join(',')}'`;
    url += '&rowlimit=';
    url += `'${searchConfig.resultsCount}'`;
    url += '&sourceId=';
    url += `'${searchConfig.resultSourceId}'`;
    url += '&startrow=';
    url += `'${searchConfig.startRow ? searchConfig.startRow : 0}'`;
    //url += searchConfig.sort.length ? `&sortlist='${sort}'` : '';

    url += '&clienttype=\'ContentSearchRegular\'';

    const promise = await new Promise<ISearchResults>((resolve) => {
      pnpService.getContext().spHttpClient.get(url, SPHttpClient.configurations.v1)
        .then((res: SPHttpClientResponse) => res.json())
        .then((searchResult) => {
          if (searchResult.PrimaryQueryResult) {
            const resultRows = searchResult.PrimaryQueryResult.RelevantResults.Table.Rows;
            resultRows.map((elt) => {
              const result: ISearchResult = {};
              elt.Cells.map((item) => {
                result[item.Key] = item.Value;
              });
              relevantResults.push(result);
            });
          }
          resolve({ RelevantResults: relevantResults, TotalRows: searchResult.PrimaryQueryResult.RelevantResults.TotalRows });
        });
    });
    return promise;
  }
  public static  formatManagedMetadataField(tags: string):IManagedMetadataValue[] {
    const tagsArray = [];
    if (!tags) return tagsArray;
    if (tags.indexOf(';')) {
      tags.split(';').forEach((value) => {
        const splittedValue = value.split('|#');
        if (splittedValue[0] === 'L0') {
          tagsArray.push({
            label: splittedValue[1].split('|')[1],
            termID: splittedValue[1].split('|')[0]
          });
        }
      });
    } else {
      const splittedValue = tags.split('|#');
      if (splittedValue[0] === 'L0') {
        tagsArray.push({
          label: splittedValue[1].split('|')[1],
          termID: splittedValue[1].split('|')[0]
        });
      }
    }
    return tagsArray;
  }
  public static formatUserField(userValue: string):IUserValue[] | IUserValue {
    if (!userValue) return;
    const userArray = [];
    if (userValue.indexOf(';')) {
      userValue.split(';').forEach((value) => {
        const tag : IUserValue = {
          email:  value.split('|')[0],
          name: value.split('|')[1]
        };
        userArray.push(tag);
      });
    } else {
      const tag : IUserValue = {
        email:  userValue.split('|')[0],
        name: userValue.split('|')[1]
      };
      userArray.push(tag);
    }
    return userArray.length === 1 ?  userArray[0] : userArray;
  }
  public static formatLookupField(lookupValue: string):ILookupValue {
    const lookup : ILookupValue = {
      id: '',
      value: ''
    };
    if (lookupValue.indexOf(';#') > -1) {
      lookup.id = lookupValue.split(';#')[0];
      lookup.value = lookupValue.split(';#')[1];
    }
    return lookup;
  }
  public static formatUrlField(urlField: string):IUrlValue {
    if (isEmpty(urlField)) return;
    const separatorPos = urlField.lastIndexOf(',');
    if (separatorPos === -1) return;

    const urlValue : IUrlValue = {
      title:  urlField.substr(separatorPos + 1).trim(),
      url: urlField.substr(0, separatorPos)
    };

    return urlValue;

  }
  public static formatCalculatedField(calculatedField: string):string {
    if (isEmpty(calculatedField)) return;
    const separatorPos = calculatedField.lastIndexOf(';#');
    if (separatorPos === -1) return;

    return calculatedField.substr(separatorPos + 2).trim();

  }
  public static formatDateField(dateField):Date {
    return new Date(dateField);
  }
  public static formatNumberField(numberField):number {
    return parseInt(numberField);
  }
  public static formatBooleanField(boolField):boolean {
    return (boolField as string)?.toLowerCase() === 'true';
  }
  public static formatField(value: string, mapping:ISearchMapping) {
    let formattedValue = {};
    switch (mapping.type) {
    case 'lookup':{ formattedValue = this.formatLookupField(value); break; }
    case 'managedMetadata': { formattedValue = this.formatManagedMetadataField(value); break; }
    case 'user':{ formattedValue = this.formatUserField(value); break; }
    case 'date':{ formattedValue = this.formatDateField(value); break; }
    case 'number':{ formattedValue = this.formatNumberField(value); break; }
    case 'calculated':{ formattedValue = this.formatCalculatedField(value); break; }
    case 'boolean':{ formattedValue = this.formatBooleanField(value); break; }

    default: { formattedValue = (value ?? '').toString(); }
    }
    return formattedValue;
  }

}
export default SearchHelper;

