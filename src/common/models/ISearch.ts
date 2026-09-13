import { ISort } from '@pnp/sp/search';
import { FieldType } from './SPEntities';

export interface ISearchResult {
    [key: string]: string;
}
export interface ISearchResults {
    RelevantResults: ISearchResult[];
    TotalRows?: number;
}
export interface ISearchMapping {
    searchProperty:string;
    field:string;
    type:FieldType;
  }
export interface ISearchConfig {
    query: string;
    sort?: ISort[];
    resultsCount: number;
    startRow?:number;
    selectedProperties: string[];
    queryTemplate: string;
    resultSourceId: string;
    enableQueryRules: boolean;
}
