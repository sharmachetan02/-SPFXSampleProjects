export type FieldType = 'number' | 'string' | 'user' | 'date' | 'html' | 'managedMetadata' | 'lookup' | 'calculated' | 'boolean';
export interface ILookupValue {
    id: string;
    value: string;
}
export interface IUserValue {
    name: string;
    email: string;
}
export interface IManagedMetadataValue {
    label: string;
    termID: string;
}
export interface IUrlValue {
    title: string;
    url: string;
}
export interface ITaxonomyTerm {
    id: string;
    labels: {
      name: string;
      languageTag: string;
      isDefault: boolean;
    }[];
  }
