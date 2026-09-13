/* eslint-disable @typescript-eslint/no-use-before-define, no-use-before-define */
import { IDropdownOption } from "@fluentui/react";
import { RecurrenceData } from "../components/Recurrence Pattern Picker/RecurrencePatternPicker.types";
import { IReletedItem } from "../../webparts/genericForm/components/MA Requirement Generic/MARequirementGeneric.types";
export interface Iterm {
  Label: string;
  TermGuid: string;
}
export interface IBaseItem {
  Title?: string;
  Id?: number;
  Created?: Date;
  Modified?: Date;
}
export interface IContentType {
  Name?: string;
  Id?: string;
}
export interface Authority {
  // Define properties for Authority as needed
  Title?: string;
  Id?: number;
}

// Define DP interface or type
export interface DP extends IBaseItem {
  // Add properties as needed
  Title?: string;
  Id?: number;
  Description?: string;
}
export interface IItem {
  // Add properties as needed
  key?: string;
  text?: string;
}
export interface IItemOption extends IItem {
  id?: string;
  type?: string;
}
export interface ISanctionCategory extends IBaseItem {
  Summary?: string;
  Color?: string;
}
export interface ICountry extends IBaseItem {
  Identifier?: string;
  Grouping?: Iterm[];
}
export interface IMarket extends IBaseItem {
  Country?: ICountry,
  Grouping?: Iterm[];
  Identifier?: string;
  Capacity?: string;
  SanctionCategory?: ISanctionCategory;
  MaPriority?: string[];
  IsSanctioned?: boolean;
  Comments?: string;
  DocumentSpace?: string;

}
export interface IContentType {
  Name?: string;
  Id?: string;
}
export interface Authority {
  // Define properties for Authority as needed
  Title?: string;
  Id?: number;
}

// Define DP interface or type
export interface DP extends IBaseItem {
  // Add properties as needed
  Title?: string;
  Id?: number;
  Description?: string;
}
export interface IItem {
  // Add properties as needed
  key?: string;
  text?: string;
}
export interface IItemOption extends IItem {
  id?: string;
  type?: string;
}
export interface ISanctionCategory extends IBaseItem {
  Summary?: string;
  Color?: string;
}
export interface ICountry extends IBaseItem {
  Grouping?: Iterm[];
  Identifier?: string;
  Capacity?: string;
  SanctionCategory?: ISanctionCategory;
  MaPriority?: string[];
  IsSanctioned?: boolean;
  Comments?: string;
}
export interface IContact extends IBaseItem {
  Id?: number;
  Authority?: Authority[];
  DP?: DP[];
  LegalSvcProvider?: ILegalSvcProvider[];
  TP?: ITP[];
  JobTitle?: string;
  FirstName?: string;
  LastName?: string;
  Department?: string;
  Email?: string;
  PhoneNumber?: string;
  Summary?: string;
  Comments?: string;
  ContentType?: IContentType;
  Item?: IItemOption;
  EutelsatContact?: boolean;

}
export interface ILegalSvcProvider extends IBaseItem {
  Owners?: IContact[];
  Countries?: ICountry[];
  Status?: string;
  ContentType?: IContentType;
  Summary?: string;
  Comments?: string;
  Address?: string;
  ContentTypeName?: string;
  DocumentSpace?: string;
}
export interface ITP extends IBaseItem {
  Owners?: IContact[];
  Markets?: IMarket[];
  Status?: string;
  Summary?: string;
  Comments?: string;
  Address?: string;
  Portal?: string;
  Title?: string;
  DocumentSpace?: string;
}
export interface IEutelsatEntity extends IBaseItem {
  Owners?: IContact[];
  Country: ICountry;
  Status?: string;
  Summary?: string;
  Comments?: string;
  Address?: string;
}
export interface IDP extends IBaseItem {
  Owners?: IContact[];
  Contacts?: IContact[];
  Status?: string;
  Summary?: string;
  Comments?: string;
  Address?: string;
  Portal?: string;
  DocumentSpace?: string;
}

export interface IAuthority extends IBaseItem {
  Portal?: string;
  PortalDescription?: string;
  ContentType?: IContentType;
  Summary?: string;
  Address?: string;
  Comments?: string;
  Grouping?: Iterm;
  Countries?: ICountry[];
  Owners?: IContact[];
  Identifier?: string;
  DocumentSpace?: string;

}

export interface IContact extends IBaseItem {
  Id?: number;
  Authority?: Authority[];
  DP?: DP[];
  LegalSvcProvider?: ILegalSvcProvider[];
  TP?: ITP[];
  JobTitle?: string;
  FirstName?: string;
  LastName?: string;
  Department?: string;
  Email?: string;
  PhoneNumber?: string;
  Summary?: string;
  Comments?: string;
  ContentType?: IContentType;
  Item?: IItemOption;
}
export interface ILegalSvcProvider extends IBaseItem {
  Owners?: IContact[];
  Contacts?: IContact[];
  Countries?: ICountry[];
  Status?: string;
  ContentType?: IContentType;
  Summary?: string;
  Comments?: string;
  Address?: string;
}
export interface IEutelsatEntity extends IBaseItem {
  Owners?: IContact[];
  Country: ICountry;
  Status?: string;
  Summary?: string;
  Comments?: string;
  Address?: string;
  DocumentSpace?: string;
}
export interface IDP extends IBaseItem {
  Owners?: IContact[];
  Contacts?: IContact[];
  Status?: string;
  Summary?: string;
  Comments?: string;
  Address?: string;
  Portal?: string;
}

export interface IAuthority extends IBaseItem {
  Portal?: string;
  PortalDescription?: string;
  ContentType?: IContentType;
  Summary?: string;
  Address?: string;
  Comments?: string;
  Grouping?: Iterm;
  Countries?: ICountry[];
  Owners?: IContact[];
  Created?: Date;
  Modified?: Date;
  DocumentSpace?: string;
}

export interface ISNP extends IBaseItem {
  Country?: ICountry;
  Summary?: string;
  City?: string;
  ContentType?: IContentType;
  TeleportPartner?: ITP;
  Status?: string;
  EutEntity?: IEutelsatEntity;
  Comments?: string;
  DocumentSpace?: string;
}


export interface IFEE extends IBaseItem {
  Id?: number;
  Category?: string;
  Summary?: string;
  ContentType?: IContentType;
  Item?: IItemOption;
  //itemChoices: IDropdownOption[];
  ChargedBy?: string;
  ChargeByItem?: IItemOption;
  Currency?: ICurrency;
  VatFreeCost?: string;
  Vat?: string;
  VatRate?: string;
  CostType?: string;
  PoNonPO?: string;
  Status?: string;
  DueDate?: Date;
  Recurrence?: string;
  Frequency?: string;
  RecurranceFee?: boolean;
  recurrencePattern?: RecurrenceData;
  RecurrencPatternId?: number;
  Comments?: string;
  CostTotal?: string; //cutsom column for total
  DocumentSpace?: string;
}
export interface ICurrency extends IBaseItem {
  test?: string;
}

export interface IPayment extends IBaseItem {
  Summary: string;
  Fees?: IFEE[];
  Item?: IItemOption;
  Beneficiary?: string;
  BeneficiaryItem?: IItemOption;
  Currency?: ICurrency;
  Contacts?: IContact[];
  VatFreeCost?: string;
  Vat?: string;
  VatRate?: string;
  Status?: string;
  DueDate?: Date;
  DatePaid?: Date;
  Comments?: string;
  CostTotal?: string; //cutsom column for total
  DocumentSpace?: string;
}
export interface IMARequirementStatus extends IBaseItem {
  test?: string;
}

export interface IMARequirement extends IBaseItem {
  MaVertical?: IDropdownOption[];
  GeoLeo?: string;
  Comments?: string;
  Type?: IContentType;
  RequirementType?: string;
  Summary?: string;
  Markets?: IMarket[];
  ResponsibleParty?: string;
  ResponsiblePartyItem?: IItemOption;
  EutelsatOwner?: IContact[];
  Contacts?: IContact[];
  Synthesis?: IMARequirementStatus;
  RagStatus?: string;
  ApplicationDate?: Date;
  EstimatedDate?: Date;
  EffectiveDate?: Date;
  ExpirationDate?: Date;
  Letter1SentDate?: Date;
  Letter2SentDate?: Date;
  Letter1Milestone?: Date;
  Letter2Milestone?: Date;
  ResponseDate?: Date;
  ResponseSummary?: string;
  RecurrencePattern?: RecurrenceData;
  RequestedPartyItem?: IAuthority;
  RequestedPartyContacts?: IContact[];
  Reference?: string;
  TrialDemo?: boolean;
  RenewalTerm?: string;
  ReletedItems?: IReletedItem[];
  RecurrencPatternId?: number;
  Frequency?: string;
  DocumentSpace?: string;
  Snps?: ISNP[];
  RequestedPartyType?: string;
  ResponsiblePartyContacts?: IContact[];
}

export interface IMarketReadiness extends IBaseItem {
  Market?: IMarket;
  ResponsiblePartyType?: string;
  ResponsiblePartyItem?: IBaseItem;
  EutelsatOwners: IContact[];
  Verticals: IBaseItem[];

  AllVerticalRAGStatus: string;
  AllVerticalEstimatedDate: Date;
  AllVerticalEffectiveDate: Date;
  AllVerticalComments: string;

  SpaceRAGStatus: string;
  SpaceEstimatedDate: Date;
  SpaceEffectiveDate: Date;
  SpaceComments: string;

  LandFixedRAGStatus: string;
  LandFixedEstimatedDate: Date;
  LandFixedEffectiveDate: Date;
  LandFixedComments: string;

  LandMobilityRAGStatus: string;
  LandMobilityEstimatedDate: Date;
  LandMobilityEffectiveDate: Date;
  LandMobilityComments: string;

  MaritimeRAGStatus: string;
  MaritimeEstimatedDate: Date;
  MaritimeEffectiveDate: Date;
  MaritimeComments: string;

  AviationRAGStatus: string;
  AviationEstimatedDate: Date;
  AviationEffectiveDate: Date;
  AviationComments: string;
}

export interface ISNPReadiness extends IBaseItem {
  Snp?: ISNP;
  ResponsiblePartyType?: string;
  ResponsiblePartyItem?: IBaseItem;
  EutelsatOwners?: IContact[];
  RAGStatus?: string;
  EstimatedDate?: Date;
  EffectiveDate?: Date;
  Comments?: string;
}
export interface ILibrary extends IBaseItem {
  Download?: string;
  FileType?: string;
  Classification?: string;
  Edit?: string;
  Preview?: string;
  Open?: string;
  Delete?: string;
  IsFolder?: boolean;
  ServerRelativeUrl?: string;
  UniqueId?: string;
  Children?: ILibrary[];
  HasChildren?: boolean;
  Summary?: string;
  Version?: string;
  //Keywords?: { Id: number; Title: string }[];
  Keywords?: IKeyword[];
  Name?: string;
  //Type?: { Id: number; Title: string };
  Type?: IFileType;
  Actions?: string;
  Category?: IFileType;
}


export interface ILibrary extends IBaseItem {
  Download?: string;
  FileType?: string;
  Classification?: string;
  Edit?: string;
  Preview?: string;
  Open?: string;
  Delete?: string;
  IsFolder?: boolean;
  ServerRelativeUrl?: string;
  UniqueId?: string;
  Children?: ILibrary[];
  IsChildren?: boolean;
  Summary?: string;
  Version?: string;
  //Keywords?: { Id: number; Title: string }[];
  Keywords?: IKeyword[];
  Name?: string;
  //Type?: { Id: number; Title: string };
  Type?: IFileType;
  Actions?: string;
  Category?: IFileType;
}
export interface IDocument {
  name: string;
  extension: string;
  download: string;
  contentTypeId: string;
  uniqueId: string;
  sheetUrl: string;
  ServerRelativeUrl?: string;
  Title: string;
  Name?: string;
}
export interface IFileType extends IBaseItem {
  Id?: number;
  Title?: string;
  Category?: string;
}
export interface IKeyword extends IBaseItem {
  Id?: number;
  Title?: string;
}
