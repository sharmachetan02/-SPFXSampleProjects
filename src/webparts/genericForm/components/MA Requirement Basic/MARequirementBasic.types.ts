import { IDropdownOption } from '@fluentui/react/lib/Dropdown';
export interface ICategorizedDropdownOption extends IDropdownOption {
  maVerticals?: string[];
  category?:string;
}

export interface MARequirementBasicData {
  maVertical: IDropdownOption[];
  maVerticalChoices: IDropdownOption[];
  snps: IDropdownOption[];
  snpsChoices: IDropdownOption[];
  geoLeo: string;
  geoLeoChoices: IDropdownOption[];
  comments: string;
  name: string;
  type: string;
  typeChoices: IDropdownOption[];
  requirementType: string;
  requirementTypeChoices: ICategorizedDropdownOption[];
  summary: string;
  markets: IDropdownOption[];
  marketsChoices: IDropdownOption[];
  responsibleParty: string;
  responsiblePartyChoices: IDropdownOption[];
  responsiblePartyItem: IDropdownOption;
  responsiblePartyItemChoices: IDropdownOption[];
  eutelsatOwners: IDropdownOption[];
  eutelsatOwnersChoices: IDropdownOption[];
  responsiblePartyContacts: IDropdownOption[];
  responsiblePartyContactsChoices: IDropdownOption[];
  synthesisStatus: string;
  synthesisStatusChoices: IDropdownOption[];
  ragStatus: string;
  ragStatusChoices: IDropdownOption[];
}

export interface MARequirementBasicProps {
  maRequirementBasicData?: Partial<MARequirementBasicData>;
  onChange: (data: MARequirementBasicData) => void;
  errors?: { [key: string]: string };
}

export interface MARequirementBasicState {
  maVertical: IDropdownOption[];
  maVerticalChoices: IDropdownOption[];
  snps: IDropdownOption[];
  snpsChoices: IDropdownOption[];
  geoLeo: string;
  geoLeoChoices: IDropdownOption[];
  comments: string;
  name: string;
  type: string;
  typeChoices: IDropdownOption[];
  requirementType: string;
  requirementTypeChoices: ICategorizedDropdownOption[];
  summary: string;
  markets: IDropdownOption[];
  marketsChoices: IDropdownOption[];
  responsibleParty: string;
  responsiblePartyChoices: IDropdownOption[];
  responsiblePartyItem: IDropdownOption;
  responsiblePartyItemChoices: IDropdownOption[];
  eutelsatOwners: IDropdownOption[];
  eutelsatOwnersChoices: IDropdownOption[];
  responsiblePartyContacts: IDropdownOption[];
  responsiblePartyContactsChoices: IDropdownOption[];
  synthesisStatus: string;
  synthesisStatusChoices: IDropdownOption[];
  ragStatus: string;
  ragStatusChoices: IDropdownOption[];
}
