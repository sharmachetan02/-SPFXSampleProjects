import { IDropdownOption } from "office-ui-fabric-react";
import { PNPService } from "../../../../common/services/PNPService";
import { ITaxonomyTerm } from "../../../../common/models/SPEntities";
import { UserService } from "../../../../common/services/UserContextService";

export interface ISanctionOption extends IDropdownOption {
  color: string;
  summary: string;
}

export interface IMarketProps {
  pnpService: PNPService;
  userService: UserService;
  itemId?: number;
  callback?: (string?) => void;
}

export interface IMarketState {
  isFormReady: boolean;
  name: string;
  country: IDropdownOption;
  countryChoices: IDropdownOption[];
  groupings?: ITaxonomyTerm[];
  isSanctioned: boolean;
  sanctionCategory: ISanctionOption;
  sanctionCategoryChoices: ISanctionOption[];
  maPriorities: IDropdownOption[];
  maPrioritiesChoices: IDropdownOption[];
  maCapacity: string;
  comments: string;
  errors: { [key: string]: string };
}


