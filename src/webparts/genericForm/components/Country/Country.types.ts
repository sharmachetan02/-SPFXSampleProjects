import { IDropdownOption } from "office-ui-fabric-react";
import { PNPService } from "../../../../common/services/PNPService";
import { ITaxonomyTerm } from "../../../../common/models/SPEntities";
import { UserService } from "../../../../common/services/UserContextService";

export interface ISanctionOption extends IDropdownOption {
  color: string;
  summary: string;
}

export interface ICountryProps {
  pnpService: PNPService;
  userService: UserService;
  itemId?: number;
  callback?: (string?) => void;
}

export interface ICountryyState {
  isFormReady: boolean;
  name: string;
  identifier: string;
  groupings?: ITaxonomyTerm[];
  errors: { [key: string]: string };
}


