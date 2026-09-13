import { IDropdownOption } from "office-ui-fabric-react";
import { PNPService } from "../../../../common/services/PNPService";
import { ITaxonomyTerm } from "../../../../common/models/SPEntities";
import { UserService } from "../../../../common/services/UserContextService";




export interface IAuthorityProps {
  pnpService: PNPService;
  userService: UserService;
  itemId?: number;
  callback?: (string?) => void;
}



export interface IAuthorityyState {
  isFormReady: boolean;
  name: string;
  portal: string;
  type: string;
  typeChoices: IDropdownOption[];
  summary: string;
  address: string;
  grouping?: ITaxonomyTerm;
  comments: string;
  countries: IDropdownOption[];
  countriesChoices: IDropdownOption[];
  eutOwners: IDropdownOption[];
  eutOwnersChoices: IDropdownOption[];
  errors: { [key: string]: string };
}
