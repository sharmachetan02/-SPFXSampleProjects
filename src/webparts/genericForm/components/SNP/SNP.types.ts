import { IDropdownOption } from "office-ui-fabric-react";
import { PNPService } from "../../../../common/services/PNPService";
import { UserService } from "../../../../common/services/UserContextService";




export interface ISNPProps {
  pnpService: PNPService;
  userService: UserService;
  itemId?: number;
  callback?: (string?) => void;
}



export interface ISNPState {
  isFormReady: boolean;
  name: string;
  type: string;
  typeChoices: IDropdownOption[];
  summary: string;
  city: string;
  status: string;
  statusChoices: IDropdownOption[];
  comments: string;
  country: IDropdownOption;
  countryChoices: IDropdownOption[];
  eutelsatEntity: IDropdownOption;
  eutelsatEntityChoices: IDropdownOption[];
  teleportPartner: IDropdownOption;
  teleportPartnerChoices: IDropdownOption[];
  errors: { [key: string]: string };
}
