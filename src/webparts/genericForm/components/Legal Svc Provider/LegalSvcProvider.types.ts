import { IDropdownOption } from 'office-ui-fabric-react';
import { PNPService } from '../../../../common/services/PNPService';
import { UserService } from '../../../../common/services/UserContextService';

export interface ILegalSvcProviderProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (string?) => void;
}

export interface ILegalSvcProviderState {
    isFormReady: boolean;
    name: string;
    summary: string;
    address: string;
    comments: string;
    status: string;
    statusChoices: IDropdownOption[];
    countries: IDropdownOption[];
    countryChoices: IDropdownOption[];
    errors: { [key: string]: string };
    typeChoices: IDropdownOption[];
    type: string;
    eutelsatOwnersChoices: IDropdownOption[];
    eutelsatOwners: IDropdownOption[];
}
