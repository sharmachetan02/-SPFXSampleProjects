import { IDropdownOption } from 'office-ui-fabric-react';
import { PNPService } from '../../../../common/services/PNPService';
import { UserService } from '../../../../common/services/UserContextService';

export interface IDistributionPartnerProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (string?) => void;
}

export interface IDistributionPartnerState {
    isFormReady: boolean;
    name: string;
    summary: string;
    address: string;
    comments: string;
    status: string;
    statusChoices: IDropdownOption[];
    errors: { [key: string]: string };
    portal: string;
    eutelsatOwnersChoices: IDropdownOption[];
    eutelsatOwners: IDropdownOption[];
}
