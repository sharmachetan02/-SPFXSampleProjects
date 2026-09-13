import { IDropdownOption } from 'office-ui-fabric-react';
import { PNPService } from '../../../../common/services/PNPService';
import { FormType } from '../Form.types';
import { UserService } from '../../../../common/services/UserContextService';

export interface IEutelsatEntityProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    preselectedItemId?: number;
    preselectedItemType?: FormType;
    callback?: (string?) => void;
}

export interface IEutelsatEntityState {
    isFormReady: boolean;
    name: string;
    summary: string;
    address: string;
    comments: string;
    status: string;
    statusChoices: IDropdownOption[];
    country: IDropdownOption;
    countryChoices: IDropdownOption[];
    errors: { [key: string]: string };
    openItems: string[];
}
