import { IDropdownOption } from 'office-ui-fabric-react';
import { PNPService } from '../../../../common/services/PNPService';
import { FormOrigin } from '../Form.types';
import { UserService } from '../../../../common/services/UserContextService';

export interface IContactProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (string?) => void;
    preselectedItemId?: number;
    preselectedType?: string;
    formOrigin: FormOrigin;
}

export interface IContactState {
    isFormReady: boolean;
    title: string;
    firstname: string;
    lastname: string;
    summary: string;
    comments: string;
    errors: { [key: string]: string };
    typeChoices: IDropdownOption[];
    type: string;
    jobtitle: string;
    department: string;
    email: string;
    phonenumber: string;
    item: IDropdownOption;
    itemChoices: IDropdownOption[];
    isItemPreselected: boolean;

}
