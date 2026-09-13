import { IDropdownOption } from 'office-ui-fabric-react';
import { PNPService } from '../../../../common/services/PNPService';
import { FormOrigin } from '../Form.types';
import { UserService } from '../../../../common/services/UserContextService';

export interface IPaymentProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (boolean) => void;
    preselectedItemId?: number;
    formOrigin: FormOrigin;

}

export interface IPaymentState {
    isFormReady: boolean;
    name: string;
    summary: string;
    fees: IDropdownOption[];
    feesChoices: IDropdownOption[];
    beneficiary: string;
    beneficiaryChoices: IDropdownOption[];
    beneficiaryItem: IDropdownOption;
    beneficiaryItemChoices: IDropdownOption[];
    currency: IDropdownOption;
    currencyChoices: IDropdownOption[];
    contacts: IDropdownOption[];
    contactsChoices: IDropdownOption[];
    vatFreeCost: string;
    vat: string;
    vatRate: string;
    status: string;
    statusChoices: IDropdownOption[];
    dueDate: Date;
    datePaid: Date;
    comments: string;
    errors: { [key: string]: string };

}
