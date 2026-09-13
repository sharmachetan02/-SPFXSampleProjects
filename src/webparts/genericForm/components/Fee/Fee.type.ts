import { IDropdownOption } from 'office-ui-fabric-react';
import { PNPService } from '../../../../common/services/PNPService';
import { RecurrenceData } from '../../../../common/components/Recurrence Pattern Picker/RecurrencePatternPicker.types';
import { UserService } from '../../../../common/services/UserContextService';

export interface IFeeProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (boolean) => void;
    preselectedItemType?: string;
    preselectedItemId?: number;
}

export interface IFeeState {
    isFormReady: boolean;
    name: string;
    category: string;
    categoryChoices: IDropdownOption[];
    summary: string;
    typeChoices: IDropdownOption[];
    type: string;
    item: IDropdownOption;
    itemChoices: IDropdownOption[];
    chargedBy: string;
    chargedByChoices: IDropdownOption[];
    chargeByItem: IDropdownOption;
    chargeByItemChoices: IDropdownOption[];
    currency: IDropdownOption;
    currencyChoices: IDropdownOption[];
    vatFreeCost: string;
    vat: string;
    vatRate: string;
    costType: string;
    costTypeChoices: IDropdownOption[];
    poNonPO: string;
    poNonPOChoices: IDropdownOption[];
    status: string;
    statusChoices: IDropdownOption[];
    dueDate: Date;
    recurrenceFee: boolean;
    recurrencePattern: RecurrenceData;
    recurrencPatternId?: number;
    comments: string;
    errors: { [key: string]: string };

}
