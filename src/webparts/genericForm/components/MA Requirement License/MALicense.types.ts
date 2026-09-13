
//import { IDropdownOption } from 'office-ui-fabric-react';
import { RecurrenceData } from '../../../../common/components/Recurrence Pattern Picker/RecurrencePatternPicker.types';
import { PNPService } from '../../../../common/services/PNPService';
import { IDropdownOption, } from '@fluentui/react/lib/Dropdown';
import { MARequirementBasicData } from '../MA Requirement Basic/MARequirementBasic.types';
import { UserService } from '../../../../common/services/UserContextService';

export interface ILicenseProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (boolean) => void;
    formType?: string;
    formItemId?: number;
}


export interface ILicenseState {
    maRequirementBasicData: MARequirementBasicData;
    applicationDate: Date;
    estimatedDate: Date;
    effectiveDate: Date;
    recurrencePattern: RecurrenceData;
    recurrencPatternId?: number;
    isFormReady: boolean;
    errors: { [key: string]: string };
    requestedPartyChoices: IDropdownOption[];
    requestedParty: string;
    requestedPartyItem: IDropdownOption;
    requestedPartyItemChoices: IDropdownOption[];
    requestedPartyContacts: IDropdownOption[];
    requestedPartyContactsChoices: IDropdownOption[];
    reference: string;
    trialDemo: boolean;
    renewalTerm: string;
    renewalTermChoices: IDropdownOption[];
    expirationDate: Date;

}
