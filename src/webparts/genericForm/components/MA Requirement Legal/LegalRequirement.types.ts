import { IDropdownOption } from '@fluentui/react';
import { PNPService } from '../../../../common/services/PNPService';
import { MARequirementBasicData } from '../MA Requirement Basic/MARequirementBasic.types';
import { UserService } from '../../../../common/services/UserContextService';


export interface ILegalRequirementProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (boolean) => void;
    formType?: string;
    formItemId?: number;
}
export interface ILegalRequirementState {
    maRequirementBasicData: MARequirementBasicData;
    applicationDate: Date;
    estimatedDate: Date;
    effectiveDate: Date;
    isFormReady: boolean;
    errors: { [key: string]: string };
    requestedPartyChoices: IDropdownOption[];
    requestedParty: string;
    requestedPartyItem: IDropdownOption;
    requestedPartyItemChoices: IDropdownOption[];
    requestedPartyContacts: IDropdownOption[];
    requestedPartyContactsChoices: IDropdownOption[];
}
