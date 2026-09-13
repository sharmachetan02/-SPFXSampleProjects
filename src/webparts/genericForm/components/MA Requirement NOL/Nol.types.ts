import { IDropdownOption } from '@fluentui/react';
import { PNPService } from '../../../../common/services/PNPService';
import { MARequirementBasicData } from '../MA Requirement Basic/MARequirementBasic.types';
import { UserService } from '../../../../common/services/UserContextService';
export interface INolRequirementProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (boolean) => void;
    formType?: string;
    formItemId?: number;
}

export interface INolRequirementState {
    maRequirementBasicData: MARequirementBasicData;
    letter1SentDate: Date;
    letter2SentDate: Date;
    letter1Milestone: Date;
    letter2Milestone: Date;
    responseDate: Date;
    responseSummary: string;
    isFormReady: boolean;
    errors: { [key: string]: string };
    requestedPartyChoices: IDropdownOption[];
    requestedParty: string;
    requestedPartyItem: IDropdownOption;
    requestedPartyItemChoices: IDropdownOption[];
    requestedPartyContacts: IDropdownOption[];
    requestedPartyContactsChoices: IDropdownOption[];
}
