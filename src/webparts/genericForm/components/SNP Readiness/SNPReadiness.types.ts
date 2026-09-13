import { IDropdownOption } from '@fluentui/react';
import { PNPService } from '../../../../common/services/PNPService';
import { UserService } from '../../../../common/services/UserContextService';

export interface ISNPReadinessProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (boolean) => void;
    preselectedItemId?: number;
}

export interface ISNPReadinessState {
    isFormReady: boolean;
    name: string;
    responsiblePartyTypeChoices: IDropdownOption[];
    responsiblePartyType: string;
    responsiblePartyItemChoices: IDropdownOption[];
    responsiblePartyItem: IDropdownOption;
    errors: { [key: string]: string };
    eutelsatOwnersChoices: IDropdownOption[];
    eutelsatOwners: IDropdownOption[];
    readinessRAGStatus: string;
    readinessRAGStatusChoices: IDropdownOption[];
    readinessEstimatedDate?: Date;
    readinessEffectiveDate?: Date;
    readinessComments?: string;
    snpItem: IDropdownOption;
    snpItemChoices: IDropdownOption[];
}
