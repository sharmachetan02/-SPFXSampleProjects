import { IDropdownOption } from '@fluentui/react';
import { PNPService } from '../../../../common/services/PNPService';
import { UserService } from '../../../../common/services/UserContextService';

export interface IMarketReadinessProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (boolean) => void;
    preselectedItemId?: number;
}

export interface IMarketReadinessState {
    isFormReady: boolean;
    name: string;
    market: IDropdownOption;
    marketChoices: IDropdownOption[];
    responsiblePartyTypeChoices: IDropdownOption[];
    responsiblePartyType: string;
    responsiblePartyItemChoices: IDropdownOption[];
    responsiblePartyItem: IDropdownOption;
    errors: { [key: string]: string };
    eutelsatOwnersChoices: IDropdownOption[];
    eutelsatOwners: IDropdownOption[];
    verticals: IDropdownOption[];
    verticalsChoices: IDropdownOption[];
    accordionOpenItems: string[];
    verticalRAGStatus: string;
    verticalRAGStatusChoices: IDropdownOption[];
    verticalEstimatedDate: Date;
    verticalEffectiveDate: Date;
    verticalComments: string;

    spaceRAGStatus: string;
    spaceRAGStatusChoices: IDropdownOption[];
    spaceEstimatedDate: Date;
    spaceEffectiveDate: Date;
    spaceComments: string;

    landFixedRAGStatus: string;
    landFixedRAGStatusChoices: IDropdownOption[];
    landFixedEstimatedDate: Date;
    landFixedEffectiveDate: Date;
    landFixedComments: string;

    landMobilityRAGStatus: string;
    landMobilityRAGStatusChoices: IDropdownOption[];
    landMobilityEstimatedDate: Date;
    landMobilityEffectiveDate: Date;
    landMobilityComments: string;

    maritimeRAGStatus: string;
    maritimeRAGStatusChoices: IDropdownOption[];
    maritimeEstimatedDate: Date;
    maritimeEffectiveDate: Date;
    maritimeComments: string;

    aviationRAGStatus: string;
    aviationRAGStatusChoices: IDropdownOption[];
    aviationEstimatedDate: Date;
    aviationEffectiveDate: Date;
    aviationComments: string;
}
