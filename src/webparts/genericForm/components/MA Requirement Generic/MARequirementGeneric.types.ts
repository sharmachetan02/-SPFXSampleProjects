
import { IDropdownOption, } from '@fluentui/react';
import { PNPService } from '../../../../common/services/PNPService';
import { MARequirementBasicData } from '../MA Requirement Basic/MARequirementBasic.types';
import { UserService } from '../../../../common/services/UserContextService';


export interface IMARequirementGenericProps {
    pnpService: PNPService;
    userService: UserService;
    itemId?: number;
    callback?: (boolean) => void;
    formType?: string;
    formItemId?: number;
}

export interface IReletedItem {
    Id?: string;
    relationShip: string;
    releted: string;
    reletedChoices: IDropdownOption[];
    reletedItem: IDropdownOption[];
    reletedItemChoices?: IDropdownOption[];
    datasheetType?: string;
    legalSvcProvider?: { Id: string; Title: string }[];
    authority?: { Id: string; Title: string }[];
    distributionPartner?: { Id: string; Title: string }[];
    snp?: { Id: string; Title: string }[];
}
export interface IMARequirementGenericState {
    maRequirementBasicData: MARequirementBasicData;
    applicationDate: Date;
    estimatedDate: Date;
    effectiveDate: Date;
    reletedItems: IReletedItem[];
    isFormReady: boolean;
    errors: { [key: string]: string };

}
