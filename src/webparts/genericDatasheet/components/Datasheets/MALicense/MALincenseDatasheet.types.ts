import { IFEE, IMARequirement } from "../../../../../common/models/IBusiness";
import { PNPService } from "../../../../../common/services/PNPService";
import { UserService } from "../../../../../common/services/UserContextService";
import { IMessageBanner } from "../../../../genericForm/components/Form.types";
export interface IPaginationProperties {
    isEnabled: boolean;
    showTopPagination?: boolean;
    showBottomPagination?: boolean;
}
export interface IMALicenseRequirementDatasheetProps {
    pnpService: PNPService;
    userService: UserService;
    itemId: number;
    pagination?: IPaginationProperties;
}

export interface IMALicenseRequirementDatasheetState {
    isDatasheetReady: boolean;
    showEditFormDialog: boolean;
    showDeleteDialog: boolean;
    item: IMARequirement;
    feeViewItems?: IFEE[];
    errorMessage: string | null;
    messageBanner: IMessageBanner;
    isFormProcessing: boolean;
    isConfirmButtonDisabled: boolean;
}
