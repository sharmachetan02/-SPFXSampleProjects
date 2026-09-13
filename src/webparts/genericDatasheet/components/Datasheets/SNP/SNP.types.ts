import { IFEE, IMARequirement, ISNP, ISNPReadiness } from "../../../../../common/models/IBusiness";
import { PNPService } from "../../../../../common/services/PNPService";
import { UserService } from "../../../../../common/services/UserContextService";
import { IMessageBanner } from "../../../../genericForm/components/Form.types";
export interface IPaginationProperties {
    isEnabled: boolean;
    showTopPagination?: boolean;
    showBottomPagination?: boolean;
}
export interface ISNPDatasheetProps {
    pnpService: PNPService;
      userService: UserService;
    itemId: number;
    pagination?: IPaginationProperties;
}

export interface ISNPDatasheetState {
    isDatasheetReady: boolean;
    showEditFormDialog: boolean;
    showDeleteDialog: boolean;
    item: ISNP;
    feeViewItems?: IFEE[];
    maRequirements?: IMARequirement[];
    snpReadinessItems: ISNPReadiness[];
    errorMessage: string | null;
    messageBanner: IMessageBanner;
    isFormProcessing: boolean;
    isConfirmButtonDisabled: boolean;
}
