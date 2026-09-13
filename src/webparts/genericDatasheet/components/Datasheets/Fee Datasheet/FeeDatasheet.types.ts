import { IFEE, IPayment } from "../../../../../common/models/IBusiness";
import { PNPService } from "../../../../../common/services/PNPService";
import { UserService } from "../../../../../common/services/UserContextService";
import { IMessageBanner } from "../../../../genericForm/components/Form.types";

export interface IPaginationProperties {
    isEnabled: boolean;
    showTopPagination?: boolean;
    showBottomPagination?: boolean;
}
export interface IFeeDatasheetProps {
    pnpService: PNPService;
    userService: UserService;
    itemId: number;
    pagination?: IPaginationProperties;
}

export interface IFeeDatasheetState {
    isDatasheetReady: boolean;
    showEditFormDialog: boolean;
    showDeleteDialog: boolean;
    item: IFEE;
    paymentViewItems: IPayment[];
    errorMessage: string | null;
    messageBanner: IMessageBanner;
    isFormProcessing: boolean;
    isConfirmButtonDisabled: boolean;
}
