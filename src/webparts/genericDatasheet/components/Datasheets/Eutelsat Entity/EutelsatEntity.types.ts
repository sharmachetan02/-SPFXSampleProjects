import { IEutelsatEntity } from "../../../../../common/models/IBusiness";
import { PNPService } from "../../../../../common/services/PNPService";
import { UserService } from "../../../../../common/services/UserContextService";
import { IMessageBanner } from "../../../../genericForm/components/Form.types";
export interface IEutelsatEntityDatasheetProps {
    pnpService: PNPService;
    userService: UserService;
    itemId: number;
}

export interface IEutelsatEntityDatasheetState {
    isDatasheetReady: boolean;
    showEditFormDialog: boolean;
    showDeleteDialog: boolean;
    item: IEutelsatEntity;
    errorMessage: string | null;
    messageBanner: IMessageBanner;
    isFormProcessing: boolean;
    isConfirmButtonDisabled: boolean;

}
