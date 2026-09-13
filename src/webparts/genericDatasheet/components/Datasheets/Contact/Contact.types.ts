import { IContact } from "../../../../../common/models/IBusiness";
import { PNPService } from "../../../../../common/services/PNPService";
import { IMessageBanner } from "../../../../genericForm/components/Form.types";
import { UserService } from "../../../../../common/services/UserContextService";
export interface IContactDatasheetProps {
    pnpService: PNPService;
    userService: UserService;
    itemId: number;
}

export interface IContactDatasheetState {
    isDatasheetReady: boolean;
    showEditFormDialog: boolean;
    showDeleteDialog: boolean;
    item: IContact;
    errorMessage: string | null;
    messageBanner: IMessageBanner;
    isFormProcessing: boolean;
    isConfirmButtonDisabled: boolean;

}
