import { IContact, IDP, IMARequirement } from "../../../../../common/models/IBusiness";
import { PNPService } from "../../../../../common/services/PNPService";
import { UserService } from "../../../../../common/services/UserContextService";
import { IMessageBanner } from "../../../../genericForm/components/Form.types";

export interface IPaginationProperties {
    isEnabled: boolean;
    showTopPagination?: boolean;
    showBottomPagination?: boolean;

}
export interface IDPDatasheetProps {
  pnpService: PNPService;
  userService: UserService;
  itemId: number;
  pagination?: IPaginationProperties;
}

export interface IDPDatasheetState {
  isDatasheetReady: boolean;
  showEditFormDialog: boolean;
  showDeleteDialog: boolean;
  item: IDP;
  maReqViewItems?: IMARequirement[];
  contactViewItems: IContact[];
  errorMessage: string | null;
    messageBanner: IMessageBanner;
    isFormProcessing: boolean;
    isConfirmButtonDisabled: boolean;
}
