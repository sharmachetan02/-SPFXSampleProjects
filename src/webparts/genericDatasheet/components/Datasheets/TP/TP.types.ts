import { IContact, IMARequirement, ISNP, ITP } from "../../../../../common/models/IBusiness";
import { PNPService } from "../../../../../common/services/PNPService";
import { UserService } from "../../../../../common/services/UserContextService";
import { IMessageBanner } from "../../../../genericForm/components/Form.types";
export interface IPaginationProperties {
  isEnabled: boolean;
  showTopPagination?: boolean;
  showBottomPagination?: boolean;
}
export interface ITPDatasheetProps {
  pnpService: PNPService;
  userService: UserService;
  itemId: number;
  pagination?: IPaginationProperties;
}

export interface ITPDatasheetState {
  isDatasheetReady: boolean;
  showEditFormDialog: boolean;
  showDeleteDialog: boolean;
  item: ITP;
  snpViewItems: ISNP[];
  maReqViewItems?: IMARequirement[];
  errorMessage: string | null;
  contactViewItems?: IContact[];
  messageBanner: IMessageBanner;
  isFormProcessing: boolean;
  isConfirmButtonDisabled: boolean;
}
