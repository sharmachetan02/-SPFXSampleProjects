import { IAuthority, IContact, IFEE, IMARequirement, IPayment } from "../../../../../common/models/IBusiness";
import { PNPService } from "../../../../../common/services/PNPService";
import { IMessageBanner } from "../../../../genericForm/components/Form.types";
import { UserService } from "../../../../../common/services/UserContextService";

export interface IPaginationProperties {
  isEnabled: boolean;
  showTopPagination?: boolean;
  showBottomPagination?: boolean;
}
export interface IAuthorityDatasheetProps {
  pnpService: PNPService;
  userService: UserService;
  itemId: number;
  pagination?: IPaginationProperties;
}
export interface IAuthorityDatasheetState {
  isDatasheetReady: boolean;
  showEditFormDialog: boolean;
  showDeleteDialog: boolean;
  item: IAuthority;
  maReqViewItems?: IMARequirement[];
  feeViewItems: IFEE[];
  paymentViewItems: IPayment[];
  contactViewItems: IContact[];
  errorMessage: string | null;
  messageBanner: IMessageBanner;
  isFormProcessing: boolean;
  isConfirmButtonDisabled: boolean;
}


