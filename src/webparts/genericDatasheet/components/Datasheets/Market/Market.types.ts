import { IMARequirement, IMarket, IMarketReadiness } from "../../../../../common/models/IBusiness";
import { PNPService } from "../../../../../common/services/PNPService";
import { UserService } from "../../../../../common/services/UserContextService";
import { IMessageBanner } from "../../../../genericForm/components/Form.types";
export interface IPaginationProperties {
  isEnabled: boolean;
  showTopPagination?: boolean;
  showBottomPagination?: boolean;
}
export interface IMarketDatasheetProps {
  pnpService: PNPService;
    userService: UserService;
  itemId: number;
  pagination?: IPaginationProperties;
}
export interface IMarketDatasheetState {
  isDatasheetReady: boolean;
  showEditFormDialog: boolean;
  showDeleteDialog: boolean;
  item: IMarket;
  maRequirements: IMARequirement[];
  marketReadinessItems: IMarketReadiness[];
  errorMessage: string | null;
  messageBanner: IMessageBanner;
  isFormProcessing: boolean;
  isConfirmButtonDisabled: boolean;
}
