import { IAuthority, ICountry, IEutelsatEntity, ILegalSvcProvider, IMarket } from "../../../../../common/models/IBusiness";
import { PNPService } from "../../../../../common/services/PNPService";
import { IMessageBanner } from "../../../../genericForm/components/Form.types";
import { UserService } from "../../../../../common/services/UserContextService";
export interface IPaginationProperties {
  isEnabled: boolean;
  showTopPagination?: boolean;
  showBottomPagination?: boolean;
}
export interface ICountryDatasheetProps {
  pnpService: PNPService;
  userService: UserService;
  itemId: number;
  pagination?: IPaginationProperties;
}
export interface ICountryDatasheetState {
  isDatasheetReady: boolean;
  showEditFormDialog: boolean;
  showDeleteDialog: boolean;
  item: ICountry;
  legalSvcProviders: ILegalSvcProvider[];
  authorities: IAuthority[];
  snps: IAuthority[];
  eutsatEntities: IEutelsatEntity[];
  marketItems: IMarket[];
  errorMessage: string | null;
  messageBanner: IMessageBanner;
  isFormProcessing: boolean;
  isConfirmButtonDisabled: boolean;

}
