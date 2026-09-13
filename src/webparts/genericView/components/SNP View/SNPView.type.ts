import { PNPService } from '../../../../common/services/PNPService';
import { ISNP } from '../../../../common/models/IBusiness';
import { IFilter } from '../../../../common/models/IFilter';
import { IColumn } from '@fluentui/react';
import { ISort } from '@pnp/sp/search';
import { UserService } from '../../../../common/services/UserContextService';

export interface IPaginationProperties {
  isEnabled: boolean;
  showTopPagination?: boolean;
  showBottomPagination?: boolean;
}
export interface ISNPViewProps {
  pnpService: PNPService;
  userService: UserService;
  columns?: string[];
  filters?: string[];
  sort?: ISort;
  viewItems?: ISNP[];
  pagination?: IPaginationProperties;
}

export interface ISNPViewState {
  isViewReady: boolean;
  columns: IColumn[];
  viewItems: ISNP[];
  filters: IFilter[];
  sort: ISort;
  searchKeyword: string;
  pageIndex: number;
  pageSize: number;
  isFilterPanelOpen?: boolean;
}

