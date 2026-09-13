import { PNPService } from '../../../../common/services/PNPService';
import { IFEE } from '../../../../common/models/IBusiness';
import { IFilter } from '../../../../common/models/IFilter';
import { IColumn, IDropdownOption } from '@fluentui/react';
import { ISort } from '@pnp/sp/search';
import { UserService } from '../../../../common/services/UserContextService';

export interface IPaginationProperties {
  isEnabled: boolean;
  showTopPagination?: boolean;
  showBottomPagination?: boolean;
}
export interface IFEEViewProps {
  pnpService: PNPService;
  userService: UserService;
  columns?: string[];          // allowed column fieldNames
  filters?: string[];          // allowed filter names
  sort?: ISort;                // initial sort
  viewItems?: IFEE[];
  pagination?: IPaginationProperties;    // items from outside
}

export interface IItemOption extends IDropdownOption {
  id: string;
  type: string;
}

export interface IFEEViewState {
  isViewReady: boolean;
  columns: IColumn[];
  viewItems: IFEE[];
  filters: IFilter[];
  sort: ISort;
  searchKeyword: string;
  pageIndex: number;
  pageSize: number;
  isFilterPanelOpen?: boolean;
}

