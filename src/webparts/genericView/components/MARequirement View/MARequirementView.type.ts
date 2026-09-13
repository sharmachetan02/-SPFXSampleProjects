import { PNPService } from '../../../../common/services/PNPService';
import { IMARequirement } from '../../../../common/models/IBusiness';
import { IFilter } from '../../../../common/models/IFilter';
import { IColumn, IDropdownOption } from '@fluentui/react';
import { ISort } from '@pnp/sp/search';
import { UserService } from '../../../../common/services/UserContextService';

export interface IPaginationProperties {
  isEnabled: boolean;
  showTopPagination?: boolean;
  showBottomPagination?: boolean;
}

export interface IGroupingProperties {
  isEnabled: boolean;
  groupBy: string;
}

export interface IMARequirementViewProps {
  pnpService: PNPService;
  userService: UserService;
  columns?: string[];          // allowed column fieldNames
  filters?: string[];          // allowed filter names
  sort?: ISort;                // initial sort
  viewItems?: IMARequirement[];      // items from outside
  pagination?: IPaginationProperties;
  grouping?: IGroupingProperties;
}

export interface ISanctionOption extends IDropdownOption {
  color: string;
  summary: string;
}

export interface IGroupingOption extends IDropdownOption {
  type: string;
}

export interface IItemOption extends IDropdownOption {
  id: string;
  type: string;
}
export interface IMARequirementViewState {
  isViewReady: boolean;
  columns: IColumn[];
  viewItems: IMARequirement[];
  filters: IFilter[];
  sort: ISort;
  searchKeyword: string;
  pageIndex: number;
  pageSize: number;
  isFilterPanelOpen?: boolean;
}
