import { PNPService } from '../../../../common/services/PNPService';
import { IEutelsatEntity } from '../../../../common/models/IBusiness';
import { IFilter } from '../../../../common/models/IFilter';
import { IColumn } from '@fluentui/react';
import { ISort } from '@pnp/sp/search';
import { UserService } from '../../../../common/services/UserContextService';

export interface IPaginationProperties {
  isEnabled: boolean;
  showTopPagination?: boolean;
  showBottomPagination?: boolean;
}
export interface IEutelsatEntityViewProps {
  pnpService: PNPService;
  userService: UserService;
  columns?: string[];          // allowed column fieldNames
  filters?: string[];          // allowed filter names
  sort?: ISort;                // initial sort
  viewItems?: IEutelsatEntity[];      // items from outside
  pagination?: IPaginationProperties; // pagination params
}
export interface IEutelsatEntityViewState {
  isViewReady: boolean;
  columns: IColumn[];
  viewItems: IEutelsatEntity[]
  filters: IFilter[];
  sort: ISort;
  searchKeyword: string;
  pageIndex: number;
  pageSize: number;
  isFilterPanelOpen?: boolean;

}
