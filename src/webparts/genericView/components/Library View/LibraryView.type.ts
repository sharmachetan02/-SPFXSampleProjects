import { PNPService } from '../../../../common/services/PNPService';
import { ILibrary } from '../../../../common/models/IBusiness';
import { IFilter } from '../../../../common/models/IFilter';
import { IBreadcrumbItem, IColumn, IDropdownOption } from '@fluentui/react';
import { ISort } from '@pnp/sp/search';
import { UserService } from '../../../../common/services/UserContextService';
import { IMessageBanner } from '../../../genericForm/components/Form.types';

export interface IPaginationProperties {
  isEnabled: boolean;
  showTopPagination?: boolean;
  showBottomPagination?: boolean;
}
export interface ILibraryViewProps {
  pnpService: PNPService;
  userService: UserService;
  columns?: string[];          // allowed column fieldNames
  filters?: string[];          // allowed filter names
  sort?: ISort;                // initial sort
  // viewItems?: ILibrary[];      // items from outside
  folderPath: string;
  pagination?: IPaginationProperties; // pagination params
}

export interface ISanctionOption extends IDropdownOption {
  color: string;
  summary: string;
}

export interface IGroupingOption extends IDropdownOption {
  type: string;
}

export interface ILibraryViewState {
  isViewReady: boolean;
  columns: IColumn[];
  viewItems: ILibrary[];
  filters: IFilter[];
  sort: ISort;
  searchKeyword: string;
  pageIndex: number;
  pageSize: number;
  currentFolderPath: string;
  breadcrumbs: IBreadcrumbItem[];
  isDeleteDialogOpen: boolean;
  isDeleteDisabled?: boolean;
  documentToDelete: ILibrary;
  isPreviewDocSource: boolean;
  previewDocSource: string;
  previewDocTitle: string;
  isFolderVisible: boolean;
  isFileVisible: boolean;
  editFromId: number;
  //fileTypeCatListItems?: IFileType[];
  isFilterPanelOpen?: boolean;
  messageBanner: IMessageBanner;
  isFormProcessing: boolean;
  isConfirmButtonDisabled: boolean;
}

