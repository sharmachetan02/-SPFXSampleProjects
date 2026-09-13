import { PNPService } from '../../../common/services/PNPService';
import { UserService } from '../../../common/services/UserContextService';
import { FormType } from '../../genericForm/components/Form.types';

export interface IHelloWorldProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  pnpService: PNPService;
  userService: UserService;
}
export interface IHelloWorldState {
  showFormDialog: boolean;
  itemType: FormType;
  itemId: number;
  preselectedItemId: number;
  preselectedItemType: FormType;
}
