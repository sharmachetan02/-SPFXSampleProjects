import { PNPService } from '../../../../common/services/PNPService';
import { UserService } from '../../../../common/services/UserContextService';

export interface IGenericFormProps {
  pnpService: PNPService;
  userService: UserService
}
