import { PNPService } from '../../../../../common/services/PNPService';
import { IPayment } from '../../../../../common/models/IBusiness';
import { UserService } from '../../../../../common/services/UserContextService';
import { IMessageBanner } from '../../../../genericForm/components/Form.types';

export interface IPaymentDatasheetProps {
    pnpService: PNPService;
      userService: UserService;
    itemId: number;
}

export interface IPaymentDatasheetState {
    isDatasheetReady: boolean;
    showEditFormDialog: boolean;
    showDeleteDialog: boolean;
    item: IPayment;
    errorMessage: string | null;
    messageBanner: IMessageBanner;
    isFormProcessing: boolean;
    isConfirmButtonDisabled: boolean;
}
