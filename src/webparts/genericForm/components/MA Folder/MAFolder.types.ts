import { PNPService } from '../../../../common/services/PNPService';
import { UserService } from '../../../../common/services/UserContextService';

export interface IMAFolderProps {
    //context: WebPartContext;
    pnpService: PNPService;
    userService: UserService
    currentFolderPath?: string; // For updating an existing folder
    itemId?: number; // Optional: for folder item reference
    callback?: (string?) => void;
}

export interface IMAFolderState {
    isFormReady: boolean;
    name: string;
    exitingFolderName?: string;
    errors: { [key: string]: string };
}
