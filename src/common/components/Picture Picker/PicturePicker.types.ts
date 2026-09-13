import { PNPService } from '../../services/PNPService';

export interface IPicturePickerProps {
    onChange: (url: string) => void;
    pnpService : PNPService;
}

export interface IPicturePickerState {
    maxSizeExceeded: boolean;
    url:string;
}
