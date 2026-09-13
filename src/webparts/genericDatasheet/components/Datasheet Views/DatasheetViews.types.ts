import { IMarketReadiness, ISNPReadiness, IContact, IPayment, IFEE } from "../../../../common/models/IBusiness";
import { PNPService } from "../../../../common/services/PNPService";
import { IMessageBanner } from "../../../genericForm/components/Form.types";
import { UserService } from "../../../../common/services/UserContextService";

export interface ISNPReadinessDatasheetViewProps {
    items: ISNPReadiness[];
    pnpService: PNPService;
    userService: UserService;
    refresh?: () => void;
    renderField?: (item: ISNPReadiness, key: string) => JSX.Element;
    snpItemId?: number;
}

export interface IFeeDatasheetViewProps {
    items: IFEE[];
    pnpService: PNPService;
    userService: UserService;
    refresh?: () => void;
    renderField?: (item: IFEE, key: string) => JSX.Element;
    relatedItemId?: number;
    relatedItemType?: string;
}

export interface IFeeDatasheetViewState {
    showFormDialog: boolean;
    itemId?: number;
}

export interface ISNPReadinessDatasheetViewState {
    showFormDialog: boolean;
    itemId: number;
    showDeleteDialog?: boolean;
    isDeleteDisabled?: boolean;
    documentToDelete?: ISNPReadiness;
    messageBanner: IMessageBanner;
    isFormProcessing: boolean;
    isConfirmButtonDisabled: boolean;
}


export interface IMarketReadinessDatasheetViewProps {
    items: IMarketReadiness[];
    pnpService: PNPService;
    userService: UserService;
    refresh?: () => void;
    renderField?: (item: IMarketReadiness, key: string) => JSX.Element;
    marketItemId?: number;
}

export interface IMarketReadinessDatasheetViewState {
    showFormDialog: boolean;
    itemId: number;
    showDeleteDialog?: boolean;
    isDeleteDisabled?: boolean;
    documentToDelete?: IMarketReadiness;
    messageBanner: IMessageBanner;
    isFormProcessing: boolean;
    isConfirmButtonDisabled: boolean;
}



export interface IPaymentsDatasheetViewProps {
    items: IPayment[];
    pnpService: PNPService;
    userService: UserService;
    refresh?: () => void;
    renderField?: (item: IPayment, key: string) => JSX.Element;
    paymentItemId?: number;
}

export interface IPaymentsDatasheetViewState {
    showFormDialog: boolean;
    itemId: number;
}

export interface IContactDatasheetViewProps {
    contacts: IContact[];
    pnpService: PNPService;
    userService: UserService;
    refresh?: () => void;
    relatedItemId?: number;
    relatedItemType?: string;
}

export interface IContactDatasheetViewState {
    showFormDialog: boolean;
    itemId?: number;
}
