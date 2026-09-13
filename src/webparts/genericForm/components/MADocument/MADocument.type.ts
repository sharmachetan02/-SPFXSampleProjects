
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PNPService } from '../../../../common/services/PNPService';
//import { IFilePickerResult } from '@pnp/spfx-controls-react';
//import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { IAttachmentFileInfo } from "@pnp/sp/attachments";
import { UserService } from '../../../../common/services/UserContextService';
export interface IMADocumentProps {
    pnpService: PNPService;
    userService: UserService;
    filepath?: string;
    itemId?: number;
    callback?: (boolean) => void;
}

export interface IMADocumentState {
    isFormReady: boolean;
    name: string;
    title: string;
    summary: string;
    comments: string;
    document: IAttachmentFileInfo[];
    classification: string;
    classificationChoices: IDropdownOption[];
    category: string;
    categoryChoices: IDropdownOption[];
    fileType: string;
    fileTypeChoices: IDropdownOption[];
    version: string;
    keywords: IDropdownOption[];
    keywordsChoices: IDropdownOption[];
    extension?: string;
    filepath?: string;
    errors: { [key: string]: string };

}
