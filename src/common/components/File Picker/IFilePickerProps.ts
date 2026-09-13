import { IAttachmentFileInfo } from "@pnp/sp/attachments";
export interface IFilePickerProps {
    onChange: (files: IAttachmentFileInfo[]) => void;
    files:IAttachmentFileInfo[];
    multiple:boolean;
    disabled:boolean;
}
