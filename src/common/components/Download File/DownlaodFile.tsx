import * as React from 'react';
import { saveAs } from 'file-saver';
import { PNPService } from '../../services/PNPService';
import { CommandButton } from '@fluentui/react';
import { IDocument } from '../../models/IBusiness';

export interface IDownloadFileProps {
    document: IDocument;
    pnpService: PNPService;

}

export default class DownloadFile extends React.Component<IDownloadFileProps> {
    constructor(props: Readonly<IDownloadFileProps>) {
        super(props);
    }

    public render() {
        return (
            <CommandButton
                text="Download"
                style={{
                    width: '100%',
                    color: '#052d53',
                    fontSize: 14,
                    fontWeight: 'bold',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'flex-start',
                    //margin: '0px 5px',
                    backgroundColor: 'transparent',
                }}
                iconProps={{ iconName: 'Download', style: { fontSize: 14, fontWeight: 'bold', color: '#052d53' } }}
                onClick={this.downloadDocumentVersion}
                onMouseEnter={(e) => {
                    const target = e.currentTarget as HTMLElement;
                    target.style.backgroundColor = '#f0f0f0';
                    target.style.width = '100%';

                }}
                onMouseLeave={(e) => {
                    const target = e.currentTarget as HTMLElement;
                    target.style.backgroundColor = 'transparent';

                }}
            />
        );
    }
    private downloadDocumentVersion = async () => {
        const { document } = this.props;
        const serverRelativeUrl: string = document.ServerRelativeUrl;
        const blob = await this.props.pnpService.getFileByUrl(serverRelativeUrl).getBlob();
        saveAs(blob, `${document.Name}`);
    };
}
