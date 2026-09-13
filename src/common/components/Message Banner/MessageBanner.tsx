import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';

export type MessageType = 'success' | 'error' | 'warning' | 'info';

export interface IMessageBannerProps {
    message: string;
    type?: MessageType;
    visible: boolean;
    onDismiss?: () => void;
}

export const MessageBanner: React.FC<IMessageBannerProps> = ({
    message,
    type = 'info',
    visible,
    onDismiss
}) => {
    if (!visible) return null;

    const getMessageBarType = (msgType: MessageType): MessageBarType => {
        switch (msgType) {
            case 'success': return MessageBarType.success;
            case 'error': return MessageBarType.error;
            case 'warning': return MessageBarType.warning;
            case 'info':
            default: return MessageBarType.info;
        }
    };

    return (
        <MessageBar
            messageBarType={getMessageBarType(type)}
            isMultiline={false}
            onDismiss={onDismiss}
        >
            {message}
        </MessageBar>
    );
};
