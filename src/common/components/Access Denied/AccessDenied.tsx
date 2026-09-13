
// AccessDeniedMessage.tsx
import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react';

interface AccessDeniedMessageProps {
    message: string;
}

class AccessDeniedMessage extends React.Component<AccessDeniedMessageProps> {
    render() {
        const { message } = this.props;
        return (
            <MessageBar messageBarType={MessageBarType.error}>
                {message}
            </MessageBar>
        );
    }
}

export default AccessDeniedMessage;
