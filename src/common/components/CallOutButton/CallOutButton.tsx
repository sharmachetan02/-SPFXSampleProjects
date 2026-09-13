// CallOutButton.tsx
import * as React from 'react';
import { ActionButton, Callout } from '@fluentui/react';
import styles from './CallOutButton.module.scss'; // adjust path

interface CallOutButtonProps {

    icon: string;
    name: string;
    children: React.ReactNode;

}

interface CallOutButtonState {
    isVisible: boolean;
}

export class CallOutButton extends React.Component<CallOutButtonProps, CallOutButtonState> {
    private buttonId: string;

    constructor(props: CallOutButtonProps) {
        super(props);
        this.state = {
            isVisible: false,
        };
        this.buttonId = `Button-${Date.now()}-${Math.random().toString(36).substring(2, 8)}`;
    }

    private toggleCallout = () => {
        this.setState((prev) => ({ isVisible: !prev.isVisible }));
    };

    private closeCallout = () => {
        this.setState({ isVisible: false });
    };

    render(): JSX.Element {
        const { icon, name, children } = this.props;
        const { isVisible } = this.state;

        return (
            <>
                <ActionButton
                    // className={styles.callOutButton}
                    className={styles.callOutDescription}
                    id={this.buttonId}
                    iconProps={{ iconName: icon }}
                    onClick={this.toggleCallout}
                >
                    {name}
                </ActionButton>

                {isVisible && (
                    <Callout
                        className={styles.callout}
                        role="dialog"
                        gapSpace={0}
                        target={`#${this.buttonId}`}
                        setInitialFocus
                        onDismiss={this.closeCallout}
                    >
                        {children}
                    </Callout>
                )}
            </>
        );
    }
}

