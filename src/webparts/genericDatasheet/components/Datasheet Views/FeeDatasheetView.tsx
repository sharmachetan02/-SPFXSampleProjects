import * as React from 'react';
import { ActionButton, Modal } from '@fluentui/react';
import styles from '../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import { Form, FormType } from '../../../genericForm/components/Form.types';
import Fee from '../../../genericForm/components/Fee/Fee';
import FeeView from '../../../genericView/components/Fees View/FeesView';
import { IFeeDatasheetViewProps, IFeeDatasheetViewState } from './DatasheetViews.types';
import { UserRole } from '../../../../common/models/Enums';

export default class FeeDatasheetView extends React.Component<IFeeDatasheetViewProps, IFeeDatasheetViewState> {
    constructor(props: IFeeDatasheetViewProps) {
        super(props);
        this.state = {
            showFormDialog: false,
            itemId: undefined
        };
    }
    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };
    public render(): React.ReactElement<IFeeDatasheetViewProps> {
        const { items, userService } = this.props;
        const {
            showFormDialog
        } = this.state;
        return (
            <>
                <div className={styles.readinessContainer}>
                    {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.readinessNewButton}>
                        <ActionButton iconProps={{ iconName: 'Add' }} onClick={this.openFeeNewForm}>
                            {strings.NewFeeButtonLabel}
                        </ActionButton>
                    </div>}
                    <FeeView
                        userService={userService}
                        pnpService={this.props.pnpService}
                        columns={['Name', 'Summary', 'CostTotal', 'Status', 'DueDate']}
                        filters={['Search', 'Status']}
                        viewItems={items}
                    />

                </div>
                {showFormDialog &&
                    this.renderFormDialog()
                }
            </>

        );
    }

    private openFeeNewForm = () => {
        this.setState({
            showFormDialog: true,
            itemId: undefined,
        });
    };

    private closeFormDialog = () => {
        this.setState({ showFormDialog: false });
    };

    private refresh = async (refreshData) => {
        this.closeFormDialog();
        if (refreshData) {
            await this.props.refresh();
        }
    };

    private renderFormDialog() {
        const {
            pnpService,
            relatedItemId,
            relatedItemType,
            userService
        } = this.props;
        const {
            showFormDialog,
        } = this.state;
        const formConfig = Form.config.find((f) => f.type.toLowerCase() === FormType.Fee);
        const formControl = (<Fee pnpService={pnpService} userService={userService}
            preselectedItemType={relatedItemType}
            preselectedItemId={relatedItemId}
            callback={() => this.refresh(true)}></Fee>);
        const formTitle = `New ${formConfig.name}`;
        return (<Modal
            isOpen={showFormDialog}
            onDismiss={this.closeFormDialog}
            containerClassName={styles.modalContainer}
            isBlocking={true} >
            <div>
                <div className={styles.modalHeader}>
                    {formTitle}
                    <ActionButton
                        iconProps={{ iconName: 'Cancel' }}
                        onClick={this.closeFormDialog}
                        className={styles.iconButton}
                    >Close</ActionButton>
                </div>
                <div className={styles.modalBody}>
                    {formControl}
                </div>
            </div>
        </Modal>);
    }
}
