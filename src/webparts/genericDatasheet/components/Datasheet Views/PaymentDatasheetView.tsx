
import * as React from 'react';
//import { ISNPReadiness } from '../../../../common/models/IBusiness';
//import { MarketReadinessResponsibleParty } from '../../../../common/models/Enums';
//import { Accordion, AccordionItem, AccordionHeader, AccordionPanel } from '../../../../common/components/Accordion/Accordion';
import * as strings from 'GenericDatasheetStrings';
import { ActionButton, Modal } from '@fluentui/react';
import styles from '../Datasheet.module.scss';
import { Form, FormOrigin, FormType } from '../../../genericForm/components/Form.types';
import Payment from '../../../genericForm/components/Payment/Payment';
import PaymentView from '../../../genericView/components/Payment View/PaymentView';
import { IPaymentsDatasheetViewProps, IPaymentsDatasheetViewState } from './DatasheetViews.types';
import { UserRole } from '../../../../common/models/Enums';

export default class PaymentDatasheetView extends React.Component<IPaymentsDatasheetViewProps, IPaymentsDatasheetViewState> {
    constructor(props: IPaymentsDatasheetViewProps) {
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
    public render(): React.ReactElement<IPaymentsDatasheetViewProps> {
        const { items, userService } = this.props;
        const {
            showFormDialog
        } = this.state;
        return (
            <>
                <div className={styles.readinessContainer}>
                    {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.readinessNewButton}>
                        <ActionButton iconProps={{ iconName: 'Add' }} onClick={this.openSNPMarketReadinessNewForm}>
                            {strings.PaymentLabelNewButton}
                        </ActionButton>
                    </div>}
                    <PaymentView
                        pnpService={this.props.pnpService}
                        userService={userService}
                        columns={['Title', 'Summary', 'Fees', 'CostTotal', 'Status', 'Item', 'DatePaid']}
                        filters={['Search', 'Status', 'Item']}
                        viewItems={items}
                    //pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}
                    />

                </div>
                {showFormDialog &&
                    this.renderFormDialog()
                }
            </>

        );
    }

    private openSNPMarketReadinessNewForm = () => {
        this.setState({
            showFormDialog: true,
            itemId: undefined,
        });
    };
    /*
    private openSNPMarketReadinessEditForm = (itemId) => {
        this.setState({ showFormDialog: true, itemId });
    };*/

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
            userService,
            paymentItemId
        } = this.props;
        const {
            showFormDialog,
        } = this.state;
        const formConfig = Form.config.find((f) => f.type.toLowerCase() === FormType.Payment);

        const formControl = (<Payment pnpService={pnpService} userService={userService}
            preselectedItemId={paymentItemId} callback={() => this.refresh(true)} formOrigin={FormOrigin.Datasheet}></Payment>);
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


