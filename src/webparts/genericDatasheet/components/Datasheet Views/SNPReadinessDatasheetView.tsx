import * as React from 'react';
import { ISNPReadiness } from '../../../../common/models/IBusiness';
import { MarketReadinessResponsibleParty, UserRole } from '../../../../common/models/Enums';
import { Accordion, AccordionItem, AccordionHeader, AccordionPanel } from '../../../../common/components/Accordion/Accordion';
import * as strings from 'GenericDatasheetStrings';
import { Label, ActionButton, Modal, Icon, Dialog, DialogType, PrimaryButton, DefaultButton, DialogFooter, Stack, SpinnerSize, Spinner } from '@fluentui/react';
import styles from '../Datasheet.module.scss';
import { Form, FormType, IMessageBanner } from '../../../genericForm/components/Form.types';
import SNPMarketReadiness from '../../../genericForm/components/SNP Readiness/SNPReadiness';
import { ISNPReadinessDatasheetViewProps, ISNPReadinessDatasheetViewState } from './DatasheetViews.types';
import { Consts } from '../../../../common/consts/Consts';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { MessageBanner } from '../../../../common/components/Message Banner/MessageBanner';

export default class SNPReadinessDatasheetView extends React.Component<ISNPReadinessDatasheetViewProps, ISNPReadinessDatasheetViewState> {
    constructor(props: ISNPReadinessDatasheetViewProps) {
        super(props);
        this.state = {
            showFormDialog: false,
            itemId: undefined,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'error',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        };
    }
    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };
    private renderReadinessItem = (readinessItem: ISNPReadiness): JSX.Element => {
        let responsibleTitle = "";
        let responsibleIcon = "";

        switch (readinessItem.ResponsiblePartyType) {
            case MarketReadinessResponsibleParty.Overall:
                responsibleTitle = readinessItem.ResponsiblePartyType;
                responsibleIcon = "Country";
                break;
            case MarketReadinessResponsibleParty.Eutelsat:
                responsibleTitle = readinessItem.ResponsiblePartyType;
                responsibleIcon = "Satellite";
                break;
            case MarketReadinessResponsibleParty.DP:
                responsibleTitle = `${readinessItem.ResponsiblePartyType} - ${readinessItem.ResponsiblePartyItem?.Title}`;
                responsibleIcon = "Agreement";
                break;
            case MarketReadinessResponsibleParty.TP:
                responsibleTitle = `${readinessItem.ResponsiblePartyType} - ${readinessItem.ResponsiblePartyItem?.Title}`;
                responsibleIcon = "Dish";
                break;
        }

        return (
            <AccordionItem key={responsibleTitle} value={responsibleTitle}>
                <AccordionHeader icon={responsibleIcon}>
                    {responsibleTitle}
                </AccordionHeader>
                <AccordionPanel>
                    {this.renderReadinessDetails(readinessItem)}
                </AccordionPanel>
            </AccordionItem>
        );
    };

    private renderReadinessDetails = (readinessItem: ISNPReadiness): JSX.Element => {
        const left = ["RAGStatus", "EstimatedDate", "EffectiveDate", "EutelsatOwners"];
        const right = ["Comments"];
        if (readinessItem.ResponsiblePartyType === MarketReadinessResponsibleParty.Overall) {
            right.push("ResponsiblePartyItem");
        }
        return (
            <>
                {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.readinessCommandsBar}>
                    <ActionButton iconProps={{ iconName: 'Edit' }} onClick={() => this.openSNPMarketReadinessEditForm(readinessItem.Id)}>
                        {strings.SnpReadinessLabelEditButton}
                    </ActionButton>
                    <ActionButton iconProps={{ iconName: 'Delete' }} onClick={() => this.openDeleteDialog(readinessItem)}>
                        {strings.SnpReadinessLabelDeleteButton}
                    </ActionButton>
                </div>}
                <div className={styles.datasheetInfoContainer}>
                    <div className={styles.datasheetInfoContainerLeft}>
                        {left.map((field) => (
                            <div key={field} className="ms-Grid">
                                <div className={`ms-Grid-row ${styles.datasheetInfoRow}`}>
                                    <div className="ms-Grid-col ms-sm3 ms-md3">
                                        <div className={styles.datasheetInfoLabel}>
                                            <Label>{strings[`SnpReadinessLabel${field}`]}:</Label>
                                        </div>
                                    </div>
                                    <div className="ms-Grid-col ms-sm9 ms-md9">
                                        {this.props.renderField(readinessItem, field)}
                                    </div>
                                </div>
                            </div>
                        ))}
                    </div>
                    <div className={styles.datasheetInfoContainerRight}>
                        {right.map((field) => (
                            <div key={field} className="ms-Grid">
                                <div className={`ms-Grid-row ${styles.datasheetInfoRow}`}>
                                    <div className="ms-Grid-col ms-sm3 ms-md3">
                                        <div className={styles.datasheetInfoLabel}>
                                            <Label>{strings[`SnpReadinessLabel${field}`]}:</Label>
                                        </div>
                                    </div>
                                    <div className="ms-Grid-col ms-sm9 ms-md9">
                                        {this.props.renderField(readinessItem, field)}
                                    </div>
                                </div>
                            </div>
                        ))}
                    </div>
                </div>
            </>

        );
    };

    public render(): React.ReactElement<ISNPReadinessDatasheetViewProps> {
        const { items } = this.props;
        const sortedSNPReadinessItems = items.sort((a, b) => {
            const order = [MarketReadinessResponsibleParty.Overall, MarketReadinessResponsibleParty.Eutelsat, MarketReadinessResponsibleParty.DP, MarketReadinessResponsibleParty.TP];
            return order.indexOf(a.ResponsiblePartyType as MarketReadinessResponsibleParty) - order.indexOf(b.ResponsiblePartyType as MarketReadinessResponsibleParty);
        });
        const {
            showFormDialog,
            showDeleteDialog
        } = this.state;
        return (
            <>
                <div className={styles.readinessContainer}>
                    {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.readinessNewButton}>
                        <ActionButton iconProps={{ iconName: 'Add' }} onClick={this.openSNPMarketReadinessNewForm}>
                            {strings.SnpReadinessLabelNewButton}
                        </ActionButton>
                    </div>}
                    {items && items.length > 0 ? (
                        <Accordion multiple collapsible>
                            {sortedSNPReadinessItems.map((item) => this.renderReadinessItem(item))}
                        </Accordion>
                    ) : (
                        <div className={styles.noDataMessage}>
                            <Icon iconName="Remove" className={styles.mutedIcon} />
                            <span>{strings.DataSheetNoRecordFound}</span>
                            <Icon iconName="Remove" className={styles.mutedIcon} />
                        </div>
                    )}
                </div>
                {showFormDialog &&
                    this.renderFormDialog()
                }
                {showDeleteDialog && this.renderDeleteDialog()}
            </>

        );
    }

    private openSNPMarketReadinessNewForm = () => {
        this.setState({
            showFormDialog: true,
            itemId: undefined,
        });
    };

    private openSNPMarketReadinessEditForm = (itemId) => {
        this.setState({ showFormDialog: true, itemId });
    };

    private closeFormDialog = () => {
        this.setState({ showFormDialog: false, showDeleteDialog: false, documentToDelete: null, isDeleteDisabled: false });

    };
    private openDeleteDialog = (item: ISNPReadiness) => {
        this.setState({
            showDeleteDialog: true, documentToDelete: item,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'error',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        });

    };
    private deleteDataSheet = async () => {
        try {
            this.setState({ isFormProcessing: true });
            const { pnpService } = this.props;
            this.setState({ isDeleteDisabled: true });
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.MA_READINESS_URL);
            await pnpService.getListItems(listUrl).getById(this.state.documentToDelete?.Id).recycle();
            const messageBanner: IMessageBanner = {
                message: strings.DeleteSuccessMessage,
                type: 'success',
                visible: true
            };
            this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
            setTimeout(() => { this.refresh(true); }, 2000);

        } catch (error) {
            console.error(strings.DeleteErrorMessage, error);
            const messageBanner: IMessageBanner = {
                message: strings.DeleteErrorMessage,
                type: 'error',
                visible: true
            };
            this.setState({ isFormProcessing: false, messageBanner });
        }
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
            snpItemId
        } = this.props;
        const {
            showFormDialog,
            itemId,
        } = this.state;
        let formControl, formTitle = null;
        const formConfig = Form.config.find((f) => f.type.toLowerCase() === FormType.SNPMarketReadiness);
        if (itemId) {
            formControl = (<SNPMarketReadiness pnpService={pnpService} userService={userService} itemId={itemId} callback={() => this.refresh(true)}></SNPMarketReadiness>);
            formTitle = `Edit ${formConfig.name}`;
        } else {
            formControl = (<SNPMarketReadiness pnpService={pnpService} userService={userService}
                preselectedItemId={snpItemId} callback={() => this.refresh(true)}></SNPMarketReadiness>);
            formTitle = `New ${formConfig.name}`;
        }

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

    private renderDeleteDialog() {
        const { documentToDelete } = this.state;
        return (
            <Dialog
                hidden={!documentToDelete}
                onDismiss={this.closeFormDialog}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: documentToDelete?.Title || 'Delete'
                }}
                modalProps={{ isBlocking: true }}
                styles={{
                    main: {
                        selectors: {
                            '@media (min-width: 480px)': {
                                maxWidth: '410px !important', // Override media query
                            },
                        },

                    }
                }}
            >
                {this.state.messageBanner.visible && (
                    <div className={styles.topMessageBanner}>
                        <MessageBanner
                            message={this.state.messageBanner.message}
                            type={this.state.messageBanner.type}
                            visible={this.state.messageBanner.visible}
                        />
                    </div>
                )}
                {this.state.messageBanner.type !== 'success' && (
                    <span>{strings.LibraryViewRecycleBinConfirmation}</span>
                )}
                <DialogFooter>
                    <PrimaryButton
                        disabled={this.state.isConfirmButtonDisabled}
                        onClick={this.deleteDataSheet} >
                        {this.state.isFormProcessing ? (
                            <Stack horizontal verticalAlign="center">
                                <Spinner size={SpinnerSize.xSmall} labelPosition="right" />
                                <span style={{ marginLeft: 8 }}>Processing</span>
                            </Stack>
                        ) : (
                            strings.DataSheetConfirmButton
                        )}
                    </PrimaryButton>
                    <DefaultButton
                        onClick={this.closeFormDialog}
                        text="Back"
                    />
                </DialogFooter>
            </Dialog>
        );
    }

}
