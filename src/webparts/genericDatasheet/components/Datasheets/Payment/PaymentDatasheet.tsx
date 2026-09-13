import * as React from 'react';
import { IPaymentDatasheetProps, IPaymentDatasheetState } from './PaymentDatasheet.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import {
    ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, ITooltipHostStyles, Label,
    Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost
} from '@fluentui/react';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { FormOrigin, FormType, IMessageBanner } from '../../../../genericForm/components/Form.types';
import { IPayment, IContact, IFEE } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { RichText } from '../../../../../common/components/RichText/RichText';
import { StatusPill } from '../../../../genericView/components/View.utility';
import { renderObjectPills } from '../../Datasheet.utility';
import Payment from '../../../../genericForm/components/Payment/Payment';
import LibraryView from '../../../../genericView/components/Library View/LibraryView';
import { UserRole } from '../../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../../common/components/Access Denied/AccessDenied';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';
import DomHelper from '../../../../../common/helpers/DomHelper';

export default class PaymentDatasheet extends React.Component<
    IPaymentDatasheetProps,
    IPaymentDatasheetState
> {


    constructor(props: Readonly<IPaymentDatasheetProps>) {
        super(props);
        this.state = {
            isDatasheetReady: false,
            showEditFormDialog: false,
            showDeleteDialog: false,
            item: null,
            errorMessage: null,
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
    public async componentDidMount() {
        if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner, UserRole.Admin)) {
            await this.init();
            this.setState({ isDatasheetReady: true });
        }
    }

    public initItem = async () => {
        const { pnpService, itemId } = this.props;
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.PAYMENT.SUMMARY,
            Consts.FIELDS.PAYMENT.STATUS,
            Consts.FIELDS.PAYMENT.COMMENTS,
            Consts.FIELDS.PAYMENT.VATRATE,
            Consts.FIELDS.PAYMENT.DUEDATE,
            Consts.FIELDS.PAYMENT.PAYMENTPAID,
            Consts.FIELDS.PAYMENT.VATFREECOST,
            Consts.FIELDS.PAYMENT.VAT,
            Consts.FIELDS.PAYMENT.CURRENCY,
            Consts.FIELDS.PAYMENT.BENEFICIARY,
            Consts.FIELDS.FEE.CHARGEDBY,
            Consts.FIELDS.COMMON.CREATION_TIME,
            Consts.FIELDS.COMMON.MODIFICATION_TIME,
            Consts.FIELDS.COMMON.DOCUMENTS_SPACE,
            `${Consts.FIELDS.PAYMENT.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.PAYMENT.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.PAYMENT.CURRENCY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.CURRENCY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.PAYMENT.CURRENCY}/${Consts.FIELDS.PAYMENT.CURRENCY_ISOCODE}`,
            `${Consts.FIELDS.PAYMENT.FEES}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.FEES}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.PAYMENT.CONTACTS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.CONTACTS}/${Consts.FIELDS.COMMON.TITLE}`
        ];

        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.PAYMENT_URL);
        try {
            const item = await pnpService
                .getListByUrl(listUrl)
                .items.getById(itemId)
                .select(...selectedFields)
                .expand(
                    Consts.FIELDS.PAYMENT.AUTHORITY,
                    Consts.FIELDS.PAYMENT.CURRENCY,
                    Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER,
                    Consts.FIELDS.PAYMENT.FEES,
                    Consts.FIELDS.PAYMENT.CONTACTS,
                    Consts.FIELDS.PAYMENT.TELEPORTPARTNER
                )();

            let beneficiaryItem = null;
            switch (item[Consts.FIELDS.FEE.CHARGEDBY]) {
                case 'Law Firm':
                case 'Legal Representative':
                    beneficiaryItem = {
                        id: item[Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID] + "-" + FormType.LegalServiceProvider,
                        text: item[Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.TITLE],
                        type: FormType.LegalServiceProvider
                    };
                    break;
                case 'Administration':
                case 'Organization':
                case 'Regulator':
                    beneficiaryItem = {
                        id: item[Consts.FIELDS.PAYMENT.AUTHORITY][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.PAYMENT.AUTHORITY][Consts.FIELDS.COMMON.ID] + "-" + FormType.Authority,
                        text: item[Consts.FIELDS.PAYMENT.AUTHORITY][Consts.FIELDS.COMMON.TITLE],
                        type: FormType.Authority,
                    };
                    break;
                case 'Teleport Partner':
                    beneficiaryItem = {
                        id: item[Consts.FIELDS.PAYMENT.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.PAYMENT.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + FormType.TP,
                        text: item[Consts.FIELDS.PAYMENT.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
                        type: FormType.TP,
                    };
                    break;
                default:
                    break;
            }

            const payment: IPayment = {
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                Summary: item[Consts.FIELDS.PAYMENT.SUMMARY],
                Status: item[Consts.FIELDS.PAYMENT.STATUS],
                Comments: item[Consts.FIELDS.PAYMENT.COMMENTS],
                VatFreeCost: item[Consts.FIELDS.PAYMENT.VATFREECOST],
                Vat: item[Consts.FIELDS.PAYMENT.VAT],
                VatRate: item[Consts.FIELDS.PAYMENT.VATRATE],
                Currency: item[Consts.FIELDS.PAYMENT.CURRENCY] ? item[Consts.FIELDS.PAYMENT.CURRENCY][Consts.FIELDS.PAYMENT.CURRENCY_ISOCODE] : null,
                DueDate: item[Consts.FIELDS.PAYMENT.DUEDATE] ? new Date(item[Consts.FIELDS.PAYMENT.DUEDATE]) : null,
                DatePaid: item[Consts.FIELDS.PAYMENT.PAYMENTPAID] ? new Date(item[Consts.FIELDS.PAYMENT.PAYMENTPAID]) : null,
                Fees: item[Consts.FIELDS.PAYMENT.FEES] || [],
                Contacts: item[Consts.FIELDS.PAYMENT.CONTACTS] || [],
                BeneficiaryItem: beneficiaryItem,
                Beneficiary: item[Consts.FIELDS.PAYMENT.BENEFICIARY],
                Created: item[Consts.FIELDS.COMMON.CREATION_TIME],
                Modified: item[Consts.FIELDS.COMMON.MODIFICATION_TIME],
                DocumentSpace: item[Consts.FIELDS.COMMON.DOCUMENTS_SPACE]?.Url
            };
            return payment;
        } catch (error) {
            if (error instanceof Error && error.message.includes(strings.DataSheetItemDoesNotExist)) {
                const errorMessage = strings.DataSheetItemDoesNotExist;
                this.setState({ errorMessage });
                return null;

            }
            throw error;
        }
    };

    private init = async () => {
        let item: IPayment;
        const promises: Promise<void>[] = [];
        promises.push(this.initItem().then(async (opts) => {
            item = opts as IPayment;
        }));
        await Promise.all(promises);
        this.setState({ item });
    };

    private renderDatasheetInfo = (): JSX.Element | null => {
        const { item } = this.state;
        if (!item) {
            return null;
        }
        const leftColumns = ['Summary', 'Fees', 'Status', 'DueDate', 'DatePaid'];
        const rightColumns = ['Beneficiary', 'BeneficiaryItem', 'Contacts', 'VatFreeCost', 'Vat', 'Comments'];

        return (
            <div className={`${styles.datasheetInfoContainer}`}>
                <div className={styles.datasheetInfoContainerLeft}>
                    {leftColumns.map((key) => this.renderDatasheetInfoRow(key))}
                </div>
                <div className={styles.datasheetInfoContainerRight}>
                    {rightColumns.map((key) => key && this.renderDatasheetInfoRow(key))}
                </div>
            </div>
        );
    };

    private renderDatasheetInfoRow = (key: string) => {
        const { item } = this.state;
        const labelKey = `PaymentDataSheet${key}`;

        return (
            <div className="ms-Grid" key={key}>
                <div className={`ms-Grid-row ${styles.datasheetInfoRow}`}>
                    <div className="ms-Grid-col ms-sm12 ms-md12">
                        <div className="ms-Grid-col ms-sm3 ms-md3">
                            <div className={styles.datasheetInfoLabel}>
                                <Label>{`${strings[labelKey] || key}:`}</Label>
                            </div>
                        </div>
                        <div className="ms-Grid-col ms-sm9 ms-md9">
                            <div>
                                {this.renderDatasheetInfoValue(item, key)}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        );
    };

    private renderDatasheetInfoValue = (item: IPayment, fieldName: string) => {
        const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        const value = fieldName ? item[fieldName] : undefined;

        if (value === null || value === undefined || (Array.isArray(value) && value.length === 0)) {
            return (
                <TooltipHost content="No data" styles={tooltipStyles}>
                    <span className={styles.detailsListValue}>
                        <Icon iconName="Remove" className={styles.mutedIcon} />
                        <span className={styles.srOnly}>No data</span>
                    </span>
                </TooltipHost>
            );
        }

        if (fieldName === 'Summary') {
            //return <RichText isEditMode={false} value={value} />;
            const cleanSummary = DomHelper.cleanRichHtml(item[fieldName]);
            if (!cleanSummary) return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
            return <RichText isEditMode={false} value={cleanSummary} />;
        }

        if (fieldName === 'Comments') {
            //return <RichText isEditMode={false} value={value} />;
            const cleanComments = DomHelper.cleanRichHtml(item[fieldName]);
            if (!cleanComments) return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
            return <RichText isEditMode={false} value={cleanComments} />;
        }

        if (fieldName === 'Status') {
            return StatusPill(String(value));
        }

        if (fieldName === 'Fees') {
            return renderObjectPills(
                value as IFEE[],
                (fee) => (
                    <span key={fee.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(FormType.Fee)}&itemId=${fee.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{fee.Title}</a>
                        </span>
                    </span>
                )
            );
        }

        if (fieldName === 'Contacts') {
            return renderObjectPills(
                value as IContact[],
                (contact) => (
                    <span key={contact.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(FormType.Contact)}&itemId=${contact.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{contact.Title}</a>
                        </span>
                    </span>
                )
            );
        }
        if (fieldName === 'Beneficiary') {
            if (value) {
                return <span className={styles.detailsListValue}>{String(value)}</span>;
            }
        }
        if (fieldName === 'BeneficiaryItem') {
            /*if (item.BeneficiaryItem) {
                return (
                    <a
                        data-interception="on"
                        rel="noreferrer"
                        title={item.BeneficiaryItem.text}
                        onClick={(e) => {
                            e.preventDefault();
                            const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(item.BeneficiaryItem.type)}&itemId=${item.BeneficiaryItem.id}`;
                            UrlHelper.navigate(url, false);
                        }}
                    >
                        {item.BeneficiaryItem.text}
                    </a>
                );
            }*/
            if (item.BeneficiaryItem) {
                const itemValue = value;
                const viewType = itemValue.type;
                if (itemValue && viewType) {
                    return (<div className={styles.termPills}>
                        <span key={itemValue.key} className={styles.termPill}>
                            <span className={styles.termPillText}>
                                <a
                                    data-interception="on"
                                    rel="noreferrer"
                                    onClick={(e) => {
                                        e.preventDefault();
                                        const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${viewType}&itemId=${itemValue.id}`;
                                        UrlHelper.navigate(url, false);
                                    }}
                                    title={String(itemValue?.text)}>{String(itemValue?.text)}</a>
                            </span>
                        </span>

                    </div>
                    );
                }
            }


            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        if (fieldName === 'VatFreeCost') {
            const currencyCode = item?.Currency && value > 0 ? `${item.Currency}`.trim() : '';
            return <span className={styles.detailsListValue}>{String(value)} {currencyCode}</span>;
        }

        if (fieldName === 'Vat') {
            const currencyCode = item?.Currency && value > 0 ? `${item.Currency}`.trim() : '';
            if (item.VatRate) {
                return <span className={styles.detailsListValue}>{String(value)} {currencyCode} {value > 0 ? `(${item.VatRate}%)` : ''}</span>;
            }
            return <span className={styles.detailsListValue}>{String(value)} {currencyCode}</span>;
        }

        if (fieldName === 'DueDate') {
            if (value instanceof Date) {
                const formatedDate = UtilHelper.formatDate(value, 'LL', 'en-us');
                return <span className={styles.detailsListValue}>{`${formatedDate}`}</span>;
            }
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        if (fieldName === 'DatePaid') {
            if (value instanceof Date) {
                const formatedDate = UtilHelper.formatDate(value, 'LL', 'en-us');
                return <span className={styles.detailsListValue}>{`${formatedDate}`}</span>;
            }
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        if (typeof value === 'string' || typeof value === 'number') {
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        return <span className={styles.detailsListValue}>{String(value)}</span>;
    };

    public render(): React.ReactElement<IPaymentDatasheetProps> {
        const {
            showEditFormDialog,
            showDeleteDialog,
            item,
            errorMessage
        } = this.state;
        const created = UtilHelper.formatDate(item?.Created, 'LL', 'en-us');
        const modified = UtilHelper.formatDate(item?.Modified, 'LL', 'en-us');
        const { pnpService, userService } = this.props;
        if (!this.hasAnyRole(UserRole.Owner, UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Admin)) {
            return (
                <AccessDeniedMessage message={strings.FormAccessDeniedMessage}> </AccessDeniedMessage>
            );
        }
        if (errorMessage) {
            if (errorMessage.includes(strings.DataSheetItemDoesNotExist)) {
                return (
                    <div className={styles.noDataMessage}>
                        <Icon iconName="Remove" className={styles.mutedIcon} />
                        <span>{strings.DataSheetNoRecordFound || 'No record found'}</span>
                        <Icon iconName="Remove" className={styles.mutedIcon} />
                    </div>
                );
            }
        }
        return (
            <>
                {this.state && this.state.isDatasheetReady ? (
                    <div>
                        <div className={styles.datasheetHeader}>
                            <div className={styles.datasheetTypeContainer}>
                                <h5 className={styles.datasheetType}>
                                    Payment
                                </h5>
                            </div>
                            {this.hasAnyRole(UserRole.Owner, UserRole.FinanceContributor, UserRole.Admin) && <div className={styles.datasheetCommandsContainer}>
                                <div className={styles.datasheetCommand} title={strings.DataSheetEditButton}>
                                    <a href="#" onClick={this.openEditForm}>
                                        <Icon iconName="Edit" />
                                    </a>
                                </div>
                                {showEditFormDialog && this.renderEditFormDialog()}
                                <div className={styles.datasheetCommand} title={strings.DataSheetDeleteButton}>
                                    <a href="#" onClick={this.openDeleteDialog}>
                                        <Icon iconName="Delete" />
                                    </a>
                                </div>
                                {showDeleteDialog && this.renderDeleteDialog()}
                            </div>}
                        </div>
                        <div className={styles.datasheetTitleContainer}>
                            <Icon iconName="DateTime" className={styles.dateIcon} />
                            <span className={styles.dateText}>
                                {`${strings.DataSheetLabelCreatedOn} ${created} / ${strings.DataSheetLabelUpdatedOn} ${modified}`}
                            </span>
                            <div className={styles.titleText}>
                                <h1>{item?.Title}</h1>
                            </div>
                        </div>
                        <div className={styles.tabsContainer}>
                            <Pivot linkSize="large">
                                <PivotItem headerText="Info" itemIcon="Info">
                                    <div className={styles.section}>
                                        <Icon iconName="TaskManager" className="sectionIcon" />
                                        <h3>{strings.DataSheetSectionProperties}</h3>
                                    </div>
                                    {this.renderDatasheetInfo()}
                                    <div className={styles.section}>
                                        <Icon iconName="Library" className="sectionIcon" />
                                        <h3 >{strings.DataSheetSectionDocs}</h3>
                                    </div>
                                    <LibraryView
                                        folderPath={item.DocumentSpace}
                                        pnpService={pnpService}
                                        userService={userService}
                                        pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}>
                                    </LibraryView>
                                </PivotItem>
                            </Pivot>
                        </div>
                    </div>
                ) : (
                    <div className="ms-Grid-col ms-sm12 ms-md12">
                        <div className={styles.spinnerContainer}>
                            <Spinner label="Loading..." />
                        </div>
                    </div>
                )}
            </>
        );
    }

    private openEditForm = () => {
        this.setState({ showEditFormDialog: true });
    };

    private closeEditForm = async (refreshData = false) => {
        this.setState({ showEditFormDialog: false });
        if (refreshData) {
            await this.init();
        }
    };

    private openDeleteDialog = () => {
        this.setState({ showDeleteDialog: true });
    };

    private closeDeleteDialog = () => {
        this.setState({ showDeleteDialog: false });
    };

    private deleteDataSheet = async () => {
        try {
            this.setState({ isFormProcessing: true });
            const { pnpService, itemId } = this.props;
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.PAYMENT_URL);
            await pnpService.getListItems(listUrl).getById(itemId).recycle();
            const messageBanner: IMessageBanner = {
                message: strings.DeleteSuccessMessage,
                type: 'success',
                visible: true
            };
            this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
            setTimeout(() => { window.location.href = Config.SITE.ROOT_WEB_URL; }, 2000);
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

    private renderEditFormDialog = () => {
        const { showEditFormDialog } = this.state;
        const { pnpService, userService, itemId } = this.props;

        const editFormControl = (
            <Payment
                pnpService={pnpService}
                userService={userService}
                itemId={itemId}
                callback={async () => {
                    await this.closeEditForm(true);
                }}
                formOrigin={FormOrigin.Datasheet}
            />
        );
        const editFormTitle = strings.PaymentDataSheetEditFormTitle || 'Edit Payment';

        return (
            <Modal
                isOpen={showEditFormDialog}
                onDismiss={() => this.closeEditForm(true)}
                containerClassName={styles.modalContainer}
                isBlocking={true}
            >
                <div>
                    <div className={styles.modalHeader}>
                        {editFormTitle}
                        <ActionButton
                            iconProps={{ iconName: 'Cancel' }}
                            onClick={() => this.closeEditForm()}
                            className={styles.iconButton}
                        >
                            {strings.DataSheetCancelButton}
                        </ActionButton>
                    </div>
                    <div className={styles.modalBody}>
                        {editFormControl}
                    </div>
                </div>
            </Modal>
        );
    };

    private renderDeleteDialog = () =>
    (<Dialog
        hidden={false}
        onDismiss={this.closeDeleteDialog}
        dialogContentProps={{
            type: DialogType.normal,
            title: strings.DataSheetDeleteModalTitle
        }}
        modalProps={{
            isBlocking: true
        }}
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
        <div>
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
                <span>{strings.DataSheetDeleteModalMessage}</span>
            )}
            <DialogFooter>
                <PrimaryButton
                    disabled={this.state.isConfirmButtonDisabled}
                    onClick={this.deleteDataSheet}
                //text={strings.DataSheetConfirmButton}
                >
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
                    onClick={this.closeDeleteDialog}
                    text={strings.DataSheetBackButton}
                />
            </DialogFooter>
        </div>
    </Dialog>);
}
