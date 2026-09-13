import * as React from 'react';
import { IContactDatasheetProps, IContactDatasheetState } from './Contact.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import {
    ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, ITooltipHostStyles, Label,
    LayerHost, Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost

} from '@fluentui/react';
import Contact from '../../../../genericForm/components/Contact/Contact';
import { Datasheet, DatasheetType, IDatasheetConfig } from '../../Datasheet.types';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { IContact } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { renderArrayPills } from '../../Datasheet.utility';
import { RichText } from '../../../../../common/components/RichText/RichText';
import { ContactType, UserRole } from '../../../../../common/models/Enums';
import { FormOrigin, IMessageBanner } from '../../../../genericForm/components/Form.types';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';
import AccessDeniedMessage from '../../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../../common/helpers/DomHelper';

export default class ContactDatasheet extends React.Component<
    IContactDatasheetProps,
    IContactDatasheetState
> {
    private config: IDatasheetConfig;
    private typeChoices: { key: string; text: string }[] = [
        { key: Consts.CONTENT_TYPES.EUTELSAL_CONTACT, text: ContactType.Eutelsat },
        { key: Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT, text: ContactType.Administration },
        { key: Consts.CONTENT_TYPES.ORGANIZATION_CONTACT, text: ContactType.Organization },
        { key: Consts.CONTENT_TYPES.REGULATOR_CONTACT, text: ContactType.Regulator },
        { key: Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT, text: ContactType.TP },
        { key: Consts.CONTENT_TYPES.LAW_FIRM_CONTACT, text: ContactType.LawFirm },
        { key: Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE_CONTACT, text: ContactType.LegalRepresentative },
        { key: Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT, text: ContactType.DP }
    ];
    constructor(props: Readonly<IContactDatasheetProps>) {
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
        this.config = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.Contact);
        //console.log("Teleport partner view config:", this.config);
    }

    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };
    public async componentDidMount() {
        if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner,
            UserRole.Contributor, UserRole.Visitor, UserRole.Admin)) {
            await this.init();
            this.setState({ isDatasheetReady: true });
        }

    }
    public initItem = async () => {
        const { pnpService, itemId } = this.props;
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.COMMON.CREATION_TIME,
            Consts.FIELDS.COMMON.MODIFICATION_TIME,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
            Consts.FIELDS.CONTACT.FIRST_NAME,
            Consts.FIELDS.CONTACT.LAST_NAME,
            Consts.FIELDS.CONTACT.COMMENTS,
            Consts.FIELDS.CONTACT.SUMMARY,
            Consts.FIELDS.CONTACT.DEPARTMENT,
            Consts.FIELDS.CONTACT.JOB_TITLE,
            Consts.FIELDS.CONTACT.EMAIL,
            Consts.FIELDS.CONTACT.PHONE_NUMBER,
            `${Consts.FIELDS.CONTACT.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.CONTACT.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.CONTACT.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.CONTACT.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,

        ];
        const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.Contact).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        try {
            const item = await this.props.pnpService.getListItems(listUrl).
                getById(itemId).
                select(...selectedFields).
                expand(Consts.FIELDS.CONTACT.AUTHORITY,
                    Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER,
                    Consts.FIELDS.CONTACT.TELEPORT_PARTNER,
                    Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER)();
            const filterTypeChoices = this.typeChoices.find((type) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1)?.key;
            let itemValue = null;
            let isEuetelsatContact = false;
            if (filterTypeChoices === Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT ||
                filterTypeChoices === Consts.CONTENT_TYPES.ORGANIZATION_CONTACT ||
                filterTypeChoices === Consts.CONTENT_TYPES.REGULATOR_CONTACT) {
                const authority = item[Consts.FIELDS.CONTACT.AUTHORITY];
                if (authority && authority[Consts.FIELDS.COMMON.ID] && authority[Consts.FIELDS.COMMON.TITLE]) {
                    itemValue = {
                        id: authority[Consts.FIELDS.COMMON.ID],
                        key: authority[Consts.FIELDS.COMMON.ID] + '-' + DatasheetType.Authority,
                        text: authority[Consts.FIELDS.COMMON.TITLE],
                        type: DatasheetType.Authority
                    };
                }
            } else if (filterTypeChoices === Consts.CONTENT_TYPES.LAW_FIRM_CONTACT ||
                filterTypeChoices === Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE_CONTACT) {
                const legalSvcProvider = item[Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER];
                if (legalSvcProvider && legalSvcProvider[Consts.FIELDS.COMMON.ID] && legalSvcProvider[Consts.FIELDS.COMMON.TITLE]) {
                    itemValue = {
                        id: legalSvcProvider[Consts.FIELDS.COMMON.ID],
                        key: legalSvcProvider[Consts.FIELDS.COMMON.ID] + '-' + DatasheetType.LegalSvcProvider,
                        text: legalSvcProvider[Consts.FIELDS.COMMON.TITLE],
                        type: DatasheetType.LegalSvcProvider
                    };

                }
            } else if (filterTypeChoices === Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT) {
                const distributionPartner = item[Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER];
                if (distributionPartner && distributionPartner[Consts.FIELDS.COMMON.ID] && distributionPartner[Consts.FIELDS.COMMON.TITLE]) {
                    itemValue = {
                        id: distributionPartner[Consts.FIELDS.COMMON.ID],
                        key: distributionPartner[Consts.FIELDS.COMMON.ID] + '-' + DatasheetType.DistributionPartner,
                        text: distributionPartner[Consts.FIELDS.COMMON.TITLE],
                        type: DatasheetType.DistributionPartner
                    };

                }
            } else if (filterTypeChoices === Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT) {
                const teleportPartner = item[Consts.FIELDS.CONTACT.TELEPORT_PARTNER];
                if (teleportPartner && teleportPartner[Consts.FIELDS.COMMON.ID] && teleportPartner[Consts.FIELDS.COMMON.TITLE]) {
                    itemValue = {
                        id: teleportPartner[Consts.FIELDS.COMMON.ID],
                        key: teleportPartner[Consts.FIELDS.COMMON.ID] + '-' + DatasheetType.TeleportPartner,
                        text: teleportPartner[Consts.FIELDS.COMMON.TITLE],
                        type: DatasheetType.TeleportPartner
                    };

                }
            } else {
                isEuetelsatContact = true;
            }
            const contact: IContact = {
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                ContentType: {
                    Id: item[Consts.FIELDS.COMMON.CONTENTTYPE_ID],
                    Name: this.getTypeChoiceTextByKey(item[Consts.FIELDS.COMMON.CONTENTTYPE_ID])
                },
                Created: item[Consts.FIELDS.COMMON.CREATION_TIME],
                Modified: item[Consts.FIELDS.COMMON.MODIFICATION_TIME],
                FirstName: item[Consts.FIELDS.CONTACT.FIRST_NAME] || '',
                LastName: item[Consts.FIELDS.CONTACT.LAST_NAME] || '',
                Summary: item[Consts.FIELDS.CONTACT.SUMMARY],
                Comments: item[Consts.FIELDS.CONTACT.COMMENTS],
                Department: item[Consts.FIELDS.CONTACT.DEPARTMENT],
                JobTitle: item[Consts.FIELDS.CONTACT.JOB_TITLE] || '',
                Email: item[Consts.FIELDS.CONTACT.EMAIL] || '',
                PhoneNumber: item[Consts.FIELDS.CONTACT.PHONE_NUMBER],
                Item: itemValue,
                EutelsatContact: isEuetelsatContact

            };
            return contact;
        } catch (error) {
            if (error instanceof Error && error.message.includes(strings.DataSheetItemDoesNotExist)) {
                const errorMessage = strings.DataSheetItemDoesNotExist;
                this.setState({ errorMessage });
                return null;

            }
            throw error;
        }
    };

    private getTypeChoiceTextByKey = (key: string | number): string | undefined => {
        const keyStr = String(key);
        const found = this.typeChoices.find((type) => keyStr.indexOf(String(type.key)) !== -1);
        return found ? found.text : "";
    };
    private init = async () => {
        let item: IContact;
        const promises: Promise<void>[] = [];
        promises.push(this.initItem().then((opts) => { item = opts as IContact; }));
        //promises.push(this.initSanctionCategoriesChoices().then((opts) => { sanctions = opts as ISanctionOption[]; }));
        await Promise.all(promises);

        /*const sanctionCategory = sanctions.find((cat) => cat.key.toString() === item.SanctionCategory.Id.toString());
        item.SanctionCategory = {
          Id: item.SanctionCategory.Id,
          Title: sanctionCategory.text,
          Color: sanctionCategory.color,
          Summary: sanctionCategory.summary
        };*/
        this.setState({
            item,
        });
        //console.log('item', item);
    };

    private renderDatasheetInfo = (): JSX.Element | null => {
        const { item } = this.state;
        if (!item) return null;
        const leftColumns = ['FirstName', 'LastName', 'Summary', !item.EutelsatContact ? 'Item' : null, 'EutelsatContact'];
        const rightColumns = ['Department', 'JobTitle', 'Email', 'PhoneNumber', 'Comments'];

        return (
            <div className={`${styles.datasheetInfoContainer}`}>
                <div className={styles.datasheetInfoContainerLeft}>
                    {leftColumns.map((key) =>
                        key && (this.renderDatasheetInfoRow(key))
                    )}
                </div>
                <div className={styles.datasheetInfoContainerRight}>
                    {rightColumns.map((key) =>
                        key && this.renderDatasheetInfoRow(key)
                    )}
                </div>
            </div>
        );
    };

    private renderDatasheetInfoRow = (key: string) => {
        const {
            item
        } = this.state;
        const labelKey = `ContactDataSheetInfoLabel${key}`;
        return (
            <div className="ms-Grid">
                <div className={`ms-Grid-row ${styles.datasheetInfoRow}`}>
                    <div className="ms-Grid-col ms-sm12 ms-md12">
                        <div className="ms-Grid-col ms-sm3 ms-md3">
                            <div className={styles.datasheetInfoLabel}>
                                <Label>{`${strings[labelKey]}:`}</Label>
                            </div>
                        </div>
                        <div className="ms-Grid-col ms-sm9 ms-md9">
                            <div>
                                {this.renderDatasheetInfoValue(item, key)}
                            </div>
                        </div>
                    </div>
                </div>
            </div>);
    };


    private renderDatasheetInfoValue = (item, fieldName) => {
        //const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        //const value = fieldName ? item[fieldName] : undefined;
        /*const swatchStyle = (color?: string) => ({
          width: 12,
          height: 12,
          borderRadius: 3,
          background: color || '#888',
          display: 'inline-block',
          verticalAlign: 'middle' as const,
          marginRight: 6,
          border: '1px solid rgba(0,0,0,.1)',
        });*/

        const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        const itemDatasheetUrl = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(this.config.type)}&itemId=${item.Id}`;
        //if (!column) return <span />;
        //const fieldName = column.fieldName as keyof IContact | undefined;
        const value = fieldName ? item[fieldName] : undefined;
        if (value === null || value === undefined || (Array.isArray(value) && value.length === 0)) {
            return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
        }
        if (fieldName === 'Title') {
            return (
                <a
                    data-interception="on"
                    rel="noreferrer"
                    onClick={(e) => {
                        e.preventDefault();
                        UrlHelper.navigate(itemDatasheetUrl, false);
                    }}
                >
                    {String(value)}
                </a>
            );
        }
        if (fieldName === 'ContentType') {
            const ct = value as { Id: string; Name: string };
            return <span className={styles.detailsListValue}>{String(ct?.['Name'])}</span>;
        }
        if (fieldName === 'Type') {
            return (
                <div className={styles.tooltipHostContainer} title="">

                    <span className={styles.tooltipHostText}> {value?.Name}</span>
                </div>
            );
        }
        if (fieldName === 'Item') {
            const itemValue = value;
            const viewType = itemValue.type;
            return (<div className={styles.termPills}>
                <span key={itemValue.key} className={styles.termPill}>
                    <span className={styles.termPillText}>
                        <a
                            data-interception="on"
                            rel="noreferrer"
                            onClick={(e) => {
                                e.preventDefault();
                                const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(viewType)}&itemId=${itemValue.id}`;
                                UrlHelper.navigate(url, false);
                            }}
                        >
                            {String(itemValue?.text)}
                        </a>
                    </span>
                </span>
            </div>
            );
        }
        if (fieldName === 'Comments' || fieldName === 'Address' || fieldName === 'Summary') {
           // return <RichText isEditMode={false} value={item[fieldName]} />;
            const cleanValue = DomHelper.cleanRichHtml(item[fieldName]);
            if (!cleanValue) return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
            return <RichText isEditMode={false} value={cleanValue} />;
        }
        if (typeof value === 'boolean') {
            const iconName = value ? 'AcceptMedium' : 'Cancel';
            const iconClass = value ? styles.iconGreen : styles.iconRed;
            return <div className={styles.detailsListValue}><Icon iconName={iconName} className={iconClass} /></div>;
        }
        if (value instanceof Date) {
            return <span className={styles.detailsListValue}>{UtilHelper.formatDate(value, 'LL', 'en-us')}</span>;
        }
        if (typeof value === 'string' || typeof value === 'number') {
            return <span className={styles.detailsListValue} title={String(value)}>{String(value)}</span>;
        }
        if (Array.isArray(value) && value?.every((v) => typeof v === 'string' || typeof v === 'number')) {
            return renderArrayPills(value as (string | number)[]);
        }

    };

    public render(): React.ReactElement<IContactDatasheetProps> {
        const {
            showEditFormDialog,
            showDeleteDialog,
            item,
            errorMessage
        } = this.state;
        const created = UtilHelper.formatDate(item?.Created, 'LL', 'en-us');
        const modified = UtilHelper.formatDate(item?.Modified, 'LL', 'en-us');
        if (!this.hasAnyRole(UserRole.Contributor, UserRole.Owner, UserRole.Visitor,
            UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Admin)) {
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
                    </div>
                );
            }
        }
        return (
            <>
                {
                    this.state && this.state.isDatasheetReady ?

                        <div>
                            <div className={styles.datasheetHeader}>
                                <div className={styles.datasheetTypeContainer}>
                                    < h5 className={styles.datasheetType} >
                                        {strings.ContactDataSheetTitle}
                                    </h5 >
                                </div>
                                {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.datasheetCommandsContainer}>
                                    {<div className={styles.datasheetCommand} title={strings.DataSheetEditButton}>
                                        <a href="#" onClick={this.openEditForm} >
                                            <Icon iconName="Edit" />
                                        </a>
                                    </div>
                                    }
                                    {showEditFormDialog &&
                                        this.renderEditFormDialog()
                                    }
                                    {<div className={styles.datasheetCommand} title={strings.DataSheetDeleteButton}>
                                        <a href="#" className="" onClick={this.openDeleteDialog}>
                                            <Icon iconName="Delete" />
                                        </a>
                                    </div>}
                                    {showDeleteDialog && this.renderDeleteDialog()}
                                </div>}
                            </div>
                            <div className={styles.datasheetTitleContainer}>
                                <Icon iconName='DateTime' className={styles.dateIcon} />
                                <span className={styles.dateText} >
                                    {`${strings.DataSheetLabelCreatedOn} ${created} / ${strings.DataSheetLabelUpdatedOn} ${modified}`}
                                </span>
                                <div className={styles.titleText}>
                                    <h1>{item?.Title}</h1>
                                </div>
                            </div>
                            <div className={styles.tabsContainer}>
                                <Pivot linkSize="large" >
                                    <PivotItem headerText="Info" itemIcon="Info" >
                                        <div className={styles.section}>
                                            <Icon iconName="TaskManager" className="sectionIcon" />
                                            <h3 >{strings.DataSheetSectionProperties}</h3>
                                        </div>
                                        <LayerHost
                                            id='sanctionCategoryLayer'
                                            style={{ position: 'relative', zIndex: 500000 }} // higher than page chrome
                                        />
                                        {this.renderDatasheetInfo()}
                                    </PivotItem>
                                </Pivot>
                            </div>
                        </div > :
                        <div className="ms-Grid-col ms-sm12 ms-md12">
                            <div className={styles.spinnerContainer}>
                                <Spinner label={'Loading...'} />
                            </div>
                        </div>
                }
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
            const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.Contact).list;
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
            await pnpService.getListItems(listUrl).getById(itemId).recycle();
            const messageBanner: IMessageBanner = {
                message: strings.DeleteSuccessMessage,
                type: 'success',
                visible: true
            };
            this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
            //setTimeout(() => { window.location.href = Config.SITE.ROOT_WEB_URL; }, 2000);
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

        const {
            showEditFormDialog,
        } = this.state;
        const {
            pnpService,
            userService,
            itemId
        } = this.props;

        const editFormControl = (<Contact pnpService={pnpService}
            userService={userService} itemId={itemId} callback={async () => { await this.closeEditForm(true); }} formOrigin={FormOrigin.Datasheet}></Contact>);
        const editFormTitle = strings.ContactDataSheetEditFormTitle;

        return (<Modal
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
                    >{strings.DataSheetCancelButton}</ActionButton>
                </div>
                <div className={styles.modalBody}>
                    {editFormControl}
                </div>
            </div>
        </Modal>);
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
