import * as React from 'react';
import { IDPDatasheetProps, IDPDatasheetState } from './DPDatasheet.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import {
    ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, IDropdownOption, ITooltipHostStyles, Label,
    Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost

} from '@fluentui/react';
import { Datasheet, DatasheetType } from '../../Datasheet.types';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { IContact, IContentType, IDP, IItemOption, IMARequirement } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { renderArrayPills, renderObjectPills } from '../../Datasheet.utility';
import { RichText } from '../../../../../common/components/RichText/RichText';
import { StatusPill } from '../../../../genericView/components/View.utility';
import ContactDatasheetView from '../../Datasheet Views/ContactDatasheetView';

import DistributionPartner from '../../../../genericForm/components/Distribution Partner/DistributionPartner';

import MARequirementView from '../../../../genericView/components/MARequirement View/MARequirementView';
import { IViewConfig, View, ViewType } from '../../../../genericView/components/View.types';
import { MARequirementResponsibleParty, MARequirementType, UserRole } from '../../../../../common/models/Enums';
import LibraryView from '../../../../genericView/components/Library View/LibraryView';
import AccessDeniedMessage from '../../../../../common/components/Access Denied/AccessDenied';
import { IMessageBanner } from '../../../../genericForm/components/Form.types';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';
import DomHelper from '../../../../../common/helpers/DomHelper';

export default class DPDatasheet extends React.Component<
    IDPDatasheetProps,
    IDPDatasheetState
> {

    private responsibleItemChoices: IItemOption[] = [];
    private requestedItemChoices: IDropdownOption[] = [];
    private config: IViewConfig;

    private typeChoices: IDropdownOption[] = [
        { key: Consts.CONTENT_TYPES.MA_RequirementCredential, text: MARequirementType.Credential },
        { key: Consts.CONTENT_TYPES.MA_RequirementGeneric, text: MARequirementType.Generic },
        { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_NOL, text: MARequirementType.NOL },
        { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_LEGAL, text: MARequirementType.Legal },
    ];
    constructor(props: Readonly<IDPDatasheetProps>) {
        super(props);
        this.state = {
            isDatasheetReady: false,
            showEditFormDialog: false,
            showDeleteDialog: false,
            item: null,
            maReqViewItems: null,
            contactViewItems: [],
            errorMessage: null,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'error',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        };
        this.config = View.config.find((f) => f.type.toLowerCase() === ViewType.MARequirement);
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
            Consts.FIELDS.DISTRIBUTION_PARTNER.SUMMARY,
            Consts.FIELDS.DISTRIBUTION_PARTNER.ADDRESS,
            Consts.FIELDS.DISTRIBUTION_PARTNER.STATUS,
            Consts.FIELDS.DISTRIBUTION_PARTNER.COMMENTS,
            Consts.FIELDS.DISTRIBUTION_PARTNER.PORTAL,
            Consts.FIELDS.COMMON.CREATION_TIME,
            Consts.FIELDS.COMMON.MODIFICATION_TIME,
            Consts.FIELDS.COMMON.DOCUMENTS_SPACE,
            Consts.FIELDS.DISTRIBUTION_PARTNER.OWNERS_ID,
            `${Consts.FIELDS.DISTRIBUTION_PARTNER.OWNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.DISTRIBUTION_PARTNER.OWNER}/${Consts.FIELDS.COMMON.TITLE}`
        ];

        const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.DistributionPartner).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        try {
            const item = await pnpService
                .getListByUrl(listUrl)
                .items.getById(itemId)
                .select(...selectedFields)
                .expand(Consts.FIELDS.DISTRIBUTION_PARTNER.OWNER)();

            const dp: IDP = {
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                Summary: item[Consts.FIELDS.DISTRIBUTION_PARTNER.SUMMARY],
                Address: item[Consts.FIELDS.DISTRIBUTION_PARTNER.ADDRESS],
                Status: item[Consts.FIELDS.DISTRIBUTION_PARTNER.STATUS],
                Comments: item[Consts.FIELDS.DISTRIBUTION_PARTNER.COMMENTS],
                Portal: item[Consts.FIELDS.DISTRIBUTION_PARTNER.PORTAL],
                Owners: item[Consts.FIELDS.DISTRIBUTION_PARTNER.OWNER] || [],
                Created: item[Consts.FIELDS.COMMON.CREATION_TIME],
                Modified: item[Consts.FIELDS.COMMON.MODIFICATION_TIME],
                DocumentSpace: item[Consts.FIELDS.COMMON.DOCUMENTS_SPACE]?.Url
            };
            return dp;

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
        let item: IDP;
        let maReqItems: IMARequirement[];
        let contactItems: IContact[];
        const promises: Promise<void>[] = [];
        promises.push(this.initItem().then((opts) => { item = opts as IDP; }));
        promises.push(this.fetchContactDetails().then((opts) => { contactItems = opts as IContact[]; }));
        promises.push(this.initMAReqViewItems().then((opts) => { maReqItems = opts as IMARequirement[]; }));
        await Promise.all(promises);
        this.setState({
            item,
            contactViewItems: contactItems,
            maReqViewItems: maReqItems
        });

    };
    public initMAReqViewItems = async () => {
        const { itemId } = this.props;
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
            Consts.FIELDS.MAGENERIC.SUMMARY,
            `${Consts.FIELDS.MAGENERIC.MARKETS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.MARKETS}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.VERTICALS,
            `${Consts.FIELDS.MAGENERIC.SYNTHESIS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.SYNTHESIS}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.RAGSTATUS,
            `${Consts.FIELDS.MAGENERIC.SNPS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.SNPS}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY,
            `${Consts.FIELDS.MAGENERIC.EUTELSATENTITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.EUTELSATENTITY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.CONTACTS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.CONTACTS}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.EUTELSATOWNERS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.EUTELSATOWNERS}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.GEOLEO,
            Consts.FIELDS.MAGENERIC.COMMENTS,
            Consts.FIELDS.MAGENERIC.APPLICATIONDATE,
            Consts.FIELDS.MAGENERIC.ESTIMATEDDATE,
            Consts.FIELDS.MAGENERIC.EFFECTIVEDATE,
            Consts.FIELDS.MAGENERIC.EXPIRATIONDATE,
            `${Consts.FIELDS.MAGENERIC.RELATEDITEMS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.RELATEDITEMS}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
        ];
        const expand = [Consts.FIELDS.MAGENERIC.MARKETS, Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE,
        Consts.FIELDS.MAGENERIC.SYNTHESIS, Consts.FIELDS.MAGENERIC.SNPS, Consts.FIELDS.MAGENERIC.EUTELSATENTITY, Consts.FIELDS.MAGENERIC.TELEPORTPARTNER,
        Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER, Consts.FIELDS.MAGENERIC.CONTACTS, Consts.FIELDS.MAGENERIC.EUTELSATOWNERS,
        Consts.FIELDS.MAGENERIC.RELATEDITEMS, Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT, Consts.FIELDS.MAGENERIC.AUTHORITY
        ];
        const filterQuery = `${Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER_ID} eq ${itemId}`;
        const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.config.list);
        const items = await this.props.pnpService.getListItems(listUrl).
            select(...selectedFields)
            .expand(...expand)
            .filter(filterQuery)
            .orderBy(Consts.FIELDS.COMMON.TITLE)();
        const viewItems: IMARequirement[] = [];
        items.map((item) => {
            const rawApplicationDate = item[Consts.FIELDS.MAGENERIC.APPLICATIONDATE];
            const ApplicationDate: Date | null = rawApplicationDate ? new Date(rawApplicationDate) : null;
            const rawEffectiveDate = item[Consts.FIELDS.MAGENERIC.EFFECTIVEDATE];
            const EffectiveDate: Date | null = rawEffectiveDate ? new Date(rawEffectiveDate) : null;
            const rawExpirationDate = item[Consts.FIELDS.MAGENERIC.EXPIRATIONDATE];
            const ExpirationDate: Date | null = rawExpirationDate ? new Date(rawExpirationDate) : null;
            const responsibleParty = item[Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY];
            let responsiblePartyItem = null;
            switch (responsibleParty) {
                case MARequirementResponsibleParty.Eutelsat:
                    if (item?.[Consts.FIELDS.MAGENERIC.EUTELSATENTITY]?.[Consts.FIELDS.COMMON.ID] && item?.[Consts.FIELDS.MAGENERIC.EUTELSATENTITY]?.[Consts.FIELDS.COMMON.TITLE]) {
                        responsiblePartyItem = {
                            id: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID] + "-" + ViewType.EutelsatEntity,
                            text: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.TITLE],
                            type: ViewType.EutelsatEntity,
                        };
                        this.responsibleItemChoices.push(responsiblePartyItem);
                    }
                    break;
                case MARequirementResponsibleParty.TP:
                    if (item?.[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER]?.[Consts.FIELDS.COMMON.ID] &&
                        item?.[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER]?.[Consts.FIELDS.COMMON.TITLE]) {
                        responsiblePartyItem = {
                            id: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + ViewType.TeleportPartner,
                            text: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
                            type: ViewType.TeleportPartner,
                        };
                        this.responsibleItemChoices.push(responsiblePartyItem);
                    }
                    break;
                case MARequirementResponsibleParty.DP:
                    if (item?.[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.ID] && item?.[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.TITLE]) {
                        responsiblePartyItem = {
                            id: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID] + "-" + ViewType.DistributionPartner,
                            text: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE],
                            type: ViewType.DistributionPartner,
                        };
                        this.responsibleItemChoices.push(responsiblePartyItem);
                    }
                    break;
                default:
                    break;
            }
            const Type: IContentType = this.typeChoices
                .map((type) => ({ Id: type.key + "", Name: type.text }))
                .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1) || { Id: '', Name: '' };
            const authority = item?.[Consts.FIELDS.MAGENERIC.AUTHORITY];
            const RequestedPartyItem = authority ? {
                Id: authority[Consts.FIELDS.COMMON.ID],
                Title: authority[Consts.FIELDS.COMMON.TITLE]
            } : null;
            if (RequestedPartyItem) {
                this.requestedItemChoices.push({
                    key: RequestedPartyItem.Id,
                    text: RequestedPartyItem.Title
                });
            }
            viewItems.push({
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                MaVertical: item[Consts.FIELDS.MAGENERIC.VERTICALS]?.map((v) => ({
                    key: v,
                    text: v
                })),
                Type,
                Markets: item[Consts.FIELDS.MAGENERIC.MARKETS]?.map((market) => ({
                    Id: market[Consts.FIELDS.COMMON.ID],
                    Title: market[Consts.FIELDS.COMMON.TITLE]
                })) || [],
                Summary: item[Consts.FIELDS.PAYMENT.SUMMARY],
                ResponsiblePartyItem: responsiblePartyItem,
                RequestedPartyItem,
                ApplicationDate,
                ExpirationDate,
                EffectiveDate,
                RagStatus: item[Consts.FIELDS.MAGENERIC.RAGSTATUS],
                Synthesis: (item?.[Consts.FIELDS.MAGENERIC.SYNTHESIS]?.[Consts.FIELDS.COMMON.ID] &&
                    item?.[Consts.FIELDS.MAGENERIC.SYNTHESIS]?.[Consts.FIELDS.COMMON.TITLE]) ? {
                    Id: item[Consts.FIELDS.MAGENERIC.SYNTHESIS][Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.MAGENERIC.SYNTHESIS][Consts.FIELDS.COMMON.TITLE]
                } : null,
            });
        });
        return viewItems;
    };
    private fetchContactDetails = async () => {
        const { pnpService, itemId } = this.props;
        if (!itemId) {
            return;
        }
        const contactList = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.CONTACT_URL);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.CONTACT.JOB_TITLE,
            Consts.FIELDS.CONTACT.EMAIL,
            Consts.FIELDS.CONTACT.PHONE_NUMBER,
            Consts.FIELDS.CONTACT.DEPARTMENT,
            Consts.FIELDS.CONTACT.FIRST_NAME,
            Consts.FIELDS.CONTACT.LAST_NAME,
            Consts.FIELDS.CONTACT.SUMMARY,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID
        ];

        const filterQuery = `${Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER_ID} eq ${itemId}`;
        const contacts = await pnpService
            .getListByUrl(contactList)
            .items.select(...selectedFields)
            .filter(filterQuery)();

        const contactDetails: IContact[] = [];
        for (const contact of contacts) {
            contactDetails.push({
                Id: contact[Consts.FIELDS.COMMON.ID],
                Title: contact[Consts.FIELDS.COMMON.TITLE],
                JobTitle: contact[Consts.FIELDS.CONTACT.JOB_TITLE],
                FirstName: contact[Consts.FIELDS.CONTACT.FIRST_NAME],
                LastName: contact[Consts.FIELDS.CONTACT.LAST_NAME],
                Department: contact[Consts.FIELDS.CONTACT.DEPARTMENT],
                Email: contact[Consts.FIELDS.CONTACT.EMAIL],
                PhoneNumber: contact[Consts.FIELDS.CONTACT.PHONE_NUMBER],
                Summary: contact[Consts.FIELDS.CONTACT.SUMMARY]
            });
        }
        return contactDetails;
    };

    private renderDatasheetInfo = (): JSX.Element | null => {
        const leftColumns = ['Summary', 'Address', 'Status'];
        const rightColumns = ['Portal', 'Owners', 'Comments'];

        return (
            <div className={`${styles.datasheetInfoContainer}`}>
                <div className={styles.datasheetInfoContainerLeft}>
                    {leftColumns.map((key) =>
                        (this.renderDatasheetInfoRow(key))
                    )}
                </div>
                <div className={styles.datasheetInfoContainerRight}>
                    {rightColumns.map((key) =>
                        (this.renderDatasheetInfoRow(key))
                    )}
                </div>
            </div>
        );
    };

    private renderDatasheetInfoRow = (key: string) => {
        const {
            item
        } = this.state;
        const labelKey = `DPDataSheetInfoLabel${key}`;
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

    private renderDatasheetInfoValue = (item: IDP, fieldName: string) => {
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

        if (fieldName === 'Address') {
            //return <RichText isEditMode={false} value={value} />;
            const cleanAddress = DomHelper.cleanRichHtml(item[fieldName]);
            if (!cleanAddress) return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
            return <RichText isEditMode={false} value={cleanAddress} />;
        }

        if (fieldName === 'Status') {
            return StatusPill(String(value));
        }
        if (fieldName === 'Portal') {
            return <a data-interception="off" href={String(value.Url)} target="_blank" rel="noreferrer">Link</a>;
        }
        if (fieldName === 'Owners') {
            return renderObjectPills(
                value as IContact[],
                (owner) => (
                    <span key={owner.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a data-interception="on" rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Contact)}&itemId=${owner.Id}`;
                                    UrlHelper.navigate(url, false);
                                }} title={owner.Title} >{owner.Title}</a>
                        </span>
                    </span>
                )
            );
        }

        if (typeof value === 'string' || typeof value === 'number') {
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        if (Array.isArray(value) && value?.every((v) => typeof v === 'string' || typeof v === 'number')) {
            return renderArrayPills(value as (string | number)[]);
        }

        return <span className={styles.detailsListValue}>{String(value)}</span>;
    };



    public render(): React.ReactElement<IDPDatasheetProps> {
        const {
            showEditFormDialog,
            showDeleteDialog,
            item,
            maReqViewItems,
            errorMessage
        } = this.state;
        const { pnpService, userService } = this.props;
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
                {this.state && this.state.isDatasheetReady ?

                    <div>
                        <div className={styles.datasheetHeader}>
                            <div className={styles.datasheetTypeContainer}>
                                <h5 className={styles.datasheetType}>
                                    {strings.DPDataSheetTitle}
                                </h5>
                            </div>
                            {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.datasheetCommandsContainer}>
                                {<div className={styles.datasheetCommand} title={strings.DataSheetEditButton}>
                                    <a href="#" onClick={this.openEditForm}>
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
                                    <div className={styles.section}>
                                        <Icon iconName="People" className="sectionIcon" />
                                        <h3>{strings.DataSheetSectionContacts}</h3>
                                    </div>
                                    <ContactDatasheetView
                                        contacts={this.state.contactViewItems}
                                        pnpService={pnpService}
                                        userService={userService}
                                        refresh={() => this.init()}
                                        relatedItemId={item?.Id}
                                        relatedItemType={strings.DPDatasheetTypeDistributionPartner}
                                    />
                                </PivotItem>
                                <PivotItem headerText="MA Requirements" itemIcon="Bullseye" >
                                    <MARequirementView pnpService={pnpService} userService={userService}
                                        columns={['Title', 'Summary', 'MaVertical', 'Type', 'Countries', 'RequestedPartyItem', 'RagStatus', 'Synthesis',
                                            'Geo/Leo', 'ApplicationDate', 'EffectiveDate', 'ExpirationDate']}
                                        filters={['Search', 'MaVertical', 'Type', 'Countries', 'RequestedPartyItem', 'RagStatus', 'Synthesis', 'Geo/Leo']}
                                        viewItems={maReqViewItems}
                                        pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}
                                    ></MARequirementView>
                                </PivotItem>
                            </Pivot>
                        </div>
                    </div> :
                    <div className="ms-Grid-col ms-sm12 ms-md12">
                        <div className={styles.spinnerContainer}>
                            <Spinner label="Loading..." />
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
            const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.DistributionPartner).list;
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
        const { showEditFormDialog } = this.state;
        const { pnpService, userService, itemId } = this.props;



        const editFormControl = (
            <DistributionPartner
                pnpService={pnpService}
                userService={userService}
                itemId={itemId}
                callback={async () => {
                    await this.closeEditForm(true);
                }}
            />
        );

        const editFormTitle = strings.DPDataSheetEditFormTitle;

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
