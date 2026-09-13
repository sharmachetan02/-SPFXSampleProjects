import * as React from 'react';
import { IFeeDatasheetProps, IFeeDatasheetState } from './FeeDatasheet.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import {
    ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, ITooltipHostStyles, Label,
    MessageBar,
    MessageBarType,
    Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost
} from '@fluentui/react';
import Fee from '../../../../genericForm/components/Fee/Fee';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { FormType, IMessageBanner } from '../../../../genericForm/components/Form.types';
import { IFEE, IItemOption, IPayment } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { RichText } from '../../../../../common/components/RichText/RichText';
import { StatusPill } from '../../../../genericView/components/View.utility';
//import PaymentView from '../../../genericView/components/Payment View/PaymentView';
import { IViewConfig, View } from '../../../../genericView/components/View.types';
import LibraryView from '../../../../genericView/components/Library View/LibraryView';
import PaymentDatasheetView from '../../Datasheet Views/PaymentDatasheetView';
import { FeeChargedBy, UserRole } from '../../../../../common/models/Enums';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';
import { DatasheetType } from '../../Datasheet.types';
import DomHelper from '../../../../../common/helpers/DomHelper';

export default class FeeDatasheet extends React.Component<
    IFeeDatasheetProps,
    IFeeDatasheetState
> {
    private paymentConfig: IViewConfig;
    private beneficiaryItemChoices: IItemOption[] = [];
    constructor(props: Readonly<IFeeDatasheetProps>) {
        super(props);
        this.state = {
            isDatasheetReady: false,
            showEditFormDialog: false,
            showDeleteDialog: false,
            item: null,
            paymentViewItems: [],
            errorMessage: null,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'error',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        };
        this.paymentConfig = View.config.find((f) => f.type.toLowerCase() === DatasheetType.Payment);
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
            Consts.FIELDS.FEE.SUMMARY,
            Consts.FIELDS.FEE.STATUS,
            Consts.FIELDS.FEE.COMMENTS,
            Consts.FIELDS.COMMON.CREATION_TIME,
            Consts.FIELDS.COMMON.MODIFICATION_TIME,
            Consts.FIELDS.COMMON.DOCUMENTS_SPACE,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
            Consts.FIELDS.FEE.CATEGORY,
            Consts.FIELDS.FEE.VATRATE,
            Consts.FIELDS.FEE.DUEDATE,
            Consts.FIELDS.FEE.VATFREECOST,
            Consts.FIELDS.FEE.VAT,
            Consts.FIELDS.FEE.COSTTYPE,
            Consts.FIELDS.FEE.PONONPO,
            Consts.FIELDS.FEE.CHARGEDBY,
            Consts.FIELDS.FEE.FEETYPE,
            Consts.FIELDS.FEE.VATFREE,
            Consts.FIELDS.FEE.REPEATINGFEE,
            `${Consts.FIELDS.FEE.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.CURRENCY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.CURRENCY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.CURRENCY}/${Consts.FIELDS.FEE.CURRENCY_ISOCODE}`,
            `${Consts.FIELDS.FEE.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.SNP}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.SNP}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.MA_REQUIREMENT}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.MA_REQUIREMENT}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.FEE.RECURRENCEPATTERN_ID
        ];

        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.FEE_URL);
        try {
            const item = await pnpService
                .getListByUrl(listUrl)
                .items.getById(itemId)
                .select(...selectedFields)
                .expand(
                    Consts.FIELDS.FEE.AUTHORITY,
                    Consts.FIELDS.FEE.CURRENCY,
                    Consts.FIELDS.FEE.LEGALSERVICEPROVIDER,
                    Consts.FIELDS.FEE.SNP,
                    Consts.FIELDS.FEE.MA_REQUIREMENT,
                    Consts.FIELDS.FEE.TELEPORTPARTNER
                )();
            const chargedByItem: IItemOption[] = [];
            if (item[Consts.FIELDS.FEE.CHARGEDBY]) {
                switch (item[Consts.FIELDS.FEE.CHARGEDBY]) {
                    case FeeChargedBy.LawFirm:
                    case FeeChargedBy.LegalRepresentative:
                        chargedByItem.push({
                            id: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.LegalSvcProvider,
                            text: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.TITLE],
                            type: DatasheetType.LegalSvcProvider

                        });
                        break;
                    case FeeChargedBy.Administration:
                    case FeeChargedBy.Organization:
                    case FeeChargedBy.Regulator:
                        chargedByItem.push({
                            id: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.Authority,
                            text: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.TITLE],
                            type: DatasheetType.Authority,
                        });
                        break;
                    case FeeChargedBy.TeleportPartner:
                        chargedByItem.push({
                            id: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.TeleportPartner,
                            text: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
                            type: DatasheetType.TeleportPartner
                        });
                        break;
                    default:
                        break;
                }
            }

            const fee: IFEE = {
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                Summary: item[Consts.FIELDS.FEE.SUMMARY],
                Category: item[Consts.FIELDS.FEE.CATEGORY],
                Status: item[Consts.FIELDS.FEE.STATUS],
                Comments: item[Consts.FIELDS.FEE.COMMENTS],
                VatFreeCost: item[Consts.FIELDS.FEE.VATFREECOST],
                Vat: item[Consts.FIELDS.FEE.VAT],
                VatRate: item[Consts.FIELDS.FEE.VATRATE],
                Currency: item[Consts.FIELDS.FEE.CURRENCY]?.[Consts.FIELDS.FEE.CURRENCY_ISOCODE] ?? '',
                CostType: item[Consts.FIELDS.FEE.COSTTYPE],
                PoNonPO: item[Consts.FIELDS.FEE.PONONPO],
                ChargedBy: item[Consts.FIELDS.FEE.CHARGEDBY],
                ChargeByItem: chargedByItem.length > 0 ? chargedByItem[0] : null,
                DueDate: item[Consts.FIELDS.FEE.DUEDATE] ? new Date(item[Consts.FIELDS.FEE.DUEDATE]) : null,
                RecurrencPatternId: item[Consts.FIELDS.FEE.RECURRENCEPATTERN_ID],
                Item: item[Consts.FIELDS.FEE.SNP] ? {
                    key: item[Consts.FIELDS.FEE.SNP][Consts.FIELDS.COMMON.ID],
                    text: item[Consts.FIELDS.FEE.SNP][Consts.FIELDS.COMMON.TITLE],
                    type: FormType.SNP
                } : item[Consts.FIELDS.FEE.MA_REQUIREMENT] ? {
                    key: item[Consts.FIELDS.FEE.MA_REQUIREMENT][Consts.FIELDS.COMMON.ID],
                    text: item[Consts.FIELDS.FEE.MA_REQUIREMENT][Consts.FIELDS.COMMON.TITLE],
                    type: FormType.MARequirementGeneric
                } : null,
                Created: item[Consts.FIELDS.COMMON.CREATION_TIME],
                Modified: item[Consts.FIELDS.COMMON.MODIFICATION_TIME],
                DocumentSpace: item[Consts.FIELDS.COMMON.DOCUMENTS_SPACE]?.Url,
                RecurranceFee: item[Consts.FIELDS.FEE.REPEATINGFEE] || false,
            };
            return fee;
        } catch (error) {
            if (error instanceof Error && error.message.includes(strings.DataSheetItemDoesNotExist)) {
                const errorMessage = strings.DataSheetItemDoesNotExist;
                this.setState({ errorMessage });
                return null;

            }
            throw error;
        }
    };

    private fetchRecurrencePatternFrequency = async (recurrencePatternId: number): Promise<string> => {
        if (!recurrencePatternId) {
            return '';
        }

        try {
            const { pnpService } = this.props;
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.RECURRENCE_PATTERN_URL);
            const item = await pnpService
                .getListByUrl(listUrl)
                .items.getById(recurrencePatternId)
                .select(
                    Consts.FIELDS.RECURRENCE_PATTERN.FREQUENCY,
                    Consts.FIELDS.RECURRENCE_PATTERN.WEEKDAY,
                    Consts.FIELDS.RECURRENCE_PATTERN.MONTH,
                    Consts.FIELDS.RECURRENCE_PATTERN.DAY,
                    Consts.FIELDS.RECURRENCE_PATTERN.INTERVAL,
                    Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE
                )()
                .catch(() => null);

            if (!item) {
                return '';
            }

            return this.formatRecurrencePattern(item);
        } catch (error) {
            console.error('Error fetching recurrence pattern frequency:', error);
            return '';
        }
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private formatRecurrencePattern = (item: Record<string, any>): string => {
        const frequency = item[Consts.FIELDS.RECURRENCE_PATTERN.FREQUENCY] || '';
        const weekDay = item[Consts.FIELDS.RECURRENCE_PATTERN.WEEKDAY] || '';
        const month = item[Consts.FIELDS.RECURRENCE_PATTERN.MONTH] || '';
        const day = item[Consts.FIELDS.RECURRENCE_PATTERN.DAY] || 1;
        const interval = item[Consts.FIELDS.RECURRENCE_PATTERN.INTERVAL] || 1;
        const endDate = item[Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE] || null;

        let result = '';

        switch (frequency) {
            case 'Weekly':
                result = `${weekDay} of every ${interval} week${interval > 1 ? 's' : ''}`;
                break;
            case 'Monthly':
                result = `Day ${day} of every ${interval} month${interval > 1 ? 's' : ''}`;
                break;
            case 'Quarterly':
                result = `Day ${day} of every 3 months`;
                break;
            case 'Half Annual':
                result = `Day ${day} of every 6 months`;
                break;
            case 'Annual':
                result = `Every ${month} ${day} of every 1 year`;
                break;
            case 'Multi Annual':
                result = `Every ${month} ${day} of every ${interval} year${interval > 1 ? 's' : ''}`;
                break;
            default:
                result = frequency;
        }

        if (frequency && endDate) {
            const date = UtilHelper.formatDate(endDate, 'LL', 'en-us');
            result += ` and expired on ${date} `;
        }

        return result;
    };

    private init = async () => {
        let item: IFEE;
        let paymentsItems: IPayment[];
        const promises: Promise<void>[] = [];
        promises.push(this.initItem().then(async (opts) => {
            item = opts as IFEE;
            if (item.RecurrencPatternId) {
                item.Frequency = await this.fetchRecurrencePatternFrequency(item.RecurrencPatternId);
            }
        }));

        promises.push(this.initPaymentViewItems().then((opts) => { paymentsItems = opts as IPayment[]; }));
        await Promise.all(promises);
        this.setState({
            item,
            paymentViewItems: paymentsItems

        });

    };

    public initPaymentViewItems = async () => {
        const { itemId } = this.props;

        if (!itemId) {
            return;
        }

        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.PAYMENT.SUMMARY,
            Consts.FIELDS.PAYMENT.STATUS,
            Consts.FIELDS.PAYMENT.COMMENTS,
            Consts.FIELDS.PAYMENT.SUMMARY,
            Consts.FIELDS.PAYMENT.VATRATE,
            Consts.FIELDS.PAYMENT.DUEDATE,
            Consts.FIELDS.PAYMENT.PAYMENTPAID,
            Consts.FIELDS.PAYMENT.VATFREECOST,
            Consts.FIELDS.PAYMENT.VAT,
            Consts.FIELDS.PAYMENT.VATRATE,
            Consts.FIELDS.PAYMENT.BENEFICIARY,
            Consts.FIELDS.FEE.CHARGEDBY,
            `${Consts.FIELDS.PAYMENT.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.PAYMENT.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.PAYMENT.CURRENCY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.CURRENCY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.PAYMENT.CURRENCY}/${Consts.FIELDS.FEE.CURRENCY_ISOCODE}`,
            `${Consts.FIELDS.PAYMENT.FEES}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.FEES}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.PAYMENT.CONTACTS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.PAYMENT.CONTACTS}/${Consts.FIELDS.COMMON.TITLE}`
        ];
        const expand = [Consts.FIELDS.PAYMENT.AUTHORITY, Consts.FIELDS.PAYMENT.CURRENCY, Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER,
        Consts.FIELDS.PAYMENT.FEES, Consts.FIELDS.PAYMENT.CONTACTS, Consts.FIELDS.PAYMENT.TELEPORTPARTNER];
        const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.paymentConfig.list);
        const filterQuery = `${Consts.FIELDS.PAYMENT.FEES_ID} eq ${itemId}`;
        const items = await this.props.pnpService.getListItems(listUrl).
            select(...selectedFields)
            .expand(...expand)
            .filter(filterQuery)
            .orderBy(Consts.FIELDS.COMMON.TITLE)();
        const viewItems: IPayment[] = [];
        items.map((item) => {
            const rawDueDate = item[Consts.FIELDS.PAYMENT.DUEDATE];
            const DueDate: Date | null = rawDueDate ? new Date(rawDueDate) : null;
            const rawDatePaid = item[Consts.FIELDS.PAYMENT.PAYMENTPAID];
            const DatePaid: Date | null = rawDatePaid ? new Date(rawDatePaid) : null;
            let beneficiaryItem = null;
            switch (item[Consts.FIELDS.FEE.CHARGEDBY]) {
                case FeeChargedBy.LawFirm:
                case FeeChargedBy.LegalRepresentative:
                    beneficiaryItem = {
                        id: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID] +
                            "-" + DatasheetType.LegalSvcProvider,
                        text: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.TITLE],
                        type: DatasheetType.LegalSvcProvider
                    };
                    this.beneficiaryItemChoices.push(beneficiaryItem);
                    break;
                case FeeChargedBy.Administration:
                case FeeChargedBy.Organization:
                case FeeChargedBy.Regulator:
                    beneficiaryItem = {
                        id: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.Authority,
                        text: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.TITLE],
                        type: DatasheetType.Authority,
                    };
                    this.beneficiaryItemChoices.push(beneficiaryItem);
                    break;
                case FeeChargedBy.TeleportPartner:
                    beneficiaryItem = {
                        id: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.TeleportPartner,
                        text: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
                        type: DatasheetType.TeleportPartner,
                    };
                    this.beneficiaryItemChoices.push(beneficiaryItem);
                    break;
                default:
                    break;
            }
            const vatFreeCost = parseInt(item[Consts.FIELDS.PAYMENT.VATFREECOST]);
            const vat = parseInt(item[Consts.FIELDS.PAYMENT.VAT]);
            const total = (!isNaN(vatFreeCost) ? vatFreeCost : 0) + (!isNaN(vat) ? vat : 0);
            const currencyCode = item[Consts.FIELDS.PAYMENT.CURRENCY]?.[Consts.FIELDS.FEE.CURRENCY_ISOCODE] ?? '';
            viewItems.push({
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                Fees: item[Consts.FIELDS.PAYMENT.FEES]?.map((country) => ({
                    Id: country[Consts.FIELDS.COMMON.ID],
                    Title: country[Consts.FIELDS.COMMON.TITLE]
                })) || [],
                Summary: item[Consts.FIELDS.PAYMENT.SUMMARY],
                VatFreeCost: item[Consts.FIELDS.PAYMENT.VATFREECOST],
                Vat: item[Consts.FIELDS.PAYMENT.VAT],
                Status: item[Consts.FIELDS.PAYMENT.STATUS],
                CostTotal: total ? total.toString() + " " + currencyCode : '0',
                DueDate,
                DatePaid,
                Item: beneficiaryItem
            });
        });
        return viewItems;
    };

    private renderDatasheetInfo = (): JSX.Element | null => {
        const { item } = this.state;
        if (!item) {
            return null;
        }
        const leftColumns = ['Summary', 'Category', 'Item', 'Status', item.RecurranceFee ? 'Frequency' : 'DueDate', 'CostType'];
        const rightColumns = ['ChargedBy', 'ChargeByItem', 'VatFreeCost', 'Vat', 'PoNonPO', 'Comments'];

        return (
            <div className={`${styles.datasheetInfoContainer}`}>
                <div className={styles.datasheetInfoContainerLeft}>
                    {leftColumns.map((key) => key && this.renderDatasheetInfoRow(key))}
                </div>
                <div className={styles.datasheetInfoContainerRight}>
                    {rightColumns.map((key) => key && this.renderDatasheetInfoRow(key))}
                </div>
            </div>
        );
    };

    private renderDatasheetInfoRow = (key: string) => {
        const { item } = this.state;
        const labelKey = `FeeDataSheet${key}`;

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

    private renderDatasheetInfoValue = (item: IFEE, fieldName: string) => {
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

        if (fieldName === 'Item') {
            /*if (item.Item) {
                return (
                    <a
                        data-interception="on"
                        rel="noreferrer"
                        onClick={(e) => {
                            e.preventDefault();
                            const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(item.Item.type)}&itemId=${item.Item.key}`;
                            UrlHelper.navigate(url, false);
                        }}
                    >
                        {item.Item.text}
                    </a>
                );
            }*/
            const itemValue = value;
            const datasheetType = itemValue.type;
            if (itemValue && datasheetType) {
                return (<div className={styles.termPills}>
                    <span key={itemValue.key} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${item.Item.type}&itemId=${item.Item.key}`;
                                    UrlHelper.navigate(url, false);
                                }}
                                title={String(itemValue?.text)}>{String(itemValue?.text)}</a>
                        </span>
                    </span>

                </div>
                );
            }
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        if (fieldName === 'ChargedBy') {
            /*if (item.ChargedBy) {
                return (
                    <a
                        data-interception="on"
                        rel="noreferrer"
                        onClick={(e) => {
                            e.preventDefault();
                            const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(item.ChargeByItem.type)}&itemId=${item.ChargeByItem.id}`;
                            UrlHelper.navigate(url, false);
                        }}
                    >
                        {item.ChargeByItem.text}
                    </a>
                );
            }*/
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }
        if (fieldName === 'ChargeByItem') {
            const itemValue = value;
            const datasheetType = itemValue.type;
            if (itemValue && datasheetType) {
                return (<div className={styles.termPills}>
                    <span key={itemValue.key} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${datasheetType}&itemId=${item.ChargeByItem.id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                                rel="noreferrer" title={String(itemValue?.text)}>{String(itemValue?.text)}</a>
                        </span>
                    </span>

                </div>
                );
            }
        }

        if (fieldName === 'Vat') {
            //const itemValue = value;
            const currencyCode = item?.Currency && value > 0 ? `${item.Currency}`.trim() : '';

            if (item.VatRate) {
                return <span className={styles.detailsListValue}>{String(value)} {currencyCode} {value > 0 ? `(${item.VatRate}%)` : ''}</span>;
            }
            return <span className={styles.detailsListValue}>{String(value)} {currencyCode}</span>;
        }

        if (fieldName === 'CostType') {
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        if (fieldName === 'PoNonPO') {
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        if (fieldName === 'Frequency') {
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        if (fieldName === 'DueDate') {
            if (value instanceof Date) {

                const formatedDate = UtilHelper.formatDate(value, 'LL', 'en-us');
                return <span className={styles.detailsListValue}>{`${formatedDate}`}</span>;
            }
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        if (fieldName === 'VatFreeCost') {
            const currencyCode = item?.Currency && value > 0 ? `${item.Currency}`.trim() : '';
            return <span className={styles.detailsListValue}>{String(value)} {currencyCode}</span>;
        }

        if (typeof value === 'string' || typeof value === 'number') {
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }

        return <span className={styles.detailsListValue}>{String(value)}</span>;
    };

    public render(): React.ReactElement<IFeeDatasheetProps> {
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
                <MessageBar messageBarType={MessageBarType.error}>
                    {strings.FormAccessDeniedMessage}
                </MessageBar>
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
                {this.state && this.state.isDatasheetReady ? (
                    <div>
                        <div className={styles.datasheetHeader}>
                            <div className={styles.datasheetTypeContainer}>
                                <h5 className={styles.datasheetType}>
                                    Fee
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
                                {this.hasAnyRole(UserRole.Owner, UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Admin) && <PivotItem headerText="Payments" itemIcon="Bank">
                                    <PaymentDatasheetView
                                        items={this.state.paymentViewItems}
                                        renderField={(item, key) => this.renderDatasheetInfoValue(item, key)}
                                        pnpService={pnpService}
                                        userService={userService}
                                        refresh={this.init}
                                        paymentItemId={item.Id}

                                    />
                                </PivotItem>}
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
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.FEE_URL);
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
            <Fee
                pnpService={pnpService}
                userService={userService}
                itemId={itemId}
                callback={async () => {
                    await this.closeEditForm(true);
                }}
            />
        );
        const editFormTitle = strings.FeeDataSheeteditFormTitle || 'Edit Fee';

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
