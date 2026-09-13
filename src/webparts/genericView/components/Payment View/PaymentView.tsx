import * as React from 'react';
import { IPaymentViewState, IPaymentViewProps, IItemOption } from './PaymentView.type';
import styles from '../View.module.scss';
import { IViewConfig, View, ViewType } from '../View.types';
import 'react-widgets/styles.css';
import { Config } from '../../../../common/config/Config';
import {
    ConstrainMode, DetailsList, DetailsListLayoutMode, IColumn, Icon, IDropdownOption, ITooltipHostStyles, Label,
    Panel,
    PanelType,
    PrimaryButton, SearchBox, SelectionMode, Spinner, TooltipHost
} from '@fluentui/react';
import { IFilter, LogicalOperator } from '../../../../common/models/IFilter';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import * as strings from 'GenericViewPartStrings';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { Consts } from '../../../../common/consts/Consts';
import { IPayment } from '../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../common/helpers/Util';
import { SortDirection } from '@pnp/sp/search';
import { renderArrayPills, renderObjectPills, StatusPill } from '../View.utility';
import AppliedFilters from '../AppliedFilters/AppliedFilters';
import Pager from '../Pager/Pager';
import { CallOutButton } from '../../../../common/components/CallOutButton/CallOutButton';
import { RichText } from '../../../../common/components/RichText/RichText';
import { ContactType, UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';

export default class PaymentView extends React.Component<
    IPaymentViewProps,
    IPaymentViewState
> {
    private config: IViewConfig;
    private allViewItems: IPayment[] = [];
    private beneficiaryItemChoices: IItemOption[] = [];
    private feesTypeChoices: IDropdownOption[] = [];

    constructor(props: Readonly<IPaymentViewProps>) {
        super(props);
        this.state = {
            isViewReady: false,
            columns: [],
            filters: [],
            sort: null,
            searchKeyword: '',
            viewItems: [],
            pageIndex: 0,
            pageSize: 25,
            isFilterPanelOpen: false
        };
        this.config = View.config.find((f) => f.type.toLowerCase() === ViewType.Payment);
    }

    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };
    public async componentDidMount() {
        if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner, UserRole.Admin)) {
            await this.init();
            this.setState({ isViewReady: true });
        }
    }

    public async componentDidUpdate(prev: IPaymentViewProps) {
        if (
            prev.columns !== this.props.columns ||
            prev.filters !== this.props.filters ||
            prev.sort !== this.props.sort ||
            prev.viewItems !== this.props.viewItems
        ) {
            await this.init();
        }
    }

    private setSearchText = (searchKeyword?: string) => {
        const { filters } = this.state;
        const searchFilter = filters.find((f) => f.name === 'Search');
        searchFilter.selectedValues = searchKeyword ? [searchKeyword] : [];
        this.setState({ filters, searchKeyword });
    };

    private initColumns(): IColumn[] {
        const allColumns = [{
            key: 'Title',
            name: strings.PaymentViewTitleColumnLabel,
            fieldName: 'Title',
            minWidth: 300,
            maxWidth: 300,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Summary',
            name: strings.PaymentViewSummaryColumnLabel,
            fieldName: 'Summary',
            minWidth: 90,
            maxWidth: 90,
            isResizable: true,
        },
        {
            key: 'Fees',
            name: strings.PaymentViewFeesTypeColumnLabel,
            fieldName: 'Fees',
            minWidth: 180,
            maxWidth: 180,
            isResizable: true,
        },
        {
            key: 'Item',
            name: strings.PaymentViewItemColumnLabel,
            fieldName: 'Item',
            minWidth: 180,
            maxWidth: 180,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'CostTotal',
            name: strings.PaymentViewCostTotalColumnLabel,
            fieldName: 'CostTotal',
            minWidth: 90,
            maxWidth: 90,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Status',
            name: strings.PaymentViewStatusColumnLabel,
            fieldName: 'Status',
            minWidth: 90,
            maxWidth: 90,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'DueDate',
            name: strings.PaymentViewDueDateColumnLabel,
            fieldName: 'DueDate',
            minWidth: 120,
            maxWidth: 120,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'DatePaid',
            name: strings.PaymentViewDateDateColumnLabel,
            fieldName: 'DatePaid',
            minWidth: 120,
            maxWidth: 120,
            isResizable: true,
            onColumnClick: this.onColumnClick
        }
        ];
        return this.props.columns ?
            allColumns.filter((c) => this.isColumnEnabled(c.key)) :
            allColumns;
    }

    private initFeesChoices = async () => {
        /*const {
            pnpService
        } = this.props;
        const countriesChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.FEE_URL);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE
        ];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        items.forEach((item) => {
            countriesChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
        });
        return countriesChoices;*/

        const feesTypeChoices: IDropdownOption[] = [];
        this.feesTypeChoices.forEach((item) => {
            if (!feesTypeChoices.find((i) => i.key === item.key)) {
                feesTypeChoices.push(item);
            }
        });
        feesTypeChoices.sort((a, b) => a.text?.localeCompare(b.text));
        return feesTypeChoices;
    };

    public initStatusChoices = async () => {
        const {
            pnpService
        } = this.props;
        let statusChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.PAYMENT.STATUS)
            .select('Choices')();
        statusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
        return statusChoices;
    };

    public initBeneficiaryItemChoices = async (): Promise<IItemOption[]> => {
        const itemChoices: IItemOption[] = [];
        this.beneficiaryItemChoices.forEach((item) => {
            if (!itemChoices.find((i) => i.key === item.key && i.type === item.type)) {
                itemChoices.push(item);
            }
        });
        itemChoices.sort((a, b) => a.text?.localeCompare(b.text));
        return itemChoices;
    };

    public initViewItems = async () => {
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
        const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.config.list);
        const items = await this.props.pnpService.getListItems(listUrl).
            select(...selectedFields)
            .expand(...expand)
            .orderBy(Consts.FIELDS.COMMON.TITLE)();
        const viewItems: IPayment[] = [];
        items.map((item) => {
            const rawDueDate = item[Consts.FIELDS.PAYMENT.DUEDATE];
            const DueDate: Date | null = rawDueDate ? new Date(rawDueDate) : null;
            const rawDatePaid = item[Consts.FIELDS.PAYMENT.PAYMENTPAID];
            const DatePaid: Date | null = rawDatePaid ? new Date(rawDatePaid) : null;
            let beneficiaryItem = null;
            item[Consts.FIELDS.PAYMENT.FEES]?.forEach((fee) => {
                this.feesTypeChoices.push({ key: fee[Consts.FIELDS.COMMON.ID], text: fee[Consts.FIELDS.COMMON.TITLE] });
            });
            switch (item[Consts.FIELDS.FEE.CHARGEDBY]) {
                case ContactType.LawFirm:
                case ContactType.LegalRepresentative:
                    beneficiaryItem = {
                        id: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID] +
                            "-" + ViewType.LegalRepresentative,
                        text: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.TITLE],
                        type: ViewType.LegalRepresentative
                    };
                    this.beneficiaryItemChoices.push(beneficiaryItem);
                    break;
                case ContactType.Administration:
                case ContactType.Organization:
                case ContactType.Regulator:
                    beneficiaryItem = {
                        id: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID] + "-" + ViewType.Authority,
                        text: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.TITLE],
                        type: ViewType.Authority,
                    };
                    this.beneficiaryItemChoices.push(beneficiaryItem);
                    break;
                case ContactType.TP:
                    beneficiaryItem = {
                        id: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + ViewType.TeleportPartner,
                        text: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
                        type: ViewType.TeleportPartner,
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

    private initFilters = (): IFilter[] => {
        let allFilters: IFilter[] = [];
        const searchFilter: IFilter = {
            name: 'Search',
            displayName: strings.PaymentViewTitleFilterLabel,
            placeHolder: strings.PaymentViewTitleFilterPlaceHolder,
            fields: ['Title', 'Item', 'CostTotal', 'Summary', 'Status', 'Fees'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.Or,
        };
        const FeesTypeFilter: IFilter = {
            name: 'Fees',
            displayName: strings.PaymentViewFeesTypeFilterLabel,
            placeHolder: strings.PaymentViewFeesTypeFilterPlaceHolder,
            fields: ['Fees'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const ItemTypeFilter: IFilter = {
            name: 'Item',
            displayName: strings.PaymentViewItemFilterLabel,
            placeHolder: strings.PaymentViewItemFilterPlaceHolder,
            fields: ['Item'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const StatusTypeFilter: IFilter = {
            name: 'Status',
            displayName: strings.PaymentViewStatusFilterLabel,
            placeHolder: strings.PaymentViewStatusFilterPlaceHolder,
            fields: ['Status'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        allFilters = [searchFilter, FeesTypeFilter, ItemTypeFilter, StatusTypeFilter];
        return this.props.filters ? allFilters.filter((f) => this.isFilterEnabled(f.name)) : allFilters;
    };

    private init = async () => {
        const columns = this.initColumns();
        let filters = this.initFilters();
        let items: IPayment[];
        if (this.props.viewItems) {
            items = this.props.viewItems.slice();
            this.props.viewItems.forEach((itm) => {
                if (itm.Item) {
                    this.beneficiaryItemChoices.push({
                        id: itm.Item.id,
                        key: itm.Item.key,
                        text: itm.Item.text,
                        type: itm.Item.type
                    });

                }
                if (itm.Fees) {
                    itm.Fees.forEach((fee) => {
                        this.feesTypeChoices.push({ key: fee.Id, text: fee.Title });
                    });
                }
            });

        } else {
            items = await this.initViewItems();
        }

        let feesType: IDropdownOption[] = [];
        let status: IDropdownOption[] = [];
        let itemType: IItemOption[] = [];
        const promises: Promise<void>[] = [];
        if (this.isFilterEnabled('Fees')) {
            promises.push(this.initFeesChoices().then((opts) => { feesType = opts as IDropdownOption[]; }));
        }
        if (this.isFilterEnabled('Item')) {
            promises.push(this.initBeneficiaryItemChoices().then((opts) => { itemType = opts as IItemOption[]; }));
        }
        if (this.isFilterEnabled('Status')) {
            promises.push(this.initStatusChoices().then((opts) => { status = opts as IDropdownOption[]; }));
        }
        await Promise.all(promises);
        filters = filters.map((f) => {
            if (f.name === 'Fees') return { ...f, values: feesType };
            if (f.name === 'Item') return { ...f, values: itemType };
            if (f.name === 'Status') return { ...f, values: status };
            return f;
        });
        const sort = this.props.sort ?? { Property: 'Title', Direction: SortDirection.Ascending };
        const sortedCols = columns.map((c) => ({
            ...c,
            isSorted: c.key === sort.Property || c.fieldName === sort.Property,
            isSortedDescending: sort.Direction === SortDirection.Descending
        }));
        this.allViewItems = items;
        this.setState({
            columns: sortedCols,
            filters,
            sort,
            pageIndex: 0,
            viewItems: items
        });

    };

    private isFilterEnabled = (name: string) =>
        !this.props.filters || this.props.filters.includes(name);

    private isColumnEnabled = (keyOrField: string) =>
        !this.props.columns || this.props.columns.includes(keyOrField);

    public applyFilters = (): IPayment[] => {
        const { filters } = this.state;
        let filtered = this.allViewItems;
        filters.forEach((filter) => {
            filtered = filtered.filter((item) =>
                filter.fields.some((field) => {
                    const selected = filter.selectedValues || [];
                    if (selected.length === 0) return true;
                    const assertions = selected.map((f) => {
                        if (filter.name === 'Search') {
                            const needle = UtilHelper.toNormalForm(f?.toString()?.toLowerCase());
                            if (!needle) return true;
                            if (field === 'Fees') {
                                return (item[field] || []).some((c) => {
                                    const hay = UtilHelper.toNormalForm(c?.Title?.toString()?.toLowerCase());
                                    return hay?.includes(needle) === true;
                                });
                            } else if (field === 'Item') {
                                // If Item is an object (IItem), search in Title and Type
                                const itemValue = item[field] as IItemOption;
                                if (!itemValue) return false;
                                const hayTitle = UtilHelper.toNormalForm(itemValue.text?.toString()?.toLowerCase());
                                return (hayTitle?.includes(needle) === true);
                            } else {
                                const hay = UtilHelper.toNormalForm(item[field]?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            }
                        }

                        if (filter.name === 'Fees') {
                            return (item[field] || []).some((t) => String(t.Id) === String((f as IDropdownOption).key));
                        }
                        if (filter.name === 'Status') {
                            return String(item[field]?.toString()) === String((f as IDropdownOption).key);
                        }
                        if (filter.name === 'Item') {
                            const itemValue = item[field];
                            if (!itemValue) return false;
                            return String(itemValue.key) === String((f as IDropdownOption).key) && itemValue.type === (f as IItemOption).type;
                        }
                        return true;
                    });
                    return filter.operatorBetweenValues === LogicalOperator.And ?
                        assertions.every(Boolean) :
                        assertions.some(Boolean);
                })
            );
        });
        return filtered;
    };

    private clearFilters = (): void => {
        let { filters } = this.state;
        filters = filters.map((filter) => {
            filter.selectedValues = [];
            return filter;
        });
        this.setState({ filters, searchKeyword: '', pageIndex: 0 });
    };

    private sort = (items) => {
        const { sort } = this.state;
        if (!sort) return items;
        const key = sort.Property as keyof IPayment;
        const desc = sort.Direction === SortDirection.Descending;
        return items.slice().sort((a, b) => {
            let av = a[key];
            let bv = b[key];
            if (key === 'DueDate' || key === 'DatePaid') {
                const distantPast = new Date(-1970);
                const dateA = a[key] ? new Date(a[key]) : distantPast;
                const dateB = b[key] ? new Date(b[key]) : distantPast;
                return desc ? dateB.getTime() - dateA.getTime() : dateA.getTime() - dateB.getTime();
            }
            if (key === 'Item') {
                av = av?.text ?? '';
                bv = bv?.text ?? '';
            }
            const aStr = ((av || '') as string).toLowerCase();
            const bStr = ((bv || '') as string).toLowerCase();
            if (aStr === bStr) return 0;
            const cmp = aStr > bStr ? 1 : -1;
            return desc ? -cmp : cmp;
        });
    };

    private renderItemColumn = (item, index: number, column: IColumn) => {
        const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        const itemDatasheetUrl = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(this.config.type)}&itemId=${item.Id}`;
        if (!column) return <span />;
        const fieldName = column.fieldName as keyof IPayment | undefined;
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
        if (fieldName === 'Fees') {
            const feesListType = View.config.find((f) => f.type.toLowerCase() === ViewType.FEE).type;
            return renderObjectPills(
                value,
                (item) => (
                    <span className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a onClick={(e) => {
                                e.preventDefault();
                                const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(feesListType)}&itemId=${item?.Id}`;
                                UrlHelper.navigate(url, false);
                            }}
                                rel="noreferrer" >{item?.Title}</a>
                        </span>
                    </span>
                )
            );
        }
        if (fieldName === 'Item') {
            const itemValue = value;
            const viewType = itemValue.type;
            return (<div className={styles.termPills}>
                <span key={itemValue.key} className={styles.termPill}>
                    <span className={styles.termPillText}>
                        <a
                            onClick={(e) => {
                                e.preventDefault();
                                const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${viewType}&itemId=${itemValue.id}`;
                                UrlHelper.navigate(url, false);
                            }}
                            rel="noreferrer" title={String(itemValue?.text)}>{String(itemValue?.text)}</a>
                    </span>
                </span>
            </div>
            );
        }
        if (fieldName === 'Status') {
            return StatusPill(String(value));
        }
        if (fieldName === 'Summary') {
            /*return (
                <CallOutButton icon={'View'} name={'Preview'} >
                    <RichText isEditMode={false} value={item[fieldName]} />
                </CallOutButton>
            );*/
            const cleanSummary = DomHelper.cleanRichHtml(item[fieldName]);
            if (!cleanSummary) return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
            return (
                <CallOutButton icon={'View'} name={'Preview'} >
                    <RichText isEditMode={false} value={cleanSummary} />
                </CallOutButton>
            );
        }
        if (fieldName === 'CostTotal') {
            return (
                <div className={styles.tooltipHostContainer} title="">
                    <span className={styles.tooltipHostText}> {value}</span>
                </div>
            );
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
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }
        if (Array.isArray(value) && value?.every((v) => typeof v === 'string' || typeof v === 'number')) {
            return renderArrayPills(value as (string | number)[]);
        }
    };

    private onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        let { columns } = this.state;
        let isSortedDescending = column.isSortedDescending;
        if (column.isSorted) {
            isSortedDescending = !isSortedDescending;
        }
        columns = columns.map((col) => {
            col.isSorted = col.key === column.key;
            if (col.isSorted) {
                col.isSortedDescending = isSortedDescending;
            }
            return col;
        });
        const sort = {
            Property: column.key,
            Direction: isSortedDescending ? SortDirection.Descending : SortDirection.Ascending
        };

        this.setState({
            sort,
            columns,
            pageIndex: 0
        });
    };

    private onChangeFeestype = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const contentTypeFilter = filters.find((f) => f.name === 'Fees');
        contentTypeFilter.selectedValues = options;
        filters = filters.map(((f) => ((f.name === contentTypeFilter.name) ? contentTypeFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeItem = (options: IItemOption[]) => {
        let { filters } = this.state;
        const itemFilter = filters.find((f) => f.name === 'Item');
        itemFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === itemFilter.name) ? itemFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeStatus = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const countryFilter = filters.find((f) => f.name === 'Status');
        countryFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === countryFilter.name) ? countryFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    public refreshFilters = (filters: IFilter[]) => {
        const { searchKeyword } = this.state;
        const searchFilter = filters.find((f) => f.name === 'Search');
        this.setState({ filters, searchKeyword: searchFilter.selectedValues.length > 0 ? searchKeyword : '' });
    };

    public onSearch = (filter: IFilter, newValue): void => {
        let { filters } = this.state;
        filter.selectedValues = [].concat(newValue);
        filters = filters.map(((f) => ((f.name === filter.name) ? filter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    public onChange = (filter: IFilter, searchKeyword): void => {
        let { filters } = this.state;
        if (!searchKeyword) {
            filter.selectedValues = [];
            filters = filters.map(((f) => ((f.name === filter.name) ? filter : f)));
            this.setState({ filters, searchKeyword });
        } else {
            this.setState({ searchKeyword });
        }

    };

    // Go to a specific page (0-based)
    private setPage = (page: number) => {
        this.setState({ pageIndex: Math.max(0, page) });
    };

    // Change page size and reset to first page
    private setPageSize = (size: number) => {
        const pageSize = Math.max(1, size || 1);
        this.setState({ pageSize, pageIndex: 0 });
    };

    private openFilterPanel = () => this.setState({ isFilterPanelOpen: true });
    private dismissFilterPanel = () => this.setState({ isFilterPanelOpen: false });

    public render(): React.ReactElement<IPaymentViewProps> {
        const { filters, columns, searchKeyword, pageIndex, pageSize, isFilterPanelOpen } = this.state;
        const feesTypeFilter = filters.find((f) => f.name === 'Fees');
        const itemFilter = filters.find((f) => f.name === 'Item');
        const statusTypeFilter = filters.find((f) => f.name === 'Status');
        const searchFilter = filters.find((f) => f.name === 'Search');
        let viewItems = this.applyFilters();
        viewItems = this.sort(viewItems);
        const pageCount = Math.max(1, Math.ceil(viewItems.length / pageSize));
        const safeIndex = Math.min(pageIndex, pageCount - 1);
        const start = safeIndex * pageSize;
        const end = start + pageSize;
        const pagedItems = viewItems.slice(start, end);
        const itemsToDisplay = this.props.pagination?.isEnabled ? pagedItems : viewItems;
        if (!this.hasAnyRole(UserRole.Owner, UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Admin)) {
            return (
                <AccessDeniedMessage message={strings.FormAccessDeniedMessage}> </AccessDeniedMessage>
            );
        }
        return (
            <>
                {
                    this.state && this.state.isViewReady ?
                        <div className={styles.listView}>
                            {this.allViewItems && this.allViewItems.length > 0 && (
                                <>
                                    <div className={styles.applyFiltersBtn}>

                                        <PrimaryButton
                                            text="Filters"
                                            iconProps={{ iconName: 'Filter' }}
                                            className={`${styles.clearButton} "form-button"`}
                                            styles={{
                                                icon: { order: 0, marginRight: 4 }, // moves icon after text
                                            }}
                                            onClick={() => this.openFilterPanel()}
                                        />

                                    </div>

                                    <Panel                               // Always open
                                        isOpen={isFilterPanelOpen}
                                        // Let users interact with page behind it
                                        isBlocking={false}
                                        // Dismiss is disabled (always on)
                                        onDismiss={() => this.dismissFilterPanel()}
                                        type={PanelType.smallFixedFar} // pick: smallFixedFar / smallFixedNear / medium / large
                                        headerText={'Filters'}
                                        closeButtonAriaLabel="Close"
                                        // Optional: remove light dismiss since we never close
                                        isLightDismiss={false}

                                        styles={{
                                            scrollableContent: {
                                                height: '100%'
                                            },
                                            main: {
                                                selectors: {
                                                    '@media (min-width: 480px)': {
                                                        width: '300px !important', // Override media query
                                                    },
                                                },

                                            }
                                        }}
                                    >

                                        <div className={styles.containerPaneldiv}>
                                            <div className={styles.clearFiltersBtn}>
                                                <PrimaryButton
                                                    text="Clear Filters"
                                                    title="Clear"
                                                    disabled={false}
                                                    onClick={this.clearFilters}
                                                    className={`${styles.clearButton} "form-button"`}
                                                    iconProps={{ iconName: 'ClearFilter' }}
                                                    styles={{
                                                        icon: { order: 0, marginRight: 4 }, // moves icon after text
                                                    }}
                                                /></div>
                                            {/* Search */}
                                            {
                                                this.isFilterEnabled('Search') &&
                                                <>
                                                    <div>
                                                        <Label className={styles.filterLabel}>{searchFilter.displayName}</Label>
                                                    </div>

                                                    <div className={styles.panelChildDiv}>
                                                        <SearchBox
                                                            placeholder={searchFilter.placeHolder}
                                                            autoComplete='off'
                                                            value={searchKeyword}
                                                            onSearch={(newValue) => this.onSearch(searchFilter, newValue)}
                                                            onChange={(_, newValue) => this.onChange(searchFilter, newValue)}
                                                            onClear={() => this.setSearchText('')} />
                                                    </div>
                                                </>
                                            }
                                            {/* Fees Type */}
                                            {
                                                this.isFilterEnabled('Fees') &&
                                                <>

                                                    <div >
                                                        <Label className={styles.filterLabel}>{feesTypeFilter.displayName}</Label>
                                                    </div>
                                                    <div className={styles.panelChildDiv}>
                                                        <MultiselectWrapper
                                                            placeholder={feesTypeFilter.placeHolder}
                                                            data={feesTypeFilter?.values as IItemOption[]}
                                                            dataKey={(item: IItemOption) => item.key}
                                                            textField={(item: IItemOption) => item.text}
                                                            value={feesTypeFilter?.selectedValues}
                                                            showSelectedItemsInList
                                                            showPlaceholderWithValues
                                                            renderTagValue={() => null}
                                                            filter="contains"


                                                            onChange={this.onChangeFeestype} />
                                                    </div>
                                                </>
                                            }
                                            {/* Item */}
                                            {
                                                this.isFilterEnabled('Item') &&
                                                <>
                                                    <div >
                                                        <Label className={styles.filterLabel}>{itemFilter.displayName}</Label>
                                                    </div>

                                                    <div className={styles.panelChildDiv}>
                                                        <MultiselectWrapper
                                                            placeholder={itemFilter.placeHolder}
                                                            data={itemFilter?.values as IDropdownOption[]}
                                                            dataKey={(item: IDropdownOption) => item.key}
                                                            textField={(item: IDropdownOption) => item.text}
                                                            value={itemFilter?.selectedValues}
                                                            showSelectedItemsInList
                                                            showPlaceholderWithValues
                                                            renderTagValue={() => null}
                                                            filter="contains"


                                                            onChange={this.onChangeItem} />
                                                    </div>
                                                </>
                                            }
                                            {/* Status */}
                                            {
                                                this.isFilterEnabled('Status') &&
                                                <>
                                                    <div>
                                                        <Label className={styles.filterLabel}>{statusTypeFilter.displayName}</Label>
                                                    </div>
                                                    <div className={styles.panelChildDiv}>
                                                        <MultiselectWrapper
                                                            placeholder={statusTypeFilter.placeHolder}
                                                            data={statusTypeFilter?.values as IDropdownOption[]}
                                                            dataKey={(item: IDropdownOption) => item.key}
                                                            textField={(item: IDropdownOption) => item.text}
                                                            value={statusTypeFilter?.selectedValues}
                                                            filter="contains"
                                                            showSelectedItemsInList
                                                            showPlaceholderWithValues
                                                            renderTagValue={() => null}


                                                            onChange={this.onChangeStatus} />
                                                    </div>
                                                </>
                                            }


                                        </div>
                                    </Panel>
                                    <div className={styles.tableContainer}>
                                        <AppliedFilters
                                            filters={filters}
                                            callback={this.refreshFilters}
                                        />
                                    </div>
                                </>
                            )}
                            <div className={styles.listViewContainer}>
                                {viewItems.length > 0 ? (<div className={styles.detailsList}>
                                    {this.props.pagination?.isEnabled && this.props.pagination.showTopPagination && (
                                        <Pager
                                            total={viewItems.length}
                                            pageIndex={pageIndex}
                                            pageSize={pageSize}
                                            onPageChange={this.setPage}
                                            onPageSizeChange={this.setPageSize}
                                        />)}

                                    <DetailsList
                                        items={itemsToDisplay}
                                        columns={columns}
                                        layoutMode={DetailsListLayoutMode.fixedColumns}
                                        constrainMode={ConstrainMode.horizontalConstrained}
                                        selectionMode={SelectionMode.none}
                                        onRenderItemColumn={this.renderItemColumn}
                                        onShouldVirtualize={() => false}
                                    />
                                    {this.props.pagination?.isEnabled && this.props.pagination.showBottomPagination && (
                                        <Pager
                                            total={viewItems.length}
                                            pageIndex={pageIndex}
                                            pageSize={pageSize}
                                            onPageChange={this.setPage}
                                            onPageSizeChange={this.setPageSize}
                                        />
                                    )}
                                </div>) : (
                                    <div className={styles.noDataMessage}>
                                        <Icon iconName="Remove" className={styles.mutedIcon} />
                                        <span>{strings.NoItemsMessage || 'No record found'}</span>
                                        <Icon iconName="Remove" className={styles.mutedIcon} />
                                    </div>
                                )
                                }
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
}

