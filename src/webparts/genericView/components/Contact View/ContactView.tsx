import * as React from 'react';
import { IContactViewProps, IContactViewState, IItemOption } from './ContactView.types';
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
import { IContentType, IContact, IItem } from '../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../common/helpers/Util';
import { SortDirection } from '@pnp/sp/search';
import { renderArrayPills } from '../View.utility';
import AppliedFilters from '../AppliedFilters/AppliedFilters';
import Pager from '../Pager/Pager';
import { CallOutButton } from '../../../../common/components/CallOutButton/CallOutButton';
import { RichText } from '../../../../common/components/RichText/RichText';
import { ContactType, UserRole } from '../../../../common/models/Enums';
import DomHelper from '../../../../common/helpers/DomHelper';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';

export default class ContactView extends React.Component<
    IContactViewProps,
    IContactViewState
> {
    private config: IViewConfig;
    private allViewItems: IContact[] = [];
    private typeChoices: IDropdownOption[] = [
        { key: Consts.CONTENT_TYPES.EUTELSAL_CONTACT, text: ContactType.Eutelsat },
        { key: Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT, text: ContactType.Administration },
        { key: Consts.CONTENT_TYPES.ORGANIZATION_CONTACT, text: ContactType.Organization },
        { key: Consts.CONTENT_TYPES.REGULATOR_CONTACT, text: ContactType.Regulator },
        { key: Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT, text: ContactType.TP },
        { key: Consts.CONTENT_TYPES.LAW_FIRM_CONTACT, text: ContactType.LawFirm },
        { key: Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE_CONTACT, text: ContactType.LegalRepresentative },
        { key: Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT, text: ContactType.DP }
    ];
    private itemChoices: IItemOption[] = [];
    constructor(props: Readonly<IContactViewProps>) {
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
        this.config = View.config.find((f) => f.type.toLowerCase() === ViewType.Contact);
    }
    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };
    public async componentDidMount() {
        if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner,
            UserRole.Contributor, UserRole.Visitor, UserRole.Admin)) {
            await this.init();
            this.setState({ isViewReady: true });
        }
    }

    public async componentDidUpdate(prev: IContactViewProps) {
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
        const allColumns = [
            {
                key: 'Title',
                name: strings.ContactViewTitleColumnLabel,
                fieldName: 'Title',
                minWidth: 300,
                maxWidth: 300,
                isResizable: true,
                onColumnClick: this.onColumnClick
            },
            {
                key: 'Summary',
                name: strings.ContactViewSummaryColumnLabel,
                fieldName: 'Summary',
                minWidth: 70,
                maxWidth: 70,
                isResizable: true,
            },
            {
                key: 'Item',
                name: strings.ContactViewItemColumnLabel,
                fieldName: 'Item',
                minWidth: 210,
                maxWidth: 210,
                isResizable: true,
                onColumnClick: this.onColumnClick
            },
            {
                key: 'JobTitle',
                name: strings.ContactViewJobTitleColumnLabel,
                fieldName: 'JobTitle',
                minWidth: 210,
                maxWidth: 210,
                isResizable: true,
                onColumnClick: this.onColumnClick
            },
            {
                key: 'ContentType',
                name: strings.ContactViewContentTypeColumnLabel,
                fieldName: 'ContentType',
                minWidth: 110,
                maxWidth: 110,
                isResizable: true,
                onColumnClick: this.onColumnClick
            },
            {
                key: 'Email',
                name: strings.ContactViewEmailColumnLabel,
                fieldName: 'Email',
                minWidth: 180,
                maxWidth: 180,
                isResizable: true,
                onColumnClick: this.onColumnClick
            },
            {
                key: 'PhoneNumber',
                name: strings.ContactViewPhoneNumberColumnLabel,
                fieldName: 'PhoneNumber',
                minWidth: 120,
                maxWidth: 120,
                isResizable: true,
            }

        ];
        return this.props.columns ?
            allColumns.filter((c) => this.isColumnEnabled(c.key)) :
            allColumns;
    }

    private initContentTypeChoices = async (): Promise<IDropdownOption[]> => {
        let contentTypeChoices: IDropdownOption[] = [];
        contentTypeChoices = this.typeChoices;
        contentTypeChoices.sort((a, b) => a.text?.localeCompare(b.text));
        return contentTypeChoices;
    };

    private initItemChoices = async (): Promise<IItemOption[]> => {
        const itemChoices: IItemOption[] = [];
        this.itemChoices.forEach((item) => {
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
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
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
        const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.config.list);
        const items = await this.props.pnpService.getListItems(listUrl).
            select(...selectedFields).
            expand(Consts.FIELDS.CONTACT.AUTHORITY,
                Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER,
                Consts.FIELDS.CONTACT.TELEPORT_PARTNER,
                Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER).
            orderBy(Consts.FIELDS.COMMON.TITLE)();
        const viewItems: IContact[] = [];
        items.map((item) => {
            const filterTypeChoices = this.typeChoices.find((type) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1)?.key;
            let itemValue = null;
            if (filterTypeChoices === Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT ||
                filterTypeChoices === Consts.CONTENT_TYPES.ORGANIZATION_CONTACT ||
                filterTypeChoices === Consts.CONTENT_TYPES.REGULATOR_CONTACT) {
                const authority = item[Consts.FIELDS.CONTACT.AUTHORITY];
                if (authority && authority[Consts.FIELDS.COMMON.ID] && authority[Consts.FIELDS.COMMON.TITLE]) {
                    itemValue = {
                        id: authority[Consts.FIELDS.COMMON.ID],
                        key: authority[Consts.FIELDS.COMMON.ID] + '-' + ViewType.Authority,
                        text: authority[Consts.FIELDS.COMMON.TITLE],
                        type: ViewType.Authority
                    };
                    this.itemChoices.push({
                        id: authority[Consts.FIELDS.COMMON.ID],
                        key: itemValue?.key,
                        text: itemValue?.text,
                        type: itemValue?.type
                    });
                }
            } else if (filterTypeChoices === Consts.CONTENT_TYPES.LAW_FIRM_CONTACT ||
                filterTypeChoices === Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE_CONTACT) {
                const legalSvcProvider = item[Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER];
                if (legalSvcProvider && legalSvcProvider[Consts.FIELDS.COMMON.ID] && legalSvcProvider[Consts.FIELDS.COMMON.TITLE]) {
                    itemValue = {
                        id: legalSvcProvider[Consts.FIELDS.COMMON.ID],
                        key: legalSvcProvider[Consts.FIELDS.COMMON.ID] + '-' + ViewType.LegalSvcProvider,
                        text: legalSvcProvider[Consts.FIELDS.COMMON.TITLE],
                        type: ViewType.LegalSvcProvider
                    };
                    this.itemChoices.push({
                        id: legalSvcProvider[Consts.FIELDS.COMMON.ID],
                        key: itemValue?.key,
                        text: itemValue?.text,
                        type: itemValue?.type

                    });
                }
            } else if (filterTypeChoices === Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT) {
                const distributionPartner = item[Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER];
                if (distributionPartner && distributionPartner[Consts.FIELDS.COMMON.ID] && distributionPartner[Consts.FIELDS.COMMON.TITLE]) {
                    itemValue = {
                        id: distributionPartner[Consts.FIELDS.COMMON.ID],
                        key: distributionPartner[Consts.FIELDS.COMMON.ID] + '-' + ViewType.DistributionPartner,
                        text: distributionPartner[Consts.FIELDS.COMMON.TITLE],
                        type: ViewType.DistributionPartner
                    };
                    this.itemChoices.push({
                        id: distributionPartner[Consts.FIELDS.COMMON.ID],
                        key: itemValue?.key,
                        text: itemValue?.text,
                        type: itemValue?.type

                    });
                }
            } else if (filterTypeChoices === Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT) {
                const teleportPartner = item[Consts.FIELDS.CONTACT.TELEPORT_PARTNER];
                if (teleportPartner && teleportPartner[Consts.FIELDS.COMMON.ID] && teleportPartner[Consts.FIELDS.COMMON.TITLE]) {
                    itemValue = {
                        id: teleportPartner[Consts.FIELDS.COMMON.ID],
                        key: teleportPartner[Consts.FIELDS.COMMON.ID] + '-' + ViewType.TeleportPartner,
                        text: teleportPartner[Consts.FIELDS.COMMON.TITLE],
                        type: ViewType.TeleportPartner
                    };
                    this.itemChoices.push({
                        id: teleportPartner[Consts.FIELDS.COMMON.ID],
                        key: itemValue?.key,
                        text: itemValue?.text,
                        type: itemValue?.type

                    });
                }
            }
            viewItems.push({
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                ContentType: {
                    Id: item[Consts.FIELDS.COMMON.CONTENTTYPE_ID],
                    Name: this.getTypeChoiceTextByKey(item[Consts.FIELDS.COMMON.CONTENTTYPE_ID])
                },
                Summary: item[Consts.FIELDS.CONTACT.SUMMARY],
                Comments: item[Consts.FIELDS.CONTACT.COMMENTS],
                Department: item[Consts.FIELDS.CONTACT.DEPARTMENT],
                JobTitle: item[Consts.FIELDS.CONTACT.JOB_TITLE] || '',
                Email: item[Consts.FIELDS.CONTACT.EMAIL] || '',
                PhoneNumber: item[Consts.FIELDS.CONTACT.PHONE_NUMBER],
                Item: itemValue
            });
        });
        return viewItems;
    };

    private initFilters = (): IFilter[] => {
        let allFilters: IFilter[] = [];
        const searchFilter: IFilter = {
            name: 'Search',
            displayName: strings.CountryViewSearchFilterLabel,
            placeHolder: strings.CountryViewSearchFilterPlaceHolder,
            fields: ['Title', 'ContentType', 'Item', 'JobTitle', 'Email', 'PhoneNumber', 'Summary', 'Email'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.Or,
        };
        const contentTypeFilter: IFilter = {
            name: 'ContentType',
            displayName: strings.ContactViewContentTypeFilterLabel,
            placeHolder: strings.ContactViewContentTypeFilterPlaceHolder,
            fields: ['ContentType'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const itemFilter: IFilter = {
            name: 'Item',
            displayName: strings.ContactViewItemFilterLabel,
            placeHolder: strings.ContactViewItemFilterPlaceHolder,
            fields: ['Item'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        allFilters = [searchFilter, contentTypeFilter, itemFilter];
        return this.props.filters ? allFilters.filter((f) => this.isFilterEnabled(f.name)) : allFilters;
    };

    private getTypeChoiceTextByKey = (key: string | number): string | undefined => {
        const keyStr = String(key);
        const found = this.typeChoices.find((type) => keyStr.indexOf(String(type.key)) !== -1);
        return found ? found.text : "";
    };

    private init = async () => {
        const columns = this.initColumns();
        let filters = this.initFilters();
        let items: IContact[];
        if (this.props.viewItems) {
            this.itemChoices = [];
            this.props.viewItems.forEach((itm) => {
                if (itm.Item) {
                    this.itemChoices.push({
                        id: itm.Item.id,
                        key: itm.Item.key,
                        text: itm.Item.text,
                        type: itm.Item.type
                    });
                }
            });
            items = this.props.viewItems.slice();
        } else {
            items = await this.initViewItems();
        }
        let contentTypes: IDropdownOption[] = [];
        let itemChoices: IItemOption[] = [];
        const promises: Promise<void>[] = [];
        if (this.isFilterEnabled('ContentType')) {
            promises.push(this.initContentTypeChoices().then((opts) => { contentTypes = opts as IDropdownOption[]; }));
        }
        if (this.isFilterEnabled('Item')) {
            promises.push(this.initItemChoices().then((opts) => { itemChoices = opts as IItemOption[]; }));
        }
        await Promise.all(promises);
        filters = filters.map((f) => {
            if (f.name === 'ContentType') return { ...f, values: contentTypes };
            if (f.name === 'Item') return { ...f, values: itemChoices };
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

    public applyFilters = (): IContact[] => {
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
                            if (field === 'ContentType') {
                                const hay = UtilHelper.toNormalForm((item[field] as IContentType)?.Name?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            } else if (field === 'Item') {
                                const itemValue = item[field] as IItem;
                                if (!itemValue) return false;
                                const hay = UtilHelper.toNormalForm(itemValue?.text?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            } else {
                                if (!item[field]) return false;
                                const hay = UtilHelper.toNormalForm(item[field]?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            }
                        }

                        if (filter.name === 'ContentType') {
                            return String((item[field] as IContentType).Id).indexOf(String((f as IDropdownOption).key)) !== -1;
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

    private sort = (items: IContact[]): IContact[] => {
        const { sort } = this.state;
        if (!sort) return items;
        const key = sort.Property as keyof IContact;
        const desc = sort.Direction === SortDirection.Descending;
        return items.slice().sort((a: IContact, b: IContact) => {
            let av = a[key];
            let bv = b[key];
            if (key === 'ContentType') {
                av = (av as IContentType)?.Name ?? '';
                bv = (bv as IContentType)?.Name ?? '';
            }
            if (key === 'Item') {
                av = (av as IItem)?.text ?? '';
                bv = (bv as IItem)?.text ?? '';
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
        const fieldName = column.fieldName as keyof IContact | undefined;
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
            const ct = value as IDropdownOption;
            return <span className={styles.detailsListValue}>{String(ct?.['Name'])}</span>;
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
                            title={String(itemValue?.text)}
                            onClick={(e) => {
                                e.preventDefault();
                                const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${viewType}&itemId=${itemValue.id}`;
                                UrlHelper.navigate(url, false);
                            }}
                        >{String(itemValue?.text)}</a>
                    </span>
                </span>
            </div>
            );
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

    private onChangeContentType = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const contentTypeFilter = filters.find((f) => f.name === 'ContentType');
        contentTypeFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === contentTypeFilter.name) ? contentTypeFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeItem = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const itemFilter = filters.find((f) => f.name === 'Item');
        itemFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === itemFilter.name) ? itemFilter : f)));
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

    public render(): React.ReactElement<IContactViewProps> {
        const { filters, columns, searchKeyword, pageIndex, pageSize, isFilterPanelOpen } = this.state;
        const contentTypeFilter = filters.find((f) => f.name === 'ContentType');
        const itemFilter = filters.find((f) => f.name === 'Item');
        const searchFilter = filters.find((f) => f.name === 'Search');
        let viewItems = this.applyFilters();
        viewItems = this.sort(viewItems);
        const pageCount = Math.max(1, Math.ceil(viewItems.length / pageSize));
        const safeIndex = Math.min(pageIndex, pageCount - 1);
        const start = safeIndex * pageSize;
        const end = start + pageSize;
        const pagedItems = viewItems.slice(start, end);
        const itemsToDisplay = this.props.pagination?.isEnabled ? pagedItems : viewItems;

        if (!this.hasAnyRole(UserRole.Contributor, UserRole.Owner, UserRole.Visitor,
            UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Admin)) {
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

                                            {/* Content Type */}
                                            {
                                                this.isFilterEnabled('ContentType') &&
                                                <>
                                                    <div >
                                                        <Label className={styles.filterLabel}>{contentTypeFilter.displayName}</Label>
                                                    </div>
                                                    <div className={styles.panelChildDiv}>
                                                        <MultiselectWrapper
                                                            placeholder={contentTypeFilter.placeHolder}
                                                            data={contentTypeFilter?.values as IDropdownOption[]}
                                                            dataKey={(item: IDropdownOption) => item.key}
                                                            textField={(item: IDropdownOption) => item.text}
                                                            value={contentTypeFilter?.selectedValues}
                                                            showSelectedItemsInList
                                                            showPlaceholderWithValues
                                                            renderTagValue={() => null}
                                                            filter="contains"
                                                            onChange={this.onChangeContentType} />
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
                                                            dataKey={(item: IItemOption) => `${item.key}_${item.type}`}
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
