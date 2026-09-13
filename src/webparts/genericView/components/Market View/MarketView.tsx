import * as React from 'react';
import { IMarketViewState, IMarketViewProps, IGroupingOption, ISanctionOption } from './MarketView.type';
import styles from '../View.module.scss';
import { IViewConfig, View, ViewType } from '../View.types';
import 'react-widgets/styles.css';
import { TaxonomyHelper } from '../../../../common/helpers/TaxonomyHelper';
import { Config } from '../../../../common/config/Config';
import {
    ConstrainMode, DetailsList, DetailsListLayoutMode, IColumn, Icon, IDropdownOption, ITooltipHostStyles, Label,
    LayerHost, Panel, PanelType, PrimaryButton, SearchBox, SelectionMode, Spinner, TooltipHost
} from '@fluentui/react';
import { IFilter, LogicalOperator } from '../../../../common/models/IFilter';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import * as strings from 'GenericViewPartStrings';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { Consts } from '../../../../common/consts/Consts';
import { RenderItemProp } from 'react-widgets/cjs/List';
import { IMarket, ISanctionCategory, Iterm } from '../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../common/helpers/Util';
import { SortDirection } from '@pnp/sp/search';
import { renderArrayPills, renderObjectPills } from '../View.utility';
import AppliedFilters from '../AppliedFilters/AppliedFilters';
import Pager from '../Pager/Pager';
import { UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
export default class MarketView extends React.Component<
    IMarketViewProps,
    IMarketViewState
> {
    private config: IViewConfig;
    private allViewItems: IMarket[] = [];
    constructor(props: Readonly<IMarketViewProps>) {
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
        this.config = View.config.find((f) => f.type.toLowerCase() === ViewType.Market);
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

    public async componentDidUpdate(prev: IMarketViewProps) {
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
            name: strings.MarketViewTitleColumnLabel,
            fieldName: 'Title',
            minWidth: 360,
            maxWidth: 360,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Country',
            name: strings.MarketViewCountryColumnLabel,
            fieldName: 'Country',
            minWidth: 150,
            maxWidth: 150,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Grouping',
            name: strings.MarketViewGroupingColumnLabel,
            fieldName: 'Grouping',
            minWidth: 260,
            maxWidth: 260,
            isResizable: true,
        },
        {
            key: 'SanctionCategory',
            name: strings.MarketViewSanctionCategoryColumnLabel,
            fieldName: 'SanctionCategory',
            minWidth: 150,
            maxWidth: 150,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Priority',
            name: strings.MarketViewPriorityColumnLabel,
            fieldName: 'MaPriority',
            minWidth: 260,
            maxWidth: 260,
            isResizable: true,
        }
        ];
        return this.props.columns ?
            allColumns.filter((c) => this.isColumnEnabled(c.key)) :
            allColumns;
    }

    private initCountryChoices = (items: IMarket[]): IDropdownOption[] => {
        const countryChoices: IDropdownOption[] = [];
        const countryMap = new Map<number, string>();
        items.forEach((item) => {
            if (item.Country && item.Country.Id && item.Country.Title && !countryMap.has(item.Country.Id)) {
                countryMap.set(item.Country.Id, item.Country.Title);
            }
        });
        countryMap.forEach((title, id) => {
            countryChoices.push({
                key: id,
                text: title
            });
        });
        return countryChoices.sort((a, b) => (a.text as string)?.localeCompare(b.text as string));
    };

    private initSanctionCategoriesChoices = async () => {
        const {
            pnpService
        } = this.props;
        const sanctionCategoriesChoices: ISanctionOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.SANCTION_CATEGORY_URL);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.SANCTION_CATEGORY.COLOR,
            Consts.FIELDS.SANCTION_CATEGORY.SUMMARY
        ];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        items.forEach((item) => {
            const sanctionCategory: ISanctionOption = {
                key: item[Consts.FIELDS.COMMON.ID],
                text: item[Consts.FIELDS.COMMON.TITLE],
                color: item[Consts.FIELDS.SANCTION_CATEGORY.COLOR],
                summary: item[Consts.FIELDS.SANCTION_CATEGORY.SUMMARY]
            };
            sanctionCategoriesChoices.push(sanctionCategory);
        });
        return sanctionCategoriesChoices;
    };

    public initMaPrioritiesChoices = async (): Promise<IDropdownOption[]> => {
        const {
            pnpService
        } = this.props;
        let maPrioritiesChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MARKET.MARKET_ACCESS_PRIORITY)
            .select('Choices')();
        maPrioritiesChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
        return maPrioritiesChoices;
    };

    public initGroupingsChoices = async (): Promise<IGroupingOption[]> => {
        const {
            pnpService
        } = this.props;
        const terms = await TaxonomyHelper.getTermSetTerms(pnpService, Config.TAXONOMY.GROUPING_TERMSETID);
        const groupingsChoices: IGroupingOption[] = [];
        terms.map((item) => {
            item.children.map((child) => {
                groupingsChoices.push({
                    key: child.id,
                    text: child.labels[0].name,
                    type: item.labels[0].name
                });
            });
        });
        return groupingsChoices;
    };

    public initViewItems = async () => {
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.MARKET.COUNTRY,
            Consts.FIELDS.MARKET.GROUPING,
            Consts.FIELDS.MARKET.MARKET_ACCESS_PRIORITY,
            Consts.FIELDS.MARKET.IS_SANCTIONED,
            Consts.FIELDS.MARKET.COMMENTS,
            `${Consts.FIELDS.MARKET.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MARKET.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MARKET.SANCTION_CATEGORY}/${Consts.FIELDS.COMMON.ID}`,
        ];
        const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.config.list);
        const expand = [Consts.FIELDS.MARKET.COUNTRY, Consts.FIELDS.MARKET.SANCTION_CATEGORY];
        const items = await this.props.pnpService.getListItems(listUrl).
            select(...selectedFields).
            expand(...expand).
            orderBy(Consts.FIELDS.COMMON.TITLE)();
        const viewItems: IMarket[] = [];
        items.map((item) => {
            viewItems.push({
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                Comments: item[Consts.FIELDS.MARKET.COMMENTS],
                Grouping: item[Consts.FIELDS.MARKET.GROUPING],
                Country: {
                    Id: item[Consts.FIELDS.MARKET.COUNTRY]?.[Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.MARKET.COUNTRY]?.[Consts.FIELDS.COMMON.TITLE]
                },
                SanctionCategory: {
                    Id: item[Consts.FIELDS.MARKET.SANCTION_CATEGORY][Consts.FIELDS.COMMON.ID]
                },
                MaPriority: item[Consts.FIELDS.MARKET.MARKET_ACCESS_PRIORITY],
                IsSanctioned: item[Consts.FIELDS.MARKET.IS_SANCTIONED]
            });
        });
        return viewItems;
    };

    private initFilters = (): IFilter[] => {
        let allFilters: IFilter[] = [];
        const searchFilter: IFilter = {
            name: 'Search',
            displayName: strings.MarketViewSearchFilterLabel,
            placeHolder: strings.MarketViewSearchFilterPlaceHolder,
            fields: ['Title', 'Country', 'SanctionCategory', 'MaPriority', 'Grouping'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.Or,
        };
        const countryFilter: IFilter = {
            name: 'Country',
            displayName: strings.MarketViewCountryFilterLabel,
            placeHolder: strings.MarketViewCountryFilterPlaceHolder,
            fields: ['Country'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const groupingFilter: IFilter = {
            name: 'Grouping',
            displayName: strings.MarketViewGroupingFilterLabel,
            placeHolder: strings.MarketViewGroupingFilterPlaceHolder,
            fields: ['Grouping'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const priorityFilter: IFilter = {
            name: 'Priority',
            displayName: strings.MarketViewPriorityFilterLabel,
            placeHolder: strings.MarketViewPriorityFilterPlaceHolder,
            fields: ['MaPriority'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const SanctionCategory: IFilter = {
            name: 'SanctionCategory',
            displayName: strings.MarketViewSanctionCategoryFilterLabel,
            placeHolder: strings.MarketViewSanctionCategoryFilterPlaceHolder,
            fields: ['SanctionCategory'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        allFilters = [searchFilter, countryFilter, groupingFilter, priorityFilter, SanctionCategory];
        return this.props.filters ? allFilters.filter((f) => this.isFilterEnabled(f.name)) : allFilters;
    };

    private init = async () => {
        const columns = this.initColumns();
        let filters = this.initFilters();
        let items: IMarket[];
        if (this.props.viewItems) {
            items = this.props.viewItems.slice();
        } else {
            items = await this.initViewItems();
        }
        let priorities: IDropdownOption[] = [];
        let sanctions: ISanctionOption[] = [];
        let groupings: IDropdownOption[] = [];
        let countries: IDropdownOption[] = [];
        const promises: Promise<void>[] = [];
        if (this.isFilterEnabled('Priority')) {
            promises.push(this.initMaPrioritiesChoices().then((opts) => { priorities = opts as IDropdownOption[]; }));
        }
        if (this.isFilterEnabled('Grouping')) {
            promises.push(this.initGroupingsChoices().then((opts) => { groupings = opts as IDropdownOption[]; }));
        }
        promises.push(this.initSanctionCategoriesChoices().then((opts) => { sanctions = opts as ISanctionOption[]; }));
        await Promise.all(promises);
        countries = this.initCountryChoices(items);
        items.map((item) => {
            const sanctionCategory = sanctions.find((cat) => cat.key.toString() === item.SanctionCategory.Id.toString());
            item.SanctionCategory = {
                Id: item.SanctionCategory.Id,
                Title: sanctionCategory.text,
                Color: sanctionCategory.color,
                Summary: sanctionCategory.summary
            };
        });
        filters = filters.map((f) => {
            if (f.name === 'Priority') return { ...f, values: priorities };
            if (f.name === 'Country') return { ...f, values: countries };
            if (f.name === 'SanctionCategory') return { ...f, values: sanctions };
            if (f.name === 'Grouping') return { ...f, values: groupings };
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

    public applyFilters = (): IMarket[] => {
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
                            if (field === 'Country') {
                                const hay = UtilHelper.toNormalForm((item[field] as any)?.Title?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            } else if (field === 'SanctionCategory') {
                                const hay = UtilHelper.toNormalForm(item[field]?.Title?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            } else if (field === 'Grouping') {
                                return (item[field] || []).some((t) => {
                                    const hay = UtilHelper.toNormalForm(t?.Label?.toString()?.toLowerCase());
                                    return hay?.includes(needle) === true;
                                });
                            } else if (field === 'MaPriority') {
                                return (item[field] || []).some((t) => {
                                    const hay = UtilHelper.toNormalForm(t?.toString()?.toLowerCase());
                                    return hay?.includes(needle) === true;
                                });
                            } else {
                                const hay = UtilHelper.toNormalForm(item[field]?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            }
                        }
                        if (filter.name === 'Priority') {
                            return (item[field] || []).some((t) => String(t) === String((f as IDropdownOption).key));
                        }
                        if (filter.name === 'Country') {
                            return String((item[field] as any)?.Id) === String((f as IDropdownOption).key);
                        }
                        if (filter.name === 'SanctionCategory') {
                            return String((item[field] as ISanctionCategory).Id) === String((f as IDropdownOption).key);
                        }
                        if (filter.name === 'Grouping') {
                            return (item[field] || []).some((t) => t.Label === (f as IDropdownOption).text);
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
        const key = sort.Property as keyof IMarket;
        const desc = sort.Direction === SortDirection.Descending;
        return items.slice().sort((a, b) => {
            let av = a[key];
            let bv = b[key];
            if (key === 'Country') {
                av = av?.Title ?? '';
                bv = bv?.Title ?? '';
            } else if (key === 'SanctionCategory') {
                av = av?.Title ?? '';
                bv = bv?.Title ?? '';
            }
            const aStr = ((av || '') as string).toLowerCase();
            const bStr = ((bv || '') as string).toLowerCase();
            if (aStr === bStr) return 0;
            const cmp = aStr > bStr ? 1 : -1;
            return desc ? -cmp : cmp;
        });
    };

    private swatchStyle = (color?: string) => ({
        width: 12,
        height: 12,
        borderRadius: 3,
        background: color || '#888',
        display: 'inline-block',
        verticalAlign: 'middle' as const,
        marginRight: 6,
        border: '1px solid rgba(0,0,0,.1)',
    });

    private renderItemColumn = (item, index: number, column: IColumn) => {
        const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        const itemDatasheetUrl = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(this.config.type)}&itemId=${item.Id}`;
        if (!column) return <span />;
        const fieldName = column.fieldName as keyof IMarket | undefined;
        const value = fieldName ? item[fieldName] : undefined;
        if (value === null || value === undefined || (Array.isArray(value) && value.length === 0)) {
            return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
        }
        if (fieldName === 'Country') {
            const country = value as any;
            return (
                <span className={styles.detailsListValue}>{country?.Title}</span>
            );
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
        if (fieldName === 'Grouping') {
            return renderObjectPills(
                value as Iterm[],
                (term) => (
                    <span key={term.TermGuid} className={styles.termPill}>
                        <span className={styles.termPillText}>{term.Label}</span>
                    </span>
                )
            );
        }
        if (fieldName === 'SanctionCategory') {
            const sc = value as ISanctionCategory;
            return (
                <TooltipHost content={<div className={styles.tooltipHost} dangerouslySetInnerHTML={{ __html: sc.Summary }}></div>}
                    styles={tooltipStyles} calloutProps={{
                        layerProps: { hostId: 'sanctionCategoryLayer' },
                        styles: { root: { zIndex: 500002 } }
                    }}>
                    <div className={styles.tooltipHostContainer} title="">
                        <span style={this.swatchStyle(sc.Color)} />
                        <span className={styles.tooltipHostText}> {sc.Title}</span>
                    </div>
                </TooltipHost>
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

    private onChangeCountry = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const countryFilter = filters.find((f) => f.name === 'Country');
        countryFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === countryFilter.name) ? countryFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeGrouping = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const groupingFilter = filters.find((f) => f.name === 'Grouping');
        groupingFilter.selectedValues = options;
        filters = filters.map(((f) => ((f.name === groupingFilter.name) ? groupingFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeMaPriorities = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const priorityFilter = filters.find((f) => f.name === 'Priority');
        priorityFilter.selectedValues = options;
        filters = filters.map(((f) => ((f.name === priorityFilter.name) ? priorityFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeSanctionCategory = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const sanctionCategoryFilter = filters.find((f) => f.name === 'SanctionCategory');
        sanctionCategoryFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === sanctionCategoryFilter.name) ? sanctionCategoryFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private renderListItem: RenderItemProp<ISanctionOption> = ({ item }) => {
        if (!item) return null;
        const color = (item as ISanctionOption).color ?? '#888';
        const swatch = (c: string) => ({
            width: 12,
            height: 12,
            borderRadius: 3,
            background: c,
            display: 'inline-block',
            verticalAlign: 'middle' as const,
            marginRight: 5,
            border: '1px solid rgba(0,0,0,.1)',
        });
        return (

            <div style={{ verticalAlign: 'middle' }}>
                <span style={swatch(color)} />
                <span style={{ verticalAlign: 'middle' }}>{item.text}</span>
            </div>
        );
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

    public render(): React.ReactElement<IMarketViewProps> {
        const { filters, columns, searchKeyword, pageIndex, pageSize, isFilterPanelOpen } = this.state;
        const groupingFilter = filters.find((f) => f.name === 'Grouping');
        const countryFilter = filters.find((f) => f.name === 'Country');
        const priorityFilter = filters.find((f) => f.name === 'Priority');
        const sanctionCategoryFilter = filters.find((f) => f.name === 'SanctionCategory');
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

                                                    <div >
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
                                            {/* Country */}
                                            {
                                                this.isFilterEnabled('Country') &&
                                                <>
                                                    <div>
                                                        <Label className={styles.filterLabel}>{countryFilter.displayName}</Label>
                                                    </div>
                                                    <div className={styles.panelChildDiv}>
                                                        <MultiselectWrapper
                                                            placeholder={countryFilter.placeHolder}
                                                            data={countryFilter?.values as IDropdownOption[]}
                                                            dataKey={(item: IDropdownOption) => item.key}
                                                            textField={(item: IDropdownOption) => item.text}
                                                            value={countryFilter?.selectedValues}
                                                            filter="contains"
                                                            showSelectedItemsInList
                                                            showPlaceholderWithValues
                                                            renderTagValue={() => null}
                                                            onChange={this.onChangeCountry} />
                                                    </div>
                                                </>
                                            }
                                            {/* Grouping */}
                                            {
                                                this.isFilterEnabled('Grouping') &&
                                                <>
                                                    <div >
                                                        <Label className={styles.filterLabel}>{groupingFilter.displayName}</Label>
                                                    </div>
                                                    <div className={styles.panelChildDiv}>
                                                        <MultiselectWrapper
                                                            placeholder={groupingFilter.placeHolder}
                                                            data={groupingFilter?.values as IDropdownOption[]}
                                                            dataKey={(item: IDropdownOption) => item.key}
                                                            textField={(item: IDropdownOption) => item.text}
                                                            value={groupingFilter?.selectedValues}
                                                            filter="contains"
                                                            showSelectedItemsInList
                                                            showPlaceholderWithValues
                                                            renderTagValue={() => null}
                                                            groupBy={(type) => (type as IGroupingOption).type}
                                                            onChange={this.onChangeGrouping} />
                                                    </div>
                                                </>
                                            }
                                            {/* Sanction Category */}
                                            {
                                                this.isFilterEnabled('SanctionCategory') &&
                                                <>
                                                    <div >
                                                        <Label className={styles.filterLabel}>{sanctionCategoryFilter.displayName}</Label>
                                                    </div>
                                                    <div className={styles.panelChildDiv}>
                                                        <MultiselectWrapper
                                                            placeholder={sanctionCategoryFilter.placeHolder}
                                                            data={sanctionCategoryFilter?.values as IDropdownOption[]}
                                                            dataKey={(item: IDropdownOption) => item.key}
                                                            textField={(item: IDropdownOption) => item.text}
                                                            value={sanctionCategoryFilter?.selectedValues}
                                                            filter="contains"
                                                            renderListItem={this.renderListItem}
                                                            showSelectedItemsInList
                                                            showPlaceholderWithValues
                                                            renderTagValue={() => null}
                                                            onChange={this.onChangeSanctionCategory} />
                                                    </div>
                                                </>
                                            }
                                            {/* Priority */}
                                            {
                                                this.isFilterEnabled('Priority') &&
                                                <>

                                                    <div >
                                                        <Label className={styles.filterLabel}>{priorityFilter.displayName}</Label>
                                                    </div>
                                                    <div className={styles.panelChildDiv}>
                                                        <MultiselectWrapper
                                                            placeholder={priorityFilter.placeHolder}
                                                            data={priorityFilter?.values as IDropdownOption[]}
                                                            dataKey={(item: IDropdownOption) => item.key}
                                                            textField={(item: IDropdownOption) => item.text}
                                                            value={priorityFilter?.selectedValues}
                                                            showSelectedItemsInList
                                                            showPlaceholderWithValues
                                                            renderTagValue={() => null}
                                                            filter="contains"


                                                            onChange={this.onChangeMaPriorities} />
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
                                    <LayerHost
                                        id='sanctionCategoryLayer'
                                        style={{ position: 'relative', zIndex: 500000 }} // higher than page chrome
                                    />
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

