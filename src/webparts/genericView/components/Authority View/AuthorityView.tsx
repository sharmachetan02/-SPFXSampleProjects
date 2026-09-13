import * as React from 'react';
import { IAuthorityViewState, IAuthorityViewProps } from './AuthorityView.type';
import styles from '../View.module.scss';
import { IViewConfig, View, ViewType } from '../View.types';
import 'react-widgets/styles.css';
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
import { IAuthority, IContentType } from '../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../common/helpers/Util';
import { SortDirection } from '@pnp/sp/search';
import { renderArrayPills, renderObjectPills } from '../View.utility';
import AppliedFilters from '../AppliedFilters/AppliedFilters';
import Pager from '../Pager/Pager';
import { CallOutButton } from '../../../../common/components/CallOutButton/CallOutButton';
import { RichText } from '../../../../common/components/RichText/RichText';
import { AuthorityType, UserRole } from '../../../../common/models/Enums';
import DomHelper from '../../../../common/helpers/DomHelper';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';

export default class AuthorityView extends React.Component<
    IAuthorityViewProps,
    IAuthorityViewState
> {
    private config: IViewConfig;
    private allViewItems: IAuthority[] = [];
    private typeChoices: IDropdownOption[] = [
        { key: Consts.CONTENT_TYPES.ADMINISTRATION, text: AuthorityType.Administration },
        { key: Consts.CONTENT_TYPES.ORGANIZATION, text: AuthorityType.Organization },
        { key: Consts.CONTENT_TYPES.REGULATOR, text: AuthorityType.Regulator }
    ];

    constructor(props: Readonly<IAuthorityViewProps>) {
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
        this.config = View.config.find((f) => f.type.toLowerCase() === ViewType.Authority);
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

    public async componentDidUpdate(prev: IAuthorityViewProps) {
        if (
            prev.columns !== this.props.columns ||
            prev.filters !== this.props.filters ||
            prev.sort !== this.props.sort ||
            prev.viewItems !== this.props.viewItems
        ) {
            await this.init(); // re-apply filtered columns/filters and items
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
            name: strings.AuthorityViewTitleColumnLabel,
            fieldName: 'Title',
            minWidth: 360,
            maxWidth: 360,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Countries',
            name: strings.AuthorityViewCountriesColumnLabel,
            fieldName: 'Countries',
            minWidth: 300,
            maxWidth: 300,
            isResizable: true,
        },
        {
            key: 'ContentType',
            name: strings.AuthorityViewContentTypeColumnLabel,
            fieldName: 'ContentType',
            minWidth: 120,
            maxWidth: 120,
            isResizable: true,
            onColumnClick: this.onColumnClick,
        },
        {
            key: 'Summary',
            name: strings.AuthorityViewSummaryColumnLabel,
            fieldName: 'Summary',
            minWidth: 120,
            maxWidth: 120,
            isResizable: true
        }
        ];
        return this.props.columns ?
            allColumns.filter((c) => this.isColumnEnabled(c.key)) :
            allColumns;
    }

    private initCountriesChoices = async () => {
        const {
            pnpService
        } = this.props;
        const countriesChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.COUNTRY_URL);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE
        ];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        items.forEach((item) => {
            countriesChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
        });
        return countriesChoices;
    };

    private initTypeChoices = async (): Promise<IDropdownOption[]> => this.typeChoices;

    public initViewItems = async () => {
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.AUTHORITY.SUMMARY,
            Consts.FIELDS.AUTHORITY.ADDRESS,
            Consts.FIELDS.AUTHORITY.GROUPING,
            Consts.FIELDS.AUTHORITY.COMMENTS,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
            Consts.FIELDS.AUTHORITY.PORTAL,
            `${Consts.FIELDS.AUTHORITY.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.AUTHORITY.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.AUTHORITY.ETU_OWNERS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.AUTHORITY.ETU_OWNERS}/${Consts.FIELDS.COMMON.TITLE}`
        ];
        const expand = [Consts.FIELDS.AUTHORITY.COUNTRY, Consts.FIELDS.AUTHORITY.ETU_OWNERS];
        const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.config.list);
        const items = await this.props.pnpService.getListItems(listUrl).
            select(...selectedFields)
            .expand(...expand)
            .orderBy(Consts.FIELDS.COMMON.TITLE)();
        const viewItems: IAuthority[] = [];

        items.map((item) => {
            const ContentType: IContentType = this.typeChoices
                .map((type) => ({ Id: type.key + "", Name: type.text }))
                .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1);

            viewItems.push({
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                ContentType,
                Summary: item[Consts.FIELDS.AUTHORITY.SUMMARY],
                Address: item[Consts.FIELDS.AUTHORITY.ADDRESS],
                Comments: item[Consts.FIELDS.AUTHORITY.COMMENTS],
                Countries: item[Consts.FIELDS.AUTHORITY.COUNTRY]?.map((country) => ({
                    Id: country[Consts.FIELDS.COMMON.ID],
                    Title: country[Consts.FIELDS.COMMON.TITLE]
                })) || [],
                Owners: item[Consts.FIELDS.AUTHORITY.ETU_OWNERS]?.map((Owner) => ({
                    Id: Owner[Consts.FIELDS.COMMON.ID],
                    Title: Owner[Consts.FIELDS.COMMON.TITLE]
                })) || [],
                Grouping: item[Consts.FIELDS.COUNTRY.GROUPING],
                Portal: item[Consts.FIELDS.AUTHORITY.PORTAL]?.Url?.trim() || '',
            });
        });
        return viewItems;

    };

    private initFilters = (): IFilter[] => {
        let allFilters: IFilter[] = [];

        const searchFilter: IFilter = {
            name: 'Search',
            displayName: strings.AuthorityViewSearchFilterLabel,
            placeHolder: strings.AuthortyViewSearchFilterPlaceHolder,
            fields: ['Title', 'Countries', 'ContentType', 'Summary'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.Or,
        };
        const countriesFilter: IFilter = {
            name: 'Countries',
            displayName: strings.AuthorityViewCountriesFilterLabel,
            placeHolder: strings.AuthorityViewCountriesFilterPlaceHolder,
            fields: ['Countries'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const ContentTypeFilter: IFilter = {
            name: 'ContentType',
            displayName: strings.AuthorityViewContentTypeFilterLabel,
            placeHolder: strings.AuthorityViewContentTypeFilterPlaceHolder,
            fields: ['ContentType'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };

        allFilters = [searchFilter, ContentTypeFilter, countriesFilter,];
        return this.props.filters ? allFilters.filter((f) => this.isFilterEnabled(f.name)) : allFilters;
    };

    private init = async () => {
        const columns = this.initColumns();
        let filters = this.initFilters();
        let items: IAuthority[];
        if (this.props.viewItems) {
            items = this.props.viewItems.slice();
        } else {
            items = await this.initViewItems();
        }

        let contentType: IDropdownOption[] = [];
        let country: IDropdownOption[] = [];
        const promises: Promise<void>[] = [];
        if (this.isFilterEnabled('Countries')) {
            promises.push(this.initCountriesChoices().then((opts) => { country = opts as IDropdownOption[]; }));
        }
        if (this.isFilterEnabled('ContentType')) {
            promises.push(this.initTypeChoices().then((opts) => { contentType = opts as IDropdownOption[]; }));
        }
        await Promise.all(promises);
        filters = filters.map((f) => {
            if (f.name === 'Countries') return { ...f, values: country };
            if (f.name === 'ContentType') return { ...f, values: contentType };
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

    public applyFilters = (): IAuthority[] => {
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
                            } else if (field === 'Countries') {
                                return (item[field] || []).some((c) => {
                                    const hay = UtilHelper.toNormalForm(c?.Title?.toString()?.toLowerCase());
                                    return hay?.includes(needle) === true;
                                });
                            } else {
                                const hay = UtilHelper.toNormalForm(item[field]?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            }
                        }
                        if (filter.name === 'Countries') {
                            return (item[field] || []).some((t) => String(t.Id) === String((f as IDropdownOption).key));
                        }
                        if (filter.name === 'ContentType') {
                            return item[field]?.Id === String((f as IDropdownOption).key);
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

        const key = sort.Property as keyof IAuthority;
        const desc = sort.Direction === SortDirection.Descending;

        return items.slice().sort((a, b) => {
            let av = a[key];
            let bv = b[key];

            if (key === 'ContentType') {
                av = av?.Name ?? '';
                bv = bv?.Name ?? '';
            }
            const aStr = (av || '').toLowerCase();
            const bStr = (bv || '').toLowerCase();

            if (aStr === bStr) return 0;
            const cmp = aStr > bStr ? 1 : -1;
            return desc ? -cmp : cmp;
        });
    };

    private renderItemColumn = (item, index: number, column: IColumn) => {
        const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        const itemDatasheetUrl = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(this.config.type)}&itemId=${item.Id}`;
        if (!column) return <span />;
        const fieldName = column.fieldName as keyof IAuthority | undefined;
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
            return (
                <span className={styles.detailsListValue}>{String(value?.Name)}</span>
            );
        }
        if (fieldName === 'Countries') {
            const countryListType = View.config.find((f) => f.type.toLowerCase() === ViewType.Country).type;
            return renderObjectPills(
              value,
              (item) => (
                <span className={styles.termPill}>
                  <span className={styles.termPillText}>
                    <a data-interception="off"
                      rel="noreferrer"
                      onClick={(e) => {
                        e.preventDefault();
                        const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(countryListType)}&itemId=${item?.Id}`;
                        UrlHelper.navigate(url, false);
                      }}
                    >{item?.Title}</a>
                  </span>
                </span>
              )
            );
        }
        if (fieldName === 'Summary') {
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

    private onChangetype = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const contentTypeFilter = filters.find((f) => f.name === 'ContentType');
        contentTypeFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === contentTypeFilter.name) ? contentTypeFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeCountries = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const countryFilter = filters.find((f) => f.name === 'Countries');
        countryFilter.selectedValues = options;
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


    public render(): React.ReactElement<IAuthorityViewProps> {
        const { filters, columns, searchKeyword, pageIndex, pageSize, isFilterPanelOpen } = this.state;
        const countryFilter = filters.find((f) => f.name === 'Countries');
        const contentTypeFilter = filters.find((f) => f.name === 'ContentType');
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
                            <div className={styles.applyFiltersBtn}>
                                {this.allViewItems && this.allViewItems.length > 0 &&
                                    <PrimaryButton
                                        text="Filters"
                                        iconProps={{ iconName: 'Filter' }}
                                        className={`${styles.clearButton} "form-button"`}
                                        styles={{
                                            icon: { order: 0, marginRight: 4 }, // moves icon after text
                                        }}
                                        onClick={() => this.openFilterPanel()}
                                    />
                                }
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
                                        this.isFilterEnabled('Search') && <>
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
                                            </div></>
                                    }
                                    {/* Content Type */}
                                    {
                                        this.isFilterEnabled('ContentType') && <>
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


                                                            onChange={this.onChangetype}


                                                />
                                            </div>
                                        </>
                                    }
                                    {/* Country */}
                                    {
                                        this.isFilterEnabled('Countries') &&
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


                                                            onChange={this.onChangeCountries} />
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
