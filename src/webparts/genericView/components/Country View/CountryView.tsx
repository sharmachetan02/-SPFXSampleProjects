import * as React from 'react';
import { ICountryViewState, ICountryViewProps, IGroupingOption } from './CountryView.type';
import styles from '../View.module.scss';
import { IViewConfig, View, ViewType } from '../View.types';
import 'react-widgets/styles.css';
import { TaxonomyHelper } from '../../../../common/helpers/TaxonomyHelper';
import { Config } from '../../../../common/config/Config';
import {
    ConstrainMode, DetailsList, DetailsListLayoutMode, IColumn, Icon, IDropdownOption, ITooltipHostStyles, Label,
    Panel, PanelType, PrimaryButton, SearchBox, SelectionMode, Spinner, TooltipHost
} from '@fluentui/react';
import { IFilter, LogicalOperator } from '../../../../common/models/IFilter';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import * as strings from 'GenericViewPartStrings';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { Consts } from '../../../../common/consts/Consts';
import { ICountry, Iterm } from '../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../common/helpers/Util';
import { SortDirection } from '@pnp/sp/search';
import { renderArrayPills, renderObjectPills } from '../View.utility';
import AppliedFilters from '../AppliedFilters/AppliedFilters';
import Pager from '../Pager/Pager';
import { UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';

export default class CountryView extends React.Component<
    ICountryViewProps,
    ICountryViewState
> {
    private config: IViewConfig;
    private allViewItems: ICountry[] = [];
    constructor(props: Readonly<ICountryViewProps>) {
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
        this.config = View.config.find((f) => f.type.toLowerCase() === ViewType.Country);
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

    public async componentDidUpdate(prev: ICountryViewProps) {
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
            name: strings.CountryViewTitleColumnLabel,
            fieldName: 'Title',
            minWidth: 360,
            maxWidth: 360,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Identifier',
            name: strings.CountryViewIdentifierColumnLabel,
            fieldName: 'Identifier',
            minWidth: 150,
            maxWidth: 150,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Grouping',
            name: strings.CountryViewGroupingColumnLabel,
            fieldName: 'Grouping',
            minWidth: 260,
            maxWidth: 260,
            isResizable: true,
        }
        ];
        return this.props.columns ?
            allColumns.filter((c) => this.isColumnEnabled(c.key)) :
            allColumns;
    }

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
            Consts.FIELDS.COUNTRY.IDENTIFIER,
            Consts.FIELDS.COUNTRY.GROUPING,
        ];
        const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.config.list);
        const items = await this.props.pnpService.getListItems(listUrl).
            select(...selectedFields).
            orderBy(Consts.FIELDS.COMMON.TITLE)();
        const viewItems: ICountry[] = [];
        items.map((item) => {
            viewItems.push({
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                Identifier: item[Consts.FIELDS.COUNTRY.IDENTIFIER],
                Grouping: item[Consts.FIELDS.COUNTRY.GROUPING]
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
            fields: ['Title', 'Identifier', 'Grouping'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.Or,
        };
        const groupingFilter: IFilter = {
            name: 'Grouping',
            displayName: strings.CountryViewGroupingFilterLabel,
            placeHolder: strings.CountryViewGroupingFilterPlaceHolder,
            fields: ['Grouping'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const identifierFilter: IFilter = {
            name: 'Identifier',
            displayName: strings.CountryViewIdentifierFilterLabel,
            placeHolder: strings.CountryViewIdentifierFilterPlaceHolder,
            fields: ['Identifier'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        allFilters = [searchFilter, identifierFilter, groupingFilter];
        return this.props.filters ? allFilters.filter((f) => this.isFilterEnabled(f.name)) : allFilters;
    };

    private init = async () => {
        const columns = this.initColumns();
        let filters = this.initFilters();
        let items: ICountry[];
        if (this.props.viewItems) {
            items = this.props.viewItems.slice();
        } else {
            items = await this.initViewItems();
        }
        let identifiers: IDropdownOption[] = [];
        let groupings: IDropdownOption[] = [];
        const promises: Promise<void>[] = [];
        if (this.isFilterEnabled('Identifier')) {
            identifiers = items.map((country) => ({ key: country.Identifier, text: country.Identifier }));
        }
        if (this.isFilterEnabled('Grouping')) {
            promises.push(this.initGroupingsChoices().then((opts) => { groupings = opts as IDropdownOption[]; }));
        }
        await Promise.all(promises);
        filters = filters.map((f) => {
            if (f.name === 'Identifier') return { ...f, values: identifiers };
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

    public applyFilters = (): ICountry[] => {
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
                            else if (field === 'Grouping') {
                                return (item[field] || []).some((t) => {
                                    const hay = UtilHelper.toNormalForm(t?.Label?.toString()?.toLowerCase());
                                    return hay?.includes(needle) === true;
                                });
                            } else {
                                const hay = UtilHelper.toNormalForm(item[field]?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            }
                        }
                        if (filter.name === 'Identifier') {
                            return String((item[field])).indexOf(String((f as IDropdownOption).key)) !== -1;
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
        const key = sort.Property as keyof ICountry;
        const desc = sort.Direction === SortDirection.Descending;
        return items.slice().sort((a, b) => {
            const av = a[key];
            const bv = b[key];
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
        const fieldName = column.fieldName as keyof ICountry | undefined;
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
        if (fieldName === 'Grouping') {
            return renderObjectPills(
                value as Iterm[],
                (item) => (
                    <span className={styles.termPill}>
                        <span className={styles.termPillText}>{item.Label}</span>
                    </span>
                )
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

    private onChangeGrouping = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const groupingFilter = filters.find((f) => f.name === 'Grouping');
        groupingFilter.selectedValues = options;
        filters = filters.map(((f) => ((f.name === groupingFilter.name) ? groupingFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeIdentifier = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const identifierFilter = filters.find((f) => f.name === 'Identifier');
        identifierFilter.selectedValues = options;
        filters = filters.map(((f) => ((f.name === identifierFilter.name) ? identifierFilter : f)));
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

    public render(): React.ReactElement<ICountryViewProps> {
        const { filters, columns, searchKeyword, pageIndex, pageSize, isFilterPanelOpen } = this.state;
        const groupingFilter = filters.find((f) => f.name === 'Grouping');
        const identifierFilter = filters.find((f) => f.name === 'Identifier');
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

                                            {/* Priority */}
                                            {
                                                this.isFilterEnabled('Priority') &&
                                                <>

                                                    <div >
                                                        <Label className={styles.filterLabel}>{identifierFilter.displayName}</Label>
                                                    </div>
                                                    <div className={styles.panelChildDiv}>
                                                        <MultiselectWrapper
                                                            placeholder={identifierFilter.placeHolder}
                                                            data={identifierFilter?.values as IDropdownOption[]}
                                                            dataKey={(item: IDropdownOption) => item.key}
                                                            textField={(item: IDropdownOption) => item.text}
                                                            value={identifierFilter?.selectedValues}
                                                            showSelectedItemsInList
                                                            showPlaceholderWithValues
                                                            renderTagValue={() => null}
                                                            filter="contains"


                                                            onChange={this.onChangeIdentifier} />
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
