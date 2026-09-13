import * as React from 'react';
import { IMARequirementViewState, IMARequirementViewProps, IItemOption } from './MARequirementView.type';
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
import { IContentType, IMARequirement } from '../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../common/helpers/Util';
import { SortDirection } from '@pnp/sp/search';
import { renderArrayPills, renderObjectPills, StatusPill } from '../View.utility';
import AppliedFilters from '../AppliedFilters/AppliedFilters';
import Pager from '../Pager/Pager';
import { CallOutButton } from '../../../../common/components/CallOutButton/CallOutButton';
import { RichText } from '../../../../common/components/RichText/RichText';
import { DatasheetType } from '../../../genericDatasheet/components/Datasheet.types';
import { ContactType, MARequirementType, UserRole } from '../../../../common/models/Enums';
import DomHelper from '../../../../common/helpers/DomHelper';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';

export default class MARequirementView extends React.Component<
    IMARequirementViewProps,
    IMARequirementViewState
> {
    private config: IViewConfig;
    private allViewItems: IMARequirement[] = [];
    private responsibleItemChoices: IItemOption[] = [];
    private requestedItemChoices: IDropdownOption[] = [];
    private typeChoices: IDropdownOption[] = [
        { key: Consts.CONTENT_TYPES.MA_RequirementCredential, text: MARequirementType.Credential },
        { key: Consts.CONTENT_TYPES.MA_RequirementGeneric, text: MARequirementType.Generic },
        { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_NOL, text: MARequirementType.NOL },
        { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_LEGAL, text: MARequirementType.Legal },
    ];

    constructor(props: Readonly<IMARequirementViewProps>) {
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
            this.setState({ isViewReady: true });
        }
    }

    public async componentDidUpdate(prev: IMARequirementViewProps) {
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
            name: strings.MARequirementViewTitleColumnLabel,
            fieldName: 'Title',
            minWidth: 180,
            maxWidth: 180,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Markets',
            name: strings.MARequirmentViewMarketsColumnLabel,
            fieldName: 'Markets',
            minWidth: 105,
            maxWidth: 105,
            isResizable: true,
        },
        {
            key: 'Summary',
            name: strings.MArequimentViewSummaryColumnLabel,
            fieldName: 'Summary',
            minWidth: 60,
            maxWidth: 60,
            isResizable: true,
        },
        {
            key: 'MaVertical',
            name: strings.MARequimentViewMaVerticalColumnLabel,
            fieldName: 'MaVertical',
            minWidth: 90,
            maxWidth: 90,
            isResizable: true,
        },
        {
            key: 'Type',
            name: strings.MARequirmentViewTypeColumnLabel,
            fieldName: 'Type',
            minWidth: 55,
            maxWidth: 55,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'ResponsiblePartyItem',
            name: strings.MARequirmentViewResponsiblePartyilterColumnLabel,
            fieldName: 'ResponsiblePartyItem',
            minWidth: 85,
            maxWidth: 85,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'RequestedPartyItem',
            name: strings.MARequirementViewRequestedPartyContactsColumnLabel,
            fieldName: 'RequestedPartyItem',
            minWidth: 110,
            maxWidth: 110,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'RagStatus',
            name: strings.MARequirementViewRagStatusColumnLabel,
            fieldName: 'RagStatus',
            minWidth: 100,
            maxWidth: 100,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Synthesis',
            name: strings.MARequirementViewSynthesisColumnLabel,
            fieldName: 'Synthesis',
            minWidth: 120,
            maxWidth: 120,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'ApplicationDate',
            name: strings.MARequirementViewApplicationDateColumnLabel,
            fieldName: 'ApplicationDate',
            minWidth: 115,
            maxWidth: 115,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'EffectiveDate',
            name: strings.MARequirementViewEffectiveDateColumnLabel,
            fieldName: 'EffectiveDate',
            minWidth: 100,
            maxWidth: 100,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'ExpirationDate',
            name: strings.MARequirementViewExpirationDateColumnLabel,
            fieldName: 'ExpirationDate',
            minWidth: 110,
            maxWidth: 110,
            isResizable: true,
            onColumnClick: this.onColumnClick
        }
        ];
        return this.props.columns ?
            allColumns.filter((c) => this.isColumnEnabled(c.key)) :
            allColumns;
    }

    private initTypeChoices = async (): Promise<IDropdownOption[]> => this.typeChoices;

    public initMaVerticalChoices = async () => {
        const {
            pnpService
        } = this.props;
        let maVerticalChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MAGENERIC.VERTICALS)
            .select('Choices')();
        maVerticalChoices = data.Choices?.map((choice) => ({ key: choice, text: choice }));
        return maVerticalChoices;
    };
    private initRagStatusChoices = async () => {
        const {
            pnpService
        } = this.props;
        let ragStatusChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MAGENERIC.RAGSTATUS)
            .select('Choices')();
        ragStatusChoices = data.Choices?.map((choice) => ({ key: choice, text: choice }));
        return ragStatusChoices;
    };

    private initSynthesisStatusChoices = async () => {
        const {
            pnpService
        } = this.props;
        const synthesisChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.REQUIREMENT_STATUS_URL);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE
        ];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        items.forEach((item) => {
            synthesisChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
        });
        return synthesisChoices;
    };

    private initMarketsChoices = async () => {
        const {
            pnpService
        } = this.props;
        const marketsChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.MARKET_URL);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE
        ];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        items.forEach((item) => {
            marketsChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
        });
        return marketsChoices;
    };

    public initRequestedItemChoices = async (): Promise<IDropdownOption[]> => {
        const itemChoices: IDropdownOption[] = [];
        this.requestedItemChoices.forEach((item) => {
            if (!itemChoices.find((i) => i.key === item.key)) {
                itemChoices.push(item);
            }
        });
        itemChoices.sort((a, b) => a.text?.localeCompare(b.text));
        return itemChoices;
    };

    public initResplonsibleItemChoices = async (): Promise<IDropdownOption[]> => {
        const itemChoices: IItemOption[] = [];
        this.responsibleItemChoices.forEach((item) => {
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
        const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.config.list);
        const items = await this.props.pnpService.getListItems(listUrl).
            select(...selectedFields)
            .expand(...expand)
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
                case ContactType.Eutelsat:
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
                case ContactType.TP:
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
                case ContactType.DP:
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
                ResponsibleParty: responsibleParty,
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

    private initFilters = (): IFilter[] => {
        let allFilters: IFilter[] = [];
        const searchFilter: IFilter = {
            name: 'Search',
            displayName: strings.SNPViewSearchFilterLabel,
            placeHolder: strings.SNPViewSearchFilterPlaceHolder,
            fields: ['Title', 'MaVertical', 'Type', 'Markets', 'ResponsiblePartyItem', 'RequestedPartyItem',
                'RagStatus', 'Synthesis'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.Or,
        };
        const MaVerticalTypeFilter: IFilter = {
            name: 'MaVertical',
            displayName: strings.MARequimentViewMaVerticalFilterLabel,
            placeHolder: strings.MARequimentViewMaVerticalFilterPlaceholder,
            fields: ['MaVertical'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const typeFilter: IFilter = {
            name: 'Type',
            displayName: strings.MARequirmentViewTypeFilterLabel,
            placeHolder: strings.MARequirmentViewTypeColumnPlaceholder,
            fields: ['Type'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const marketsFilter: IFilter = {
            name: 'Markets',
            displayName: strings.MARequirmentViewMarketsFilterLabel,
            placeHolder: strings.MARequirmentViewMarketsColumnPlaceholder,
            fields: ['Markets'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const responsiblePartyItemFilter: IFilter = {
            name: 'ResponsiblePartyItem',
            displayName: strings.MARequirmentViewResponsiblePartyilterLabel,
            placeHolder: strings.MARequirmentViewResponsiblePartyilterPlaceholder,
            fields: ['ResponsiblePartyItem'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const RequestedPartyItemFilter: IFilter = {
            name: 'RequestedPartyItem',
            displayName: strings.MARequirementViewRequestedPartyContactsFilterLabel,
            placeHolder: strings.MARequirementViewRequestedPartyFilterPlaceholder,
            fields: ['RequestedPartyItem'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const ragStatusFilter: IFilter = {
            name: 'RagStatus',
            displayName: strings.MARequirementViewRagStatusFilterLabel,
            placeHolder: strings.MARequirementViewRagStatusFilterPlaceholder,
            fields: ['RagStatus'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const synthesisFilter: IFilter = {
            name: 'Synthesis',
            displayName: strings.MARequirementViewSynthesisFilterLabel,
            placeHolder: strings.MARequirementViewSynthesisFilterPlaceholder,
            fields: ['Synthesis'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        allFilters = [searchFilter, MaVerticalTypeFilter, typeFilter, marketsFilter,
            responsiblePartyItemFilter, RequestedPartyItemFilter, ragStatusFilter, synthesisFilter];
        return this.props.filters ? allFilters.filter((f) => this.isFilterEnabled(f.name)) : allFilters;
    };

    private init = async () => {
        const columns = this.initColumns();
        let filters = this.initFilters();
        let items: IMARequirement[];
        if (this.props.viewItems) {
            items = this.props.viewItems.slice();
            this.props.viewItems.forEach((itm) => {
                if (itm.ResponsiblePartyItem) {
                    this.responsibleItemChoices.push({
                        id: itm.ResponsiblePartyItem.id,
                        key: itm.ResponsiblePartyItem.key,
                        text: itm.ResponsiblePartyItem.text,
                        type: itm.ResponsiblePartyItem.type
                    });
                }
            });
            this.props.viewItems.forEach((itm) => {
                if (itm.RequestedPartyItem) {
                    this.requestedItemChoices.push({
                        key: itm.RequestedPartyItem.Id,
                        text: itm.RequestedPartyItem.Title
                    });
                }
            });
        } else {
            items = await this.initViewItems();
        }

        let MaVerticalType: IDropdownOption[] = [];
        let marketsType: IDropdownOption[] = [];
        let contentType: IItemOption[] = [];
        let responsiblePartyItemType: IDropdownOption[] = [];
        let RequestedPartyItemType: IDropdownOption[] = [];
        let RagStatusType: IDropdownOption[] = [];
        let synthesisType: IDropdownOption[] = [];
        const promises: Promise<void>[] = [];

        if (this.isFilterEnabled('MaVertical')) {
            promises.push(this.initMaVerticalChoices().then((opts) => { MaVerticalType = opts as IDropdownOption[]; }));
        }
        if (this.isFilterEnabled('Type')) {
            promises.push(this.initTypeChoices().then((opts) => { contentType = opts as IItemOption[]; }));
        }
        if (this.isFilterEnabled('Markets')) {
            promises.push(this.initMarketsChoices().then((opts) => { marketsType = opts as IDropdownOption[]; }));
        }
        if (this.isFilterEnabled('ResponsiblePartyItem')) {
            promises.push(this.initResplonsibleItemChoices().then((opts) => { responsiblePartyItemType = opts as IItemOption[]; }));
        }
        if (this.isFilterEnabled('RequestedPartyItem')) {
            promises.push(this.initRequestedItemChoices().then((opts) => { RequestedPartyItemType = opts as IItemOption[]; }));
        }
        if (this.isFilterEnabled('RagStatus')) {
            promises.push(this.initRagStatusChoices().then((opts) => { RagStatusType = opts as IDropdownOption[]; }));
        }
        if (this.isFilterEnabled('Synthesis')) {
            promises.push(this.initSynthesisStatusChoices().then((opts) => { synthesisType = opts as IDropdownOption[]; }));
        }
        await Promise.all(promises);

        filters = filters.map((f) => {
            if (f.name === 'MaVertical') return { ...f, values: MaVerticalType };
            if (f.name === 'Type') return { ...f, values: contentType };
            if (f.name === 'Markets') return { ...f, values: marketsType };
            if (f.name === 'ResponsiblePartyItem') return { ...f, values: responsiblePartyItemType };
            if (f.name === 'RequestedPartyItem') return { ...f, values: RequestedPartyItemType };
            if (f.name === 'RagStatus') return { ...f, values: RagStatusType };
            if (f.name === 'Synthesis') return { ...f, values: synthesisType };
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

    public applyFilters = (): IMARequirement[] => {
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

                            if (field === 'Markets') {
                                return (item[field] || []).some((c) => {
                                    const hay = UtilHelper.toNormalForm(c?.Title?.toString()?.toLowerCase());
                                    return hay?.includes(needle) === true;
                                });
                            } else if (field === 'Type') {
                                const hay = UtilHelper.toNormalForm((item[field] as IContentType)?.Name?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            } else if (field === 'MaVertical') {
                                return (item[field] || []).some((c) => {
                                    const hay = UtilHelper.toNormalForm(c?.text?.toString()?.toLowerCase());
                                    return hay?.includes(needle) === true;
                                });
                            } else if (field === 'Synthesis') {
                                const hay = UtilHelper.toNormalForm(item[field]?.Title?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            } else if (field === 'ResponsiblePartyItem') {
                                const itemValue = item[field];
                                if (!itemValue) return false;
                                const hayTitle = UtilHelper.toNormalForm(itemValue.text?.toString()?.toLowerCase());
                                return (hayTitle?.includes(needle) === true);
                            } else if (field === 'RequestedPartyItem') {
                                const hay = UtilHelper.toNormalForm(item[field]?.Title?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            } else {
                                const hay = UtilHelper.toNormalForm(item[field]?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            }
                        }

                        if (filter.name === 'Markets') {
                            return (item[field] || []).some((t) => String(t?.Id) === String((f as IDropdownOption).key));
                        }
                        if (filter.name === 'Synthesis') {
                            return String(item[field]?.Id) === String((f as IDropdownOption).key);
                        }
                        if (filter.name === 'RagStatus') {
                            return String(item[field]) === String((f as IDropdownOption).key);
                        }
                        if (filter.name === 'Type') {
                            return String(item[field]?.Id) === String((f as IDropdownOption).key);
                        }
                        if (filter.name === 'MaVertical') {
                            return (item[field] || []).some((t) => String(t?.key) === String((f as IDropdownOption).key));
                        }
                        if (filter.name === 'ResponsiblePartyItem') {
                            const itemValue = item[field];
                            if (!itemValue) return false;
                            return String(itemValue.key) === String((f as IDropdownOption).key) && itemValue.type === (f as IItemOption).type;
                        }
                        if (filter.name === 'RequestedPartyItem') {
                            return String(item[field]?.Id) === String((f as IDropdownOption).key);
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
        const key = sort.Property as keyof IMARequirement;
        const desc = sort.Direction === SortDirection.Descending;
        return items.slice().sort((a, b) => {
            if (key === 'ExpirationDate' || key === 'EffectiveDate' || key === 'ApplicationDate') {
                const distantPast = new Date(-1970);
                const dateA = a[key] ? new Date(a[key]) : distantPast;
                const dateB = b[key] ? new Date(b[key]) : distantPast;
                return desc ? dateB.getTime() - dateA.getTime() : dateA.getTime() - dateB.getTime();
            }
            let av = a[key];
            let bv = b[key];
            if (key === 'ResponsiblePartyItem') {
                av = av?.text ?? '';
                bv = bv?.text ?? '';
            }
            if (key === 'RequestedPartyItem' || key === 'Synthesis') {
                av = av?.Title ?? '';
                bv = bv?.Title ?? '';
            }
            if (key === 'Type') {
                av = av?.Name ?? '';
                bv = bv?.Name ?? '';
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
        if (!column) return <span />;
        const fieldName = column.fieldName as keyof IMARequirement | undefined;
        const value = fieldName ? item[fieldName] : undefined;
        let viewType = null;
        switch (item?.Type?.Name) {
            case MARequirementType.Generic:
                viewType = DatasheetType.MARequirementGeneric;
                break;
            case MARequirementType.NOL:
                viewType = DatasheetType.MARequirementNol;
                break;
            case MARequirementType.Legal:
                viewType = DatasheetType.MARequirementLegal;
                break;
            case MARequirementType.Credential:
                viewType = DatasheetType.MARequirementLicense;
                break;
        }
        const itemDatasheetUrl = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${viewType}&itemId=${item.Id}`;
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
        if (fieldName === 'Type') {
            return (
                <div className={styles.tooltipHostContainer} title="">

                    <span className={styles.tooltipHostText}> {value?.Name}</span>
                </div>
            );
        }

        if (fieldName === 'ResponsiblePartyItem') {
            const itemValue = value;
            const viewType = itemValue.type;
            return (<div className={styles.termPills}>
                <span key={itemValue.key} className={styles.termPill}>
                    <span className={styles.termPillText} key={itemValue.key}>
                        <a
                            data-interception="on"
                            rel="noreferrer"
                            title={String(itemValue?.text)}
                            onClick={(e) => {
                                e.preventDefault();
                                const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${viewType}&itemId=${itemValue?.id}`;
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
        if (fieldName === 'RagStatus') {
            return StatusPill(String(value));
        }
        if (fieldName === 'Markets') {
            const feesListType = View.config.find((f) => f.type.toLowerCase() === ViewType.Market).type;
            return renderObjectPills(
                value,
                (market) => (
                    <span key={market.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(feesListType)}&itemId=${market?.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{market?.Title}</a>
                        </span>
                    </span>
                )
            );
        }
        if (fieldName === 'RequestedPartyItem') {
            const itemValue = value;
            if (!itemValue.Title) {
                return (<TooltipHost content='No data' styles={tooltipStyles}>
                    <span className={styles.detailsListValue}>
                        <Icon iconName="Remove" className={styles.mutedIcon} />
                        <span className={styles.srOnly}>No data</span>
                    </span>
                </TooltipHost>);
            }
            const viewType = ViewType.Authority;
            return (<div className={styles.termPills}>
                <span key={itemValue.Id} className={styles.termPill}>
                    <span className={styles.termPillText}>
                        <a
                            data-interception="on"
                            rel="noreferrer"
                            title={String(itemValue?.Title)}
                            onClick={(e) => {
                                e.preventDefault();
                                const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${viewType}&itemId=${itemValue.Id}`;
                                UrlHelper.navigate(url, false);
                            }}
                        >{String(itemValue?.Title)}</a>
                    </span>
                </span>

            </div>
            );
        }
        if (fieldName === 'Synthesis') {
            return (
                <div className={styles.termPills}>
                    <span key={(value)?.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            {(value)?.Title}
                        </span>
                    </span>

                </div>
            );
        }
        if (fieldName === 'MaVertical') {
            return renderObjectPills(
                value,
                (item) => (
                    <span className={styles.termPill}>
                        <span className={styles.termPillText}>
                            {item?.text}
                        </span>
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
        const { grouping } = this.props;
        let { columns } = this.state;
        let isSortedDescending = column.isSortedDescending;
        if (grouping && grouping?.isEnabled) return;
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

    private onChangeMaVertical = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const contentTypeFilter = filters.find((f) => f.name === 'MaVertical');
        contentTypeFilter.selectedValues = options;
        filters = filters.map(((f) => ((f.name === contentTypeFilter.name) ? contentTypeFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeType = (options: IItemOption[]) => {
        let { filters } = this.state;
        const itemFilter = filters.find((f) => f.name === 'Type');
        itemFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === itemFilter.name) ? itemFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeSynthesis = (options: IItemOption[]) => {
        let { filters } = this.state;
        const itemFilter = filters.find((f) => f.name === 'Synthesis');
        itemFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === itemFilter.name) ? itemFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeRequestedPartyItem = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const contentTypeFilter = filters.find((f) => f.name === 'RequestedPartyItem');
        contentTypeFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === contentTypeFilter.name) ? contentTypeFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeRagStatus = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const marketsFilter = filters.find((f) => f.name === 'RagStatus');
        marketsFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === marketsFilter.name) ? marketsFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeResponsiblePartyItem = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const marketsFilter = filters.find((f) => f.name === 'ResponsiblePartyItem');
        marketsFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === marketsFilter.name) ? marketsFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private onChangeMarkets = (options: IItemOption[]) => {
        let { filters } = this.state;
        const itemFilter = filters.find((f) => f.name === 'Markets');
        itemFilter.selectedValues = options;
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

    public buildGroups = (items, fieldName) => {
        const groups = [];
        let currentGroup = null;
        items.forEach((item, index) => {
            const value = item[fieldName];
            if (!currentGroup || currentGroup.name !== value) {
                // close previous group
                if (currentGroup) {
                    currentGroup.count = index - currentGroup.startIndex;
                }
                // start new group
                currentGroup = {
                    key: value,
                    name: value,
                    startIndex: index,
                    count: 0
                };
                groups.push(currentGroup);
            }
        });
        // last group
        if (currentGroup) {
            currentGroup.count = items.length - currentGroup.startIndex;
        }
        return groups;
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

    public render(): React.ReactElement<IMARequirementViewProps> {
        const { filters, columns, searchKeyword, pageIndex, pageSize, isFilterPanelOpen } = this.state;
        const maVerticalFilter = filters.find((f) => f.name === 'MaVertical');
        const typeFilter = filters.find((f) => f.name === 'Type');
        const MarketsFilter = filters.find((f) => f.name === 'Markets');
        const responsiblePartyItemFilter = filters.find((f) => f.name === 'ResponsiblePartyItem');
        const RequestedPartyItemFilter = filters.find((f) => f.name === 'RequestedPartyItem');
        const RagStatusFilter = filters.find((f) => f.name === 'RagStatus');
        const synthesisFilter = filters.find((f) => f.name === 'Synthesis');
        const searchFilter = filters.find((f) => f.name === 'Search');
        let viewItems = this.applyFilters();
        viewItems = this.sort(viewItems);
        const pageCount = Math.max(1, Math.ceil(viewItems.length / pageSize));
        const safeIndex = Math.min(pageIndex, pageCount - 1);
        const start = safeIndex * pageSize;
        const end = start + pageSize;
        const pagedItems = viewItems.slice(start, end);
        const itemsToDisplay = this.props.pagination?.isEnabled ? pagedItems : viewItems;
        const groupsToDisplay =
            this.props.grouping?.isEnabled ?
                this.buildGroups(itemsToDisplay, this.props.grouping.groupBy) :
                undefined;

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
                                                />

                                            </div>
                                            <div className={`${styles.tableContainer} ${styles.tableContainerForUnwrap}`}>
                                                {/* Search */}
                                                {
                                                    this.isFilterEnabled('Search') &&
                                                    <><div>
                                                        <Label className={styles.filterLabel}>{searchFilter.displayName}</Label>
                                                    </div><div className={styles.panelChildDiv}>
                                                            <SearchBox
                                                                placeholder={searchFilter.placeHolder}
                                                                autoComplete='off'
                                                                value={searchKeyword}
                                                                onSearch={(newValue) => this.onSearch(searchFilter, newValue)}
                                                                onChange={(_, newValue) => this.onChange(searchFilter, newValue)}
                                                                onClear={() => this.setSearchText('')} />
                                                        </div></>
                                                }
                                                {/* MaVertical */}
                                                {
                                                    this.isFilterEnabled('MaVertical') &&
                                                    <><div>
                                                        <Label className={styles.filterLabel}>{maVerticalFilter.displayName}</Label>
                                                    </div><div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper placeholder={maVerticalFilter.placeHolder}
                                                                data={maVerticalFilter?.values as IItemOption[]}
                                                                dataKey={(item: IItemOption) => item.key} // composite key
                                                                textField={(item: IItemOption) => item.text}
                                                                value={maVerticalFilter?.selectedValues}
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null}
                                                                filter="contains" onChange={this.onChangeMaVertical} />
                                                        </div></>
                                                }
                                                {/* Type */}
                                                {
                                                    this.isFilterEnabled('Type') &&
                                                    <><div>
                                                        <Label className={styles.filterLabel}>{typeFilter.displayName}</Label>
                                                    </div><div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper placeholder={typeFilter.placeHolder}
                                                                data={typeFilter?.values as IDropdownOption[]}
                                                                dataKey={(item: IDropdownOption) => item.key}
                                                                textField={(item: IDropdownOption) => item.text}
                                                                value={typeFilter?.selectedValues}
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null}
                                                                filter="contains" onChange={this.onChangeType} />
                                                        </div></>
                                                }
                                                {/* Markets */}
                                                {
                                                    this.isFilterEnabled('Markets') &&
                                                    <><div>
                                                        <Label className={styles.filterLabel}>{MarketsFilter.displayName}</Label>
                                                    </div><div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper placeholder={MarketsFilter.placeHolder}
                                                                data={MarketsFilter?.values as IDropdownOption[]}
                                                                dataKey={(item: IDropdownOption) => item.key}
                                                                textField={(item: IDropdownOption) => item.text}
                                                                value={MarketsFilter?.selectedValues}
                                                                filter="contains"
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null} onChange={this.onChangeMarkets} />
                                                        </div></>
                                                }
                                                {/* ResponsiblePartyItem */}
                                                {
                                                    this.isFilterEnabled('ResponsiblePartyItem') &&
                                                    <><div>
                                                        <Label className={styles.filterLabel}>{responsiblePartyItemFilter.displayName}</Label>
                                                    </div><div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper placeholder={responsiblePartyItemFilter.placeHolder}
                                                                data={responsiblePartyItemFilter?.values as IItemOption[]}
                                                                dataKey={(item: IItemOption) => item.key} // composite key
                                                                textField={(item: IItemOption) => item.text}
                                                                value={responsiblePartyItemFilter?.selectedValues}
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null}
                                                                filter="contains" onChange={this.onChangeResponsiblePartyItem} />
                                                        </div></>
                                                }
                                                {/* RequestedPartyItem */}
                                                {
                                                    this.isFilterEnabled('RequestedPartyItem') &&
                                                    <><div>
                                                        <Label className={styles.filterLabel}>{RequestedPartyItemFilter.displayName}</Label>
                                                    </div><div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper placeholder={RequestedPartyItemFilter.placeHolder}
                                                                data={RequestedPartyItemFilter?.values as IDropdownOption[]}
                                                                dataKey={(item: IDropdownOption) => item.key}
                                                                textField={(item: IDropdownOption) => item.text}
                                                                value={RequestedPartyItemFilter?.selectedValues}
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null}
                                                                filter="contains" onChange={this.onChangeRequestedPartyItem} />
                                                        </div></>
                                                }
                                                {/* RagStatus */}
                                                {
                                                    this.isFilterEnabled('RagStatus') &&
                                                    <><div>
                                                        <Label className={styles.filterLabel}>{RagStatusFilter.displayName}</Label>
                                                    </div><div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper placeholder={RagStatusFilter.placeHolder}
                                                                data={RagStatusFilter?.values as IDropdownOption[]}
                                                                dataKey={(item: IDropdownOption) => item.key}
                                                                textField={(item: IDropdownOption) => item.text}
                                                                value={RagStatusFilter?.selectedValues}
                                                                filter="contains"
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null} onChange={this.onChangeRagStatus} />
                                                        </div></>
                                                }
                                                {/* Synthesis */}
                                                {
                                                    this.isFilterEnabled('Synthesis') &&
                                                    <><div>
                                                        <Label className={styles.filterLabel}>{synthesisFilter.displayName}</Label>
                                                    </div><div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper placeholder={synthesisFilter.placeHolder}
                                                                data={synthesisFilter?.values as IItemOption[]}
                                                                dataKey={(item: IItemOption) => item.key} // composite key
                                                                textField={(item: IItemOption) => item.text}
                                                                value={synthesisFilter?.selectedValues}
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null}
                                                                filter="contains" onChange={this.onChangeSynthesis} />
                                                        </div></>
                                                }


                                            </div>
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
                                        groups={groupsToDisplay}
                                        groupProps={{ showEmptyGroups: false }}
                                        disableSelectionZone
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
                                        <span>No record found</span>
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


