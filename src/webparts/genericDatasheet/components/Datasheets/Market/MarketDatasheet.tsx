/* eslint-disable max-len */
import * as React from 'react';
import { IMarketDatasheetProps, IMarketDatasheetState } from './Market.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import * as commonStrings from 'AppCustomizerStrings';

import {
  ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, IDropdownOption, ITooltipHostStyles, Label,
  LayerHost, Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost
} from '@fluentui/react';
import { Datasheet, DatasheetType, ISanctionOption } from '../../Datasheet.types';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { IBaseItem, IContact, IContentType, ICountry, IMARequirement, IMarket, IMarketReadiness, ISanctionCategory, Iterm } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { renderArrayPills, renderObjectPills, StatusPill } from '../../Datasheet.utility';
import { RichText } from '../../../../../common/components/RichText/RichText';
import { ViewType } from '../../../../genericView/components/View.types';
import MARequirementView from '../../../../genericView/components/MARequirement View/MARequirementView';
import { SortDirection } from '@pnp/sp/search';
import LibraryView from '../../../../genericView/components/Library View/LibraryView';
import MarketReadinessDatasheetView from '../../Datasheet Views/MarketReadinessDatasheetView';
import { MarketReadinessResponsibleParty, UserRole } from '../../../../../common/models/Enums';
import Market from '../../../../genericForm/components/Market/Market';
import AccessDeniedMessage from '../../../../../common/components/Access Denied/AccessDenied';
import { IMessageBanner } from '../../../../genericForm/components/Form.types';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';
import DomHelper from '../../../../../common/helpers/DomHelper';
export default class MarketDatasheet extends React.Component<
  IMarketDatasheetProps,
  IMarketDatasheetState
> {

  constructor(props: Readonly<IMarketDatasheetProps>) {
    super(props);
    this.state = {
      isDatasheetReady: false,
      showEditFormDialog: false,
      showDeleteDialog: false,
      item: null,
      maRequirements: [],
      marketReadinessItems: [],
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
    if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner,
      UserRole.Contributor, UserRole.Visitor, UserRole.Admin)) {
      await this.init();
      this.setState({ isDatasheetReady: true });
    }
  }

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
    const items = await pnpService.getListByUrl(listUrl).items.select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();


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

  public initItem = async () => {
    const { pnpService, itemId } = this.props;
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.MARKET.GROUPING,
      Consts.FIELDS.MARKET.MARKET_ACCESS_PRIORITY,
      Consts.FIELDS.MARKET.MARKET_ACCESS_CAPACITY,
      `${Consts.FIELDS.MARKET.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.MARKET.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
      Consts.FIELDS.MARKET.IS_SANCTIONED,
      Consts.FIELDS.MARKET.COMMENTS,
      Consts.FIELDS.COMMON.CREATION_TIME,
      Consts.FIELDS.COMMON.MODIFICATION_TIME,
      Consts.FIELDS.COMMON.DOCUMENTS_SPACE,
      `${Consts.FIELDS.MARKET.SANCTION_CATEGORY}/${Consts.FIELDS.COMMON.ID}`,
    ];
    const expand = [Consts.FIELDS.MARKET.SANCTION_CATEGORY, Consts.FIELDS.MARKET.COUNTRY];
    const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.Market).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    try {
      const item = await pnpService
        .getListByUrl(listUrl)
        .items.getById(itemId)
        .select(...selectedFields)
        .expand(...expand)();
      const market: IMarket = {
        Id: item[Consts.FIELDS.COMMON.ID],
        Title: item[Consts.FIELDS.COMMON.TITLE],
        Country: {
          Id: item[Consts.FIELDS.MARKET.COUNTRY][Consts.FIELDS.COMMON.ID],
          Title: item[Consts.FIELDS.MARKET.COUNTRY][Consts.FIELDS.COMMON.TITLE]
        },
        Comments: item[Consts.FIELDS.MARKET.COMMENTS],
        Grouping: item[Consts.FIELDS.MARKET.GROUPING],
        SanctionCategory: {
          Id: item[Consts.FIELDS.MARKET.SANCTION_CATEGORY][Consts.FIELDS.COMMON.ID]
        },
        MaPriority: item[Consts.FIELDS.MARKET.MARKET_ACCESS_PRIORITY],
        Capacity: item[Consts.FIELDS.MARKET.MARKET_ACCESS_CAPACITY],
        IsSanctioned: item[Consts.FIELDS.MARKET.IS_SANCTIONED],
        Created: item[Consts.FIELDS.COMMON.CREATION_TIME],
        Modified: item[Consts.FIELDS.COMMON.MODIFICATION_TIME],
        DocumentSpace: item[Consts.FIELDS.COMMON.DOCUMENTS_SPACE]?.Url
      };
      return market;
    } catch (error) {
      if (error instanceof Error && error.message.includes(strings.DataSheetItemDoesNotExist)) {
        const errorMessage = strings.DataSheetItemDoesNotExist;
        this.setState({ errorMessage });
        return null;
      }
      throw error;
    }
  };

  public initMarketReadinessItems = async (): Promise<IMarketReadiness[]> => {
    const { pnpService, itemId } = this.props;
    const list = Consts.LISTS.MA_READINESS_URL;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.MARKET_READINESS.RESPONSIBLE_PARTY,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      Consts.FIELDS.MARKET_READINESS.VERTICALS,
      Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_RAGSTATUS,
      Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_ESTIMATED_DATE,
      Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_EFFECTIVE_DATE,
      Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_COMMENTS,
      Consts.FIELDS.MARKET_READINESS.SPACE_NW_RAGSTATUS,
      Consts.FIELDS.MARKET_READINESS.SPACE_NW_ESTIMATED_DATE,
      Consts.FIELDS.MARKET_READINESS.SPACE_NW_EFFECTIVE_DATE,
      Consts.FIELDS.MARKET_READINESS.SPACE_NW_COMMENTS,
      Consts.FIELDS.MARKET_READINESS.LAND_FIXED_RAGSTATUS,
      Consts.FIELDS.MARKET_READINESS.LAND_FIXED_ESTIMATED_DATE,
      Consts.FIELDS.MARKET_READINESS.LAND_FIXED_EFFECTIVE_DATE,
      Consts.FIELDS.MARKET_READINESS.LAND_FIXED_COMMENTS,
      Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_RAGSTATUS,
      Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_ESTIMATED_DATE,
      Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_EFFECTIVE_DATE,
      Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_COMMENTS,
      Consts.FIELDS.MARKET_READINESS.MARITIME_RAGSTATUS,
      Consts.FIELDS.MARKET_READINESS.MARITIME_ESTIMATED_DATE,
      Consts.FIELDS.MARKET_READINESS.MARITIME_EFFECTIVE_DATE,
      Consts.FIELDS.MARKET_READINESS.MARITIME_COMMENTS,
      Consts.FIELDS.MARKET_READINESS.AVIATION_RAGSTATUS,
      Consts.FIELDS.MARKET_READINESS.AVIATION_ESTIMATED_DATE,
      Consts.FIELDS.MARKET_READINESS.AVIATION_EFFECTIVE_DATE,
      Consts.FIELDS.MARKET_READINESS.AVIATION_COMMENTS,
      `${Consts.FIELDS.MARKET_READINESS.MARKET}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.MARKET_READINESS.MARKET}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.MARKET_READINESS.OWNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.MARKET_READINESS.OWNER}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const expand = [Consts.FIELDS.MARKET_READINESS.MARKET,
    Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER,
    Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER,
    Consts.FIELDS.MARKET_READINESS.OWNER];
    const items = await pnpService.getListByUrl(listUrl).items.
      filter(`${Consts.FIELDS.MARKET_READINESS.MARKET_ID} eq ${itemId}`).
      select(...selectedFields).
      expand(...expand)();

    return items.map((item): IMarketReadiness => {
      let ResponsiblePartyItem: IBaseItem | null = null;
      const ResponsiblePartyType: string = item[Consts.FIELDS.MARKET_READINESS.RESPONSIBLE_PARTY];
      if (ResponsiblePartyType === MarketReadinessResponsibleParty.TP) {
        ResponsiblePartyItem = {
          Id: item[Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER][Consts.FIELDS.COMMON.ID],
          Title: item[Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER][Consts.FIELDS.COMMON.TITLE],
        };
      }
      if (ResponsiblePartyType === MarketReadinessResponsibleParty.DP) {
        ResponsiblePartyItem = {
          Id: item[Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
          Title: item[Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE],
        };
      }
      const marketReadiness: IMarketReadiness = {
        Id: item[Consts.FIELDS.COMMON.ID],
        Title: item[Consts.FIELDS.COMMON.TITLE],
        Market: {
          Id: item[Consts.FIELDS.MARKET_READINESS.MARKET][Consts.FIELDS.COMMON.ID],
          Title: item[Consts.FIELDS.MARKET_READINESS.MARKET][Consts.FIELDS.COMMON.TITLE],
        },
        ResponsiblePartyType,
        ResponsiblePartyItem,
        EutelsatOwners: item[Consts.FIELDS.MARKET_READINESS.OWNER]?.map((etuowner) => ({
          Id: etuowner[Consts.FIELDS.COMMON.ID],
          Title: etuowner[Consts.FIELDS.COMMON.TITLE]
        })) || [],
        Verticals: item[Consts.FIELDS.MARKET_READINESS.VERTICALS]?.map((vertical) => ({
          Title: vertical
        })) || [],
        AllVerticalRAGStatus: item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_RAGSTATUS] || '',
        AllVerticalEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_ESTIMATED_DATE]) : null,
        AllVerticalEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_EFFECTIVE_DATE]) : null,
        AllVerticalComments: item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_COMMENTS] || '',

        SpaceRAGStatus: item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_RAGSTATUS] || '',
        SpaceEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_ESTIMATED_DATE]) : null,
        SpaceEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_EFFECTIVE_DATE]) : null,
        SpaceComments: item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_COMMENTS] || '',

        LandFixedRAGStatus: item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_RAGSTATUS] || '',
        LandFixedEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_ESTIMATED_DATE]) : null,
        LandFixedEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_EFFECTIVE_DATE]) : null,
        LandFixedComments: item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_COMMENTS] || '',

        LandMobilityRAGStatus: item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_RAGSTATUS] || '',
        LandMobilityEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_ESTIMATED_DATE]) : null,
        LandMobilityEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_EFFECTIVE_DATE]) : null,
        LandMobilityComments: item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_COMMENTS] || '',

        MaritimeRAGStatus: item[Consts.FIELDS.MARKET_READINESS.MARITIME_RAGSTATUS] || '',
        MaritimeEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.MARITIME_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.MARITIME_ESTIMATED_DATE]) : null,
        MaritimeEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.MARITIME_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.MARITIME_EFFECTIVE_DATE]) : null,
        MaritimeComments: item[Consts.FIELDS.MARKET_READINESS.MARITIME_COMMENTS] || '',

        AviationRAGStatus: item[Consts.FIELDS.MARKET_READINESS.AVIATION_RAGSTATUS] || '',
        AviationEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.AVIATION_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.AVIATION_ESTIMATED_DATE]) : null,
        AviationEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.AVIATION_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.AVIATION_EFFECTIVE_DATE]) : null,
        AviationComments: item[Consts.FIELDS.MARKET_READINESS.AVIATION_COMMENTS] || '',
      };
      return marketReadiness;
    });

  };

  public initMARequirementItems = async () => {
    const { pnpService, itemId } = this.props;
    const typeChoices: IDropdownOption[] = [
      { key: Consts.CONTENT_TYPES.MA_RequirementCredential, text: commonStrings.MARequirementCredentialsContentTypeName },
      { key: Consts.CONTENT_TYPES.MA_RequirementGeneric, text: commonStrings.MARequirementGenericContentTypeName },
      { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_NOL, text: commonStrings.MARequirementNolContentTypeName },
      { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_LEGAL, text: commonStrings.MARequirementLegalContentTypeName },
    ];
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
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.MA_REQUIREMENT_URL);
    const items = await pnpService.getListByUrl(listUrl).items.
      filter(`${Consts.FIELDS.MAGENERIC.MARKETS_ID} eq ${itemId}`)
      .select(...selectedFields)
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
        case commonStrings.MARequirementEutelsatTypeName:
          if (item?.[Consts.FIELDS.MAGENERIC.EUTELSATENTITY]?.[Consts.FIELDS.COMMON.ID] && item?.[Consts.FIELDS.MAGENERIC.EUTELSATENTITY]?.[Consts.FIELDS.COMMON.TITLE]) {
            responsiblePartyItem = {
              id: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID],
              key: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID] + "-" + ViewType.EutelsatEntity,
              text: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.TITLE],
              type: ViewType.EutelsatEntity,
            };
          }
          break;
        case commonStrings.MARequirementTeleportPartnerTypeName:
          if (item?.[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER]?.[Consts.FIELDS.COMMON.ID] &&
            item?.[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER]?.[Consts.FIELDS.COMMON.TITLE]) {
            responsiblePartyItem = {
              id: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
              key: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + ViewType.TeleportPartner,
              text: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
              type: ViewType.TeleportPartner,
            };
          }
          break;
        case commonStrings.MARequirementDistributionPartnerTypeName:
          if (item?.[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.ID] && item?.[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.TITLE]) {
            responsiblePartyItem = {
              id: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
              key: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID] + "-" + ViewType.DistributionPartner,
              text: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE],
              type: ViewType.DistributionPartner,
            };
          }
          break;
        default:
          break;
      }
      const Type: IContentType = typeChoices
        .map((type) => ({ Id: type.key + "", Name: type.text }))
        .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1) || { Id: '', Name: '' };
      const authority = item?.[Consts.FIELDS.MAGENERIC.AUTHORITY];
      const RequestedPartyItem = authority ? {
        Id: authority[Consts.FIELDS.COMMON.ID],
        Title: authority[Consts.FIELDS.COMMON.TITLE]
      } : null;
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

  private init = async () => {
    let item: IMarket;
    let sanctions: ISanctionOption[] = [];
    let maRequirements: IMARequirement[] = [];
    let marketReadinessItems: IMarketReadiness[] = [];
    const promises: Promise<void>[] = [];

    promises.push(this.initItem().then((opts) => { item = opts as IMarket; }));
    promises.push(this.initSanctionCategoriesChoices().then((opts) => { sanctions = opts as ISanctionOption[]; }));
    promises.push(this.initMARequirementItems().then((opts) => { maRequirements = opts as IMARequirement[]; }));
    promises.push(this.initMarketReadinessItems().then((opts) => { marketReadinessItems = opts as IMarketReadiness[]; }));
    await Promise.all(promises);
    const sanctionCategory = sanctions.find((cat) => cat.key.toString() === item.SanctionCategory.Id.toString());
    item.SanctionCategory = {
      Id: item.SanctionCategory.Id,
      Title: sanctionCategory.text,
      Color: sanctionCategory.color,
      Summary: sanctionCategory.summary
    };
    this.setState({
      item,
      maRequirements,
      marketReadinessItems
    });
  };

  private renderDatasheetInfo = (): JSX.Element | null => {
    const leftColumns = ['Country', 'Grouping', 'SanctionCategory', 'MaPriority'];
    const rightColumns = ['IsSanctioned', 'Capacity', 'Comments'];

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
    const labelKey = `MarketDataSheetInfoLabel${key}`;
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
    const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
    const value = fieldName ? item[fieldName] : undefined;
    const swatchStyle = (color?: string) => ({
      width: 12,
      height: 12,
      borderRadius: 3,
      background: color || '#888',
      display: 'inline-block',
      verticalAlign: 'middle' as const,
      marginRight: 6,
      border: '1px solid rgba(0,0,0,.1)',
    });
    if (value === null || value === undefined || (Array.isArray(value) && value.length === 0)) {
      return (<TooltipHost content='No data' styles={tooltipStyles}>
        <span className={styles.detailsListValue}>
          <Icon iconName="Remove" className={styles.mutedIcon} />
          <span className={styles.srOnly}>No data</span>
        </span>
      </TooltipHost>);
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
            <span style={swatchStyle(sc.Color)} />
            <span className={styles.tooltipHostText}> {sc.Title}</span>
          </div>
        </TooltipHost>
      );
    }
    if (fieldName === 'EutelsatOwners') {
      return renderObjectPills(
        value as IContact[],
        (owner) => (
          <span key={owner.Id} className={styles.termPill}>
            <span className={styles.termPillText}>
              <a
                data-interception="on"
                rel="noreferrer"
                onClick={(e) => {
                  e.preventDefault();
                  const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Contact)}&itemId=${owner.Id}`;
                  UrlHelper.navigate(url, false);
                }}
              >{owner.Title}</a>
            </span>
          </span>
        )
      );
    }
    if ([
      'Comments',
      'AllVerticalComments',
      'SpaceComments',
      'LandFixedComments',
      'LandMobilityComments',
      'MaritimeComments',
      'AviationComments'
    ].includes(fieldName)) {
      //return <RichText isEditMode={false} value={value} />;
      const cleanValue = DomHelper.cleanRichHtml(item[fieldName]);
      if (!cleanValue) return (<TooltipHost content='No data' styles={tooltipStyles}>
        <span className={styles.detailsListValue}>
          <Icon iconName="Remove" className={styles.mutedIcon} />
          <span className={styles.srOnly}>No data</span>
        </span>
      </TooltipHost>);
      return <RichText isEditMode={false} value={cleanValue} />;
    }
    if ([
      'AllVerticalRAGStatus',
      'SpaceRAGStatus',
      'LandFixedRAGStatus',
      'LandMobilityRAGStatus',
      'MaritimeRAGStatus',
      'AviationRAGStatus'
    ].includes(fieldName)) {
      return StatusPill(String(value));
    }
    if (fieldName === 'Country') {
      const ct = item[fieldName] as ICountry;
      return renderObjectPills(
        [ct],
        (country) => (
          <span key={country.Id} className={styles.termPill}>
            <span className={styles.termPillText}>
              <a
                data-interception="on"
                rel="noreferrer"
                onClick={(e) => {
                  e.preventDefault();
                  const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Country)}&itemId=${ct.Id}`;
                  UrlHelper.navigate(url, false);
                }}
              >{String(ct?.['Title'])}</a>
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










  public render(): React.ReactElement<IMarketDatasheetProps> {
    const {
      showEditFormDialog,
      showDeleteDialog,
      item,
      maRequirements
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
    return (
      <>
        {this.state && this.state.isDatasheetReady ?

          <div>
            <div className={styles.datasheetHeader}>
              <div className={styles.datasheetTypeContainer}>
                < h5 className={styles.datasheetType} >
                  {strings.MarketDataSheetTitle}
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
                <PivotItem headerText="Market Readiness" itemIcon="Rocket" >
                  <MarketReadinessDatasheetView
                    items={this.state.marketReadinessItems}
                    renderField={(item, key) => this.renderDatasheetInfoValue(item, key)}
                    pnpService={pnpService}
                    userService={userService}
                    refresh={this.init}
                    marketItemId={item.Id}
                  />
                </PivotItem>
                <PivotItem headerText="MA Requirements" itemIcon="Bullseye" >
                  <MARequirementView pnpService={pnpService} userService={userService}
                    columns={['Title', 'Summary', 'MaVertical', 'Type', 'RequestedPartyItem', 'ResponsiblePartyItem', 'RagStatus', 'Synthesis', 'Geo/Leo', 'ApplicationDate', 'EffectiveDate', 'ExpirationDate']}
                    filters={['Search', 'MaVertical', 'Type', 'RequestedPartyItem', 'ResponsiblePartyItem', 'RagStatus', 'Synthesis']}
                    grouping={{ isEnabled: true, groupBy: 'ResponsibleParty' }}
                    sort={{
                      Property: 'ResponsibleParty',
                      Direction: SortDirection.Ascending
                    }}
                    viewItems={maRequirements}
                    pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}>
                  </MARequirementView>
                </PivotItem>
              </Pivot>
            </div>
          </div> :
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
      const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.Market).list;
      const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
      await pnpService.getListByUrl(listUrl).items.getById(itemId).recycle();
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

    const {
      showEditFormDialog,
    } = this.state;
    const {
      pnpService,
      userService,
      itemId
    } = this.props;

    const editFormControl = (<Market pnpService={pnpService} userService={userService}
      itemId={itemId} callback={async () => { await this.closeEditForm(true); }}></Market>);
    const editFormTitle = strings.MarketDataSheetEditFormTitle;

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
