/* eslint-disable max-len */
import * as React from 'react';
import { ICountryDatasheetProps, ICountryDatasheetState } from './Country.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import * as commonStrings from 'AppCustomizerStrings';

import {
  ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, IDropdownOption, ITooltipHostStyles, Label,
  Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost

} from '@fluentui/react';
import Country from '../../../../genericForm/components/Country/Country';
import { Datasheet, DatasheetType } from '../../Datasheet.types';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { IAuthority, IContentType, ICountry, IEutelsatEntity, ILegalSvcProvider, IMarket, ISNP, Iterm } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { renderArrayPills, renderObjectPills } from '../../Datasheet.utility';
import LegalSvcProviderView from '../../../../genericView/components/Legal Svc Provider View/LegalSvcProviderView';
import AuthorityView from '../../../../genericView/components/Authority View/AuthorityView';
import SNPView from '../../../../genericView/components/SNP View/SNPView';
import { UserRole } from '../../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../../common/components/Access Denied/AccessDenied';
import EutelsatEntityView from '../../../../genericView/components/Eutelsat Entity View/EutelsatEntityView';
import MarketView from '../../../../genericView/components/Market View/MarketView';
import { IMessageBanner } from '../../../../genericForm/components/Form.types';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';

export default class CountryDatasheet extends React.Component<
  ICountryDatasheetProps,
  ICountryDatasheetState
> {

  constructor(props: Readonly<ICountryDatasheetProps>) {
    super(props);
    this.state = {
      isDatasheetReady: false,
      showEditFormDialog: false,
      showDeleteDialog: false,
      item: null,
      legalSvcProviders: [],
      authorities: [],
      snps: [],
      eutsatEntities: [],
      marketItems: [],
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

  public initItem = async () => {
    const { pnpService, itemId } = this.props;
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.COUNTRY.IDENTIFIER,
      Consts.FIELDS.COUNTRY.GROUPING,
      Consts.FIELDS.COMMON.CREATION_TIME,
      Consts.FIELDS.COMMON.MODIFICATION_TIME
    ];
    const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.Country).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    try {
      const item = await pnpService
        .getListByUrl(listUrl)
        .items.getById(itemId)
        .select(...selectedFields)();
      const country: ICountry = {
        Id: item[Consts.FIELDS.COMMON.ID],
        Title: item[Consts.FIELDS.COMMON.TITLE],
        Identifier: item[Consts.FIELDS.COUNTRY.IDENTIFIER],
        Grouping: item[Consts.FIELDS.COUNTRY.GROUPING],
        Created: item[Consts.FIELDS.COMMON.CREATION_TIME],
        Modified: item[Consts.FIELDS.COMMON.MODIFICATION_TIME],
      };
      return country;
    } catch (error) {
      if (error instanceof Error && error.message.includes(strings.DataSheetItemDoesNotExist)) {
        const errorMessage = strings.DataSheetItemDoesNotExist;
        this.setState({ errorMessage });
        return null;
      }
      throw error;
    }
  };

  public initLegalSvcProviderItems = async () => {
    const { pnpService, itemId } = this.props;
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.STATUS,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.SUMMARY,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.COMMENTS,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.ADDRESS,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), Consts.LISTS.LEGAL_SERVICE_PROVIDER_URL);
    const items = await pnpService.getListItems(listUrl).
      filter(`${Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY_ID} eq ${itemId}`).
      select(...selectedFields).
      expand(Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER, Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY).
      orderBy(Consts.FIELDS.COMMON.TITLE)();
    const viewItems: ILegalSvcProvider[] = [];
    items.map((item) => {
      viewItems.push({
        Id: item[Consts.FIELDS.COMMON.ID],
        Title: item[Consts.FIELDS.COMMON.TITLE],
        ContentType: {
          Id: item[Consts.FIELDS.COMMON.CONTENTTYPE_ID],
          Name: this.getLegalSvcProviderType(item[Consts.FIELDS.COMMON.CONTENTTYPE_ID])
        },
        Status: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.STATUS],
        Summary: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.SUMMARY],
        Comments: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.COMMENTS],
        Address: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.ADDRESS],
        Countries: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY]?.map((country) => ({
          Id: country[Consts.FIELDS.COMMON.ID],
          Title: country[Consts.FIELDS.COMMON.TITLE]
        })) || [],
        Owners: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER]?.map((etuowner) => ({
          key: etuowner[Consts.FIELDS.COMMON.ID],
          text: etuowner[Consts.FIELDS.COMMON.TITLE]
        })) || []
      });
    });
    return viewItems;

  };

  public initAuthorityItems = async () => {
    const { pnpService, itemId } = this.props;
    const typeChoices: IDropdownOption[] = [
      { key: Consts.CONTENT_TYPES.ADMINISTRATION, text: commonStrings.AdministrationContentTypeName },
      { key: Consts.CONTENT_TYPES.ORGANIZATION, text: commonStrings.OrganizationContentTypeName },
      { key: Consts.CONTENT_TYPES.REGULATOR, text: commonStrings.RegulatorContentTypeName }
    ];
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
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.AUTHORITY_URL);
    const items = await pnpService.getListItems(listUrl).
      filter(`${Consts.FIELDS.AUTHORITY.COUNTRY_ID} eq ${itemId}`).
      select(...selectedFields)
      .expand(...expand)
      .orderBy(Consts.FIELDS.COMMON.TITLE)();

    const viewItems: IAuthority[] = [];

    items.map((item) => {
      //adding filter
      const ContentType: IContentType = typeChoices
        .map((type) => ({ Id: type.key + "", Name: type.text }))
        .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1) || { Id: '', Name: '' };

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

  public initSNPItems = async () => {
    const { pnpService, itemId } = this.props;
    const typeChoices: IDropdownOption[] = [
      { key: Consts.CONTENT_TYPES.STARGATE, text: commonStrings.SnpStargateContentTypeName },
      { key: Consts.CONTENT_TYPES.EUTELSAT_SNP, text: commonStrings.SnpEutelsatContentTypeName },
      { key: Consts.CONTENT_TYPES.PARTNER_SNP, text: commonStrings.SnpPartnerContentTypeName }
    ];
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.SNP.SUMMARY,
      Consts.FIELDS.SNP.CITY,
      Consts.FIELDS.SNP.COMMENTS,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      Consts.FIELDS.SNP.STATUS,
      `${Consts.FIELDS.SNP.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.SNP.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.SNP.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.SNP.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const expand = [Consts.FIELDS.SNP.COUNTRY, Consts.FIELDS.SNP.TELEPORTPARTNER];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.SNP_URL);
    const items = await pnpService.getListItems(listUrl).
      filter(`${Consts.FIELDS.SNP.COUNTRY_ID} eq ${itemId}`).
      select(...selectedFields)
      .expand(...expand)
      .orderBy(Consts.FIELDS.COMMON.TITLE)();
    const viewItems: ISNP[] = [];
    items.map((item) => {
      const ContentType: IContentType = typeChoices
        .map((type) => ({ Id: type.key + "", Name: type.text }))
        .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1) || { Id: '', Name: '' };
      let TeleportPartner;
      if (item[Consts.FIELDS.SNP.TELEPORTPARTNER]) {
        TeleportPartner = {
          Id: item[Consts.FIELDS.SNP.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
          Title: item[Consts.FIELDS.SNP.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE]
        };
      } else {
        TeleportPartner = null;
      }
      let Country;
      if (item[Consts.FIELDS.SNP.COUNTRY]) {
        Country = {
          Id: item[Consts.FIELDS.SNP.COUNTRY][Consts.FIELDS.COMMON.ID],
          Title: item[Consts.FIELDS.SNP.COUNTRY][Consts.FIELDS.COMMON.TITLE]
        };
      } else {
        Country = null;
      }
      viewItems.push({
        Id: item[Consts.FIELDS.COMMON.ID],
        Title: item[Consts.FIELDS.COMMON.TITLE],
        ContentType,
        Summary: item[Consts.FIELDS.SNP.SUMMARY],
        City: item[Consts.FIELDS.SNP.CITY],
        Country,
        TeleportPartner,
        Status: item[Consts.FIELDS.SNP.STATUS]
      });
    });
    return viewItems;
  };
  initEutSatEntityItems = async () => {
    const { pnpService, itemId } = this.props;
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.EUT_ENTITY.STATUS,
      Consts.FIELDS.EUT_ENTITY.SUMMARY,
      Consts.FIELDS.EUT_ENTITY.COMMENTS,
      Consts.FIELDS.EUT_ENTITY.ADDRESS,
      `${Consts.FIELDS.EUT_ENTITY.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.EUT_ENTITY.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.EUTELSAT_ENTITY_URL);
    const items = await pnpService.getListItems(listUrl).
      select(...selectedFields).
      filter(`${Consts.FIELDS.EUT_ENTITY.COUNTRY_ID} eq ${itemId}`).
      expand(Consts.FIELDS.EUT_ENTITY.COUNTRY).
      orderBy(Consts.FIELDS.COMMON.TITLE)();
    const viewItems: IEutelsatEntity[] = [];
    items.map((item) => {
      viewItems.push({
        Id: item[Consts.FIELDS.COMMON.ID],
        Title: item[Consts.FIELDS.COMMON.TITLE],
        Status: item[Consts.FIELDS.EUT_ENTITY.STATUS],
        Summary: item[Consts.FIELDS.EUT_ENTITY.SUMMARY],
        Comments: item[Consts.FIELDS.EUT_ENTITY.COMMENTS],
        Address: item[Consts.FIELDS.EUT_ENTITY.ADDRESS],
        Country: {
          Id: item[Consts.FIELDS.EUT_ENTITY.COUNTRY][Consts.FIELDS.COMMON.ID],
          Title: item[Consts.FIELDS.EUT_ENTITY.COUNTRY][Consts.FIELDS.COMMON.TITLE]
        }
      });
    });
    return viewItems;
  };
  initMarketItems = async () => {
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
      const items = await pnpService
        .getListItems(listUrl)
        .select(...selectedFields)
        .filter(`${Consts.FIELDS.MARKET.COUNTRY_ID} eq ${itemId}`)
        .expand(...expand)();
      const viewItems: IMarket[] = [];
      items.forEach((item) => {
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
        viewItems.push(market);
      });
      return viewItems;
    } catch (error) {
      if (error instanceof Error && error.message.includes(strings.DataSheetItemDoesNotExist)) {
        const errorMessage = strings.DataSheetItemDoesNotExist;
        this.setState({ errorMessage });
        return null;
      }
      throw error;
    }
  };
  private getLegalSvcProviderType = (key: string | number): string | undefined => {
    const keyStr = String(key);
    const typeChoices: IDropdownOption[] = [
      { key: Consts.CONTENT_TYPES.LAW_FIRM, text: commonStrings.LawFirmContentTypeName },
      { key: Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE, text: commonStrings.LegalRepresentativeContentTypeName }
    ];
    const found = typeChoices.find((type) => keyStr.indexOf(String(type.key)) !== -1);
    return found ? found.text : undefined;
  };

  private init = async () => {
    let item: ICountry;
    let legalSvcProviders: ILegalSvcProvider[] = [];
    let authorities: IAuthority[] = [];
    let snps: IAuthority[] = [];
    let eutsatEntities: IEutelsatEntity[] = [];
    let marketItems: IMarket[] = [];
    const promises: Promise<void>[] = [];
    promises.push(this.initItem().then((opts) => { item = opts as ICountry; }));
    promises.push(this.initLegalSvcProviderItems().then((opts) => { legalSvcProviders = opts as ILegalSvcProvider[]; }));
    promises.push(this.initAuthorityItems().then((opts) => { authorities = opts as IAuthority[]; }));
    promises.push(this.initSNPItems().then((opts) => { snps = opts as ISNP[]; }));
    promises.push(this.initEutSatEntityItems().then((opts) => { eutsatEntities = opts as IEutelsatEntity[]; }));
    promises.push(this.initMarketItems().then((opts) => { marketItems = opts as IMarket[]; }));
    await Promise.all(promises);
    this.setState({
      item,
      legalSvcProviders,
      authorities,
      snps,
      eutsatEntities,
      marketItems
    });
  };

  private renderDatasheetInfo = (): JSX.Element | null => {
    const leftColumns = ['Identifier'];
    const rightColumns = ['Grouping'];
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
    const labelKey = `CountryDataSheetInfoLabel${key}`;
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

  public render(): React.ReactElement<ICountryDatasheetProps> {
    const {
      showEditFormDialog,
      showDeleteDialog,
      item,
      legalSvcProviders,
      authorities,
      snps,
      eutsatEntities,
      marketItems
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
                  {strings.CountryDataSheetTitle}
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
                  {this.renderDatasheetInfo()}
                  <div className={styles.section}>
                    <Icon iconName="Globe" className="sectionIcon" />
                    <h3 >{strings.DataSheetSectionMarketItems}</h3>
                  </div>
                  <MarketView pnpService={pnpService} userService={userService}
                    columns={['Title', 'Grouping', 'SanctionCategory', 'Priority']}
                    filters={['Search', 'Grouping', 'SanctionCategory', 'Priority']}
                    viewItems={marketItems}
                    pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}>
                  </MarketView>
                </PivotItem>
                <PivotItem headerText="SNPs" itemIcon="InternetSharing" >
                  <SNPView pnpService={pnpService} userService={userService}
                    columns={['Title', 'ContentType', 'City', 'TeleportPartner', 'Status']}
                    filters={['Search', 'ContentType', 'TeleportPartner', 'Status']}
                    viewItems={snps}
                    pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}>
                  </SNPView>
                </PivotItem>
                <PivotItem headerText={strings.DataSheetSectionLegalSvcProviders} itemIcon="Education" >
                  <LegalSvcProviderView pnpService={pnpService} userService={userService}
                    filters={['Search', 'Title', 'ContentType', 'Status']}
                    columns={['Title', 'ContentType', 'Summary', 'Status']}
                    viewItems={legalSvcProviders}
                    pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}>
                  </ LegalSvcProviderView >
                </PivotItem>
                <PivotItem headerText={strings.DataSheetSectionAuthorities} itemIcon="Bank" >
                  <AuthorityView pnpService={pnpService} userService={userService}
                    columns={['Title', 'ContentType', 'Summary']}
                    filters={['Search', 'ContentType']}
                    viewItems={authorities}
                    pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}>
                  </AuthorityView>
                </PivotItem>
                <PivotItem headerText={strings.DataSheetSectionEutelsatEntities} itemIcon="Org" >
                  <EutelsatEntityView pnpService={pnpService} userService={userService}
                    columns={['Title', 'Summary', 'Status']}
                    filters={['Search', 'Status']}
                    viewItems={eutsatEntities}
                    pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}>
                  </EutelsatEntityView>
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
      const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.Country).list;
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

    const editFormControl = (<Country pnpService={pnpService} userService={userService} itemId={itemId} callback={async () => { await this.closeEditForm(true); }}></Country>);
    const editFormTitle = strings.CountryDataSheetEditFormTitle;

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
