import * as React from 'react';
import { IAuthorityDatasheetProps, IAuthorityDatasheetState } from './Authority.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import {
  ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, IDropdownOption, ITooltipHostStyles, Label,
  LayerHost, Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost
} from '@fluentui/react';
import { Datasheet, DatasheetType } from '../../Datasheet.types';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { IAuthority, IContact, IContentType, IFEE, IItemOption, IMARequirement, IPayment } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { renderArrayPills, renderObjectPills } from '../../Datasheet.utility';
import { RichText } from '../../../../../common/components/RichText/RichText';
import MARequirementView from '../../../../genericView/components/MARequirement View/MARequirementView';
import { ViewType } from '../../../../genericView/components/View.types';
import PaymentView from '../../../../genericView/components/Payment View/PaymentView';
import FeeView from '../../../../genericView/components/Fees View/FeesView';
import Authority from '../../../../genericForm/components/Authority/Authority';
import LibraryView from '../../../../genericView/components/Library View/LibraryView';
import ContactDatasheetView from '../../Datasheet Views/ContactDatasheetView';
//import { Navigation } from 'spfx-navigation';
import { AuthorityType, MARequirementType, FeeType, FeeChargedBy, UserRole } from '../../../../../common/models/Enums';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';
import { IMessageBanner } from '../../../../genericForm/components/Form.types';
import DomHelper from '../../../../../common/helpers/DomHelper';
export default class AuthorityDatasheet extends React.Component<
  IAuthorityDatasheetProps,
  IAuthorityDatasheetState
> {

  private FeesType: IDropdownOption[] = [
    { key: Consts.CONTENT_TYPES.SNP_FEE, text: FeeType.SNP },
    { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_FEE, text: FeeType.MATypeRequirement }
  ];

  private getTaxonomyValue = (obj: any, fieldName: string) => {
    const field = Object.keys(obj).find((f) =>
      Object.prototype.hasOwnProperty.call(obj, f) && f === fieldName
    );

    if (field && obj[field] !== null) {
      const thisId = obj[field].WssId;
      const matchingTerm = obj.TaxCatchAll?.find((item: any) => item.Id === thisId);
      if (matchingTerm) {
        return {
          Label: matchingTerm.Term,
          TermGuid: obj[field].TermGuid,
          WssId: obj[field].WssId
        };
      }
    }

    // No match found, return null
    return null;
  };
  private typeChoices: IDropdownOption[] = [
    { key: Consts.CONTENT_TYPES.ADMINISTRATION, text: AuthorityType.Administration },
    { key: Consts.CONTENT_TYPES.ORGANIZATION, text: AuthorityType.Organization },
    { key: Consts.CONTENT_TYPES.REGULATOR, text: AuthorityType.Regulator }
  ];

  private maReqTypeChoices: IDropdownOption[] = [
    { key: Consts.CONTENT_TYPES.MA_RequirementCredential, text: MARequirementType.Credential },
    { key: Consts.CONTENT_TYPES.MA_RequirementGeneric, text: MARequirementType.Generic },
    { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_NOL, text: MARequirementType.NOL },
    { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_LEGAL, text: MARequirementType.Legal },
  ];
  private responsibleItemChoices: IItemOption[] = [];
  private requestedItemChoices: IDropdownOption[] = [];

  private itemChoices: IItemOption[] = [];
  private ChargeByItemChoices: IItemOption[] = [];
  private beneficiaryItemChoices: IItemOption[] = [];
  constructor(props: Readonly<IAuthorityDatasheetProps>) {
    super(props);
    this.state = {
      isDatasheetReady: false,
      showEditFormDialog: false,
      showDeleteDialog: false,
      item: null,
      maReqViewItems: [],
      feeViewItems: [],
      paymentViewItems: [],
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
  }

  private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
    const userRoles = this.props.userService.userContext?.userRoles ?? [];
    return rolesToCheck.some((role) => userRoles.includes(role));
  };

  public async componentDidMount() {
    await this.init();
    this.setState({ isDatasheetReady: true });
  }


  public initItem = async () => {
    const { pnpService, itemId } = this.props;
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.AUTHORITY.SUMMARY,
      Consts.FIELDS.AUTHORITY.ADDRESS,
      Consts.FIELDS.AUTHORITY.GROUPING,
      Consts.FIELDS.AUTHORITY.COMMENTS,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      Consts.FIELDS.AUTHORITY.PORTAL,
      Consts.FIELDS.COMMON.CREATION_TIME,
      Consts.FIELDS.COMMON.MODIFICATION_TIME,
      `${Consts.FIELDS.AUTHORITY.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.AUTHORITY.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.AUTHORITY.ETU_OWNERS}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.AUTHORITY.ETU_OWNERS}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.AUTHORITY.TaxCatchAll}/${Consts.FIELDS.AUTHORITY.Term}`,
      `${Consts.FIELDS.AUTHORITY.TaxCatchAll}/${Consts.FIELDS.AUTHORITY.Id}`,
      Consts.FIELDS.COMMON.DOCUMENTS_SPACE
    ];
    const expand = [Consts.FIELDS.AUTHORITY.COUNTRY, Consts.FIELDS.AUTHORITY.ETU_OWNERS, Consts.FIELDS.AUTHORITY.TaxCatchAll];
    const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.Authority).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), list);
    try {
      const item = await pnpService.getListByUrl(listUrl).
        items.getById(itemId)
        .select(...selectedFields)
        .expand(...expand)();



      const ContentType: IContentType = this.typeChoices
        .map((type) => ({ Id: type.key + "", Name: type.text }))
        .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1);

      // Map TaxCatchAll to Grouping with proper term name
      const taxCatchAllData = item[Consts.FIELDS.AUTHORITY.TaxCatchAll];
      const groupingTerm = taxCatchAllData && taxCatchAllData.length > 0 ?
        (this.getTaxonomyValue(item, Consts.FIELDS.AUTHORITY.GROUPING) ?? item[Consts.FIELDS.AUTHORITY.GROUPING]) :
        item[Consts.FIELDS.AUTHORITY.GROUPING];

      const authority: IAuthority = {
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
        Grouping: groupingTerm,
        Portal: item[Consts.FIELDS.AUTHORITY.PORTAL]?.Url?.trim() || '',
        PortalDescription: item[Consts.FIELDS.AUTHORITY.PORTAL]?.Description?.trim() || '',
        Created: item[Consts.FIELDS.COMMON.CREATION_TIME],
        Modified: item[Consts.FIELDS.COMMON.MODIFICATION_TIME],
        DocumentSpace: item[Consts.FIELDS.COMMON.DOCUMENTS_SPACE]?.Url
      };
      return authority;
    } catch (error) {
      if (error instanceof Error && error.message.includes(strings.DataSheetItemDoesNotExist)) {
        const errorMessage = strings.DataSheetItemDoesNotExist;
        this.setState({ errorMessage });
        return null;
      }
      throw error;
    }
  };


  public initMAReqViewItems = async () => {
    const { pnpService, itemId } = this.props;
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
    const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.MARequirement).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const items = await this.props.pnpService.getListItems(listUrl).
      select(...selectedFields)
      .expand(...expand)
      .filter(`${Consts.FIELDS.MAGENERIC.AUTHORITY}/Id eq ${itemId}`)
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
        case strings.MARequirementDatasheetTypeEutelsat:
          if (item?.[Consts.FIELDS.MAGENERIC.EUTELSATENTITY]?.[Consts.FIELDS.COMMON.ID] && item?.[Consts.FIELDS.MAGENERIC.EUTELSATENTITY]?.[Consts.FIELDS.COMMON.TITLE]) {
            responsiblePartyItem = {
              id: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID],
              key: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.EutelsatEntity,
              text: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.TITLE],
              type: DatasheetType.EutelsatEntity,
            };
            this.responsibleItemChoices.push(responsiblePartyItem);
          }
          break;
        case strings.MARequirementDatasheetTypeTeleportPartner:
          if (item?.[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER]?.[Consts.FIELDS.COMMON.ID] &&
            item?.[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER]?.[Consts.FIELDS.COMMON.TITLE]) {
            responsiblePartyItem = {
              id: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
              key: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.TeleportPartner,
              text: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
              type: DatasheetType.TeleportPartner,
            };
            this.responsibleItemChoices.push(responsiblePartyItem);
          }
          break;
        case strings.MARequirementDatasheetTypeDistributionPartner:
          if (item?.[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.ID] && item?.[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.TITLE]) {
            responsiblePartyItem = {
              id: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
              key: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.DistributionPartner,
              text: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE],
              type: DatasheetType.DistributionPartner,
            };
            this.responsibleItemChoices.push(responsiblePartyItem);
          }
          break;
        default:
          break;
      }
      const Type: IContentType = this.maReqTypeChoices
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
        GeoLeo: item[Consts.FIELDS.MAGENERIC.GEOLEO],
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

    // try {
    const filterQuery = `${Consts.FIELDS.CONTACT.AUTHORITY_ID} eq ${itemId}`;
    const contacts = await pnpService
      .getListItems(contactList).select(...selectedFields)
      .filter(filterQuery)();

    //const contactDetails: any[] = [];
    const contactDetails: IContact[] = [];
    for (const contact of contacts) {
      contactDetails.push({
        Id: contact[Consts.FIELDS.COMMON.ID],
        JobTitle: contact[Consts.FIELDS.CONTACT.JOB_TITLE],
        FirstName: contact[Consts.FIELDS.CONTACT.FIRST_NAME],
        LastName: contact[Consts.FIELDS.CONTACT.LAST_NAME],
        Department: contact[Consts.FIELDS.CONTACT.DEPARTMENT],
        Email: contact[Consts.FIELDS.CONTACT.EMAIL],
        PhoneNumber: contact[Consts.FIELDS.CONTACT.PHONE_NUMBER],
        Summary: contact[Consts.FIELDS.CONTACT.SUMMARY],
      });
    }
    //const updatedItem = { ...item, Contacts: contactDetails };
    //this.setState({ item: updatedItem });
    return contactDetails;

  };

  private initFees = async () => {
    const { itemId } = this.props;


    if (!itemId) {
      return;
    }



    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.FEE.SUMMARY,
      Consts.FIELDS.FEE.STATUS,
      Consts.FIELDS.FEE.COMMENTS,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      Consts.FIELDS.FEE.CATEGORY,
      Consts.FIELDS.FEE.VATRATE,
      Consts.FIELDS.FEE.DUEDATE,
      Consts.FIELDS.FEE.VATFREECOST,
      Consts.FIELDS.FEE.VAT,
      Consts.FIELDS.FEE.COSTTYPE,
      Consts.FIELDS.FEE.PONONPO,
      Consts.FIELDS.FEE.CHARGEDBY,
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
      `${Consts.FIELDS.FEE.RECURRENCEPATTERN}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.FEE.RECURRENCEPATTERN}/${Consts.FIELDS.COMMON.TITLE}`
    ];

    const expand = [
      Consts.FIELDS.FEE.AUTHORITY,
      Consts.FIELDS.FEE.CURRENCY,
      Consts.FIELDS.FEE.LEGALSERVICEPROVIDER,
      Consts.FIELDS.FEE.SNP,
      Consts.FIELDS.FEE.MA_REQUIREMENT,
      Consts.FIELDS.FEE.TELEPORTPARTNER,
      Consts.FIELDS.FEE.RECURRENCEPATTERN
    ];

    const filterQuery = `${Consts.FIELDS.FEE.AUTHORITY_ID} eq ${itemId}`;
    const list = Consts.LISTS.FEE_URL;
    const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), list);
    const items = await this.props.pnpService.getListItems(listUrl).
      select(...selectedFields)
      .expand(...expand)
      .filter(filterQuery)
      .orderBy(Consts.FIELDS.COMMON.TITLE)();

    const feeViewItems: IFEE[] = [];
    items.map((item) => {
      const rawDueDate = item[Consts.FIELDS.FEE.DUEDATE];
      const DueDate: Date | null = rawDueDate ? new Date(rawDueDate) : null;
      const ContentType: IContentType = this.FeesType
        .map((type) => ({ Id: type.key + "", Name: type.text }))
        .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1) || { Id: '', Name: '' };
      let itemValue = null;

      if (ContentType.Id === Consts.CONTENT_TYPES.SNP_FEE && item[Consts.FIELDS.FEE.SNP] && item[Consts.FIELDS.FEE.SNP][Consts.FIELDS.COMMON.ID]) {
        itemValue = {
          id: item[Consts.FIELDS.FEE.SNP]?.[Consts.FIELDS.COMMON.ID],
          key: item[Consts.FIELDS.FEE.SNP]?.[Consts.FIELDS.COMMON.ID] + "-" + ViewType.SNP,
          text: item[Consts.FIELDS.FEE.SNP]?.[Consts.FIELDS.COMMON.TITLE],
          type: ViewType.SNP
        };
        this.itemChoices.push({
          id: itemValue?.id,
          key: itemValue?.key,
          text: itemValue?.text,
          type: itemValue?.type
        });
      } else if (ContentType.Id === Consts.CONTENT_TYPES.MA_REQUIREMENT_FEE && item[Consts.FIELDS.FEE.MA_REQUIREMENT] && item[Consts.FIELDS.FEE.MA_REQUIREMENT][Consts.FIELDS.COMMON.ID]) {
        itemValue = {
          id: item[Consts.FIELDS.FEE.MA_REQUIREMENT]?.[Consts.FIELDS.COMMON.ID],
          key: item[Consts.FIELDS.FEE.MA_REQUIREMENT]?.[Consts.FIELDS.COMMON.ID] + "-" + ViewType.MARequirement,
          text: item[Consts.FIELDS.FEE.MA_REQUIREMENT]?.[Consts.FIELDS.COMMON.TITLE],
          type: ViewType.MARequirement
        };
        this.itemChoices.push({
          id: itemValue.id,
          key: itemValue?.key,
          text: itemValue?.text,
          type: itemValue?.type
        });
      }


      let chargeByItem = null;
      switch (item[Consts.FIELDS.FEE.CHARGEDBY]) {
        case FeeChargedBy.LawFirm:
        case FeeChargedBy.LegalRepresentative:
          chargeByItem = {
            id: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID],
            key: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID] +
              "-" + ViewType.LegalRepresentative,
            text: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.TITLE],
            type: ViewType.LegalRepresentative
          };
          this.ChargeByItemChoices.push(chargeByItem);
          break;
        case FeeChargedBy.Administration:
        case FeeChargedBy.Organization:
        case FeeChargedBy.Regulator:
          chargeByItem = {
            id: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID],
            key: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID] + "-" + ViewType.Authority,
            text: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.TITLE],
            type: ViewType.Authority,
          };
          this.ChargeByItemChoices.push(chargeByItem);
          break;
        case FeeChargedBy.TeleportPartner:
          chargeByItem = {
            id: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
            key: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + ViewType.TeleportPartner,
            text: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
            type: ViewType.TeleportPartner,
          };
          this.ChargeByItemChoices.push(chargeByItem);
          break;
        default:
          break;
      }

      const vatFreeCost = parseInt(item[Consts.FIELDS.FEE.VATFREECOST]);
      const vat = parseInt(item[Consts.FIELDS.FEE.VAT]);
      const total = (!isNaN(vatFreeCost) ? vatFreeCost : 0) + (!isNaN(vat) ? vat : 0);
      const currencyCode = item[Consts.FIELDS.FEE.CURRENCY]?.[Consts.FIELDS.FEE.CURRENCY_ISOCODE] ?? '';
      feeViewItems.push({
        Id: item[Consts.FIELDS.COMMON.ID],
        Title: item[Consts.FIELDS.COMMON.TITLE],
        Category: item[Consts.FIELDS.FEE.CATEGORY],
        ContentType,
        Item: itemValue,
        Summary: item[Consts.FIELDS.FEE.SUMMARY],
        VatFreeCost: item[Consts.FIELDS.FEE.VATFREECOST],
        Vat: item[Consts.FIELDS.FEE.VAT],
        Status: item[Consts.FIELDS.FEE.STATUS],
        CostTotal: total ? total.toString() + " " + currencyCode : '0',
        DueDate,
        ChargeByItem: chargeByItem
      });
    });
    //this.setState({ feeViewItems });
    return feeViewItems;
  };

  private initPayments = async () => {
    const { itemId } = this.props;
    // const { item } = this.state;

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
    const list = Consts.LISTS.PAYMENT_URL;
    const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), list);
    const filterQuery = `${Consts.FIELDS.PAYMENT.AUTHORITY_ID} eq ${itemId}`;
    const items = await this.props.pnpService.getListItems(listUrl).
      select(...selectedFields)
      .expand(...expand)
      .filter(filterQuery)
      .orderBy(Consts.FIELDS.COMMON.TITLE)();

    const paymentViewItems: IPayment[] = [];
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
              "-" + ViewType.LegalRepresentative,
            text: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.TITLE],
            type: ViewType.LegalRepresentative
          };
          this.beneficiaryItemChoices.push(beneficiaryItem);
          break;
        case FeeChargedBy.Administration:
        case FeeChargedBy.Organization:
        case FeeChargedBy.Regulator:
          beneficiaryItem = {
            id: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID],
            key: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID] + "-" + ViewType.Authority,
            text: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.TITLE],
            type: ViewType.Authority,
          };
          this.beneficiaryItemChoices.push(beneficiaryItem);
          break;
        case FeeChargedBy.TeleportPartner:
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

      paymentViewItems.push({
        Id: item[Consts.FIELDS.COMMON.ID],
        Title: item[Consts.FIELDS.COMMON.TITLE],
        Fees: item[Consts.FIELDS.PAYMENT.FEES]?.map((fee) => ({
          Id: fee[Consts.FIELDS.COMMON.ID],
          Title: fee[Consts.FIELDS.COMMON.TITLE]
        })) || [],
        Summary: item[Consts.FIELDS.PAYMENT.SUMMARY],
        VatFreeCost: item[Consts.FIELDS.PAYMENT.VATFREECOST],
        Vat: item[Consts.FIELDS.PAYMENT.VAT],
        Status: item[Consts.FIELDS.PAYMENT.STATUS],
        CostTotal: total ? total.toString() + " " + currencyCode : '0',
        DueDate,
        DatePaid,
        Item: beneficiaryItem,


      });
    });
    //this.setState({ payments: viewItems });
    return paymentViewItems;
  };

  private init = async () => {
    let item: IAuthority;
    let maReqItems: IMARequirement[];
    const promises: Promise<void>[] = [];
    let initFeesItems: IFEE[];
    let initPaymentItems: IPayment[];
    let contactItems: IContact[];
    if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner, UserRole.Admin)) {
      promises.push(this.initFees().then((opts) => { initFeesItems = opts as IFEE[]; }));
      promises.push(this.initPayments().then((opts) => { initPaymentItems = opts as IPayment[]; }));
    }
    promises.push(this.initItem().then((opts) => { item = opts as IAuthority; }));
    promises.push(this.initMAReqViewItems().then((opts) => { maReqItems = opts as IMARequirement[]; }));
    promises.push(this.fetchContactDetails().then((opts) => { contactItems = opts as IContact[]; }));
    await Promise.all(promises);


    this.setState({
      item,
      maReqViewItems: maReqItems,
      feeViewItems: initFeesItems,
      paymentViewItems: initPaymentItems,
      contactViewItems: contactItems
    });
    console.log('item', item);

    // ((history) => {
    //   const pushState = history.pushState;
    //   history.pushState = (state, key, path) => {
    //     pushState.apply(history, [state, key, path]);
    //     this._onUrlChange(typeof path === 'string' ? path : path.toString());
    //   };
    // })(window.history);

    // window.addEventListener('popstate', (e) => {
    //   // Currently browsing by the browser history buttons ( back / forward )
    //   // doesn't cause any effect on a sp conditionally loaded page.
    //   this._onUrlChange();

    // });
  };
  // private _onUrlChange(newpath = window.location.search): void {
  //   // any logic you might want to trigger when the query string updates
  //   // e.g. fetching data
  //   // e.g. logging the URL changes // console.log(newpath)
  //   //this.render();
  //   this.init();
  // }
  private renderDatasheetInfo = (): JSX.Element | null => {
    const leftColumns = ['Summary', 'Address', 'Grouping', 'Countries'];
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
    const labelKey = `AuthorityDataSheetInfoLabel${key}`;
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
      return (
        <div className={styles.termPills}>
          <span key={value?.TermGuid} className={styles.termPill}>
            <span className={styles.termPillText}>{value?.Label}</span>
          </span>
        </div>
      );
    }

    if (fieldName === 'Portal') {
      return (
        <a data-interception="on"
          href={value} rel="noreferrer" >{item?.PortalDescription}</a>
      );
    }
    if (fieldName === 'Countries') {
      const countryListType = DatasheetType.Country;
      return renderObjectPills(
        value,
        (country) => (
          <span key={country.Id} className={styles.termPill}>
            <span className={styles.termPillText}>
              <a
                data-interception="on"
                rel="noreferrer"
                onClick={(e) => {
                  e.preventDefault();
                  const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(countryListType)}&itemId=${country?.Id}`;
                  UrlHelper.navigate(url, false);
                }}
              >
                {country?.Title}
              </a>
            </span>
          </span>
        )
      );
    }

    if (fieldName === 'Owners') {
      const countryListType = DatasheetType.Contact;
      return renderObjectPills(
        value,
        (owner) => (
          <span key={owner.Id} className={styles.termPill}>
            <span className={styles.termPillText}>
              <a
                data-interception="off"
                rel="noreferrer"
                onClick={(e) => {
                  e.preventDefault();
                  const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(countryListType)}&itemId=${owner?.Id}`;
                  UrlHelper.navigate(url, false);
                }}
              >
                {owner?.Title}
              </a>
            </span>
          </span>
        )
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

  public render(): React.ReactElement<IAuthorityDatasheetProps> {
    const {
      showEditFormDialog,
      showDeleteDialog,
      item,
      maReqViewItems,
      feeViewItems,
      paymentViewItems,
      errorMessage
    } = this.state;
    const { pnpService, userService } = this.props;
    const created = UtilHelper.formatDate(item?.Created, 'LL', 'en-us');
    const modified = UtilHelper.formatDate(item?.Modified, 'LL', 'en-us');
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
          <>
            <div>
              <div className={styles.datasheetHeader}>
                <div className={styles.datasheetTypeContainer}>
                  < h5 className={styles.datasheetType} >
                    {item?.ContentType.Name}
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
                      relatedItemType={item?.ContentType?.Name}
                    />
                  </PivotItem>
                  {this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner, UserRole.Admin) && <PivotItem headerText="Fees & Payments" itemIcon="Bank">
                    <div className={styles.section}>
                      <Icon iconName="Money" className="sectionIcon" />
                      <h3>Fees</h3>
                    </div>
                    <FeeView
                      pnpService={pnpService}
                      userService={userService}
                      columns={['Name', 'Summary', 'ContentType', 'Item', 'CostTotal', 'Status', 'DueDate']}
                      filters={['Search', 'ContentType', 'Status', 'Item']}
                      viewItems={feeViewItems}
                      pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}
                    />
                    <div className={styles.section}>
                      <Icon iconName="PaymentCard" className="sectionIcon" />
                      <h3>Payments</h3>
                    </div>
                    <PaymentView
                      pnpService={pnpService}
                      userService={userService}
                      columns={['Title', 'Summary', 'Fees', 'CostTotal', 'Status', 'DueDate', 'DatePaid']}
                      filters={['Search', 'Fees', 'Status']}
                      viewItems={paymentViewItems}
                      pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}
                    />
                  </PivotItem>}
                  <PivotItem headerText="MA Requirements" itemIcon="Bullseye" >
                    <MARequirementView pnpService={pnpService}
                      userService={userService}
                      columns={['Title', 'Markets', 'MaVertical', 'Type', 'ResponsiblePartyItem', 'RagStatus', 'Synthesis', 'Geo/Leo', 'ApplicationDate', 'EffectiveDate', 'ExpirationDate']}
                      filters={['Search', 'MaVertical', 'Type', 'Markets', 'ResponsiblePartyItem', 'RagStatus', 'Synthesis', 'Geo/Leo']}
                      viewItems={maReqViewItems}
                      pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}
                    ></MARequirementView>
                  </PivotItem>
                </Pivot>
              </div>
            </div>
          </> :
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
      const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.Authority).list;
      const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
      await pnpService.getListItems(listUrl).getById(itemId).recycle();
      const messageBanner: IMessageBanner = {
        message: strings.DeleteSuccessMessage,
        type: 'success',
        visible: true
      };
      this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
      //setTimeout(() => { callback(window.location); }, 2000);
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

    const editFormControl = (<Authority pnpService={pnpService} userService={userService} itemId={itemId} callback={async () => { await this.closeEditForm(true); }}></Authority>);
    const editFormTitle = strings.AuthorityDataSheetEditFormTitle;

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
  (
    <div>
      <Dialog
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
              onClick={this.deleteDataSheet} >
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
      </Dialog>

    </div>);
}
