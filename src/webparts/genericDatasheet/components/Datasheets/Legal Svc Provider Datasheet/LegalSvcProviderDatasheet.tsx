import * as React from 'react';
import { ILawFirmDatasheetProps, ILawFirmDatasheetState } from './LegalSvcProviderDatasheet.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import {
  ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, IDropdownOption, ITooltipHostStyles, Label,
  Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost
} from '@fluentui/react';
import LegalSvcProvider from '../../../../genericForm/components/Legal Svc Provider/LegalSvcProvider';
import { Datasheet, DatasheetType } from '../../Datasheet.types';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { IContact, IContentType, ICountry, IFEE, ILegalSvcProvider, IPayment } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { renderArrayPills, renderObjectPills } from '../../Datasheet.utility';
import { RichText } from '../../../../../common/components/RichText/RichText';
import { StatusPill } from '../../../../genericView/components/View.utility';
import ContactDatasheetView from '../../Datasheet Views/ContactDatasheetView';
import FeeView from '../../../../genericView/components/Fees View/FeesView';
import PaymentView from '../../../../genericView/components/Payment View/PaymentView';
import { IViewConfig, View, ViewType } from '../../../../genericView/components/View.types';
import { IItemOption } from '../../../../genericView/components/Fees View/FeesView.type';
import LibraryView from '../../../../genericView/components/Library View/LibraryView';
import { FeeChargedBy, FeeType, LegalSvcProviderType, UserRole } from '../../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../../common/components/Access Denied/AccessDenied';
import { IMessageBanner } from '../../../../genericForm/components/Form.types';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';
import DomHelper from '../../../../../common/helpers/DomHelper';




export default class LegalSvcProviderDatasheet extends React.Component<
  ILawFirmDatasheetProps,
  ILawFirmDatasheetState
> {
  private typeChoices: IDropdownOption[] = [
    { key: Consts.CONTENT_TYPES.SNP_FEE, text: FeeType.SNP },
    { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_FEE, text: FeeType.MATypeRequirement }
  ];
  private itemChoices: IItemOption[] = [];
  private ChargeByItemChoices: IItemOption[] = [];

  //Payments
  private feesConfig: IViewConfig;
  private paymentConfig: IViewConfig;
  private beneficiaryItemChoices: IItemOption[] = [];

  constructor(props: Readonly<ILawFirmDatasheetProps>) {
    super(props);
    this.state = {
      isDatasheetReady: false,
      showEditFormDialog: false,
      showDeleteDialog: false,
      item: null,
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
    this.paymentConfig = View.config.find((f) => f.type.toLowerCase() === ViewType.Payment);
    this.feesConfig = View.config.find((f) => f.type.toLowerCase() === ViewType.FEE);
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
      Consts.FIELDS.LEGAL_SVC_PROVIDER.SUMMARY,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.ADDRESS,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.STATUS,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.COMMENTS,
      Consts.FIELDS.COMMON.CREATION_TIME,
      Consts.FIELDS.COMMON.MODIFICATION_TIME,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY_ID,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNERS_ID,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      Consts.FIELDS.COMMON.DOCUMENTS_SPACE,
      Consts.FIELDS.COMMON.DOCUMENTS_SPACE,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER}/${Consts.FIELDS.COMMON.TITLE}`
    ];



    const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.LegalSvcProvider).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    try {
      const item = await pnpService
        .getListByUrl(listUrl)
        .items.getById(itemId)
        .select(...selectedFields)
        .expand(Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY, Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER)();
      const contentTypeId = item?.[Consts.FIELDS.COMMON.CONTENTTYPE_ID];
      const contentTypeName = contentTypeId && contentTypeId.includes(Consts.CONTENT_TYPES.LAW_FIRM) ?
        LegalSvcProviderType.LawFirm :
        contentTypeId && contentTypeId.includes(Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE) ?
          LegalSvcProviderType.LegalRepresentative : '';

      const legalSvcProvider: ILegalSvcProvider = {
        Id: item[Consts.FIELDS.COMMON.ID],
        ContentTypeName: contentTypeName,
        Title: item[Consts.FIELDS.COMMON.TITLE],
        Summary: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.SUMMARY],
        Address: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.ADDRESS],
        Status: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.STATUS],
        Comments: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.COMMENTS],
        Countries: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY] || [],
        Owners: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER] || [],
        Created: item[Consts.FIELDS.COMMON.CREATION_TIME],
        Modified: item[Consts.FIELDS.COMMON.MODIFICATION_TIME],
        DocumentSpace: item[Consts.FIELDS.COMMON.DOCUMENTS_SPACE]?.Url
      };
      return legalSvcProvider;
    } catch (error) {
      if (error instanceof Error && error.message.includes(strings.DataSheetItemDoesNotExist)) {
        const errorMessage = strings.DataSheetItemDoesNotExist;
        this.setState({ errorMessage });
        return null;

      }
      throw error;
    }
  };

  private init = async () => {
    let item: ILegalSvcProvider;
    let contactItems: IContact[];
    let feesItems: IFEE[];
    let paymentsItems: IPayment[];
    const promises: Promise<void>[] = [];
    if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner, UserRole.Admin)) {
      promises.push(this.fetchFees().then((opts) => { feesItems = opts as IFEE[]; }));
      promises.push(this.fetchPayments().then((opts) => { paymentsItems = opts as IPayment[]; }));
    }
    promises.push(this.initItem().then((opts) => { item = opts as ILegalSvcProvider; }));
    promises.push(this.fetchContactDetails().then((opts) => { contactItems = opts as IContact[]; }));
    await Promise.all(promises);
    this.setState({
      item,
      contactViewItems: contactItems,
      feeViewItems: feesItems,
      paymentViewItems: paymentsItems
    });


  };

  private fetchFees = async () => {
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

    const filterQuery = `${Consts.FIELDS.FEE.LEGALSERVICEPROVIDER_ID} eq ${itemId}`;
    const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.feesConfig.list);
    const items = await this.props.pnpService.getListItems(listUrl).
      select(...selectedFields)
      .expand(...expand)
      .filter(filterQuery)
      .orderBy(Consts.FIELDS.COMMON.TITLE)();

    const feeViewItems: IFEE[] = [];
    items.map((item) => {
      const rawDueDate = item[Consts.FIELDS.FEE.DUEDATE];
      const DueDate: Date | null = rawDueDate ? new Date(rawDueDate) : null;
      const ContentType: IContentType = this.typeChoices
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
        case strings.ViewFeeTypeTeleportPartner:
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

  private fetchPayments = async () => {
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
    const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.paymentConfig.list);
    const filterQuery = `${Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER_ID} eq ${itemId}`;
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
        Item: beneficiaryItem
      });
    });
    //this.setState({ payments: viewItems });
    return paymentViewItems;
  };


  private renderDatasheetInfo = (): JSX.Element | null => {
    const leftColumns = ['Summary', 'Address', 'Countries'];
    const rightColumns = ['Status', 'Owners', 'Comments'];

    return (
      <div className={`${styles.datasheetInfoContainer}`}>
        <div className={styles.datasheetInfoContainerLeft}>
          {leftColumns.map((key) => this.renderDatasheetInfoRow(key))}
        </div>
        <div className={styles.datasheetInfoContainerRight}>
          {rightColumns.map((key) => this.renderDatasheetInfoRow(key))}
        </div>
      </div>
    );
  };

  private renderDatasheetInfoRow = (key: string) => {
    const { item } = this.state;
    const labelKey = `LegalSvcProviderDataSheet${key}`;

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

  private renderDatasheetInfoValue = (item: ILegalSvcProvider, fieldName: string) => {
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

    if (fieldName === 'Status') {
      return StatusPill(String(value));
    }

    if (fieldName === 'Countries') {
      return renderObjectPills(
        value as ICountry[],
        (item: ICountry) => (
          <span className={styles.termPill}>
            <span className={styles.termPillText}>
              <a
                data-interception="on"
                rel="noreferrer"
                onClick={(e) => {
                  e.preventDefault();
                  const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Country)}&itemId=${item.Id}`;
                  UrlHelper.navigate(url, false);
                }}
              >{item.Title}</a>
            </span>
          </span>
        )
      );
    }

    if (fieldName === 'Owners') {
      return renderObjectPills(
        value as IContact[],
        (item: IContact) => (
          <span className={styles.termPill}>
            <span className={styles.termPillText}>
              <a
                data-interception="on"
                rel="noreferrer"
                onClick={(e) => {
                  e.preventDefault();
                  const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Contact)}&itemId=${item.Id}`;
                  UrlHelper.navigate(url, false);
                }}
              >{item.Title}</a>
            </span>
          </span>
        )
      );
    }

    if (typeof value === 'string' || typeof value === 'number') {
      return <span className={styles.detailsListValue}>{String(value)}</span>;
    }

    if (Array.isArray(value) && value?.every((v) => typeof v === 'string' || typeof v === 'number')) {
      return renderArrayPills(value as (string | number)[]);
    }

    return <span className={styles.detailsListValue}>{String(value)}</span>;
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
    const filterQuery = `${Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER_ID} eq ${itemId}`;
    const contacts = await pnpService
      .getListItems(contactList)
      .select(...selectedFields)
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

  public render(): React.ReactElement<ILawFirmDatasheetProps> {
    const {
      showEditFormDialog,
      showDeleteDialog,
      item,
      errorMessage
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
    if (errorMessage) {
      if (errorMessage.includes(strings.DataSheetItemDoesNotExist)) {
        return (
          <div className={styles.noDataMessage}>
            <Icon iconName="Remove" className={styles.mutedIcon} />
            <span>{strings.DataSheetNoRecordFound || 'No record found'}</span>
            <Icon iconName="Remove" className={styles.mutedIcon} />
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
                  {item?.ContentTypeName}
                </h5>
              </div>
              {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.datasheetCommandsContainer}>
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
                  <div className={styles.section}>
                    <Icon iconName="People" className="sectionIcon" />
                    <h3 >{strings.DataSheetSectionContacts}</h3>
                  </div>
                  <ContactDatasheetView
                    contacts={this.state.contactViewItems}
                    pnpService={pnpService}
                    userService={userService}
                    refresh={() => this.init()}
                    relatedItemId={item?.Id}
                    relatedItemType={item?.ContentTypeName}
                  />
                </PivotItem>
                {this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner, UserRole.Admin) && <PivotItem headerText="Fees & Payments" itemIcon="Bank">
                  <div className={styles.section}>
                    <Icon iconName="Money" className="sectionIcon" />
                    <h3>Fees</h3>
                  </div>
                  <FeeView
                    pnpService={this.props.pnpService} userService={userService}
                    columns={['Name', 'Summary', 'ContentType', 'Item', 'CostTotal', 'Status', 'DueDate']}
                    filters={['Search', 'ContentType', 'Status', 'Item']}
                    viewItems={this.state.feeViewItems}
                    pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}
                  />
                  <div className={styles.section}>
                    <Icon iconName="PaymentCard" className="sectionIcon" />
                    <h3>Payments</h3>
                  </div>
                  <PaymentView
                    pnpService={this.props.pnpService} userService={userService}
                    columns={['Title', 'Summary', 'Fees', 'CostTotal', 'Status', 'DueDate', 'DatePaid']}
                    filters={['Search', 'Fees', 'Status']}
                    viewItems={this.state.paymentViewItems}
                    pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}
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
      const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.LegalSvcProvider).list;
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
    const { showEditFormDialog } = this.state;
    const { pnpService, userService, itemId } = this.props;

    const editFormControl = (
      <LegalSvcProvider
        pnpService={pnpService}
        userService={userService}
        itemId={itemId}
        callback={async () => {
          await this.closeEditForm(true);
        }}
      />
    );
    const editFormTitle = strings.LegalSvcProviderDatasheetEditLawFirm;

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
