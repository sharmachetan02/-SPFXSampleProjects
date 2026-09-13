import * as React from 'react';

import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { IPaymentProps, IPaymentState } from './Payment.type';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import { Dropdown, Fabric, Icon, IDropdownOption, Label, TextField } from 'office-ui-fabric-react';
import { Combobox } from 'react-widgets';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { Form, FormType, IFormConfig } from '../Form.types';
import strings from 'GenericFormWebPartStrings';
import { RichText } from '../../../../common/components/RichText/RichText';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { UtilHelper } from '../../../../common/helpers/Util';
import { Consts } from '../../../../common/consts/Consts';
import 'react-widgets/styles.css';
import { IItemAddResult } from '@pnp/sp/items';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { DatePicker } from '@fluentui/react';
import { createUniqueFolder, datepickerstyles } from '../Form.utility';
import { FeeChargedBy, UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';

export default class Fee extends React.Component<IPaymentProps, IPaymentState> {
  private isEditMode: boolean;
  private config: IFormConfig;
  private AllbeneficiaryItems: Map<string, IDropdownOption[]> = new Map();
  private AllContactsItems: Map<string, IDropdownOption[]> = new Map();

  constructor(props: Readonly<IPaymentProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      isFormReady: false,
      name: '',
      summary: '',
      fees: [],
      feesChoices: [],
      beneficiary: '',
      beneficiaryChoices: [],
      beneficiaryItem: null,
      beneficiaryItemChoices: [],
      currency: null, // updated to null
      currencyChoices: [],
      contacts: [],
      contactsChoices: [],
      vatFreeCost: '',
      vat: '',
      vatRate: '',
      status: '',
      statusChoices: [],
      dueDate: null,
      datePaid: null,
      comments: '',
      errors: {}
    };

    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.Payment);
    this.AllbeneficiaryItems = new Map();
    this.AllContactsItems = new Map();
  }


  private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
    const userRoles = this.props.userService.userContext?.userRoles ?? [];
    return rolesToCheck.some((role) => userRoles.includes(role));
  };
  public async componentDidMount() {
    if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin)) {
      await this.init();
      this.setState({ isFormReady: true });
    }
  }

  private async init() {
    // Wait for AllbeneficiaryItems to be populated first
    await this.fetchAllbeneficiaryItems();
    await this.fetchAllContactsItems();
    const promises = [];
    promises.push(
      this.initFeesChoices()
    );
    promises.push(
      this.initFeebeneficiaryChoices()
    );
    promises.push(
      this.initCurrencyChoices()
    );
    promises.push(
      this.initStatusChoices()
    );
    promises.push(
      this.initFee()
    );

    await Promise.all(promises);

    if (this.isEditMode)
      await this.setInitialFormValues();
    else
      this.initFee();
  }

  private initFee = async () => {
    const {
      preselectedItemId
    } = this.props;
    const {
      feesChoices
    } = this.state;
    const fees = preselectedItemId ? [feesChoices.find((c) => c.key === preselectedItemId)] : [];
    this.setState({
      fees
    });

  };


  private fetchAllbeneficiaryItems = async () => {
    const { pnpService } = this.props;
    const selectedFields = [Consts.FIELDS.COMMON.ID, Consts.FIELDS.COMMON.TITLE];

    const configMap = [
      { key: FeeChargedBy.Administration, type: FormType.Authority, contentTypeId: Consts.CONTENT_TYPES.ADMINISTRATION },
      { key: FeeChargedBy.Organization, type: FormType.Authority, contentTypeId: Consts.CONTENT_TYPES.ORGANIZATION },
      { key: FeeChargedBy.Regulator, type: FormType.Authority, contentTypeId: Consts.CONTENT_TYPES.REGULATOR },
      { key: FeeChargedBy.LawFirm, type: FormType.LegalServiceProvider, contentTypeId: Consts.CONTENT_TYPES.LAW_FIRM },
      { key: FeeChargedBy.LegalRepresentative, type: FormType.LegalServiceProvider, contentTypeId: Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE },
      { key: FeeChargedBy.TeleportPartner, type: FormType.TP, contentTypeId: null }
    ];

    const fetchPromises = configMap.map(async (config) => {
      const list = Form.config.find((f) => f.type.toLowerCase() === config.type.toLowerCase())?.list;
      const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
      let query = pnpService.getListItems(listUrl).select(...selectedFields).expand(Consts.FIELDS.COMMON.CONTENTTYPE);
      if (config.contentTypeId) {
        query = query.filter(`startswith(${Consts.FIELDS.COMMON.CONTENTTYPE_ID}, '${config.contentTypeId}')`);
      }
      const items = await query.orderBy(Consts.FIELDS.COMMON.TITLE, true)();
      const choices = items.map((item) => ({
        key: item[Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.COMMON.TITLE]
      }));
      return { key: config.key, choices };
    });

    const results = await Promise.all(fetchPromises);
    results.forEach(({ key, choices }) => {
      this.AllbeneficiaryItems.set(key, choices);
    });
  };


  private initFeesChoices = async () => {
    const {
      pnpService
    } = this.props;
    const feesChoices: IDropdownOption[] = [];
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Fee).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      feesChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    if (!this.isEditMode) {
      const fees = [];
      this.setState({ fees, feesChoices });
    } else {
      this.setState({ feesChoices });
    }
  };


  private initFeebeneficiaryChoices = async () => {

    const {
      pnpService
    } = this.props;
    let beneficiaryChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.PAYMENT.BENEFICIARY)
      .select('Choices')();
    beneficiaryChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    if (!this.isEditMode) {
      const beneficiary = beneficiaryChoices.length > 0 ? beneficiaryChoices[0].key as string : '';
      const beneficiaryItemChoices = await this.loadBeneficiaryItemChoices(beneficiary);
      this.setState({
        beneficiary, beneficiaryChoices, beneficiaryItemChoices
      });
      this.initContactChoices();

    } else {
      this.setState({
        beneficiaryChoices
      });
    }

  };

  private initCurrencyChoices = async () => {
    const {
      pnpService
    } = this.props;
    const currencyChoices: IDropdownOption[] = [];
    const list = Consts.LISTS.CURRENCY_URL;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      currencyChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    if (!this.isEditMode) {
      const currency = null;
      this.setState({ currency, currencyChoices });
    } else {
      this.setState({ currencyChoices });
    }
  };

  private fetchAllContactsItems = async () => {
    const { pnpService } = this.props;
    const selectedFields = [Consts.FIELDS.COMMON.ID, Consts.FIELDS.COMMON.TITLE,
    `${Consts.FIELDS.CONTACT.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
    `${Consts.FIELDS.CONTACT.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
    `${Consts.FIELDS.CONTACT.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
    `${Consts.FIELDS.CONTACT.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
    `${Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER}/${Consts.FIELDS.COMMON.ID}`,
    `${Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER}/${Consts.FIELDS.COMMON.TITLE}`,
    `${Consts.FIELDS.COMMON.CONTENTTYPE_ID}`
    ];

    const configMap = [
      { key: FeeChargedBy.TeleportPartner, contentTypeId: Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT },
      { key: FeeChargedBy.LawFirm, contentTypeId: Consts.CONTENT_TYPES.LAW_FIRM_CONTACT },
      { key: FeeChargedBy.LegalRepresentative, contentTypeId: Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE_CONTACT },
      { key: FeeChargedBy.Regulator, contentTypeId: Consts.CONTENT_TYPES.REGULATOR_CONTACT },
      { key: FeeChargedBy.Organization, contentTypeId: Consts.CONTENT_TYPES.ORGANIZATION_CONTACT },
      { key: FeeChargedBy.Administration, contentTypeId: Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT }
    ];

    const list = Consts.LISTS.CONTACT_URL;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const expand = [
      Consts.FIELDS.COMMON.CONTENTTYPE,
      Consts.FIELDS.CONTACT.TELEPORT_PARTNER,
      Consts.FIELDS.CONTACT.AUTHORITY,
      Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER,
    ];
    const query = pnpService.getListItems(listUrl).select(...selectedFields).expand(...expand);
    const items = await query.orderBy(Consts.FIELDS.COMMON.TITLE, true)();

    configMap.map(async (config) => {
      const filtredItems = items.filter(
        (item) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(config.contentTypeId) !== -1
      );
      const choices = filtredItems?.map((item) => ({
        key: item[Consts.FIELDS.COMMON.ID] + "",
        text: item[Consts.FIELDS.COMMON.TITLE],
        TPId: item[Consts.FIELDS.CONTACT.TELEPORT_PARTNER]?.[Consts.FIELDS.COMMON.ID] + "",
        AUTHId: item[Consts.FIELDS.CONTACT.AUTHORITY]?.[Consts.FIELDS.COMMON.ID] + "",
        SVCId: item[Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER]?.[Consts.FIELDS.COMMON.ID] + "",
      }));
      this.AllContactsItems.set(config.key, choices);
    });
    console.log('AllContactsItems', this.AllContactsItems);
  };

  private initContactChoices = async () => {

    const {
      beneficiary,

    } = this.state;
    const contactsChoices: IDropdownOption[] = this.AllContactsItems.get(beneficiary);
    //const contactAllData = this.AllContactsItems.get(beneficiary);

    // const filterKey = maRequirementBasicData?.responsiblePartyItem?.key;

    // let filteredContacts: IDropdownOption[] = [];
    // switch (maRequirementBasicData.responsibleParty) {
    //   case strings.FormMARequirementTypeEutelsat:
    //     break;
    //   case strings.FormMARequirementTypeTeleportPartner:
    //     filteredContacts = contactAllData?.filter((item) => item["TPId"] === filterKey);
    //     break;
    //   case strings.FormMARequirementTypeDistributionPartner:
    //     filteredContacts = contactAllData?.filter((item) => item["DPId"] === filterKey);
    //   // Filter the contacts by TPId matching the key

    // }
    // maRequirementBasicData.contactsChoices = filteredContacts;
    if (!this.isEditMode) {
      const contacts = [];
      this.setState({ contacts, contactsChoices });
    } else {
      this.setState({ contactsChoices });
    }
  };

  /**
   * Loads beneficiary item choices based on the beneficiary type.
   * Uses cached data from AllbeneficiaryItems map populated during initialization.
   */
  private loadBeneficiaryItemChoices = async (beneficiary: string): Promise<IDropdownOption[]> => {
    if (this.AllbeneficiaryItems.has(beneficiary)) {
      const beneficiaryItemChoices: IDropdownOption[] = this.AllbeneficiaryItems.get(beneficiary) || [];
      return beneficiaryItemChoices;
    }
    return [];
  };

  public async initStatusChoices() {
    const {
      pnpService
    } = this.props;
    let statusChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.PAYMENT.STATUS)
      .select('Choices')();
    statusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    const status = statusChoices.length > 0 ? statusChoices[0].key as string : '';
    this.setState({
      statusChoices, status
    });
  }

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Payment).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
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
      `${Consts.FIELDS.PAYMENT.FEES}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.PAYMENT.FEES}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.PAYMENT.CONTACTS}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.PAYMENT.CONTACTS}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const expand = [Consts.FIELDS.PAYMENT.AUTHORITY, Consts.FIELDS.PAYMENT.CURRENCY, Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER,
    Consts.FIELDS.PAYMENT.FEES, Consts.FIELDS.PAYMENT.CONTACTS, Consts.FIELDS.PAYMENT.TELEPORTPARTNER];
    const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();

    let currency = null;

    if (item[Consts.FIELDS.PAYMENT.CURRENCY]) {
      currency = {
        key: item[Consts.FIELDS.PAYMENT.CURRENCY][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.PAYMENT.CURRENCY][Consts.FIELDS.COMMON.TITLE]
      };
    }

    // Get beneficiary value and load corresponding choices
    const beneficiaryValue = item[Consts.FIELDS.PAYMENT.BENEFICIARY];
    const beneficiaryItemChoices: IDropdownOption[] = await this.loadBeneficiaryItemChoices(beneficiaryValue);

    // Get the ID of the beneficiary item from the appropriate field
    let beneficiaryItemId: number = null;
    switch (beneficiaryValue) {
      case FeeChargedBy.LawFirm:
      case FeeChargedBy.LegalRepresentative:
        beneficiaryItemId = item[Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID];
        break;
      case FeeChargedBy.Administration:
      case FeeChargedBy.Organization:
      case FeeChargedBy.Regulator:
        beneficiaryItemId = item[Consts.FIELDS.PAYMENT.AUTHORITY][Consts.FIELDS.COMMON.ID];
        break;
      case FeeChargedBy.TeleportPartner:
        beneficiaryItemId = item[Consts.FIELDS.PAYMENT.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID];
        break;
      default:
        break;
    }

    // Find the matching item from choices array to ensure correct reference
    const beneficiaryItem: IDropdownOption =
      beneficiaryItemId !== null ?
        beneficiaryItemChoices.find((c) => c.key === beneficiaryItemId) || null : null;

    const rawDueDate = item[Consts.FIELDS.PAYMENT.DUEDATE];
    const dueDate: Date | null = rawDueDate ? new Date(rawDueDate) : null;
    const rawDatePaid = item[Consts.FIELDS.PAYMENT.PAYMENTPAID];
    const datePaid: Date | null = rawDatePaid ? new Date(rawDatePaid) : null;

    const initialValues = {
      name: item[Consts.FIELDS.COMMON.TITLE],
      summary: item[Consts.FIELDS.PAYMENT.SUMMARY],
      comments: item[Consts.FIELDS.PAYMENT.COMMENTS],
      vatFreeCost: item[Consts.FIELDS.PAYMENT.VATFREECOST],
      vat: item[Consts.FIELDS.PAYMENT.VAT],
      vatRate: item[Consts.FIELDS.PAYMENT.VATRATE],
      dueDate,
      beneficiary: item[Consts.FIELDS.PAYMENT.BENEFICIARY],
      beneficiaryItemChoices,
      currency,
      beneficiaryItem,
      status: item[Consts.FIELDS.PAYMENT.STATUS],
      fees: item[Consts.FIELDS.PAYMENT.FEES]?.map((fee) => ({
        key: fee[Consts.FIELDS.COMMON.ID],
        text: fee[Consts.FIELDS.COMMON.TITLE]
      })) || [],
      contacts: item[Consts.FIELDS.PAYMENT.CONTACTS]?.map((contact) => ({
        key: contact[Consts.FIELDS.COMMON.ID],
        text: contact[Consts.FIELDS.COMMON.TITLE]
      })) || [],
      datePaid

    };

    this.setState({
      ...initialValues
    });
  };

  private validateForm = (): boolean => {
    const errors: { [key: string]: string } = {};
    const {
      name,
      beneficiaryItem,
      beneficiary,
      vatFreeCost,
      currency,
      fees
    } = this.state;

    if (name?.trim() === '' || !name)
      errors.name = strings.FormPaymentErrorNameRequired;

    if (beneficiary !== '' && !beneficiaryItem || !beneficiaryItem.key || beneficiaryItem.key === '') {
      errors.beneficiaryItem = strings.FormFeeChargedByItemValidationError;
    }
    if (fees && fees.length === 0) {
      errors.fees = strings.FormPaymentErrorFeeRequired;
    }

    if (!currency || !currency.key || currency.key === '') {
      errors.currency = strings.FormPaymentErrorCurrencyRequired;
    }

    if (!vatFreeCost || vatFreeCost?.toString().trim() === '') {
      errors.vatFreeCost = strings.FormPaymentErrorVATFreeCostRequired;
    }
    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

  public render(): React.ReactElement<IPaymentProps> {
    const {
      name,
      summary,
      comments,
      errors,
      isFormReady,
      beneficiary,
      beneficiaryChoices,
      beneficiaryItem,
      beneficiaryItemChoices,
      currency,
      currencyChoices,
      vatFreeCost,
      vat,
      vatRate,
      status,
      statusChoices,
      dueDate,
      datePaid,
      contacts,
      contactsChoices,
      fees,
      feesChoices
    } = this.state;
    return (
      <div className="ms-Grid-row">
        <div className="ms-Grid-col ms-sm12 ms-md12">
          {!this.hasAnyRole(UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && (
            <AccessDeniedMessage message={strings.FormAccessDeniedMessage}> </AccessDeniedMessage>
          )}
          {/* Form  */}
          {this.hasAnyRole(UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <FormWizard
            allowDuplicateItems={false}
            getDuplicateItems={this.getDuplicateItems}
            getSummary={this.getFormSummary}
            validateForm={this.validateForm}
            isEdit={this.isEditMode}
            isFormReady={isFormReady}
            callback={this.callback}
            processCreation={this.processCreation}
            processEdit={this.processEdit}
            formConfig={this.config}
            formOrigin={this.props.formOrigin}>
            <>
              {!isFormReady &&
                <div className={styles.spinnerContainer}>
                  <Spinner size={SpinnerSize.large} label={'Loading...'} />
                </div>
              }
              {isFormReady &&
                < Fabric >
                  <form>
                    {/* Name */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormPaymentLabelName}</Label>
                        </div>
                        <Tooltip content={strings.FormPaymentTooltipName} ></Tooltip>
                      </div>
                      <TextField
                        value={name}
                        onChange={this.onChangeName}
                      />
                      {errors.name &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.name}
                          </span>
                        </div>
                      }
                    </div>

                    {/* Fees */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormPaymentLabelFee}</Label>
                        </div>
                        <Tooltip content={strings.FormPaymentTooltipFee} ></Tooltip>
                      </div>
                      <MultiselectWrapper
                        disabled={this.props.preselectedItemId && fees.length > 0}
                        value={fees}
                        data={feesChoices}
                        dataKey={(item: IDropdownOption) => item?.key}
                        textField={(item: IDropdownOption) => item?.text}
                        filter="contains"
                        onChange={this.onChangeFees} />
                      {errors.fees &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.fees}
                          </span>
                        </div>
                      }
                    </div>

                    {/* Summary */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormPaymentLabelSummary}</Label>
                        </div>
                        <Tooltip content={strings.FormPaymentTooltipSummary} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
                    </div>
                    {/* Currency */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormPaymentLabelCurrency}</Label>
                        </div>
                        <Tooltip content={strings.FormPaymentTooltipCurrency} ></Tooltip>
                      </div>
                      <Combobox
                        data={currencyChoices}
                        value={currency}
                        selectIcon={
                          <span className="ms-Dropdown-caretDownWrapper">
                            <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
                          </span>
                        }
                        textField="text"
                        filter="contains"
                        key="Key"
                        onChange={this.onChangeCurrency}
                      />

                      {errors.currency &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.currency}
                          </span>
                        </div>
                      }
                    </div>
                    {/* VAT Row */}
                    <div className="ms-Grid" >
                      <div className="ms-Grid-row" >
                        <div className="ms-Grid-col ms-sm4 ms-md4">
                          {/*vatFreeCost */}

                          <div className={styles.field} >

                            <div className={styles.fieldLabel}>
                              <div className={styles.fieldLabelContainer}>
                                <Label required>{strings.FormPaymentLabelVATFreeCost}</Label>
                              </div>
                              <Tooltip content={strings.FormPaymentTooltipVATFreeCost} ></Tooltip>
                            </div>
                            <TextField
                              value={vatFreeCost}
                              onChange={this.onchangeVatFreeCost}
                              onKeyDown={this.onVatFreeCostKeyDown}
                              type="number"
                            />
                          </div>
                        </div>
                        {/*vat */}
                        <div className="ms-Grid-col ms-sm4 ms-md4">
                          <div className={styles.field} >
                            <div className={styles.fieldLabel}>
                              <div className={styles.fieldLabelContainer}>
                                <Label>{strings.FormPaymentLabelVAT}</Label>
                              </div>
                              <Tooltip content={strings.FormPaymentTooltipVAT} ></Tooltip>
                            </div>
                            <TextField
                              value={vat}
                              onChange={this.onChangeVat}
                              onKeyDown={this.onVatKeyDown}
                              type="number"
                            />
                          </div>
                        </div>

                        {/*vat Rate */}
                        <div className="ms-Grid-col ms-sm4 ms-md4">
                          <div className={styles.field} >
                            <div className={styles.fieldLabel}>
                              <div className={styles.fieldLabelContainer}>
                                <Label >{strings.FormPaymentLabelVATRate}</Label>
                              </div>
                              <Tooltip content={strings.FormPaymentTooltipVATRate} ></Tooltip>
                            </div>
                            <TextField
                              value={vatRate}
                              onChange={this.onchangeVatRate}
                              onKeyDown={this.onVatRateKeyDown}
                              type="number"
                            />
                          </div>
                        </div>
                      </div>
                      <div className="ms-Grid-row" >
                        <div className="ms-Grid-col ms-sm12 ms-md12">
                          {errors.vatFreeCost &&
                            <div className={styles.errorContainer}>
                              <Icon iconName="Error" className={styles.errorIcon} />
                              <span className={styles.errorMessage}>
                                {errors.vatFreeCost}
                              </span>
                            </div>
                          }
                        </div>
                      </div>
                    </div>

                    {/* Bneficiary */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormPaymentLabelBeneficiary}</Label>
                        </div>
                        <Tooltip content={strings.FormPaymentTooltipBeneficiary} ></Tooltip>
                      </div>
                      <Dropdown
                        selectedKey={beneficiary}
                        onChange={this.onChangebeneficiaryType}
                        options={beneficiaryChoices} />
                    </div>


                    {/* Beneficiary Item */}
                    {beneficiary !== '' &&
                      <div className={styles.field}>
                        <div className={styles.fieldLabel}>
                          <div className={styles.fieldLabelContainer}>
                            <Label required>{strings.FormPaymentLabelBeneficiaryItem}</Label>
                          </div>
                          <Tooltip content={strings.FormPaymentTooltipBeneficiaryItem} ></Tooltip>
                        </div>
                        <Combobox
                          data={beneficiaryItemChoices}
                          value={beneficiaryItem}
                          selectIcon={
                            <span className="ms-Dropdown-caretDownWrapper">
                              <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
                            </span>
                          }
                          textField="text"
                          filter="contains"
                          key="Key"
                          onChange={this.onChangeBeneficiaryItem}
                        />
                        {errors.beneficiaryItem &&
                          <div className={styles.errorContainer}>
                            <Icon iconName="Error" className={styles.errorIcon} />
                            <span className={styles.errorMessage}>
                              {errors.beneficiaryItem}
                            </span>
                          </div>
                        }
                      </div>
                    }
                    {/* Contact */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormPaymentLabelContact}</Label>
                        </div>
                        <Tooltip content={strings.FormPaymentTooltipContact} ></Tooltip>
                      </div>
                      <MultiselectWrapper
                        value={contacts}
                        data={contactsChoices}
                        dataKey={(item: IDropdownOption) => item.key}
                        textField={(item: IDropdownOption) => item.text}
                        filter="contains"
                        onChange={this.onChangeContacts} />

                    </div>


                    {/* Due date */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormPaymentLabelDueDate}</Label>
                        </div>
                        <Tooltip content={strings.FormPaymentTooltipDueDate} ></Tooltip>
                      </div>
                      <DatePicker
                        allowTextInput
                        onSelectDate={this.onchangeDueDate}
                        value={dueDate}
                        formatDate={this.formatDate}
                        styles={datepickerstyles}
                      />
                    </div>

                    {/* Status */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormPaymentLabelStatus}</Label>
                        </div>
                        <Tooltip content={strings.FormPaymentTooltipStatus} ></Tooltip>
                      </div>
                      <Dropdown

                        selectedKey={status}
                        onChange={this.onChangeStatus}
                        options={statusChoices} />
                    </div>

                    {/* Date paid */}
                    {status === "Paid" && (<div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormPaymentLabelDatePaid}</Label>
                        </div>
                        <Tooltip content={strings.FormPaymentTooltipDatePaid} ></Tooltip>
                      </div>

                      <DatePicker
                        allowTextInput
                        onSelectDate={this.onchangeDatePaid}
                        value={datePaid}
                        formatDate={this.formatDate}
                        styles={datepickerstyles}
                      />
                    </div>)}
                    {/* Comments */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormPaymentLabelComments}</Label>
                        </div>
                        <Tooltip content={strings.FormPaymentTooltipComments} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={comments} onChange={this.onChangeComments} />
                    </div>
                  </form>
                </Fabric>
              }
            </>
          </FormWizard>}
        </div>
      </div >
    );
  }

  private formatDate = (date?: Date): string => {
    if (!date) return '';
    return UtilHelper.formatDate(date, 'LL', 'en-us');
  };

  private onChangeName = (event, name: string) => {
    this.setState({ name });
  };

  private onChangeContacts = (contacts: IDropdownOption[]) => {
    this.setState({ contacts });
  };

  private onChangeFees = (fees: IDropdownOption[]) => {
    this.setState({ fees });
  };

  private onChangeSummary = (summary: string) => {
    this.setState({ summary });
    return summary;
  };

  /*private onchangeVatFreeCost = (event, vatFreeCost: string) => {
    this.setState({ vatFreeCost });
  };*/
  private onchangeVatFreeCost = (_e: any, vatFreeCost: string) => {
    if (vatFreeCost === "") return this.setState({ vatFreeCost: "" });
    const n = Math.max(0, Number(vatFreeCost) || 0);
    this.setState({ vatFreeCost: String(n) });
  };
  private onVatFreeCostKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    // Prevent invalid characters: minus, plus, E, e
    if (e.key === '-' || e.key === '+' || e.key.toLowerCase() === 'e') {
      e.preventDefault();
    }
  };
  /*private onChangeVat = (event, vat: string) => {
    this.setState({ vat });
  };*/
  private onChangeVat = (_e: any, vat: string) => {
    if (vat === "") return this.setState({ vat: "" });
    const n = Math.max(0, Number(vat) || 0);
    this.setState({ vat: String(n) });
  };
  private onVatKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    // Prevent invalid characters: minus, plus, E, e
    if (e.key === '-' || e.key === '+' || e.key.toLowerCase() === 'e') {
      e.preventDefault();
    }
  };

  /*private onchangeVatRate = (event, vatRate: string) => {
    this.setState({ vatRate });
  };*/
  private onchangeVatRate = (_e: any, vatRate: string) => {
    if (vatRate === "") return this.setState({ vatRate: "" });
    const n = Math.max(0, Number(vatRate) || 0);
    this.setState({ vatRate: String(n) });
  };
  private onVatRateKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    // Prevent invalid characters: minus, plus, E, e
    if (e.key === '-' || e.key === '+' || e.key.toLowerCase() === 'e') {
      e.preventDefault();
    }
  };


  private onchangeDueDate = (dueDate: Date) => {
    this.setState({ dueDate });
  };

  private onchangeDatePaid = (datePaid: Date) => {
    this.setState({ datePaid });
  };

  private onChangeStatus = (event, status: IDropdownOption) => {
    this.setState({ status: status.key.toString() });
  };

  private onChangebeneficiaryType = async (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ beneficiary: option.key.toString() });
    const beneficiaryItemChoices: IDropdownOption[] = await this.getbeneficiaryContentTypeId(option.key.toString());
    const contactsChoices: IDropdownOption[] = this.AllContactsItems.get(option.key.toString());

    //contactChoices
    this.setState({ beneficiaryItemChoices, beneficiaryItem: null, contactsChoices, contacts: [] });
  };

  private onChangeBeneficiaryItem = (beneficiaryItem: IDropdownOption) => {
    const { beneficiary } = this.state;
    const contactAllData = this.AllContactsItems.get(beneficiary);
    const filterKey = beneficiaryItem?.key;
    let filteredContacts: IDropdownOption[] = [];
    switch (beneficiary) {
      case FeeChargedBy.LawFirm:
      case FeeChargedBy.LegalRepresentative:
        filteredContacts = contactAllData?.filter((item) => item["SVCId"] === filterKey?.toString());
        break;
      case FeeChargedBy.TeleportPartner:
        filteredContacts = contactAllData?.filter((item) => item["TPId"] === filterKey?.toString());
        break;
      case FeeChargedBy.Regulator:
      case FeeChargedBy.Organization:
      case FeeChargedBy.Administration:
        filteredContacts = contactAllData?.filter((item) => item["AUTHId"] === filterKey?.toString());
    }
    const contactsChoices = filterKey ? filteredContacts : this.AllContactsItems.get(beneficiary);
    this.setState({ beneficiaryItem, contactsChoices, contacts: [] });
  };

  private onChangeComments = (comments: string) => {
    this.setState({ comments });
    return comments;
  };

  private getbeneficiaryContentTypeId = async (contentTypeText: string) => {
    if (this.AllbeneficiaryItems.has(contentTypeText)) {
      const cachedItems = this.AllbeneficiaryItems.get(contentTypeText) || [];
      return cachedItems;
    }
  };

  private onChangeCurrency = (currency: IDropdownOption) => {
    this.setState({ currency });
  };

  private getDuplicateItems = async (): Promise<IBaseItem[]> => {
    const duplicateItems = [];
    return duplicateItems;
  };

  private getFormSummary = () => {
    const {
      name,
      summary,
      comments,
      beneficiary,
      beneficiaryChoices,
      beneficiaryItem,
      status,
      vatFreeCost,
      vat,
      vatRate,
      contacts,
      fees,
      datePaid,
      dueDate
    } = this.state;

    const beneficiaryTextValue = beneficiaryChoices.find((choice) => choice.key === beneficiary)?.text;
    const contactsValues = contacts.map((c) => c.text).join(", ");
    const feesValues = fees.map((e) => e.text).join(", ");
    const summaryTextValue = DomHelper.cleanRichHtml(summary);
    const commentsTextValue = DomHelper.cleanRichHtml(comments);
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentDatasheetNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelName}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {name}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelFee}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {feesValues}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelSummary}</Label>
            </div>
          </div>

          <div className={styles.fieldValue}>
            {summaryTextValue.length > 0 ? <RichText isEditMode={false} value={summaryTextValue} /> : ''}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelVATFreeCost}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {vatFreeCost ? vatFreeCost : '0'}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelVAT}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {vat ? vat : '0'}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelVATRate}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {vatRate ? vatRate : '0'}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelBeneficiary}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {beneficiaryTextValue}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelBeneficiaryItem}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {beneficiaryItem.text}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelContact}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {contactsValues}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelDueDate}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {dueDate ? this.formatDate(dueDate) : ''}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelStatus}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {status}
          </div>
        </div>

        {status === "Paid" && (<div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelDatePaid}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {datePaid ? this.formatDate(datePaid) : ''}
          </div>
        </div>)}

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelComments}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {commentsTextValue.length > 0 ? <RichText isEditMode={false} value={commentsTextValue} /> : ''}
          </div>
        </div>

      </div>
    );
  };

  private callback = async (redirectUrl) => {
    const {
      callback,
    } = this.props;
    callback(redirectUrl);
  };

  private processCreation = async (): Promise<IItemAddResult> => {
    const {
      summary,
      comments,
      status,
      name,
      beneficiary,
      beneficiaryItem,
      vatFreeCost,
      vat,
      vatRate,
      currency,
      contacts,
      fees,
      datePaid,
      dueDate

    } = this.state;

    const {
      pnpService,
    } = this.props;


    const itemData = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = name;

    itemData[Consts.FIELDS.PAYMENT.FEES_ID] =
      fees.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.PAYMENT.CONTACTS_ID] =
      contacts.map((c) => (parseInt(c.key.toString())));

    if (status === "Paid" && datePaid) {
      itemData[Consts.FIELDS.PAYMENT.PAYMENTPAID] = datePaid?.toISOString();
    } else {
      itemData[Consts.FIELDS.PAYMENT.PAYMENTPAID] = null;
    }
    itemData[Consts.FIELDS.PAYMENT.SUMMARY] = summary;

    if (vatFreeCost) {
      itemData[Consts.FIELDS.PAYMENT.VATFREECOST] = vatFreeCost;
    }
    if (dueDate) {
      itemData[Consts.FIELDS.PAYMENT.DUEDATE] = dueDate?.toISOString();
    }
    if (vat) {
      itemData[Consts.FIELDS.PAYMENT.VAT] = vat;
    }
    if (vatRate) {
      itemData[Consts.FIELDS.PAYMENT.VATRATE] = vatRate;
    }

    itemData[Consts.FIELDS.PAYMENT.STATUS] = status;
    itemData[Consts.FIELDS.PAYMENT.COMMENTS] = comments;
    itemData[Consts.FIELDS.PAYMENT.CURRENCY_ID] = currency.key;
    itemData[Consts.FIELDS.PAYMENT.BENEFICIARY] = beneficiary;
    switch (beneficiary) {
      // Group 1: Law Firm & Legal Representative
      case FeeChargedBy.LawFirm:
      case FeeChargedBy.LegalRepresentative:
        itemData[Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER_ID] = beneficiaryItem.key;
        break;

      // Group 2: Organization & Regulator
      case FeeChargedBy.Administration:
      case FeeChargedBy.Organization:
      case FeeChargedBy.Regulator:
        itemData[Consts.FIELDS.PAYMENT.AUTHORITY_ID] = beneficiaryItem.key;
        break;

      // Group 4: Teleport Partner
      case FeeChargedBy.TeleportPartner:
        itemData[Consts.FIELDS.PAYMENT.TELEPORTPARTNER_ID] = beneficiaryItem.key;
        break;

      // Default: do nothing
      default:
        break;
    }

    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);
    const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
    const documentSpaceUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.DOCUMENT_SPACE_URL);
    const parentFolderUrl = `${documentSpaceUrl}${Consts.DOCUMENT_SPACE.PAYMENT_FOLDER_URL}`;
    const folderUrl = await createUniqueFolder(pnpService, parentFolderUrl, sanitizedFolderName);
    const updateData = {};
    updateData[Consts.FIELDS.COMMON.DOCUMENTS_SPACE] =
    {
      [Consts.FIELDS.COMMON.URL]: folderUrl, [Consts.FIELDS.COMMON.DESCRIPTION]: "Link"
    },
      await pnpService.getListItems(listUrl).getById(createdItem.data.ID).update(updateData);
    return createdItem;
  };

  private processEdit = async () => {
    const {
      name,
      summary,
      comments,
      status,
      beneficiary,
      beneficiaryItem,
      vatFreeCost,
      vat,
      vatRate,
      currency,
      dueDate,
      fees,
      contacts,
      datePaid
    } = this.state;

    const {
      pnpService,
      itemId
    } = this.props;

    const itemData = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = name;

    itemData[Consts.FIELDS.PAYMENT.FEES_ID] =
      fees.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.PAYMENT.CONTACTS_ID] =
      contacts.map((c) => (parseInt(c.key.toString())));
    if (status === "Paid" && datePaid) {
      itemData[Consts.FIELDS.PAYMENT.PAYMENTPAID] = datePaid?.toISOString();
    } else {
      itemData[Consts.FIELDS.PAYMENT.PAYMENTPAID] = null;
    }
    itemData[Consts.FIELDS.PAYMENT.SUMMARY] = summary;

    if (vatFreeCost) itemData[Consts.FIELDS.PAYMENT.VATFREECOST] = vatFreeCost;
    else itemData[Consts.FIELDS.PAYMENT.VATFREECOST] = '0';
    if (dueDate) {
      itemData[Consts.FIELDS.PAYMENT.DUEDATE] = dueDate?.toISOString();
    }
    if (vat) itemData[Consts.FIELDS.PAYMENT.VAT] = vat;
    else itemData[Consts.FIELDS.PAYMENT.VAT] = '0';
    if (vatRate) itemData[Consts.FIELDS.PAYMENT.VATRATE] = vatRate;
    else itemData[Consts.FIELDS.PAYMENT.VATRATE] = '0';

    itemData[Consts.FIELDS.PAYMENT.STATUS] = status;
    itemData[Consts.FIELDS.PAYMENT.COMMENTS] = comments;
    itemData[Consts.FIELDS.PAYMENT.CURRENCY_ID] = currency.key;
    itemData[Consts.FIELDS.PAYMENT.BENEFICIARY] = beneficiary;

    switch (beneficiary) {
      case FeeChargedBy.LawFirm:
      case FeeChargedBy.LegalRepresentative:
        itemData[Consts.FIELDS.PAYMENT.LEGALSERVICEPROVIDER_ID] = beneficiaryItem.key;
        break;
      case FeeChargedBy.Administration:
      case FeeChargedBy.Organization:
      case FeeChargedBy.Regulator:
        itemData[Consts.FIELDS.PAYMENT.AUTHORITY_ID] = beneficiaryItem.key;
        break;
      case FeeChargedBy.TeleportPartner:
        itemData[Consts.FIELDS.PAYMENT.TELEPORTPARTNER_ID] = beneficiaryItem.key;
        break;
      default:
        break;
    }
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
  };
}



