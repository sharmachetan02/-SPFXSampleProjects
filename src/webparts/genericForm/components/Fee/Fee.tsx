import * as React from 'react';

import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { IFeeProps, IFeeState } from './Fee.type';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import { Dropdown, Fabric, Icon, IDropdownOption, Label, TextField } from 'office-ui-fabric-react';
import { Combobox } from 'react-widgets';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { Form, FormOrigin, FormType, IFormConfig } from '../Form.types';
import strings from 'GenericFormWebPartStrings';
import { RichText } from '../../../../common/components/RichText/RichText';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { UtilHelper } from '../../../../common/helpers/Util';
import { Consts } from '../../../../common/consts/Consts';
import 'react-widgets/styles.css';

import { IItemAddResult } from '@pnp/sp/items';
//import { createUniqueFolder } from '../Form.utility';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { DatePicker, Toggle } from 'office-ui-fabric-react';
import RecurrencePatternPicker from '../../../../common/components/Recurrence Pattern Picker/RecurrencePatternPicker';
import { RecurrenceData } from '../../../../common/components/Recurrence Pattern Picker/RecurrencePatternPicker.types';
import { createUniqueFolder, datepickerstyles } from '../Form.utility';
import { FeeChargedBy, FeeType, UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';

export default class Fee extends React.Component<IFeeProps, IFeeState> {
  private isEditMode: boolean;
  private config: IFormConfig;
  private AllChargedByItems: Map<string, IDropdownOption[]> = new Map();

  constructor(props: Readonly<IFeeProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      isFormReady: false,
      name: '',
      category: '',
      categoryChoices: [],
      summary: '',
      typeChoices: [],
      type: '',
      item: null,
      itemChoices: [],
      chargedBy: '',
      chargedByChoices: [],
      chargeByItem: null,
      chargeByItemChoices: [],
      currency: null,
      currencyChoices: [],
      vatFreeCost: '',
      vat: '',
      vatRate: '',
      costType: '',
      costTypeChoices: [],
      poNonPO: '',
      poNonPOChoices: [],
      status: '',
      statusChoices: [],
      dueDate: null,
      comments: '',
      recurrenceFee: false,
      recurrencePattern: {
        frequency: '',
        weekDay: null,
        month: null,
        day: null,
        interval: null,
        endDate: null,
      },
      recurrencPatternId: 0,
      errors: {}
    };

    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.Fee);
    this.AllChargedByItems = new Map();
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

    // Wait for AllChargedByItems to be populated first
    await this.fetchAllChargedByItems();
    const promises = [];

    promises.push(
      this.initTypeAndItemChoices()
    );

    promises.push(
      this.initCategoryChoices()
    );

    promises.push(
      this.initFeeChargedByChoices()
    );

    promises.push(
      this.initCurrencyChoices()
    );

    promises.push(
      this.initCostTypeChoices()
    );

    promises.push(
      this.initPoNonPOChoices()
    );

    promises.push(
      this.initStatusChoices()
    );

    await Promise.all(promises);

    if (this.isEditMode)
      await this.setInitialFormValues();

  }

  private fetchAllChargedByItems = async () => {
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
      let query = pnpService.getListItems(listUrl).select(...selectedFields).expand(Consts.FIELDS.COMMON.CONTENTTYPE).top(5000);
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
      this.AllChargedByItems.set(key, choices);
    });
  };


  private initTypeAndItemChoices = async () => {
    const typeChoices: IDropdownOption[] = [
      { key: Consts.CONTENT_TYPES.SNP_FEE, text: FeeType.SNP },
      { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_FEE, text: FeeType.MATypeRequirement }
    ];

    if (!this.isEditMode) {
      const {
        pnpService,
        preselectedItemType,
        preselectedItemId
      } = this.props;
      const formConfigData = Form.config.find((f) => f.type.toLowerCase() === preselectedItemType);
      const type = FormType.SNP === formConfigData.type ? typeChoices[0].key as string : typeChoices[1].key as string;
      this.setState({
        typeChoices,
        type
      });

      //item bind logic
      const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), formConfigData.list);
      const selectedFields: string[] = [
        Consts.FIELDS.COMMON.ID,
        Consts.FIELDS.COMMON.TITLE
      ];

      const itemChoices: IDropdownOption[] = [];
      const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)
        .filter(`Id eq ${preselectedItemId}`)
        .top(1)();
      items.forEach((item) => {
        itemChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
      });
      const item: IDropdownOption = itemChoices.length > 0 ? itemChoices[0] : null;
      this.setState({
        itemChoices, item
      });
    } else {
      this.setState({
        typeChoices
      });


    }
  };
  private initCategoryChoices = async () => {

    const {
      pnpService
    } = this.props;
    let categoryChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.FEE.CATEGORY)
      .select('Choices')();
    categoryChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    if (!this.isEditMode) {
      const category = categoryChoices.length > 0 ? categoryChoices[0].key as string : '';
      this.setState({
        category, categoryChoices
      });
    } else {
      this.setState({
        categoryChoices
      });
    }

  };

  private initFeeChargedByChoices = async () => {

    const {
      pnpService
    } = this.props;
    let chargedByChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.FEE.CHARGEDBY)
      .select('Choices')();
    chargedByChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    if (!this.isEditMode) {
      const chargedBy = chargedByChoices.length > 0 ? chargedByChoices[0].key as string : '';
      const chargeByItemChoices = await this.loadChargeByItemChoices(chargedBy);
      this.setState({
        chargedBy, chargedByChoices, chargeByItemChoices
      });
    } else {
      this.setState({
        chargedByChoices
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
    const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true).top(5000)();
    items.forEach((item) => {
      //this.allCountries.push({ Id: item[Consts.FIELDS.COMMON.ID], Title: item[Consts.FIELDS.COMMON.TITLE], Grouping: item[Consts.FIELDS.SNP.GROUPING] });
      currencyChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    if (!this.isEditMode) {
      const currency = null;
      this.setState({ currency, currencyChoices });
    } else {
      this.setState({ currencyChoices });
    }
  };

  private initCostTypeChoices = async () => {

    const {
      pnpService
    } = this.props;
    let costTypeChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.FEE.COSTTYPE)
      .select('Choices')();
    costTypeChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    if (!this.isEditMode) {
      const costType = costTypeChoices.length > 0 ? costTypeChoices[0].key as string : '';
      this.setState({
        costType, costTypeChoices
      });
    } else {
      this.setState({
        costTypeChoices
      });
    }

  };

  private initPoNonPOChoices = async () => {

    const {
      pnpService
    } = this.props;
    let poNonPOChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.FEE.PONONPO)
      .select('Choices')();
    poNonPOChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    if (!this.isEditMode) {
      const poNonPO = poNonPOChoices.length > 0 ? poNonPOChoices[0].key as string : '';
      this.setState({
        poNonPO, poNonPOChoices
      });
    } else {
      this.setState({
        poNonPOChoices
      });
    }

  };

  public async initStatusChoices() {
    const {
      pnpService
    } = this.props;
    let statusChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.FEE.STATUS)
      .select('Choices')();
    statusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    const status = statusChoices.length > 0 ? statusChoices[0].key as string : '';
    this.setState({
      statusChoices, status
    });
  }

  private formatDate = (date?: Date): string => {
    if (!date) return '';
    return UtilHelper.formatDate(date, 'LL', 'en-us');
  };

  private formatRecurrencePattern = (recurrencePattern: RecurrenceData): string => {
    const frequency = recurrencePattern.frequency || '';
    const weekDay = recurrencePattern.weekDay || '';
    const month = recurrencePattern.month || '';
    const day = recurrencePattern.day || 1;
    const interval = recurrencePattern.interval || 1;
    const endDate = recurrencePattern.endDate || null;

    let result = '';

    switch (frequency) {
      case 'Weekly':
        result = `${weekDay} of every ${interval} week${interval > 1 ? 's' : ''}`;
        break;
      case 'Monthly':
        result = `Day ${day} of every ${interval} month${interval > 1 ? 's' : ''}`;
        break;
      case 'Quarterly':
        result = `Day ${day} of every 3 months`;
        break;
      case 'Half Annual':
        result = `Day ${day} of every 6 months`;
        break;
      case 'Annual':
        result = `Every ${month} ${day} of every 1 year`;
        break;
      case 'Multi Annual':
        result = `Every ${month} ${day} of every ${interval} year${interval > 1 ? 's' : ''}`;
        break;
      default:
        result = frequency;
    }

    if (frequency && endDate) {
      const date = UtilHelper.formatDate(endDate, 'LL', 'en-us');
      result += ` and expired on ${date} `;
    }

    return result;
  };

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Fee).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.FEE.SUMMARY,
      Consts.FIELDS.FEE.STATUS,
      Consts.FIELDS.FEE.COMMENTS,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      Consts.FIELDS.FEE.CATEGORY,
      Consts.FIELDS.FEE.SUMMARY,
      Consts.FIELDS.FEE.VATRATE,
      Consts.FIELDS.FEE.DUEDATE,
      Consts.FIELDS.FEE.VATFREECOST,
      Consts.FIELDS.FEE.VAT,
      Consts.FIELDS.FEE.VATRATE,
      Consts.FIELDS.FEE.COSTTYPE,
      Consts.FIELDS.FEE.PONONPO,
      Consts.FIELDS.FEE.CHARGEDBY,
      Consts.FIELDS.FEE.REPEATINGFEE,
      `${Consts.FIELDS.FEE.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.FEE.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,

      `${Consts.FIELDS.FEE.CURRENCY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.FEE.CURRENCY}/${Consts.FIELDS.COMMON.TITLE}`,

      `${Consts.FIELDS.FEE.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.FEE.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.TITLE}`,

      `${Consts.FIELDS.FEE.MA_REQUIREMENT}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.FEE.MA_REQUIREMENT}/${Consts.FIELDS.COMMON.TITLE}`,

      `${Consts.FIELDS.FEE.RECURRENCEPATTERN}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.FEE.RECURRENCEPATTERN}/${Consts.FIELDS.COMMON.TITLE}`,

      `${Consts.FIELDS.FEE.SNP}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.FEE.SNP}/${Consts.FIELDS.COMMON.TITLE}`,

      `${Consts.FIELDS.FEE.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.FEE.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.TITLE}`

    ];
    const expand = [Consts.FIELDS.FEE.AUTHORITY, Consts.FIELDS.FEE.CURRENCY, Consts.FIELDS.FEE.LEGALSERVICEPROVIDER,
    Consts.FIELDS.FEE.MA_REQUIREMENT, Consts.FIELDS.FEE.RECURRENCEPATTERN, Consts.FIELDS.FEE.SNP, Consts.FIELDS.FEE.TELEPORTPARTNER];
    const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
    const filterTypeChoices = "" + this.state.typeChoices.find((type) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1)?.key;

    let currency = null;

    if (item[Consts.FIELDS.FEE.CURRENCY]) {
      currency = {
        key: item[Consts.FIELDS.FEE.CURRENCY][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.FEE.CURRENCY][Consts.FIELDS.COMMON.TITLE]
      };
    }
    let itemContent: IDropdownOption = null;
    if (filterTypeChoices === Consts.CONTENT_TYPES.SNP_FEE) {
      itemContent = {
        key: item[Consts.FIELDS.FEE.SNP][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.FEE.SNP][Consts.FIELDS.COMMON.TITLE]
      };
    } else {
      itemContent = {
        key: item[Consts.FIELDS.FEE.MA_REQUIREMENT][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.FEE.MA_REQUIREMENT][Consts.FIELDS.COMMON.TITLE]
      };
    }

    let recurrencPatternId;
    if (item[Consts.FIELDS.FEE.RECURRENCEPATTERN]) {
      recurrencPatternId = item[Consts.FIELDS.FEE.RECURRENCEPATTERN][Consts.FIELDS.COMMON.ID];
    }



    // Get charged by value and load corresponding choices
    const chargedByValue = item[Consts.FIELDS.FEE.CHARGEDBY];
    const chargeByItemChoices: IDropdownOption[] = await this.loadChargeByItemChoices(chargedByValue);

    // Get the ID of the charged by item from the appropriate field
    let chargeByItemId: number = null;
    switch (chargedByValue) {
      case FeeChargedBy.LawFirm:
      case FeeChargedBy.LegalRepresentative:
        chargeByItemId = item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID];
        break;
      case FeeChargedBy.Administration:
      case FeeChargedBy.Organization:
      case FeeChargedBy.Regulator:
        chargeByItemId = item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID];
        break;
      case FeeChargedBy.TeleportPartner:
        chargeByItemId = item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID];
        break;
      default:
        break;
    }

    // Find the matching item from choices array to ensure correct reference
    const chargeByItem: IDropdownOption =
      chargeByItemId !== null ?
        chargeByItemChoices.find((c) => c.key === chargeByItemId) || null : null;
    let recurrencePattern = null;
    if (item[`${Consts.FIELDS.FEE.RECURRENCEPATTERN}`]) {
      const recurrencePatternFieldId = item[`${Consts.FIELDS.FEE.RECURRENCEPATTERN}`][Consts.FIELDS.COMMON.ID];
      recurrencePattern = await this.initEditModeRecurrence(recurrencePatternFieldId);
    }
    const rawDueDate = item[Consts.FIELDS.FEE.DUEDATE];
    const dueDate: Date | null = rawDueDate ? new Date(rawDueDate) : null;
    const initialValues = {
      name: item[Consts.FIELDS.COMMON.TITLE],
      category: item[Consts.FIELDS.FEE.CATEGORY],
      type: filterTypeChoices,
      item: itemContent,
      summary: item[Consts.FIELDS.FEE.SUMMARY],
      comments: item[Consts.FIELDS.FEE.COMMENTS],
      vatFreeCost: item[Consts.FIELDS.FEE.VATFREECOST],
      dueDate,
      vat: item[Consts.FIELDS.FEE.VAT],
      vatRate: item[Consts.FIELDS.FEE.VATRATE],
      costType: item[Consts.FIELDS.FEE.COSTTYPE],
      poNonPO: item[Consts.FIELDS.FEE.PONONPO],
      chargedBy: item[Consts.FIELDS.FEE.CHARGEDBY],
      chargeByItemChoices,
      currency,
      recurrencePattern,
      recurrenceFee: item[`${Consts.FIELDS.FEE.REPEATINGFEE}`],
      recurrencPatternId,
      chargeByItem,
      status: item[Consts.FIELDS.FEE.STATUS]
    };

    this.setState({
      ...initialValues
    });
  };

  /**
   * Loads charge by item choices based on the charged by type.
   * Uses cached data from AllChargedByItems map populated during initialization.
   */
  private loadChargeByItemChoices = async (chargedBy: string): Promise<IDropdownOption[]> => {
    if (this.AllChargedByItems.has(chargedBy)) {
      const chargeByItemChoices: IDropdownOption[] = this.AllChargedByItems.get(chargedBy) || [];
      return chargeByItemChoices;
    }
    return [];
  };

  private initEditModeRecurrence = async (itemId: number): Promise<RecurrenceData | null> => {
    if (!this.isEditMode || !itemId) return null;

    const { pnpService } = this.props;

    const list = Consts.LISTS.RECURRENCE_PATTERN_URL;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);

    const selectedFields: string[] = [
      Consts.FIELDS.RECURRENCE_PATTERN.FREQUENCY,
      Consts.FIELDS.RECURRENCE_PATTERN.WEEKDAY,
      Consts.FIELDS.RECURRENCE_PATTERN.MONTH,
      Consts.FIELDS.RECURRENCE_PATTERN.DAY,
      Consts.FIELDS.RECURRENCE_PATTERN.INTERVAL,
      Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE
    ];

    try {
      const item = await pnpService.getListItems(listUrl)
        .getById(itemId)
        .select(...selectedFields)();

      const recurrencePattern: RecurrenceData = {
        frequency: item[Consts.FIELDS.RECURRENCE_PATTERN.FREQUENCY],
        weekDay: item[Consts.FIELDS.RECURRENCE_PATTERN.WEEKDAY] ?? null,
        month: item[Consts.FIELDS.RECURRENCE_PATTERN.MONTH] ?? null,
        day: item[Consts.FIELDS.RECURRENCE_PATTERN.DAY] ?? null,
        interval: item[Consts.FIELDS.RECURRENCE_PATTERN.INTERVAL] ?? null,
        endDate: item[Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE] ? new Date(item[Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE]) : null
      };

      return recurrencePattern;
    } catch (error) {
      console.error("Error fetching recurrence pattern:", error);
      return null;
    }
  };

  private validateForm = (): boolean => {
    const errors: { [key: string]: string } = {};
    const {
      name,
      type,
      item,
      chargeByItem,
      chargedBy,
      vatFreeCost,
      currency

    } = this.state;
    if (name?.trim() === '' || !name)
      errors.name = strings.FormSNPNameValidationError;

    if (type !== '' && !item) {
      errors.relatedItem = strings.FormFeeItemValidationError;
    }

    if (chargedBy !== '' && !chargeByItem || !chargeByItem.key || chargeByItem.key === '') {
      errors.chargedByItem = strings.FormFeeChargedByItemValidationError;
    }

    if (!currency || !currency.key || currency.key === '') {
      errors.currency = strings.FormFeeCurrencyValidationError;
    }

    if (!vatFreeCost || vatFreeCost?.toString().trim() === '') {
      errors.vatFreeCost = strings.FormFeeVATFreeCostLabel;
    }
    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

  public render(): React.ReactElement<IFeeProps> {
    const {
      name,
      summary,
      comments,
      errors,
      isFormReady,
      typeChoices,
      type,
      item,
      itemChoices,
      chargedBy,
      chargedByChoices,
      chargeByItem,
      chargeByItemChoices,
      currency,
      currencyChoices,
      vatFreeCost,
      vat,
      vatRate,
      costType,
      costTypeChoices,
      poNonPO,
      poNonPOChoices,
      status,
      statusChoices,
      dueDate,
      recurrenceFee,
      category,
      categoryChoices,
      recurrencePattern
    } = this.state;

    return (
      <div className="ms-Grid-row">
        <div className="ms-Grid-col ms-sm12 ms-md12">
          {!this.hasAnyRole(UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && (
            <AccessDeniedMessage message={strings.FormAccessDeniedMessage}> </AccessDeniedMessage>
          )}
          {/* Form  */}
          {this.hasAnyRole(UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <FormWizard
            allowDuplicateItems={true}
            getDuplicateItems={this.getDuplicateItems}
            getSummary={this.getFormSummary}
            validateForm={this.validateForm}
            isEdit={this.isEditMode}
            isFormReady={isFormReady}
            callback={this.callback}
            processCreation={this.processCreation}
            processEdit={this.processEdit}
            formConfig={this.config}
            formOrigin={FormOrigin.Datasheet}>
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
                          <Label required>{strings.FormFeeNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormFeeNameTooltip} ></Tooltip>
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
                    {/* Category */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormFeeCategoryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormFeeCategoryTooltip} ></Tooltip>
                      </div>
                      <Dropdown
                        selectedKey={category}
                        onChange={this.onChangeCategory}
                        options={categoryChoices} />
                    </div>
                    {/* Summary */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormFeeSummaryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormFeeSummaryTooltip} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
                    </div>

                    {/* Type */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormFeeTypeLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormFeeTypeTooltip} ></Tooltip>
                      </div>
                      <Dropdown
                        selectedKey={type} disabled={true}
                        onChange={this.onChangeType}
                        options={typeChoices} />
                    </div>

                    {/* Item */}
                    {type !== '' &&
                      <div className={styles.field} >
                        <div className={styles.fieldLabel}>
                          <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormFeeItemLabel}</Label>
                          </div>
                          <Tooltip content={strings.FormFeeItemTooltip} ></Tooltip>
                        </div>
                        <Combobox
                          disabled={true}
                          data={itemChoices}
                          value={item}
                          selectIcon={
                            <span className="ms-Dropdown-caretDownWrapper">
                              <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
                            </span>
                          }
                          textField="text"
                          filter="contains"
                          key="key"
                          onChange={this.onChangeItem}
                        />
                        {errors.relatedItem &&
                          <div className={styles.errorContainer}>
                            <Icon iconName="Error" className={styles.errorIcon} />
                            <span className={styles.errorMessage}>
                              {errors.relatedItem}
                            </span>
                          </div>
                        }
                      </div>
                    }
                    {/* Fee Issuer Type */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormFeeChargedByLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormFeeChargedByTooltip} ></Tooltip>
                      </div>
                      <Dropdown
                        selectedKey={chargedBy}
                        onChange={this.onChangeChargedByItem}
                        options={chargedByChoices} />
                    </div>

                    {/* Fee Issuer */}
                    {chargedBy !== '' &&
                      <div className={styles.field} >
                        <div className={styles.fieldLabel}>
                          <div className={styles.fieldLabelContainer}>
                            <Label required>{strings.FormFeeChargedByItemLabel}</Label>
                          </div>
                          <Tooltip content={strings.FormFeeChargedByItemTooltip} ></Tooltip>
                        </div>
                        <Combobox
                          data={chargeByItemChoices}
                          value={chargeByItem}
                          selectIcon={
                            <span className="ms-Dropdown-caretDownWrapper">
                              <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
                            </span>
                          }
                          textField="text"
                          filter="contains"
                          onChange={this.onChangeChargeByItem}
                        />
                        {errors.chargedByItem &&
                          <div className={styles.errorContainer}>
                            <Icon iconName="Error" className={styles.errorIcon} />
                            <span className={styles.errorMessage}>
                              {errors.chargedByItem}
                            </span>
                          </div>
                        }
                      </div>
                    }
                    {/* Currency */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormFeeCurrencyLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormFeeCurrencyTooltip} ></Tooltip>
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
                        key="key"
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


                    {/*vatFreeCost */}
                    <div className="ms-Grid" >
                      <div className="ms-Grid-row" >
                        <div className="ms-Grid-col ms-sm4 ms-md4">
                          <div className={styles.field} >

                            <div className={styles.fieldLabel}>
                              <div className={styles.fieldLabelContainer}>
                                <Label required>{strings.FormFeeVATFreeCostLabel}</Label>
                              </div>
                              <Tooltip content={strings.FormFeeVatFreeTooltip} ></Tooltip>
                            </div>
                            <TextField
                              value={vatFreeCost}
                              onChange={this.onchangeVatFreeCost}
                              type="number"
                              onKeyDown={this.onVatFreeCostKeyDown}
                            />
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
                        <div className="ms-Grid-col ms-sm4 ms-md4">
                          {/*vat */}
                          <div className={styles.field} >
                            <div className={styles.fieldLabel}>
                              <div className={styles.fieldLabelContainer}>
                                <Label>{strings.FormFeeVATLabel}</Label>
                              </div>
                              <Tooltip content={strings.FormFeeVatTooltip} ></Tooltip>
                            </div>
                            <TextField
                              value={vat}
                              onChange={this.onChangeVat}
                              type="number"
                              onKeyDown={this.onVatKeyDown}
                            />
                          </div>
                        </div>
                        <div className="ms-Grid-col ms-sm4 ms-md4">
                          {/*vat Rate */}
                          <div className={styles.field} >
                            <div className={styles.fieldLabel}>
                              <div className={styles.fieldLabelContainer}>
                                <Label >{strings.FormFeeVATRateLabel}</Label>
                              </div>
                              <Tooltip content={strings.FormFeeVatRateTooltip} ></Tooltip>
                            </div>
                            <TextField
                              value={vatRate}
                              onChange={this.onchangeVatRate}
                              type="number"
                              onKeyDown={this.onVatRateKeyDown}
                            />
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="ms-Grid" >
                      <div className="ms-Grid-row" >
                        <div className="ms-Grid-col ms-sm4 ms-md4">
                          {/* cost type */}
                          <div className={styles.field}>
                            <div className={styles.fieldLabel}>
                              <div className={styles.fieldLabelContainer}>
                                <Label>{strings.FormFeeCostTypeLabel}</Label>
                              </div>
                              <Tooltip content={strings.FormFeeCostTypeTooltip} ></Tooltip>
                            </div>
                            <Dropdown
                              selectedKey={costType}
                              onChange={this.onChangeCostType}
                              options={costTypeChoices} />
                          </div>
                        </div>
                        <div className="ms-Grid-col ms-sm4 ms-md4">

                          {/* PO/Non PO */}
                          <div className={styles.field}>
                            <div className={styles.fieldLabel}>
                              <div className={styles.fieldLabelContainer}>
                                <Label >{strings.FormFeePONonPOLabel}</Label>
                              </div>
                              <Tooltip content={strings.FormFeePONonPOTooltip} ></Tooltip>
                            </div>
                            <Dropdown
                              selectedKey={poNonPO}
                              onChange={this.onChangePoNonPo}
                              options={poNonPOChoices} />
                          </div>
                        </div>
                        <div className="ms-Grid-col ms-sm4 ms-md4">
                        </div>
                      </div>
                    </div>

                    {/**Repeating fee */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormFeeRepeatingFeeLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormFeeRecurrenceTooltip} ></Tooltip>
                      </div>
                      <Toggle
                        checked={recurrenceFee}
                        onChange={this.onChangerecurrenceFee}
                      />
                    </div>

                    {recurrenceFee === true && (<RecurrencePatternPicker
                      recurrenceData={recurrencePattern}
                      onChange={this.onChangeRecurrencePattern}
                    >
                    </RecurrencePatternPicker>)}



                    {/**Due date */}
                    {recurrenceFee === false && <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormFeeDueDateLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormFeeDueDateTooltip} ></Tooltip>
                      </div>
                      <DatePicker
                        onSelectDate={this.onchangeDueDate}
                        formatDate={this.formatDate}
                        allowTextInput
                        value={dueDate}
                        styles={datepickerstyles}
                      />
                    </div>

                    }
                    {/* Status */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormSNPStatusLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormFeeStatusTooltip} ></Tooltip>
                      </div>
                      <Dropdown

                        selectedKey={status}
                        onChange={this.onChangeStatus}
                        options={statusChoices} />
                    </div>
                    {/* Comments */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormEutEntityCommentsLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormFeeTooltipComments} ></Tooltip>
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


  private onChangeName = (event, name: string) => {
    this.setState({ name });
  };

  private onChangeSummary = (summary: string) => {
    this.setState({ summary });
    return summary;
  };

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
  /*private onchangeVatFreeCost = (_e: any, vatFreeCost: string) => {
    // Prevent invalid characters: minus, plus, E, e
    if (/[-+Ee]/.test(vatFreeCost)) return this.setState({ vatFreeCost: "" });
    if (vatFreeCost === "") return this.setState({ vatFreeCost: "" });
    const n = Math.max(0, Number(vatFreeCost) || 0);
    this.setState({ vatFreeCost: String(n) });
  };*/
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
  private onChangerecurrenceFee = (event, recurrenceFee: boolean) => {
    // if (recurrenceFee) {
    //   this.setState({ recurrenceFee, dueDate: null });
    // } else {
    //   this.setState({
    //     recurrenceFee,
    //     recurrencePattern: {
    //       frequency: '',
    //       weekDay: null,
    //       month: null,
    //       day: null,
    //       interval: null,
    //       endDate: null,
    //     }
    //   });

    // }
    this.setState({ recurrenceFee });
  };

  private onchangeDueDate = (dueDate: Date) => {
    this.setState({ dueDate });
  };

  private onChangeCostType = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ costType: option.key.toString() });
  };

  private onChangePoNonPo = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ poNonPO: option.key.toString() });
  };


  private onChangeStatus = (event, status: IDropdownOption) => {
    this.setState({ status: status.key.toString() });
  };


  private onChangeRecurrencePattern = (recurrencePattern: RecurrenceData) => {
    this.setState({ recurrencePattern });
  };


  private onChangeChargedByItem = async (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ chargedBy: option.key.toString(), chargeByItem: null });
    const chargeByItemChoices: IDropdownOption[] = await this.getChargedByContentTypeId(option.key.toString());

    this.setState({ chargeByItemChoices });
  };
  private onChangeChargeByItem = (chargeByItem: IDropdownOption) => {
    this.setState({ chargeByItem });
  };

  private onChangeType = async (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ type: option.key.toString() });
    const itemChoices: IDropdownOption[] = await this.getItemsByContentTypeId(option.key.toString());
    this.setState({ itemChoices });
  };

  private onChangeCategory = async (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ category: option.key.toString() });
  };

  private onChangeComments = (comments: string) => {
    this.setState({ comments });
    return comments;
  };

  private onChangeItem = (item: IDropdownOption) => {
    this.setState({ item });
  };

  private getItemsByContentTypeId = async (contentTypeId: string) => {
    const {
      pnpService
    } = this.props;
    const itemChoices: IDropdownOption[] = [];
    const selectedContentType = contentTypeId;
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    switch (selectedContentType) {
      case Consts.CONTENT_TYPES.SNP_FEE: {
        const list = Form.config.find((f) => f.type.toLowerCase() === FormType.SNP).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const items = await pnpService.getListItems(listUrl).select(...selectedFields)
          .expand(Consts.FIELDS.COMMON.CONTENTTYPE)
          .orderBy(Consts.FIELDS.COMMON.TITLE, true).top(5000)();
        items.forEach((item) => {
          itemChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
        });
        const item = null;
        this.setState({ itemChoices, item });
        return itemChoices;
      }
      case Consts.CONTENT_TYPES.MA_REQUIREMENT_FEE: {
        const list = Consts.LISTS.MA_REQUIREMENT_URL;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const items = await pnpService.getListItems(listUrl).select(...selectedFields)
          .expand(Consts.FIELDS.COMMON.CONTENTTYPE)
          .orderBy(Consts.FIELDS.COMMON.TITLE, true).top(5000)();
        items.forEach((item) => {
          itemChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
        });
        const item = null;
        this.setState({ itemChoices, item });
        return itemChoices;
      }
      default: {
        return itemChoices;
      }
    }
  };

  private getChargedByContentTypeId = async (contentTypeText: string) => {
    if (this.AllChargedByItems.has(contentTypeText)) {
      const cachedItems = this.AllChargedByItems.get(contentTypeText) || [];
      return cachedItems;
    }
  };

  private onChangeCurrency = (currency: IDropdownOption) => {
    this.setState({ currency });
  };


  private getDuplicateItems = async (): Promise<IBaseItem[]> => {
    // const {
    //   pnpService,
    //   itemId
    // } = this.props;
    // const {
    //   name
    // } = this.state;
    // const duplicateItems: IBaseItem[] = [];
    // const list = this.config.list;
    // const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    // const selectedFields: string[] = [
    //   Consts.FIELDS.COMMON.ID,
    //   Consts.FIELDS.COMMON.TITLE,
    // ];
    // const items = await pnpService.getListItems(listUrl).select(...selectedFields).top(5000)();
    // const match = items.find((i) => (i.Title?.toLowerCase() === UtilHelper.sanitizeSharePointTitle(name).toLowerCase()) && i.Id !== itemId);
    // if (match)
    //   duplicateItems.push({ Id: match[Consts.FIELDS.COMMON.ID], Title: match[Consts.FIELDS.COMMON.TITLE] });
    // return duplicateItems;
    const duplicateItems = [];
    return duplicateItems;
  };

  private getFormSummary = () => {
    const {
      name,
      summary,
      comments,
      typeChoices,
      type,
      item,
      category,
      categoryChoices,
      chargedBy,
      chargedByChoices,
      chargeByItem,
      status,
      currency,
      vatFreeCost,
      vat,
      vatRate,
      poNonPO,
      costType,
      dueDate,
      recurrencePattern,
      recurrenceFee

    } = this.state;
    const categoryTextValue = categoryChoices.find((choice) => choice.key === category)?.text;
    const typeTextValue = typeChoices.find((choice) => choice.key === type)?.text;
    const chargedByTextValue = chargedByChoices.find((choice) => choice.key === chargedBy)?.text;
    const summaryTextValue = DomHelper.cleanRichHtml(summary);
    const commentsTextValue = DomHelper.cleanRichHtml(comments);
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeDatasheetNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {name}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeCategoryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {categoryTextValue}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeSummaryLabel}</Label>
            </div>
          </div>

          <div className={styles.fieldValue}>
            {summaryTextValue.length > 0 ? <RichText isEditMode={false} value={summaryTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeTypeLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {typeTextValue}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeItemLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {item?.text}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeChargedByLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {chargedByTextValue}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeChargedByItemLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {chargeByItem?.text}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeCurrencyLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {currency?.text}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeVATFreeCostLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {vatFreeCost}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeVATLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {vat ? vat : '0'}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeVATRateLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {vatRate ? vatRate : '0'}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeCostTypeLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {costType}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeePONonPOLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {poNonPO}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeRepeatingFeeLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {recurrenceFee ? 'Yes' : 'No'}
          </div>
        </div>
        {
          recurrenceFee === false && (
            <div className={styles.field}>
              <div className={styles.fieldLabel}>
                <div className={styles.fieldLabelContainer}>
                  <Label>{strings.FormFeeDueDateLabel}</Label>
                </div>
              </div>
              <div className={styles.fieldValue}>
                {dueDate ? this.formatDate(dueDate) : ''}
              </div>
            </div>
          )
        }
        {recurrenceFee && recurrencePattern && (
          <>
            <div className={styles.field}>
              <div className={styles.fieldLabel}>
                <div className={styles.fieldLabelContainer}>
                  <Label>{strings.FormFeeRepeatingFeeLabel}</Label>
                </div>
              </div>
              <div className={styles.fieldValue}>
                {this.formatRecurrencePattern(recurrencePattern)}
              </div>
            </div>

            {/* {recurrencePattern.month && <div className={styles.field}>
              <div className={styles.fieldLabel}>
                <div className={styles.fieldLabelContainer}>
                  <Label>{strings.MonthFieldLabel}</Label>
                </div>
              </div>
              <div className={styles.fieldValue}>
                {recurrencePattern.month ?? ''}
              </div>
            </div>
            }
            {recurrencePattern.day && <div className={styles.field}>
              <div className={styles.fieldLabel}>
                <div className={styles.fieldLabelContainer}>
                  <Label>{strings.RecurrenceDayLabel}</Label>
                </div>
              </div>
              <div className={styles.fieldValue}>
                {recurrencePattern.day ?? ''}
              </div>
            </div>
            } */}
          </>
        )}


        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPStatusLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {status}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormFeeCommentsLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {commentsTextValue.length > 0 ? <RichText isEditMode={false} value={commentsTextValue} /> : ''}
          </div>
        </div>

      </div>
    );
  };

  private callback = async () => {
    const {
      callback,
    } = this.props;
    callback(false);
  };

  private processCreation = async (): Promise<IItemAddResult> => {
    const {
      summary,
      comments,
      typeChoices,
      type,
      status,
      name,
      item,
      chargedBy,
      chargeByItem,
      vatFreeCost,
      vat,
      vatRate,
      costType,
      poNonPO,
      currency,
      dueDate,
      category,
      recurrenceFee,
      recurrencePattern

    } = this.state;

    const {
      pnpService,
    } = this.props;


    const itemData = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = name;
    itemData[Consts.FIELDS.FEE.CATEGORY] = category;

    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = typeChoices.find((choice) => choice.key === type)?.key;
    itemData[Consts.FIELDS.FEE.SUMMARY] = DomHelper.cleanRichHtml(summary);

    if (type === Consts.CONTENT_TYPES.SNP_FEE) {
      itemData[Consts.FIELDS.FEE.SNP_ID] = item.key;
    } else {
      itemData[Consts.FIELDS.FEE.MA_REQUIREMENT_ID] = item.key;
    }
    itemData[Consts.FIELDS.FEE.SUMMARY] = DomHelper.cleanRichHtml(summary);
    if (vatFreeCost) {
      itemData[Consts.FIELDS.FEE.VATFREECOST] = vatFreeCost;
    }
    itemData[Consts.FIELDS.FEE.DUEDATE] = dueDate?.toISOString();
    if (vat) {
      itemData[Consts.FIELDS.FEE.VAT] = vat;
    }
    if (vatRate) {
      itemData[Consts.FIELDS.FEE.VATRATE] = vatRate;
    }
    itemData[Consts.FIELDS.FEE.COSTTYPE] = costType;
    itemData[Consts.FIELDS.FEE.PONONPO] = poNonPO;

    itemData[Consts.FIELDS.FEE.STATUS] = status;
    itemData[Consts.FIELDS.FEE.COMMENTS] = DomHelper.cleanRichHtml(comments);
    itemData[Consts.FIELDS.FEE.CURRENCY_ID] = currency.key;
    itemData[Consts.FIELDS.FEE.CHARGEDBY] = chargedBy;
    itemData[Consts.FIELDS.FEE.REPEATINGFEE] = recurrenceFee;
    switch (chargedBy) {
      // Group 1: Law Firm & Legal Representative
      case FeeChargedBy.LawFirm:
      case FeeChargedBy.LegalRepresentative:
        itemData[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER_ID] = chargeByItem.key;
        break;

      // Group 2: Organization & Regulator
      case FeeChargedBy.Administration:
      case FeeChargedBy.Organization:
      case FeeChargedBy.Regulator:
        itemData[Consts.FIELDS.FEE.AUTHORITY_ID] = chargeByItem.key;
        break;

      // Group 4: Teleport Partner
      case FeeChargedBy.TeleportPartner:
        itemData[Consts.FIELDS.FEE.TELEPORTPARTNER_ID] = chargeByItem.key;
        break;

      // Default: do nothing
      default:
        break;
    }


    const recurrenceListUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.RECURRENCE_PATTERN_URL);
    let recurrenceItemId;
    if (recurrenceFee) {
      const recurrenceItemData = {
        [Consts.FIELDS.COMMON.TITLE]: 'Recurrence Entry',
        [Consts.FIELDS.RECURRENCE_PATTERN.FREQUENCY]: recurrencePattern.frequency,
        [Consts.FIELDS.RECURRENCE_PATTERN.WEEKDAY]: recurrencePattern.weekDay,
        [Consts.FIELDS.RECURRENCE_PATTERN.MONTH]: recurrencePattern.month,
        [Consts.FIELDS.RECURRENCE_PATTERN.DAY]: recurrencePattern.day,
        [Consts.FIELDS.RECURRENCE_PATTERN.INTERVAL]: recurrencePattern.interval,
        [Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE]: recurrencePattern.endDate?.toISOString(),
      };
      const createdRecurrenceItem: IItemAddResult = await pnpService
        .getListByUrl(recurrenceListUrl)
        .items.add(recurrenceItemData);

      recurrenceItemId = createdRecurrenceItem.data.ID;
      itemData[Consts.FIELDS.FEE.RECURRENCEPATTERN_ID] = recurrenceItemId;
    }

    //const selectedContentType = this.typeChoices.find((type) => itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1);
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);



    if (recurrenceFee && recurrenceItemId) {
      // Step 3: Update Recurrence Item with Fee ID
      await pnpService
        .getListByUrl(recurrenceListUrl)
        .items.getById(recurrenceItemId)
        .update({
          Title: `${this.config.type}-${createdItem.data.ID}`, // Make sure 'FeeID' is a valid field in Recurrence list
        });
    }

    const folderUrlMap: Record<string, string> = {
      [Consts.CONTENT_TYPES.SNP_FEE]: Consts.DOCUMENT_SPACE.FEE_SNP_FOLDER_URL,
      [Consts.CONTENT_TYPES.MA_REQUIREMENT_FEE]: Consts.DOCUMENT_SPACE.FEE_MA_REQUIREMENT_fOLDER_URL,
    };


    const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
    const documentSpaceUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.DOCUMENT_SPACE_URL);
    const parentFolderUrl = `${documentSpaceUrl}${folderUrlMap[type]}`;
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
      typeChoices,
      type,
      status,
      item,
      chargedBy,
      chargeByItem,
      vatFreeCost,
      vat,
      vatRate,
      costType,
      poNonPO,
      currency,
      dueDate,
      category,
      recurrenceFee,
      recurrencePattern,
      recurrencPatternId
    } = this.state;

    const {
      pnpService,
      itemId
    } = this.props;

    const itemData = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = name;
    itemData[Consts.FIELDS.FEE.CATEGORY] = category;
    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = typeChoices.find((choice) => choice.key === type)?.key;
    itemData[Consts.FIELDS.FEE.SUMMARY] = DomHelper.cleanRichHtml(summary);

    if (type === Consts.CONTENT_TYPES.SNP_FEE) {
      itemData[Consts.FIELDS.FEE.SNP_ID] = item.key;
    } else {
      itemData[Consts.FIELDS.FEE.MA_REQUIREMENT_ID] = item.key;
    }

    if (vatFreeCost) itemData[Consts.FIELDS.FEE.VATFREECOST] = vatFreeCost;
    else itemData[Consts.FIELDS.FEE.VATFREECOST] = '0';
    if (vat) itemData[Consts.FIELDS.FEE.VAT] = vat;
    else itemData[Consts.FIELDS.FEE.VAT] = '0';
    if (vatRate) itemData[Consts.FIELDS.FEE.VATRATE] = vatRate;
    else itemData[Consts.FIELDS.FEE.VATRATE] = '0';

    itemData[Consts.FIELDS.FEE.COSTTYPE] = costType;
    itemData[Consts.FIELDS.FEE.PONONPO] = poNonPO;
    itemData[Consts.FIELDS.FEE.STATUS] = status;
    itemData[Consts.FIELDS.FEE.COMMENTS] = DomHelper.cleanRichHtml(comments);
    itemData[Consts.FIELDS.FEE.CURRENCY_ID] = currency.key;
    itemData[Consts.FIELDS.FEE.CHARGEDBY] = chargedBy;
    itemData[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER_ID] = null;
    itemData[Consts.FIELDS.FEE.AUTHORITY_ID] = null;
    itemData[Consts.FIELDS.FEE.TELEPORTPARTNER_ID] = null;
    itemData[Consts.FIELDS.FEE.DUEDATE] = dueDate?.toISOString();
    itemData[Consts.FIELDS.FEE.REPEATINGFEE] = recurrenceFee;

    switch (chargedBy) {
      case FeeChargedBy.LawFirm:
      case FeeChargedBy.LegalRepresentative:
        itemData[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER_ID] = chargeByItem.key;
        break;
      case FeeChargedBy.Administration:
      case FeeChargedBy.Organization:
      case FeeChargedBy.Regulator:
        itemData[Consts.FIELDS.FEE.AUTHORITY_ID] = chargeByItem.key;
        break;
      case FeeChargedBy.TeleportPartner:
        itemData[Consts.FIELDS.FEE.TELEPORTPARTNER_ID] = chargeByItem.key;
        break;
      default:
        break;
    }

    const recurrenceListUrl = UrlHelper.getListWebRelativeUrl(
      pnpService.getUrl(),
      Consts.LISTS.RECURRENCE_PATTERN_URL
    );

    const recurrenceItemData = {
      [Consts.FIELDS.COMMON.TITLE]: 'Recurrence Entry',
      [Consts.FIELDS.RECURRENCE_PATTERN.FREQUENCY]: recurrencePattern?.frequency,
      [Consts.FIELDS.RECURRENCE_PATTERN.WEEKDAY]: recurrencePattern?.weekDay,
      [Consts.FIELDS.RECURRENCE_PATTERN.MONTH]: recurrencePattern?.month,
      [Consts.FIELDS.RECURRENCE_PATTERN.DAY]: recurrencePattern?.day,
      [Consts.FIELDS.RECURRENCE_PATTERN.INTERVAL]: recurrencePattern?.interval,
      [Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE]: recurrencePattern?.endDate?.toISOString(),
    };

    if (recurrenceFee && recurrencPatternId) {
      await pnpService
        .getListByUrl(recurrenceListUrl)
        .items.getById(recurrencPatternId)
        .update(recurrenceItemData);
    } else if (recurrenceFee && !recurrencPatternId) {
      const createdRecurrenceItem: IItemAddResult = await pnpService
        .getListByUrl(recurrenceListUrl)
        .items.add(recurrenceItemData);
      const recurrenceItemId = createdRecurrenceItem.data.ID;
      itemData[Consts.FIELDS.FEE.RECURRENCEPATTERN_ID] = recurrenceItemId;
    }
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);

  };

}
