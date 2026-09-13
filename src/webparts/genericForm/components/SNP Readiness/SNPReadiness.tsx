import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { ISNPReadinessProps, ISNPReadinessState } from './SNPReadiness.types';
import styles from '../Form.module.scss';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import FormWizard from '../Form Wizard/FormWizard';
import { Dropdown, Icon, IDropdownOption, Label, DatePicker } from '@fluentui/react';
import { Combobox } from 'react-widgets';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { Form, FormOrigin, FormType, IFormConfig } from '../Form.types';
import strings from 'GenericFormWebPartStrings';
import { RichText } from '../../../../common/components/RichText/RichText';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { UtilHelper } from '../../../../common/helpers/Util';
import { Consts } from '../../../../common/consts/Consts';
import 'react-widgets/styles.css';
import 'react-phone-number-input/style.css';
import { IItemAddResult } from '@pnp/sp/items';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { MarketReadinessResponsibleParty, UserRole } from '../../../../common/models/Enums';
import { datepickerstyles } from '../Form.utility';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';
export default class SNPMarketReadiness extends React.Component<
  ISNPReadinessProps,
  ISNPReadinessState
> {
  private isEditMode: boolean;
  private config: IFormConfig;
  private AllLookupItems: Map<string, IDropdownOption[]> = new Map();
  constructor(props: Readonly<ISNPReadinessProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.AllLookupItems = new Map();
    this.state = {
      name: '',
      responsiblePartyTypeChoices: [],
      responsiblePartyType: '',
      readinessRAGStatus: '',
      readinessRAGStatusChoices: [],
      errors: {},
      isFormReady: false,
      responsiblePartyItem: null,
      responsiblePartyItemChoices: [],
      eutelsatOwnersChoices: [],
      eutelsatOwners: [],
      readinessEstimatedDate: null,
      readinessEffectiveDate: null,
      readinessComments: '',
      snpItem: null,
      snpItemChoices: [],
    };
    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.SNPMarketReadiness);
  }
  private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
    const userRoles = this.props.userService.userContext?.userRoles ?? [];
    return rolesToCheck.some((role) => userRoles.includes(role));
  };
  public async componentDidMount() {
    if (this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin)) {
      await this.init();
      this.setState({ isFormReady: true });
    }
  }

  private init = async () => {
    const promises = [];
    await this.fetchAllLookupItems();

    promises.push(
      this.initResponsiblePartyChoices()
    );
    promises.push(
      this.initRAGStatusChoices()
    );

    promises.push(
      this.initEutelsatOwnersChoices()
    );

    promises.push(
      this.initSNPItem()
    );
    await Promise.all(promises);

    if (this.isEditMode)
      await this.setInitialFormValues();
    this.setDefaultItemChoices();
  };

  public initResponsiblePartyChoices = async () => {
    const {
      pnpService
    } = this.props;
    let responsiblePartyTypeChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.SNP_MA_READINESS.RESPONSIBLE_PARTY)
      .select('Choices')();
    responsiblePartyTypeChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    const responsiblePartyType = responsiblePartyTypeChoices.length > 0 ? responsiblePartyTypeChoices[0].key as string : '';
    this.setState({
      responsiblePartyTypeChoices, responsiblePartyType
    });
  };

  public initRAGStatusChoices = async () => {
    const {
      pnpService
    } = this.props;
    let readinessRAGStatusChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.SNP_MA_READINESS.RAGSTATUS)
      .select('Choices')();
    data.Choices.sort((a: string, b: string) => a?.localeCompare(b));
    readinessRAGStatusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    const readinessRAGStatus = readinessRAGStatusChoices.length > 0 ? readinessRAGStatusChoices[0].key as string : '';
    this.setState({
      readinessRAGStatusChoices, readinessRAGStatus
    });
  };

  public initEutelsatOwnersChoices = async () => {
    const {
      pnpService
    } = this.props;
    const eutelsatOwnersChoices: IDropdownOption[] = [];
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Contact).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields)
      .expand(Consts.FIELDS.COMMON.CONTENTTYPE).filter(`startswith(${Consts.FIELDS.COMMON.CONTENTTYPE_ID}, '${Consts.CONTENT_TYPES.EUTELSAL_CONTACT}')`).
      orderBy(Consts.FIELDS.COMMON.TITLE, true).top(5000)();
    items.forEach((item) => {
      eutelsatOwnersChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    const eutelsatOwners = [];
    this.setState({ eutelsatOwnersChoices, eutelsatOwners });
  };

  private initSNPItem = async () => {
    if (!this.isEditMode) {
      const {
        pnpService,
        preselectedItemId
      } = this.props;
      const formConfigData = Form.config.find((f) => f.type.toLowerCase() === FormType.SNP);
      const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), formConfigData.list);
      const selectedFields: string[] = [
        Consts.FIELDS.COMMON.ID,
        Consts.FIELDS.COMMON.TITLE
      ];

      const snpItemChoices: IDropdownOption[] = [];
      const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)
        .filter(`Id eq ${preselectedItemId}`)
        .top(1)();
      items.forEach((item) => {
        snpItemChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
      });
      const snpItem: IDropdownOption = snpItemChoices.length > 0 ? snpItemChoices[0] : null;
      this.setState({
        snpItemChoices, snpItem
      });
    }
  };

  private setDefaultItemChoices = () => {
    const { responsiblePartyType } = this.state;
    const responiblePartyItems = this.AllLookupItems.get(responsiblePartyType) || [];
    this.setState({ responsiblePartyItemChoices: responiblePartyItems });
  };

  private validateForm = async (): Promise<boolean> => {
    const errors: { [key: string]: string } = {};
    const {
      responsiblePartyType,
      responsiblePartyItem,
      eutelsatOwners,
      snpItem
    } = this.state;
    const { pnpService, preselectedItemId } = this.props;

    if (![MarketReadinessResponsibleParty.Eutelsat, MarketReadinessResponsibleParty.Overall].includes(responsiblePartyType as MarketReadinessResponsibleParty) && !responsiblePartyItem) {
      errors.responsiblePartyItem = strings.FormSNPReadinessItemValidationError;
    }

    if (!Array.isArray(eutelsatOwners) || eutelsatOwners.length === 0) {
      errors.eutelsatOwners = strings.FormSNPReadinessEutelsatOwnersValidationError;
    }

    if (!errors.responsiblePartyType && this.isEditMode === false && preselectedItemId &&
      (responsiblePartyType === MarketReadinessResponsibleParty.Overall || responsiblePartyType === MarketReadinessResponsibleParty.Eutelsat)) {
      const list = this.config.list;
      //const exists = await pnpService.get(fileUrlDocument);
      const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);

      // Check if item exists
      const existingItems = await pnpService.getListItems(listUrl).
        filter(`${Consts.FIELDS.SNP_MA_READINESS.SNP_ID} eq ${snpItem?.key} and ${Consts.FIELDS.SNP_MA_READINESS.RESPONSIBLE_PARTY} eq '${responsiblePartyType}'`)();
      if (existingItems.length > 0) {
        errors.responsiblePartyType = responsiblePartyType + " " + strings.FormSNPReadinessOverAllandEtulsatValidationError;
      }

    }
    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

  public render(): React.ReactElement<ISNPReadinessProps> {
    const {
      isFormReady,
      responsiblePartyTypeChoices,
      responsiblePartyType,
      responsiblePartyItemChoices,
      responsiblePartyItem,
      errors,
      eutelsatOwnersChoices,
      eutelsatOwners,
      readinessRAGStatus,
      readinessRAGStatusChoices,
      readinessEstimatedDate,
      readinessEffectiveDate,
      readinessComments,
      snpItemChoices,
      snpItem
    } = this.state;
    return (
      <div className="ms-Grid-row">
        <div className="ms-Grid-col ms-sm12 ms-md12">
          {!this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && (
            <AccessDeniedMessage message={strings.FormAccessDeniedMessage}> </AccessDeniedMessage>
          )}
          {/* Form  */}
          {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <FormWizard
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
            formOrigin={FormOrigin.Datasheet}>
            <>
              {!isFormReady &&
                <div className={styles.spinnerContainer}>
                  <Spinner size={SpinnerSize.large} label={'Loading...'} />
                </div>
              }
              {isFormReady &&
                <form>
                  {/* SNP */}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label>{strings.FormSNPParentSNPLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormSNPTooltipParentSNP} ></Tooltip>
                    </div>
                    <Combobox
                      disabled={true}
                      data={snpItemChoices}
                      value={snpItem}
                      selectIcon={
                        <span className="ms-Dropdown-caretDownWrapper">
                          <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
                        </span>
                      }
                      textField="text"
                      filter="contains"
                      key="key"
                    />
                  </div>

                  {/* Responsible party */}
                  <div className={styles.field} >
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label>{strings.FormSNPResponsiblePartyLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormSNPTooltipResponsibleParty} ></Tooltip>
                    </div>
                    <Dropdown
                      selectedKey={responsiblePartyType}
                      onChange={this.onChangeResponsibleParty}
                      options={responsiblePartyTypeChoices}
                      disabled={this.isEditMode} />
                    {errors.responsiblePartyType &&
                      <div className={styles.errorContainer}>
                        <Icon iconName="Error" className={styles.errorIcon} />
                        <span className={styles.errorMessage}>
                          {errors.responsiblePartyType}
                        </span>
                      </div>
                    }
                  </div>
                  {/* Item */}
                  {responsiblePartyType !== MarketReadinessResponsibleParty.Overall && responsiblePartyType !== MarketReadinessResponsibleParty.Eutelsat &&
                    <div className={styles.field} >
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormContactItemLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormContactTooltipItem} ></Tooltip>
                      </div>
                      <Combobox
                        data={responsiblePartyItemChoices}
                        value={responsiblePartyItem}
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
                      {errors.responsiblePartyItem &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.responsiblePartyItem}
                          </span>
                        </div>
                      }
                    </div>
                  }
                  {/* Eutelsat Owners*/}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label required>{strings.FormSNPEutelsatOwnersLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormSNPTooltipEutelsatOwners} ></Tooltip>
                    </div>
                    <MultiselectWrapper
                      value={eutelsatOwners}
                      data={eutelsatOwnersChoices}
                      dataKey={(item: IDropdownOption) => item.key}
                      textField={(item: IDropdownOption) => item.text}
                      filter="contains"
                      onChange={this.onChangeEutelsatOwners} />
                    {errors.eutelsatOwners &&
                      <div className={styles.errorContainer}>
                        <Icon iconName="Error" className={styles.errorIcon} />
                        <span className={styles.errorMessage}>
                          {errors.eutelsatOwners}
                        </span>
                      </div>
                    }
                  </div>
                  {/*RAG Status*/}
                  <div className={styles.field} >
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label>{strings.FormSNPRAGStatusLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormSNPTooltipRAGStatus} ></Tooltip>
                    </div>
                    <Dropdown
                      selectedKey={readinessRAGStatus}
                      onChange={this.onChangeRAGStatus}
                      options={readinessRAGStatusChoices} />
                  </div>
                  {/* Estimated Date */}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label>{strings.FormSNPEstimatedDateLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormSNPTooltipEstimatedDate} ></Tooltip>
                    </div>
                    <DatePicker
                      onSelectDate={this.onChangeEstimatedDate}
                      value={readinessEstimatedDate}
                      styles={datepickerstyles}
                      formatDate={this.formatDate}
                      allowTextInput
                    />
                  </div>
                  {/* Effective Date */}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label>{strings.FormSNPEffectiveDateLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormSNPTooltipEffectiveDate} ></Tooltip>
                    </div>
                    <DatePicker
                      allowTextInput
                      onSelectDate={this.onChangeEffectiveDate}
                      value={readinessEffectiveDate}
                      styles={datepickerstyles}
                      formatDate={this.formatDate}
                    />
                  </div>
                  {/* Comments */}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label>{strings.FormSNPCommentsLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormSNPTooltipComments} ></Tooltip>
                    </div>
                    <RichText isEditMode={true} value={readinessComments} onChange={this.onChangeComments} />
                  </div>
                </form>
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

  private getFormSummary = () => {
    const {
      snpItem,
      responsiblePartyType,
      responsiblePartyItem,
      eutelsatOwners,
      readinessRAGStatus,
      readinessEstimatedDate,
      readinessEffectiveDate,
      readinessComments
    } = this.state;
    const mareadinessText = Consts.NAMING.MA_SNP_READINESS_SUFFIX;
    const readinessCommentsTextValue = DomHelper.cleanRichHtml(readinessComments);
    let displayText = "";
    if (responsiblePartyType === MarketReadinessResponsibleParty.Overall || responsiblePartyType === MarketReadinessResponsibleParty.Eutelsat) {
      displayText = `${snpItem?.text} - ${responsiblePartyType} ${mareadinessText}`;
    } else if (responsiblePartyType === MarketReadinessResponsibleParty.DP) {
      displayText = `${snpItem?.text} - ${responsiblePartyItem?.text} ${mareadinessText}`;
    } else if (responsiblePartyType === MarketReadinessResponsibleParty.TP) {
      displayText = `${snpItem?.text} - ${responsiblePartyItem?.text} ${mareadinessText}`;
    }
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPDatasheetNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPReadinessNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {displayText}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPParentSNPLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {snpItem?.text}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPResponsiblePartyLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {responsiblePartyType}
          </div>
        </div>
        {responsiblePartyType !== MarketReadinessResponsibleParty.Overall && responsiblePartyType !== MarketReadinessResponsibleParty.Eutelsat &&
          <div className={styles.field}>
            <div className={styles.fieldLabel}>
              <div className={styles.fieldLabelContainer}>
                <Label>{strings.FormSNPResponsibleItemLabel}</Label>
              </div>
            </div>
            <div className={styles.fieldValue}>
              {responsiblePartyItem?.text}
            </div>
          </div>
        }
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPEutelsatOwnersLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {eutelsatOwners.map((c) => c.text).join(", ")}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPRAGStatusLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {readinessRAGStatus}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPEstimatedDateLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.formatDate(readinessEstimatedDate)}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPEffectiveDateLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.formatDate(readinessEffectiveDate)}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPCommentsLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {readinessCommentsTextValue.length > 0 ? <RichText isEditMode={false} value={readinessCommentsTextValue} /> : ''}
          </div>
        </div>
      </div>
    );
  };

  private processCreation = async (): Promise<IItemAddResult> => {
    const {
      snpItem,
      responsiblePartyType,
      responsiblePartyItem,
      eutelsatOwners,
      readinessRAGStatus,
      readinessEstimatedDate,
      readinessEffectiveDate,
      readinessComments
    } = this.state;

    const {
      pnpService,
    } = this.props;
    const itemData = {};
    const mareadinessText = Consts.NAMING.MA_SNP_READINESS_SUFFIX;
    let displayText = "";
    if (responsiblePartyType === MarketReadinessResponsibleParty.Overall || responsiblePartyType === MarketReadinessResponsibleParty.Eutelsat) {
      displayText = `${snpItem?.text} - ${responsiblePartyType} ${mareadinessText}`;
    } else if (responsiblePartyType === MarketReadinessResponsibleParty.DP) {
      displayText = `${snpItem?.text} - ${responsiblePartyItem?.text} ${mareadinessText}`;
    } else if (responsiblePartyType === MarketReadinessResponsibleParty.TP) {
      displayText = `${snpItem?.text} - ${responsiblePartyItem?.text} ${mareadinessText}`;
    }
    itemData[Consts.FIELDS.COMMON.TITLE] = displayText;
    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = Consts.CONTENT_TYPES.MA_READINESS_SNP;
    itemData[Consts.FIELDS.SNP_MA_READINESS.SNP_ID] = snpItem?.key;
    itemData[Consts.FIELDS.SNP_MA_READINESS.RESPONSIBLE_PARTY] = responsiblePartyType;
    if (responsiblePartyType === MarketReadinessResponsibleParty.DP)
      itemData[Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER_ID] = responsiblePartyItem?.key;
    if (responsiblePartyType === MarketReadinessResponsibleParty.TP)
      itemData[Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER_ID] = responsiblePartyItem?.key;
    itemData[Consts.FIELDS.SNP_MA_READINESS.RAGSTATUS] = readinessRAGStatus;
    itemData[Consts.FIELDS.SNP_MA_READINESS.OWNERS_ID] = eutelsatOwners.map((c) => c.key);
    itemData[Consts.FIELDS.SNP_MA_READINESS.ESTIMATED_DATE] = readinessEstimatedDate?.toISOString();
    itemData[Consts.FIELDS.SNP_MA_READINESS.EFFECTIVE_DATE] = readinessEffectiveDate?.toISOString();
    itemData[Consts.FIELDS.SNP_MA_READINESS.COMMENTS] = readinessComments;

    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);
    return createdItem;
  };

  private processEdit = async () => {
    const {
      snpItem,
      responsiblePartyType,
      responsiblePartyItem,
      eutelsatOwners,
      readinessRAGStatus,
      readinessEstimatedDate,
      readinessEffectiveDate,
      readinessComments

    } = this.state;

    const {
      pnpService,
      itemId,
    } = this.props;


    const mareadinessText = Consts.NAMING.MA_SNP_READINESS_SUFFIX;
    let displayText = "";
    if (responsiblePartyType === MarketReadinessResponsibleParty.Overall || responsiblePartyType === MarketReadinessResponsibleParty.Eutelsat) {
      displayText = `${snpItem?.text} - ${responsiblePartyType} ${mareadinessText}`;
    } else if (responsiblePartyType === MarketReadinessResponsibleParty.DP) {
      displayText = `${snpItem?.text} - ${responsiblePartyItem?.text} ${mareadinessText}`;
    } else if (responsiblePartyType === MarketReadinessResponsibleParty.TP) {
      displayText = `${snpItem?.text} - ${responsiblePartyItem?.text} ${mareadinessText}`;
    }
    const itemData = {};

    itemData[Consts.FIELDS.COMMON.TITLE] = displayText;
    itemData[Consts.FIELDS.SNP_MA_READINESS.SNP_ID] = snpItem?.key;
    itemData[Consts.FIELDS.SNP_MA_READINESS.RESPONSIBLE_PARTY] = responsiblePartyType;
    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = Consts.CONTENT_TYPES.MA_READINESS_SNP;
    if (responsiblePartyType === MarketReadinessResponsibleParty.TP) {
      itemData[Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER_ID] = responsiblePartyItem?.key;
    } else if (responsiblePartyType === MarketReadinessResponsibleParty.DP) {
      itemData[Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER_ID] = responsiblePartyItem?.key;
    }
    itemData[Consts.FIELDS.SNP_MA_READINESS.OWNERS_ID] =
      eutelsatOwners.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.SNP_MA_READINESS.RAGSTATUS] = readinessRAGStatus;
    itemData[Consts.FIELDS.SNP_MA_READINESS.ESTIMATED_DATE] = readinessEstimatedDate;
    itemData[Consts.FIELDS.SNP_MA_READINESS.EFFECTIVE_DATE] = readinessEffectiveDate;
    itemData[Consts.FIELDS.SNP_MA_READINESS.COMMENTS] = readinessComments;
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);

  };

  private onChangeEutelsatOwners = (eutelsatOwners: IDropdownOption[]) => {
    this.setState({ eutelsatOwners });
  };

  private onChangeResponsibleParty = async (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ responsiblePartyType: option.key.toString() });
    if (option.key.toString() === MarketReadinessResponsibleParty.DP || option.key.toString() === MarketReadinessResponsibleParty.TP) {
      const typeChoices = await this.getItemsByContentTypeId(option.text.toString());
      this.setState({ responsiblePartyItemChoices: typeChoices, responsiblePartyItem: null });
    }


  };

  private onChangeItem = (responsiblePartyItem: IDropdownOption) => {
    this.setState({ responsiblePartyItem });
  };

  private onChangeComments = (readinessComments: string) => {
    this.setState({ readinessComments });
    return readinessComments;
  };

  private onChangeRAGStatus = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ readinessRAGStatus: option.key.toString() });
  };

  private onChangeEstimatedDate = (readinessEstimatedDate: Date) => {
    this.setState({ readinessEstimatedDate });
    return readinessEstimatedDate;
  };

  private onChangeEffectiveDate = (readinessEffectiveDate: Date) => {
    this.setState({ readinessEffectiveDate });
    return readinessEffectiveDate;
  };

  private fetchAllLookupItems = async () => {
    const { pnpService } = this.props;
    const selectedFields = [Consts.FIELDS.COMMON.ID, Consts.FIELDS.COMMON.TITLE];

    const configMap = [
      { key: MarketReadinessResponsibleParty.TP, type: FormType.TP, contentTypeId: null },
      { key: MarketReadinessResponsibleParty.DP, type: FormType.DP, contentTypeId: null }
    ];

    const fetchPromises = configMap.map(async (config) => {
      const list = Form.config.find((f) => f.type.toLowerCase() === config.type.toLowerCase())?.list;
      const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
      let query = pnpService.getListItems(listUrl).select(...selectedFields).expand(Consts.FIELDS.COMMON.CONTENTTYPE);

      if (config.contentTypeId) {
        query = query.filter(`startswith(${Consts.FIELDS.COMMON.CONTENTTYPE_ID}, '${config.contentTypeId}')`);
      }

      const items = await query.orderBy(Consts.FIELDS.COMMON.TITLE, true).top(5000)();
      const choices = items.map((item) => ({
        key: item[Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.COMMON.TITLE]
      }));

      return { key: config.key, choices };
    });

    const results = await Promise.all(fetchPromises);

    results.forEach(({ key, choices }) => {
      this.AllLookupItems.set(key, choices);
    });


  };

  private getItemsByContentTypeId = async (contentTypeText: string) => {
    if (this.AllLookupItems.has(contentTypeText)) {
      const cachedItems = this.AllLookupItems.get(contentTypeText) || [];
      return cachedItems;
    }
  };

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.SNPMarketReadiness).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.SNP_MA_READINESS.RESPONSIBLE_PARTY,
      Consts.FIELDS.SNP_MA_READINESS.RAGSTATUS,
      Consts.FIELDS.SNP_MA_READINESS.ESTIMATED_DATE,
      Consts.FIELDS.SNP_MA_READINESS.EFFECTIVE_DATE,
      Consts.FIELDS.SNP_MA_READINESS.COMMENTS,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      `${Consts.FIELDS.SNP_MA_READINESS.SNP}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.SNP_MA_READINESS.SNP}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.SNP_MA_READINESS.OWNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.SNP_MA_READINESS.OWNER}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const expand = [Consts.FIELDS.SNP_MA_READINESS.SNP, Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER, Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER, Consts.FIELDS.SNP_MA_READINESS.OWNER];
    const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
    let snpItem: IDropdownOption = null;
    snpItem = {
      key: item[Consts.FIELDS.SNP_MA_READINESS.SNP][Consts.FIELDS.COMMON.ID],
      text: item[Consts.FIELDS.SNP_MA_READINESS.SNP][Consts.FIELDS.COMMON.TITLE],
    };

    let responsiblePartyItem: IDropdownOption | null = null;

    const responsiblePartyType: string = item[Consts.FIELDS.SNP_MA_READINESS.RESPONSIBLE_PARTY];
    if (responsiblePartyType === MarketReadinessResponsibleParty.TP) {
      responsiblePartyItem = {
        key: item[Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER][Consts.FIELDS.COMMON.TITLE],
      };
    }
    if (responsiblePartyType === MarketReadinessResponsibleParty.DP) {
      responsiblePartyItem = {
        key: item[Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE],
      };
    }
    const rawReadinessEstimatedDate = item[Consts.FIELDS.SNP_MA_READINESS.ESTIMATED_DATE];
    const readinessEstimatedDate: Date | null = rawReadinessEstimatedDate ? new Date(rawReadinessEstimatedDate) : null;
    const rawReadinessEffectiveDate = item[Consts.FIELDS.SNP_MA_READINESS.EFFECTIVE_DATE];
    const readinessEffectiveDate: Date | null = rawReadinessEffectiveDate ? new Date(rawReadinessEffectiveDate) : null;
    const initialValues = {
      name: item[Consts.FIELDS.COMMON.TITLE],
      snpItem,
      responsiblePartyType,
      responsiblePartyItem,
      eutelsatOwners: item[Consts.FIELDS.SNP_MA_READINESS.OWNER]?.map((etuowner) => ({
        key: etuowner[Consts.FIELDS.COMMON.ID],
        text: etuowner[Consts.FIELDS.COMMON.TITLE]
      })) || [],
      readinessRAGStatus: item[Consts.FIELDS.SNP_MA_READINESS.RAGSTATUS],
      readinessEstimatedDate,
      readinessEffectiveDate,
      readinessComments: item[Consts.FIELDS.SNP_MA_READINESS.COMMENTS]
    };
    this.setState({
      ...initialValues
    });
  };

  private getDuplicateItems = async (): Promise<IBaseItem[]> => {
    const duplicateItems = [];
    return duplicateItems;
  };

  private callback = async (redirectUrl) => {
    const {
      callback,
    } = this.props;
    callback(redirectUrl);
  };

}
