import * as React from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import { ILegalSvcProviderProps, ILegalSvcProviderState } from './LegalSvcProvider.types';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import { Dropdown, Fabric, Icon, IDropdownOption, Label, TextField } from 'office-ui-fabric-react';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { Form, FormOrigin, FormType, IFormConfig } from '../Form.types';
import strings from 'GenericFormWebPartStrings';
import { RichText } from '../../../../common/components/RichText/RichText';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { UtilHelper } from '../../../../common/helpers/Util';
import { Consts } from '../../../../common/consts/Consts';
import 'react-widgets/styles.css';
import { IItemAddResult } from '@pnp/sp/items';
import { createUniqueFolder } from '../Form.utility';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { LegalSvcProviderType, UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';
export default class LegalSvcProvider extends React.Component<
  ILegalSvcProviderProps,
  ILegalSvcProviderState
> {
  private isEditMode: boolean;
  private config: IFormConfig;
  private typeChoices: IDropdownOption[] = [
    { key: Consts.CONTENT_TYPES.LAW_FIRM, text: LegalSvcProviderType.LawFirm },
    { key: Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE, text: LegalSvcProviderType.LegalRepresentative }
  ];

  constructor(props: Readonly<ILegalSvcProviderProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      name: '',
      summary: '',
      address: '',
      comments: '',
      status: '',
      statusChoices: [],
      countries: [],
      countryChoices: [],
      errors: {},
      isFormReady: false,
      typeChoices: [],
      type: '',
      eutelsatOwnersChoices: [],
      eutelsatOwners: []
    };
    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.LegalServiceProvider);
    this.validateForm = this.validateForm.bind(this);
    this.getDuplicateItems = this.getDuplicateItems.bind(this);
    this.getFormSummary = this.getFormSummary.bind(this);
    this.callback = this.callback.bind(this);
    this.processCreation = this.processCreation.bind(this);
    this.processEdit = this.processEdit.bind(this);
    this.initStatusChoices = this.initStatusChoices.bind(this);
    this.initCountryChoices = this.initCountryChoices.bind(this);
    this.initTypeChoices = this.initTypeChoices.bind(this);
    this.initEutelsatOwnersChoices = this.initEutelsatOwnersChoices.bind(this);
    this.onChangeName = this.onChangeName.bind(this);
    this.onChangeStatus = this.onChangeStatus.bind(this);
    this.onChangeCountry = this.onChangeCountry.bind(this);
    this.onChangeSummary = this.onChangeSummary.bind(this);
    this.onChangeAddress = this.onChangeAddress.bind(this);
    this.onChangeComments = this.onChangeComments.bind(this);
    this.onChangeType = this.onChangeType.bind(this);
    this.onChangeEutelsatOwners = this.onChangeEutelsatOwners.bind(this);

  }
  private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
    const userRoles = this.props.userService.userContext?.userRoles ?? [];
    return rolesToCheck.some((role) => userRoles.includes(role));
  };
  public async componentDidMount() {
    if (this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner,
      UserRole.Admin)) {
      await this.init();
      this.setState({ isFormReady: true });
    }
  }

  private async init() {
    const promises = [];
    promises.push(
      this.initStatusChoices()
    );
    promises.push(
      this.initCountryChoices());
    promises.push(
      this.initTypeChoices()
    );
    promises.push(
      this.initEutelsatOwnersChoices()
    );
    await Promise.all(promises);

    if (this.isEditMode)
      await this.setInitialFormValues();

  }
  public async initTypeChoices() {

    const type = this.typeChoices.length > 0 ? this.typeChoices[0].key as string : '';
    this.setState({
      typeChoices: this.typeChoices, type
    });
  }
  public async initStatusChoices() {
    const {
      pnpService
    } = this.props;
    let statusChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.LEGAL_SVC_PROVIDER.STATUS)
      .select('Choices')();
    statusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    const status = statusChoices.length > 0 ? statusChoices[0].key as string : '';
    this.setState({
      statusChoices, status
    });
  }

  public async initCountryChoices() {
    const {
      pnpService
    } = this.props;
    const countryChoices: IDropdownOption[] = [];
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Country).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      countryChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    const countries = [];
    this.setState({ countryChoices, countries });
  }
  public async initEutelsatOwnersChoices() {
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
      orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      eutelsatOwnersChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    const eutelsatOwners = [];
    this.setState({ eutelsatOwnersChoices, eutelsatOwners });
  }

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.LegalServiceProvider).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.SUMMARY,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.ADDRESS,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.STATUS,
      Consts.FIELDS.LEGAL_SVC_PROVIDER.COMMENTS,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const expand = [Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY, Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER];
    const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
    const filterTypeChoices = "" + this.typeChoices.find((type) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1)?.key;
    const initialValues = {
      name: item[Consts.FIELDS.COMMON.TITLE],
      type: filterTypeChoices,
      summary: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.SUMMARY],
      address: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.ADDRESS],
      comments: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.COMMENTS],
      status: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.STATUS],
      countries: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY]?.map((country) => ({
        key: country[Consts.FIELDS.COMMON.ID],
        text: country[Consts.FIELDS.COMMON.TITLE]
      })) || [],
      eutelsatOwners: item[Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNER]?.map((etuowner) => ({
        key: etuowner[Consts.FIELDS.COMMON.ID],
        text: etuowner[Consts.FIELDS.COMMON.TITLE]
      })) || []
    };


    this.setState({
      ...initialValues
    });
  };

  private validateForm(): boolean {
    const errors: { [key: string]: string } = {};
    const {
      name,
      countries
    } = this.state;
    if (name?.trim() === '' || !name)
      errors.name = strings.FormEutEntityNameValidationError;


    if (!Array.isArray(countries) || countries.length === 0) {
      errors.country = strings.FormLegalSvcProviderCountryValidationError;
    }
    this.setState({ errors });
    return Object.keys(errors).length === 0;
  }

  public render(): React.ReactElement<ILegalSvcProviderProps> {
    const {
      name,
      summary,
      address,
      countries,
      countryChoices,
      status,
      statusChoices,
      comments,
      errors,
      isFormReady,
      typeChoices,
      type,
      eutelsatOwnersChoices,
      eutelsatOwners
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
            formOrigin={FormOrigin.Menu}>
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
                          <Label required>{strings.FormLegalSvcProviderNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormLegalSvcProviderTooltipName} ></Tooltip>
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
                    {/* Type */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormLegalSvcProviderTypeLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormLegalSvcProviderTooltipType} ></Tooltip>
                      </div>
                      <Dropdown
                        selectedKey={type} disabled={this.isEditMode}
                        onChange={this.onChangeType}
                        options={typeChoices} />
                    </div>

                    {/* Summary */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormLegalSvcProviderSummaryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormLegalSvcProviderTooltipSummary} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
                    </div>


                    {/* Address */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormLegalSvcProviderAddressLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormLegalSvcProviderTooltipAddress} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={address} onChange={this.onChangeAddress} />
                    </div>


                    {/* Country */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormLegalSvcProviderCountryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormLegalSvcProviderTooltipCountry} ></Tooltip>
                      </div>
                      <MultiselectWrapper
                        value={countries}
                        data={countryChoices}
                        dataKey={(item: IDropdownOption) => item.key}
                        textField={(item: IDropdownOption) => item.text}
                        filter="contains"
                        onChange={this.onChangeCountry}
                      />
                      {errors.country &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.country}
                          </span>
                        </div>
                      }
                    </div>

                    {/* Eutelsat Owners*/}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormLegalSvcProviderEutelsatOwnersLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormLegalSvcProviderTooltipEutelsatOwners} ></Tooltip>
                      </div>
                      <MultiselectWrapper
                        value={eutelsatOwners}
                        data={eutelsatOwnersChoices}
                        dataKey={(item: IDropdownOption) => item.key}
                        textField={(item: IDropdownOption) => item.text}
                        filter="contains"
                        onChange={this.onChangeEutelsatOwners} />
                    </div>
                    {/* Status */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormLegalSvcProviderStatusLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormLegalSvcProviderTooltipStatus} ></Tooltip>
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
                        <Tooltip content={strings.FormLegalSvcProviderTooltipComments} ></Tooltip>
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

  private onChangeName(event, name: string) {
    this.setState({ name });
  }

  private onChangeSummary(summary: string) {
    this.setState({ summary });
    return summary;
  }

  private onChangeAddress(address: string) {
    this.setState({ address });
    return address;
  }

  private onChangeCountry(countries: IDropdownOption[]) {
    this.setState({ countries });

  }
  private onChangeEutelsatOwners(eutelsatOwners: IDropdownOption[]) {
    this.setState({ eutelsatOwners });
  }
  private onChangeStatus(event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) {
    this.setState({ status: option.key.toString() });
  }
  private onChangeType(event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) {
    this.setState({ type: option.key.toString() });
  }
  private onChangeComments(comments: string) {
    this.setState({ comments });
    return comments;
  }

  private getDuplicateItems = async (): Promise<IBaseItem[]> => {
    const {
      pnpService,
      itemId
    } = this.props;
    const {
      name
    } = this.state;
    const duplicateItems: IBaseItem[] = [];
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields)();
    const match = items.find((i) => (i.Title?.toLowerCase() === UtilHelper.sanitizeSharePointTitle(name).toLowerCase()) && i.Id !== itemId);
    if (match)
      duplicateItems.push({ Id: match[Consts.FIELDS.COMMON.ID], Title: match[Consts.FIELDS.COMMON.TITLE] });
    return duplicateItems;
  };

  private getFormSummary() {
    const {
      name,
      summary,
      address,
      countries,
      status,
      comments,
      typeChoices,
      type,
      eutelsatOwners
    } = this.state;
    const summaryTextValue = DomHelper.cleanRichHtml(summary);
    const addressTextValue = DomHelper.cleanRichHtml(address);
    const commentsTextValue = DomHelper.cleanRichHtml(comments);
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormLegalSvcProviderDatasheetNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormLegalSvcProviderNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormLegalSvcProviderTypeLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {typeChoices.find((choice) => choice.key === type)?.text}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormLegalSvcProviderSummaryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {summaryTextValue.length > 0 ? <RichText isEditMode={false} value={summaryTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormEutEntityAddressLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {addressTextValue.length > 0 ? <RichText isEditMode={false} value={addressTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormLegalSvcProviderCountryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {countries.map((c) => c.text).join(", ")}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormLegalSvcProviderEutelsatOwnersLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {eutelsatOwners.map((c) => c.text).join(", ")}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormLegalSvcProviderStatusLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {status}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormLegalSvcProviderCommentsLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {commentsTextValue.length > 0 ? <RichText isEditMode={false} value={commentsTextValue} /> : ''}
          </div>
        </div>
      </div>
    );
  }

  private async callback(redirectUrl) {
    const {
      callback,
    } = this.props;
    callback(redirectUrl);
  }

  private async processCreation(): Promise<IItemAddResult> {
    const {
      name,
      summary,
      address,
      countries,
      status,
      comments,
      typeChoices,
      type,
      eutelsatOwners
    } = this.state;

    const {
      pnpService,
    } = this.props;
    const itemData = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = name;
    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = typeChoices.find((choice) => choice.key === type)?.key;
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.SUMMARY] = summary;
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.ADDRESS] = address;
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY_ID] = countries.map((c) => c.key);
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNERS_ID] = eutelsatOwners.map((c) => c.key);
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.STATUS] = status;
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.COMMENTS] = comments;

    const selectedContentType = this.typeChoices.find((type) => itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1);
    const list = this.config.list;
    let folderUrl: string;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);

    const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
    const documentSpaceUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.DOCUMENT_SPACE_URL);
    if (selectedContentType.key === Consts.CONTENT_TYPES.LAW_FIRM) {
      const parentFolderUrl = `${documentSpaceUrl}${Consts.DOCUMENT_SPACE.LEGAL_LAW_FIRMS_FOLDER_URL}`;
      folderUrl = await createUniqueFolder(pnpService, parentFolderUrl, sanitizedFolderName);
    } else {
      const parentFolderUrl = `${documentSpaceUrl}${Consts.DOCUMENT_SPACE.LEGAL_REPRESENTATIVES_FOLDER_URL}`;
      folderUrl = await createUniqueFolder(pnpService, parentFolderUrl, sanitizedFolderName);
    }


    const updateData = {};
    updateData[Consts.FIELDS.COMMON.DOCUMENTS_SPACE] =
    {
      [Consts.FIELDS.COMMON.URL]: folderUrl, [Consts.FIELDS.COMMON.DESCRIPTION]: "Link"
    };

    await pnpService.getListItems(listUrl).getById(createdItem.data.ID).update(updateData);
    return createdItem;


  }

  private processEdit = async () => {
    const {
      name,
      summary,
      address,
      countries,
      eutelsatOwners,
      status,
      comments,
    } = this.state;

    const {
      pnpService,
      itemId,
    } = this.props;



    const itemData = {};

    itemData[Consts.FIELDS.COMMON.TITLE] = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.SUMMARY] = summary;
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.ADDRESS] = address;
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.COUNTRY_ID] = countries.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.OWNERS_ID] =
      eutelsatOwners.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.COMMENTS] = comments;
    itemData[Consts.FIELDS.LEGAL_SVC_PROVIDER.STATUS] = status;
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
  };

}
