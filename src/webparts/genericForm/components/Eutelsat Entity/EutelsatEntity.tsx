import * as React from 'react';
import { SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { Combobox } from 'react-widgets';
import { IEutelsatEntityProps, IEutelsatEntityState } from './EutelsatEntity.types';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import { Dropdown, Fabric, Icon, IDropdownOption, Label, TextField } from 'office-ui-fabric-react';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { Form, FormOrigin, FormType, IFormConfig } from '../Form.types';
import strings from 'GenericFormWebPartStrings';
import { RichText } from '../../../../common/components/RichText/RichText';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { Consts } from '../../../../common/consts/Consts';
import 'react-widgets/styles.css';
import { IItemAddResult } from '@pnp/sp/items';
import { UtilHelper } from '../../../../common/helpers/Util';
import LoadingSpinner from '../../../../common/components/Loading Spinner/LoadingSpinner';
import { createUniqueFolder } from '../Form.utility';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';
export default class EutelsatEntity extends React.Component<
  IEutelsatEntityProps,
  IEutelsatEntityState
> {
  private isEditMode: boolean;
  private config: IFormConfig;
  constructor(props: Readonly<IEutelsatEntityProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      name: '',
      summary: '',
      address: '',
      comments: '',
      status: null,
      statusChoices: [],
      country: null,
      countryChoices: [],
      errors: {},
      isFormReady: false
    };
    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.EutelsatEntity);
    this.validateForm = this.validateForm.bind(this);
    this.getDuplicateItems = this.getDuplicateItems.bind(this);
    this.getFormSummary = this.getFormSummary.bind(this);
    this.callback = this.callback.bind(this);
    this.processCreation = this.processCreation.bind(this);
    this.processEdit = this.processEdit.bind(this);
    this.initStatusChoices = this.initStatusChoices.bind(this);
    this.initCountryChoices = this.initCountryChoices.bind(this);
    this.onChangeName = this.onChangeName.bind(this);
    this.onChangeStatus = this.onChangeStatus.bind(this);
    this.onChangeCountry = this.onChangeCountry.bind(this);
    this.onChangeSummary = this.onChangeSummary.bind(this);
    this.onChangeAddress = this.onChangeAddress.bind(this);
    this.onChangeComments = this.onChangeComments.bind(this);

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
    promises.push(
      this.initStatusChoices()
    );
    promises.push(
      this.initCountryChoices());

    await Promise.all(promises);
    if (this.isEditMode)
      await this.setInitialFormValues();
  };

  public initStatusChoices = async () => {
    const {
      pnpService
    } = this.props;
    let statusChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.EUT_ENTITY.STATUS)
      .select('Choices')();
    statusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    if (!this.isEditMode) {
      const status = statusChoices.length > 0 ? statusChoices[0].key as string : '';
      this.setState({
        statusChoices, status
      });
    } else {
      this.setState({
        statusChoices
      });
    }

  };

  public initCountryChoices = async () => {
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
    if (!this.isEditMode) {
      const country = null;
      this.setState({ countryChoices, country });
    } else {
      this.setState({ countryChoices });
    }
  };

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.EutelsatEntity).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.EUT_ENTITY.SUMMARY,
      Consts.FIELDS.EUT_ENTITY.ADDRESS,
      Consts.FIELDS.EUT_ENTITY.STATUS,
      Consts.FIELDS.EUT_ENTITY.COMMENTS,
      `${Consts.FIELDS.EUT_ENTITY.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.EUT_ENTITY.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const expand = [Consts.FIELDS.EUT_ENTITY.COUNTRY];
    const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();

    const initialValues = {
      name: item[Consts.FIELDS.COMMON.TITLE],
      summary: item[Consts.FIELDS.EUT_ENTITY.SUMMARY],
      address: item[Consts.FIELDS.EUT_ENTITY.ADDRESS],
      comments: item[Consts.FIELDS.EUT_ENTITY.COMMENTS],
      status: item[Consts.FIELDS.EUT_ENTITY.STATUS],
      country: {
        key: item[Consts.FIELDS.EUT_ENTITY.COUNTRY][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.EUT_ENTITY.COUNTRY][Consts.FIELDS.COMMON.TITLE]
      }
    };

    this.setState({
      ...initialValues
    });
  };

  private validateForm = (): boolean => {
    const errors: { [key: string]: string } = {};
    const {
      name,
      country
    } = this.state;
    if (name?.trim() === '' || !name)
      errors.name = strings.FormEutEntityNameValidationError;

    if (!country || !country.key || country.key === '')
      errors.country = strings.FormEutEntityCountryValidationError;

    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

  public render(): React.ReactElement<IEutelsatEntityProps> {
    const {
      name,
      summary,
      address,
      country,
      countryChoices,
      status,
      statusChoices,
      comments,
      errors,
      isFormReady
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
              {!isFormReady && <LoadingSpinner size={SpinnerSize.large} label={'Loading...'} className={styles.spinnerContainer} />}
              {isFormReady &&
                < Fabric >
                  <form>
                    {/* Name */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormEutEntityNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormEutEntityTooltipName} ></Tooltip>
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


                    {/* Summary */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormEutEntitySummaryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormEutEntityTooltipSummary} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
                    </div>


                    {/* Address */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormEutEntityAddressLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormEutEntityTooltipAddress} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={address} onChange={this.onChangeAddress} />
                    </div>


                    {/* Country */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormEutEntityCountryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormEutEntityTooltipCountry} ></Tooltip>
                      </div>
                      <Combobox
                        data={countryChoices}
                        value={country}
                        selectIcon={
                          <span className="ms-Dropdown-caretDownWrapper">
                            <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
                          </span>
                        }
                        textField="text"
                        filter="contains"
                        key="key"
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


                    {/* Status */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormEutEntityStatusLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormEutEntityTooltipStatus} ></Tooltip>
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
                        <Tooltip content={strings.FormEutEntityTooltipComments} ></Tooltip>
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

  private onChangeAddress = (address: string) => {
    this.setState({ address });
    return address;
  };

  private onChangeCountry = (country: IDropdownOption) => {
    this.setState({ country });
  };

  private onChangeStatus = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ status: option.key.toString() });
  };

  private onChangeComments = (comments: string) => {
    this.setState({ comments });
    return comments;
  };

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

  private getFormSummary = () => {
    const {
      name,
      summary,
      address,
      country,
      status,
      comments
    } = this.state;
    const summaryTextValue = DomHelper.cleanRichHtml(summary);
    const addressTextValue = DomHelper.cleanRichHtml(address);
    const commentsTextValue = DomHelper.cleanRichHtml(comments);
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormEutEntityTypeLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormEutEntityNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {UtilHelper.sanitizeSharePointTitle(name)}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormEutEntitySummaryLabel}</Label>
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
              <Label>{strings.FormEutEntityCountryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {country?.text}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormEutEntityStatusLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {status}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormEutEntityCommentsLabel}</Label>
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
      name,
      summary,
      address,
      country,
      status,
      comments,
    } = this.state;

    const {
      pnpService,
    } = this.props;

    const itemData = {};
    const sanitizedTitle = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.COMMON.TITLE] = sanitizedTitle;
    itemData[Consts.FIELDS.EUT_ENTITY.SUMMARY] = summary;
    itemData[Consts.FIELDS.EUT_ENTITY.ADDRESS] = address;
    itemData[Consts.FIELDS.EUT_ENTITY.COUNTRY_ID] = parseInt(country.key.toString());
    itemData[Consts.FIELDS.EUT_ENTITY.STATUS] = status;
    itemData[Consts.FIELDS.EUT_ENTITY.COMMENTS] = comments;
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);

    const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
    const documentSpaceUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.DOCUMENT_SPACE_URL);
    const parentFolderUrl = `${documentSpaceUrl}${Consts.DOCUMENT_SPACE.EUT_ENTITY_FOLDER_URL}`;
    const folderUrl = await createUniqueFolder(pnpService, parentFolderUrl, sanitizedFolderName);

    const updateData = {};
    updateData[Consts.FIELDS.COMMON.DOCUMENTS_SPACE] =
    {
      [Consts.FIELDS.COMMON.URL]: folderUrl, [Consts.FIELDS.COMMON.DESCRIPTION]: "Link"
    };

    await pnpService.getListItems(listUrl).getById(createdItem.data.ID).update(updateData);
    return createdItem;
  };

  private processEdit = async () => {
    const {
      name,
      summary,
      address,
      country,
      status,
      comments,
    } = this.state;

    const {
      pnpService,
      itemId,
    } = this.props;

    const itemData = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.EUT_ENTITY.SUMMARY] = summary;
    itemData[Consts.FIELDS.EUT_ENTITY.ADDRESS] = address;
    itemData[Consts.FIELDS.EUT_ENTITY.COUNTRY_ID] = parseInt(country.key.toString());
    itemData[Consts.FIELDS.EUT_ENTITY.STATUS] = status;
    itemData[Consts.FIELDS.EUT_ENTITY.COMMENTS] = comments;
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);

  };

}
