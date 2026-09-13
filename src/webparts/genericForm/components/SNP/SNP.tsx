import * as React from 'react';
import { SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { ISNPProps, ISNPState } from './SNP.types';
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
import { Combobox } from 'react-widgets/cjs';
import { UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';

export default class SNP extends React.Component<
  ISNPProps,
  ISNPState
> {
  private isEditMode: boolean;
  private config: IFormConfig;


  constructor(props: Readonly<ISNPProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      name: '',
      type: '',
      typeChoices: [],
      summary: '',
      city: '',
      status: '',
      statusChoices: [],
      comments: '',
      country: null,
      countryChoices: [],
      eutelsatEntity: null,
      eutelsatEntityChoices: [],
      teleportPartner: null,
      teleportPartnerChoices: [],
      errors: {},
      isFormReady: false
    };
    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.SNP);
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
      this.initEutelsatEntityChoices());
    promises.push(
      this.initCountryChoices());
    promises.push(
      this.initTeleportPartnerChoices());
    promises.push(
      this.initTypeChoices());
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
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.SNP.STATUS)
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
  private initTypeChoices = async () => {
    const typeChoices: IDropdownOption[] = [
      { key: Consts.CONTENT_TYPES.STARGATE, text: strings.FormSNPTypeStargate },
      { key: Consts.CONTENT_TYPES.EUTELSAT_SNP, text: strings.FormSNPTypeEutelsat },
      { key: Consts.CONTENT_TYPES.PARTNER_SNP, text: strings.FormSNPTypeOther }
    ];

    if (!this.isEditMode) {
      const type = typeChoices.length > 0 ? typeChoices[0].key as string : '';
      this.setState({
        typeChoices,
        type
      });
    } else {
      this.setState({
        typeChoices
      });
    }
  };

  private initCountryChoices = async () => {
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
      this.setState({ country, countryChoices });
    } else {
      this.setState({ countryChoices });
    }
  };

  public initEutelsatEntityChoices = async () => {
    const {
      pnpService
    } = this.props;
    const eutelsatEntityChoices: IDropdownOption[] = [];
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.EutelsatEntity).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields).
      //filter(`startswith(${Consts.FIELDS.COMMON.CONTENTTYPE_ID}, '${Consts.CONTENT_TYPES.EUTELSAL_CONTACT}')`).
      orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      eutelsatEntityChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    if (!this.isEditMode) {
      const eutelsatEntity = null;
      this.setState({ eutelsatEntityChoices, eutelsatEntity });
    } else {
      this.setState({ eutelsatEntityChoices });
    }
  };

  public initTeleportPartnerChoices = async () => {
    const {
      pnpService
    } = this.props;
    const teleportPartnerChoices: IDropdownOption[] = [];
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.TP).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields).
      //filter(`startswith(${Consts.FIELDS.COMMON.CONTENTTYPE_ID}, '${Consts.CONTENT_TYPES.EUTELSAL_CONTACT}')`).
      orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      teleportPartnerChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    if (!this.isEditMode) {
      const teleportPartner = null;
      this.setState({ teleportPartnerChoices, teleportPartner });
    } else {
      this.setState({ teleportPartnerChoices });
    }
  };

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const { typeChoices } = this.state;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.SNP).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);

    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.SNP.SUMMARY,
      Consts.FIELDS.SNP.CITY,
      Consts.FIELDS.SNP.COMMENTS,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      Consts.FIELDS.SNP.STATUS,
      `${Consts.FIELDS.SNP.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.SNP.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.SNP.EUTELSATENTITY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.SNP.EUTELSATENTITY}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.SNP.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.SNP.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.TITLE}`
    ];

    const expand = [Consts.FIELDS.SNP.COUNTRY, Consts.FIELDS.SNP.EUTELSATENTITY, Consts.FIELDS.SNP.TELEPORTPARTNER];

    const item = await pnpService
      .getListByUrl(listUrl)
      .items.getById(itemId)
      .select(...selectedFields)
      .expand(...expand)();



    const type = "" + typeChoices.find(
      (type) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1
    )?.key;


    let country = null;
    let eutelsatEntity = null;
    let teleportPartner = null;

    if (item[Consts.FIELDS.SNP.COUNTRY]) {
      country = {
        key: item[Consts.FIELDS.SNP.COUNTRY][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.SNP.COUNTRY][Consts.FIELDS.COMMON.TITLE]
      };
    } else {
      country = null;
    }

    if (item[Consts.FIELDS.SNP.EUTELSATENTITY]) {
      eutelsatEntity = {
        key: item[Consts.FIELDS.SNP.EUTELSATENTITY][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.SNP.EUTELSATENTITY][Consts.FIELDS.COMMON.TITLE]
      };
    } else {
      eutelsatEntity = null;
    }

    if (item[Consts.FIELDS.SNP.TELEPORTPARTNER]) {
      teleportPartner = {
        key: item[Consts.FIELDS.SNP.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.SNP.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE]
      };
    } else {
      teleportPartner = null;
    }

    const initialValues = {
      name: item[Consts.FIELDS.COMMON.TITLE],
      type,
      summary: item[Consts.FIELDS.SNP.SUMMARY],
      city: item[Consts.FIELDS.SNP.CITY],
      comments: item[Consts.FIELDS.SNP.COMMENTS],
      country,
      eutelsatEntity,
      teleportPartner,
      status: item[Consts.FIELDS.SNP.STATUS]
    };

    this.setState({
      ...initialValues
    });
  };



  private validateForm = (): boolean => {
    const errors: { [key: string]: string } = {};
    const {
      name,
      country,
      city
    } = this.state;
    if (name?.trim() === '' || !name)
      errors.name = strings.FormSNPNameValidationError;

    if (!country || !country.key || country.key === '') {
      errors.country = strings.FormSNPCountryValidationError;
    }

    if (city?.trim() === '' || !city) {
      errors.city = strings.FormSNPCityValidationError;
    }

    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

  public render(): React.ReactElement<ISNPProps> {
    const {
      name,
      type,
      typeChoices,
      summary,
      city,
      country,
      countryChoices,
      comments,
      eutelsatEntity,
      eutelsatEntityChoices,
      teleportPartner,
      teleportPartnerChoices,
      status,
      statusChoices,
      errors,
      isFormReady
    } = this.state;
    return (
      <div className="ms-Grid-row" >
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
                          <Label required>{strings.FormSNPNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormSNPTooltipName} ></Tooltip>
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
                          <Label>{strings.FormSNPSummaryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormSNPTooltipSummary} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
                    </div>
                    {/* Type */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormSNPTypeLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormSNPTooltipType} ></Tooltip>
                      </div>
                      <Dropdown
                        disabled={this.isEditMode}
                        selectedKey={type}
                        onChange={this.onChangetype}
                        options={typeChoices} />
                    </div>

                    {/* Teleport Partner */}

                    {type === Consts.CONTENT_TYPES.PARTNER_SNP && (
                      <div className={styles.field}>
                        <div className={styles.fieldLabel}>
                          <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormSNPTeleportPartnerLabel}</Label>
                          </div>
                          <Tooltip content={strings.FormSNPTooltipTeleportPartner} ></Tooltip>
                        </div>
                        <Combobox
                          data={teleportPartnerChoices}
                          value={teleportPartner}
                          selectIcon={
                            <span className="ms-Dropdown-caretDownWrapper">
                              <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
                            </span>
                          }

                          textField="text"
                          filter="contains"
                          key="key"
                          onChange={this.onChangeTelePortPartner}
                        />
                      </div>)}
                    {/* Etuelsat Entity */}
                    {type === Consts.CONTENT_TYPES.EUTELSAT_SNP && (
                      <div className={styles.field}>
                        <div className={styles.fieldLabel}>
                          <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormSNPEutelsatEntityLabel}</Label>
                          </div>
                          <Tooltip content={strings.FormSNPTooltipEutelsatEntity} ></Tooltip>
                        </div>
                        <Combobox
                          data={eutelsatEntityChoices}
                          value={eutelsatEntity}
                          selectIcon={
                            <span className="ms-Dropdown-caretDownWrapper">
                              <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
                            </span>
                          }
                          textField="text"
                          filter="contains"
                          key="key"
                          onChange={this.onChangeeutelsatOwners}
                        />
                      </div>)}




                    {/* Country */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormSNPCountryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormSNPTooltipCountry} ></Tooltip>
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


                    {/* city */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormSNPCityLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormSNPTooltipCity} ></Tooltip>
                      </div>
                      <TextField
                        value={city}
                        onChange={this.onChangeCity}
                      />
                      {errors.city &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.city}
                          </span>
                        </div>
                      }
                    </div>
                    {/* Status */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormSNPStatusLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormSNPTooltipStatus} ></Tooltip>
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
                          <Label>{strings.FormSNPCommentsLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormSNPTooltipComments} ></Tooltip>
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

  private onChangeCity = (event, city: string) => {
    this.setState({ city });
  };



  private onChangeStatus = (event, status: IDropdownOption) => {
    this.setState({ status: status.key.toString() });
  };
  private onChangeTelePortPartner = (teleportPartner: IDropdownOption) => {
    this.setState({ teleportPartner });
  };

  private onChangeCountry = (country: IDropdownOption) => {
    this.setState({ country });
  };

  private onChangeeutelsatOwners = (eutelsatEntity: IDropdownOption) => {
    this.setState({ eutelsatEntity });
  };

  private onChangetype = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ type: option.key.toString() });
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
      city,
      status,
      country,
      type,
      typeChoices,
      eutelsatEntity,
      teleportPartner,
      comments
    } = this.state;
    const typeTextValue = typeChoices.find((choice) => choice.key === type)?.text;
    const summaryTextValue = DomHelper.cleanRichHtml(summary);
    const commentsTextValue = DomHelper.cleanRichHtml(comments);
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPDatasheetLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPTypeLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {typeTextValue}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {UtilHelper.sanitizeSharePointTitle(name)}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPSummaryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {summaryTextValue.length > 0 ? <RichText isEditMode={false} value={summaryTextValue} /> : ''}
          </div>
        </div>

        {type === Consts.CONTENT_TYPES.PARTNER_SNP && (<div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPTeleportPartnerLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {teleportPartner?.text}
          </div>
        </div>)}
        {type === Consts.CONTENT_TYPES.EUTELSAT_SNP && (
          <div className={styles.field}>
            <div className={styles.fieldLabel}>
              <div className={styles.fieldLabelContainer}>
                <Label>{strings.FormSNPEutelsatEntityLabel}</Label>
              </div>
            </div>
            <div className={styles.fieldValue}>
              {eutelsatEntity?.text}
            </div>
          </div>)}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPCountryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {country?.text}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormSNPCityLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {city}
          </div>
        </div>

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
              <Label>{strings.FormSNPCommentsLabel}</Label>
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
      type,
      summary,
      city,
      country,
      eutelsatEntity,
      comments,
      teleportPartner,
      status
    } = this.state;

    const {
      pnpService,
    } = this.props;

    const itemData = {};
    const sanitizedTitle = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.COMMON.TITLE] = sanitizedTitle;
    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = type;
    itemData[Consts.FIELDS.SNP.SUMMARY] = summary;
    itemData[Consts.FIELDS.SNP.CITY] = city;
    itemData[Consts.FIELDS.SNP.COUNTRY_ID] = country?.key;
    itemData[Consts.FIELDS.SNP.EUTELSATENTITY_ID] = eutelsatEntity?.key;
    itemData[Consts.FIELDS.SNP.TELEPORTPARTNER_ID] = teleportPartner?.key;
    itemData[Consts.FIELDS.SNP.COMMENTS] = comments;
    itemData[Consts.FIELDS.SNP.STATUS] = status;



    const folderUrlMap: Record<string, string> = {
      [Consts.CONTENT_TYPES.STARGATE]: Consts.DOCUMENT_SPACE.SNP_STARGATE_FOLDER_URL,
      [Consts.CONTENT_TYPES.EUTELSAT_SNP]: Consts.DOCUMENT_SPACE.SNP_EUTELSAT_SNP_FOLDER_URL,
      [Consts.CONTENT_TYPES.PARTNER_SNP]: Consts.DOCUMENT_SPACE.SNP_PARTNER_SNP_FOLDER_URL
    };

    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);
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
      type,
      summary,
      city,
      country,
      eutelsatEntity,
      teleportPartner,
      comments,
      status,
    } = this.state;

    const {
      pnpService,
      itemId,
    } = this.props;



    const itemData = {};

    itemData[Consts.FIELDS.COMMON.TITLE] = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = type;
    itemData[Consts.FIELDS.SNP.SUMMARY] = summary;
    itemData[Consts.FIELDS.SNP.CITY] = city;
    itemData[Consts.FIELDS.SNP.COUNTRY_ID] = country?.key;
    itemData[Consts.FIELDS.SNP.EUTELSATENTITY_ID] = eutelsatEntity?.key;
    itemData[Consts.FIELDS.SNP.TELEPORTPARTNER_ID] = teleportPartner?.key;
    itemData[Consts.FIELDS.SNP.COMMENTS] = comments;
    itemData[Consts.FIELDS.SNP.STATUS] = status;


    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
  };


}

