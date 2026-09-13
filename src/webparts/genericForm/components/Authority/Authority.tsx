import * as React from 'react';
import { SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import { IAuthorityProps, IAuthorityyState } from './Authority.types';
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
import { ModernTaxonomyPicker } from "@pnp/spfx-controls-react/lib/ModernTaxonomyPicker";
import { ITermInfo } from '@pnp/spfx-controls-react/node_modules/@pnp/sp/taxonomy';
import { Config } from '../../../../common/config/Config';
import { TaxonomyService } from '../../../../common/services/TaxonomyService';
import { IBaseItem, ICountry } from '../../../../common/models/IBusiness';
import { ITaxonomyTerm } from '../../../../common/models/SPEntities';
import { AuthorityType, UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';
export default class Authority extends React.Component<
  IAuthorityProps,
  IAuthorityyState
> {
  private isEditMode: boolean;
  private config: IFormConfig;
  private allCountries: ICountry[] = [];


  constructor(props: Readonly<IAuthorityProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      name: '',
      portal: '',
      type: '',
      typeChoices: [],
      summary: '',
      address: '',
      comments: '',
      grouping: null,
      countries: [],
      countriesChoices: [],
      eutOwners: [],
      eutOwnersChoices: [],
      errors: {},
      isFormReady: false
    };
    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.Authority);
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
      this.initEutOwnerChoices());
    promises.push(
      this.initCountryChoices());
    promises.push(
      this.initTypeChoices());
    await Promise.all(promises);
    if (this.isEditMode)
      await this.setInitialFormValues();
  };

  private initTypeChoices = async () => {
    const typeChoices: IDropdownOption[] = [
      { key: Consts.CONTENT_TYPES.ADMINISTRATION, text: AuthorityType.Administration },
      { key: Consts.CONTENT_TYPES.ORGANIZATION, text: AuthorityType.Organization },
      { key: Consts.CONTENT_TYPES.REGULATOR, text: AuthorityType.Regulator }
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
    const countriesChoices: IDropdownOption[] = [];
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Country).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.AUTHORITY.GROUPING
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      this.allCountries.push({ Id: item[Consts.FIELDS.COMMON.ID], Title: item[Consts.FIELDS.COMMON.TITLE], Grouping: item[Consts.FIELDS.AUTHORITY.GROUPING] });
      countriesChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    if (!this.isEditMode) {
      const countries = [];
      this.setState({ countries, countriesChoices });
    }
    // else {
    //   this.setState({ countriesChoices });
    // }
  };

  public initEutOwnerChoices = async () => {
    const {
      pnpService
    } = this.props;
    const eutOwnersChoices: IDropdownOption[] = [];
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Contact).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields).
      filter(`startswith(${Consts.FIELDS.COMMON.CONTENTTYPE_ID}, '${Consts.CONTENT_TYPES.EUTELSAL_CONTACT}')`).
      orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      eutOwnersChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    if (!this.isEditMode) {
      const eutOwners = [];
      this.setState({ eutOwnersChoices, eutOwners });
    } else {
      this.setState({ eutOwnersChoices });
    }
  };

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const { typeChoices } = this.state;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Authority).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);

    const selectedFields: string[] = [
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

    const item = await pnpService
      .getListByUrl(listUrl)
      .items.getById(itemId)
      .select(...selectedFields)
      .expand(...expand)();

    const taxonomyService = new TaxonomyService();
    await taxonomyService.init(pnpService);

    let grouping: ITaxonomyTerm = null;
    if (item[Consts.FIELDS.AUTHORITY.GROUPING]) {
      const localizedLabelGrouping = await taxonomyService.getLocalizedTerm(
        item[Consts.FIELDS.AUTHORITY.GROUPING]?.TermGuid,
        'en-us'
      );

      grouping = {
        id: item[Consts.FIELDS.AUTHORITY.GROUPING]?.TermGuid,
        labels: [{
          name: localizedLabelGrouping,
          languageTag: 'en-US',
          isDefault: true
        }
        ]
      };
    }

    const type = "" + typeChoices.find(
      (type) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1
    )?.key;
    const initialValues = {
      name: item[Consts.FIELDS.COMMON.TITLE],
      type,
      summary: item[Consts.FIELDS.AUTHORITY.SUMMARY],
      address: item[Consts.FIELDS.AUTHORITY.ADDRESS],
      comments: item[Consts.FIELDS.AUTHORITY.COMMENTS],
      countries: item[Consts.FIELDS.AUTHORITY.COUNTRY]?.map((country) => ({
        key: country[Consts.FIELDS.COMMON.ID],
        text: country[Consts.FIELDS.COMMON.TITLE]
      })) || [],
      eutOwners: item[Consts.FIELDS.AUTHORITY.ETU_OWNERS]?.map((eutOwner) => ({
        key: eutOwner[Consts.FIELDS.COMMON.ID],
        text: eutOwner[Consts.FIELDS.COMMON.TITLE]
      })) || [],
      grouping,
      countriesChoices: this.getCountryChoices(grouping),
      portal: item[Consts.FIELDS.AUTHORITY.PORTAL]?.Url?.trim() || '',

    };

    this.setState({
      ...initialValues
    });
  };



  private validateForm = (): boolean => {
    const errors: { [key: string]: string } = {};
    const {
      name,
      countries,
      portal
    } = this.state;
    if (name?.trim() === '' || !name)
      errors.name = strings.FormAuthorityNameValidationError;

    if (!Array.isArray(countries) || countries.length === 0) {
      errors.country = strings.FormAuthorityCountryValidationError;
    }

    if (portal?.trim() && !UtilHelper.validateURL(portal)) {
      errors.portal = strings.FormAuthorityPortalValidationError;
    }



    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

  public render(): React.ReactElement<IAuthorityProps> {
    const {
      name,
      portal,
      type,
      typeChoices,
      summary,
      address,
      countries,
      countriesChoices,
      comments,
      grouping,
      eutOwners,
      eutOwnersChoices,
      errors,
      isFormReady
    } = this.state;
    return (
      <div className="ms-Grid-row" >
        <div className="ms-Grid-col ms-sm12 ms-md12">
          {/* Form  */}
          {!this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && (
            <AccessDeniedMessage message={strings.FormAccessDeniedMessage}> </AccessDeniedMessage>
          )}
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
                          <Label required>{strings.FormAuthorityNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormAuthorityTooltipName} ></Tooltip>
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
                          <Label>{strings.FormAuthorityTypeLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormAuthorityTooltipType} ></Tooltip>
                      </div>
                      <Dropdown
                        disabled={this.isEditMode}
                        selectedKey={type}
                        onChange={this.onChangetype}
                        options={typeChoices} />
                    </div>

                    {/* Summary */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormAuthoritySummaryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormAuthorityTooltipSummary} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
                    </div>


                    {/* Address */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormAuthorityAddressLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormAuthorityTooltipAddress} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={address} onChange={this.onChangeAddress} />
                    </div>
                    {/* Taxonomy Picker */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormAuthorityTaxonomyLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormAuthorityTooltipGrouping} />
                      </div>
                      <ModernTaxonomyPicker
                        allowMultipleSelections={false}
                        termSetId={Config.TAXONOMY.GROUPING_TERMSETID}
                        panelTitle={strings.FormAuthorityTaxonomyLabel}
                        label=''
                        onChange={this.onChangeGrouping}
                        context={this.props.pnpService.getContext()}
                        placeHolder={strings.FormAuthorityGroupingPlaceHolder}
                        initialValues={grouping ? [grouping] : []}

                      />
                    </div>


                    {/* Country */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormAuthorityCountryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormAuthorityTooltipCountry} ></Tooltip>
                      </div>
                      <MultiselectWrapper
                        data={countriesChoices}
                        dataKey={(item: IDropdownOption) => item.key}
                        textField={(item: IDropdownOption) => item.text}
                        value={countries}
                        filter="contains"
                        onChange={this.onChangeCountry} />
                      {errors.country &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.country}
                          </span>
                        </div>
                      }
                    </div>

                    {/* Eut Owners */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormAuthorityEutOwnersLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormAuthorityTooltipOwners} ></Tooltip>
                      </div>
                      <MultiselectWrapper
                        value={eutOwners}
                        data={eutOwnersChoices}
                        dataKey={(item: IDropdownOption) => item.key}
                        textField={(item: IDropdownOption) => item.text}
                        filter="contains"
                        onChange={this.onChangeEutOwners} />

                    </div>

                    {/* Portal */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormAuthorityPortalLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormAuthorityTooltipPortal} ></Tooltip>
                      </div>
                      <TextField
                        value={portal}
                        onChange={this.onChangePortal}
                      />
                      {errors.portal &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.portal}
                          </span>
                        </div>
                      }
                    </div>

                    {/* Comments */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormAuthorityCommentsLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormAuthorityTooltipComments} ></Tooltip>
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

  private onChangePortal = (event, portal: string) => {
    this.setState({ portal });
  };

  private onChangeSummary = (summary: string) => {
    this.setState({ summary });
    return summary;
  };

  private onChangeAddress = (address: string) => {
    this.setState({ address });
    return address;
  };


  private onChangeGrouping = (info?: ITermInfo[]): void => {
    const { grouping } = this.state;
    if (!info || info.some((term) => term === undefined) || grouping?.id === info[0]?.id) return;

    const countries: IDropdownOption[] = [];
    const countriesChoices: IDropdownOption[] = this.getCountryChoices(info[0]);
    this.setState({ grouping: info[0], countriesChoices, countries });
  };

  private onChangeCountry = (countries: IDropdownOption[]) => {
    this.setState({ countries });
  };

  private onChangeEutOwners = (eutOwners: IDropdownOption[]) => {
    this.setState({ eutOwners });
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
      portal,
      summary,
      address,
      countries,
      type,
      typeChoices,
      grouping,
      eutOwners,
      comments
    } = this.state;
    const typeTextValue = typeChoices.find((choice) => choice.key === type)?.text;
    const countriesValues = countries.map((c) => c.text).join(", ");
    const eutOwnersValues = eutOwners.map((e) => e.text).join(", ");
    const groupingValues = grouping?.labels?.[0]?.name;
    const summaryTextValue = DomHelper.cleanRichHtml(summary);
    const addressTextValue = DomHelper.cleanRichHtml(address);
    const commentsTextValue = DomHelper.cleanRichHtml(comments);
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormAuthorityDatasheetLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormAuthorityTypeLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {typeTextValue}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormAuthorityNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {UtilHelper.sanitizeSharePointTitle(name)}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormAuthoritySummaryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {summaryTextValue.length > 0 ? <RichText isEditMode={false} value={summaryTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormAuthorityAddressLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {addressTextValue.length > 0 ? <RichText isEditMode={false} value={addressTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormAuthorityTaxonomyLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {groupingValues}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormAuthorityCountryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {countriesValues}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormAuthorityOwnersLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {eutOwnersValues}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormAuthorityPortalLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {portal}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormAuthorityCommentsLabel}</Label>
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
      address,
      countries,
      grouping,
      eutOwners,
      portal,
      comments,
    } = this.state;

    const {
      pnpService,
    } = this.props;

    const itemData = {};
    const sanitizedTitle = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.COMMON.TITLE] = sanitizedTitle;
    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = type;
    itemData[Consts.FIELDS.AUTHORITY.SUMMARY] = summary;
    itemData[Consts.FIELDS.AUTHORITY.ADDRESS] = address;
    itemData[Consts.FIELDS.AUTHORITY.COUNTRY_ID] =
      countries.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.AUTHORITY.ETU_OWNERS_ID] =
      eutOwners.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.AUTHORITY.PORTAL] =
      { [Consts.FIELDS.COMMON.URL]: portal, [Consts.FIELDS.COMMON.DESCRIPTION]: "Link" };
    itemData[Consts.FIELDS.AUTHORITY.COMMENTS] = comments;

    if (grouping?.labels?.[0]?.name && grouping?.id)
      itemData[Consts.FIELDS.AUTHORITY.GROUPING] =
        { Label: grouping.labels[0].name, TermGuid: grouping.id, WssId: '-1' };

    const folderUrlMap: Record<string, string> = {
      [Consts.CONTENT_TYPES.ADMINISTRATION]: Consts.DOCUMENT_SPACE.AUTHORITY_ADMINISTRATION_FOLDER_URL,
      [Consts.CONTENT_TYPES.ORGANIZATION]: Consts.DOCUMENT_SPACE.AUTHORITY_ORGANIZATION_FOLDER_URL,
      [Consts.CONTENT_TYPES.REGULATOR]: Consts.DOCUMENT_SPACE.AUTHORITY_REGULATOR_FOLDER_URL
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
      address,
      countries,
      grouping,
      eutOwners,
      comments,
      portal
    } = this.state;

    const {
      pnpService,
      itemId,
    } = this.props;



    const itemData = {};

    itemData[Consts.FIELDS.COMMON.TITLE] = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = type;
    itemData[Consts.FIELDS.AUTHORITY.SUMMARY] = summary;
    itemData[Consts.FIELDS.AUTHORITY.ADDRESS] = address;
    itemData[Consts.FIELDS.AUTHORITY.COUNTRY_ID] = countries.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.AUTHORITY.ETU_OWNERS_ID] =
      eutOwners.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.AUTHORITY.PORTAL] =
      { [Consts.FIELDS.COMMON.URL]: portal, [Consts.FIELDS.COMMON.DESCRIPTION]: "Link" };
    itemData[Consts.FIELDS.AUTHORITY.COMMENTS] = comments;

    if (grouping?.labels?.[0]?.name && grouping?.id)
      itemData[Consts.FIELDS.AUTHORITY.GROUPING] = { Label: grouping.labels[0].name, TermGuid: grouping.id, WssId: '-1' };
    else
      itemData[Consts.FIELDS.AUTHORITY.GROUPING] = null;
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
  };

  private getCountryChoices = (info: ITaxonomyTerm) => {
    const selectedTermId = info?.id;
    let countriesChoices: IDropdownOption[];
    if (!info) {
      countriesChoices = this.allCountries
        .map((country) => ({
          key: country.Id,
          text: country.Title
        } as IDropdownOption));
    } else if (selectedTermId) {
      countriesChoices = this.allCountries
        .filter((country: ICountry) => country.Grouping?.some((term) => term.TermGuid === selectedTermId)
        ).map((country) => ({
          key: country.Id,
          text: country.Title
        } as IDropdownOption));
    }
    return countriesChoices;
  };
}

