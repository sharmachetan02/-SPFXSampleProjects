import * as React from 'react';
import PhoneInput from 'react-phone-number-input';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { IContactProps, IContactState } from './Contact.type';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import { Dropdown, Fabric, Icon, IDropdownOption, Label, TextField } from 'office-ui-fabric-react';
import { Combobox } from 'react-widgets';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { Form, FormType, IFormConfig } from '../Form.types';
import strings from 'GenericFormWebPartStrings';
import { RichText } from '../../../../common/components/RichText/RichText';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { UtilHelper } from '../../../../common/helpers/Util';
import { Consts } from '../../../../common/consts/Consts';
import 'react-widgets/styles.css';
import 'react-phone-number-input/style.css';
import { IItemAddResult } from '@pnp/sp/items';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { ContactType, UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';
export default class Contact extends React.Component<
  IContactProps,
  IContactState
> {
  private isEditMode: boolean;
  private config: IFormConfig;
  private AllLookupItems: Map<string, IDropdownOption[]> = new Map();
  constructor(props: Readonly<IContactProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.AllLookupItems = new Map();
    this.state = {
      errors: {},
      isFormReady: false,
      title: '',
      firstname: '',
      lastname: '',
      jobtitle: '',
      department: '',
      email: '',
      phonenumber: '',
      summary: '',
      comments: '',
      typeChoices: [],
      type: '',
      item: null,
      itemChoices: [],
      isItemPreselected: false
    };
    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.Contact);
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
    promises.push(this.fetchAllLookupItems());
    await Promise.all(promises);
    this.initTypeChoices();
    if (this.isEditMode) {
      await this.setInitialFormValues();
    } else {
      this.applyPreselection();
    }
  };

  private fetchAllLookupItems = async () => {
    const { pnpService } = this.props;
    const selectedFields = [Consts.FIELDS.COMMON.ID, Consts.FIELDS.COMMON.TITLE];
    const configMap = [
      { key: ContactType.Administration, type: FormType.Authority, contentTypeId: Consts.CONTENT_TYPES.ADMINISTRATION },
      { key: ContactType.Organization, type: FormType.Authority, contentTypeId: Consts.CONTENT_TYPES.ORGANIZATION },
      { key: ContactType.Regulator, type: FormType.Authority, contentTypeId: Consts.CONTENT_TYPES.REGULATOR },
      { key: ContactType.LawFirm, type: FormType.LegalServiceProvider, contentTypeId: Consts.CONTENT_TYPES.LAW_FIRM },
      { key: ContactType.LegalRepresentative, type: FormType.LegalServiceProvider, contentTypeId: Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE },
      { key: ContactType.TP, type: FormType.TP, contentTypeId: null },
      { key: ContactType.DP, type: FormType.DP, contentTypeId: null }
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

  private applyPreselection = () => {
    // const selectedType = this.getContactTypeFromParentContentType(this.props.preselectedType);
    if (this.props.preselectedType) {
      const itemChoices: IDropdownOption[] = this.AllLookupItems.get(this.props.preselectedType) || [];
      // If preselectedItemId is provided, find and preselect that item
      if (this.props.preselectedItemId) {
        const preselectedItem = itemChoices.find((choice) => choice.key === this.props.preselectedItemId);
        if (preselectedItem) {
          this.setState({ itemChoices, item: preselectedItem, isItemPreselected: true });
        } else {
          this.setState({ itemChoices });
        }
      } else {
        this.setState({ itemChoices });
      }
    }
  };

  private getContactTypeFromParentContentType = (parentContentType?: string): string | null => {
    if (!parentContentType) {
      return null;
    }
    const typeMap: { [key: string]: string } = {
      [ContactType.Eutelsat]: Consts.CONTENT_TYPES.EUTELSAL_CONTACT,
      [ContactType.LegalRepresentative]: Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE_CONTACT,
      [ContactType.LawFirm]: Consts.CONTENT_TYPES.LAW_FIRM_CONTACT,
      [ContactType.Administration]: Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT,
      [ContactType.Organization]: Consts.CONTENT_TYPES.ORGANIZATION_CONTACT,
      [ContactType.Regulator]: Consts.CONTENT_TYPES.REGULATOR_CONTACT,
      [ContactType.TP]: Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT,
      [ContactType.DP]: Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT
    };
    return typeMap[parentContentType] || null;
  };

  private initTypeChoices = async () => {
    const typeChoices: IDropdownOption[] = [
      { key: Consts.CONTENT_TYPES.EUTELSAL_CONTACT, text: ContactType.Eutelsat },
      { key: Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT, text: ContactType.Administration },
      { key: Consts.CONTENT_TYPES.ORGANIZATION_CONTACT, text: ContactType.Organization },
      { key: Consts.CONTENT_TYPES.REGULATOR_CONTACT, text: ContactType.Regulator },
      { key: Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT, text: ContactType.TP },
      { key: Consts.CONTENT_TYPES.LAW_FIRM_CONTACT, text: ContactType.LawFirm },
      { key: Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE_CONTACT, text: ContactType.LegalRepresentative },
      { key: Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT, text: ContactType.DP }
    ];
    typeChoices.sort((a, b) => a.text?.localeCompare(b.text));
    if (!this.isEditMode) {
      const selectedType = this.getContactTypeFromParentContentType(this.props.preselectedType);
      const type = selectedType || (typeChoices.length > 0 ? typeChoices[0].key as string : '');
      const typeText = typeChoices.find((t) => t.key === type).text;
      const itemChoices = this.AllLookupItems.get(typeText) || [];
      this.setState({
        typeChoices,
        type,
        itemChoices
      });
    } else {
      this.setState({
        typeChoices
      });
    }
  };

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Contact).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.CONTACT.FIRST_NAME,
      Consts.FIELDS.CONTACT.LAST_NAME,
      Consts.FIELDS.CONTACT.SUMMARY,
      Consts.FIELDS.CONTACT.JOB_TITLE,
      Consts.FIELDS.CONTACT.DEPARTMENT,
      Consts.FIELDS.CONTACT.EMAIL,
      Consts.FIELDS.CONTACT.PHONE_NUMBER,
      Consts.FIELDS.CONTACT.COMMENTS,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      `${Consts.FIELDS.CONTACT.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.CONTACT.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.CONTACT.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.CONTACT.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
    ];
    const expand = [Consts.FIELDS.CONTACT.AUTHORITY,
    Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER,
    Consts.FIELDS.CONTACT.TELEPORT_PARTNER,
    Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER];
    const contactItem = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
    const contactContentTypeId = "" + this.state.typeChoices.find((type) => contactItem[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1)?.key;
    const contactType = "" + this.state.typeChoices.find((type) => contactItem[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1)?.text;

    let item: IDropdownOption | null;
    const itemChoices: IDropdownOption[] = this.getLookupItemsByType(contactType);
    if (contactContentTypeId === Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT) {
      item = {
        key: contactItem[Consts.FIELDS.CONTACT.AUTHORITY][Consts.FIELDS.COMMON.ID],
        text: contactItem[Consts.FIELDS.CONTACT.AUTHORITY][Consts.FIELDS.COMMON.TITLE],
      };
    } else if (contactContentTypeId === Consts.CONTENT_TYPES.ORGANIZATION_CONTACT) {
      item = {
        key: contactItem[Consts.FIELDS.CONTACT.AUTHORITY][Consts.FIELDS.COMMON.ID],
        text: contactItem[Consts.FIELDS.CONTACT.AUTHORITY][Consts.FIELDS.COMMON.TITLE],
      };
    } else if (contactContentTypeId === Consts.CONTENT_TYPES.REGULATOR_CONTACT) {
      item = {
        key: contactItem[Consts.FIELDS.CONTACT.AUTHORITY][Consts.FIELDS.COMMON.ID],
        text: contactItem[Consts.FIELDS.CONTACT.AUTHORITY][Consts.FIELDS.COMMON.TITLE],
      };
    } else if (contactContentTypeId === Consts.CONTENT_TYPES.LAW_FIRM_CONTACT) {
      item = {
        key: contactItem[Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER][Consts.FIELDS.COMMON.ID],
        text: contactItem[Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER][Consts.FIELDS.COMMON.TITLE],
      };
    } else if (contactContentTypeId === Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT) {
      item = {
        key: contactItem[Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
        text: contactItem[Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE],
      };
    } else if (contactContentTypeId === Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT) {
      item = {
        key: contactItem[Consts.FIELDS.CONTACT.TELEPORT_PARTNER][Consts.FIELDS.COMMON.ID],
        text: contactItem[Consts.FIELDS.CONTACT.TELEPORT_PARTNER][Consts.FIELDS.COMMON.TITLE],
      };
    } else if (contactContentTypeId === Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE_CONTACT) {
      item = {
        key: contactItem[Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER][Consts.FIELDS.COMMON.ID],
        text: contactItem[Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER][Consts.FIELDS.COMMON.TITLE],
      };
    }
    const initialValues = {
      firstname: contactItem[Consts.FIELDS.CONTACT.FIRST_NAME],
      lastname: contactItem[Consts.FIELDS.CONTACT.LAST_NAME],
      type: contactContentTypeId,
      summary: contactItem[Consts.FIELDS.CONTACT.SUMMARY],
      item,
      itemChoices,
      jobtitle: contactItem[Consts.FIELDS.CONTACT.JOB_TITLE],
      department: contactItem[Consts.FIELDS.CONTACT.DEPARTMENT],
      email: contactItem[Consts.FIELDS.CONTACT.EMAIL],
      phonenumber: contactItem[Consts.FIELDS.CONTACT.PHONE_NUMBER],
      comments: contactItem[Consts.FIELDS.CONTACT.COMMENTS],
    };
    this.setState({
      ...initialValues
    });
  };

  private validateForm = (): boolean => {
    const errors: { [key: string]: string } = {};
    const {
      firstname,
      lastname,
      email,
      phonenumber,
      type,
      item
    } = this.state;
    if (firstname?.trim() === '' || !firstname)
      errors.firstname = strings.FormContactFirstNameValidationError;
    if (lastname?.trim() === '' || !lastname)
      errors.lastname = strings.FormContactLastNameValidationError;
    if (email?.trim() && !UtilHelper.validateEmail(email)) {
      errors.email = strings.FormContactEmailValidationError;
    }
    if (phonenumber?.trim() && !UtilHelper.validatePhoneNumber(phonenumber)) {
      errors.phonenumber = strings.FormContactPhoneNumberValidationError;
    }
    if (type !== Consts.CONTENT_TYPES.EUTELSAL_CONTACT && !item) {
      errors.relatedItem = strings.FormContactItemValidationError;
    }
    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

  public render(): React.ReactElement<IContactProps> {
    const {
      errors,
      isFormReady,
      firstname,
      lastname,
      summary,
      jobtitle,
      department,
      email,
      phonenumber,
      comments,
      typeChoices,
      type,
      item,
      itemChoices
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
                    {/* First Name */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormContactFirstNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormContactTooltipFirstName} ></Tooltip>
                      </div>
                      <TextField
                        value={firstname}
                        onChange={this.onChangeFirstName}
                      />
                      {errors.firstname &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.firstname}
                          </span>
                        </div>
                      }
                    </div>
                    {/* Last Name */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormContactLastNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormContactTooltipLastName} ></Tooltip>
                      </div>
                      <TextField
                        value={lastname}
                        onChange={this.onChangeLastName}
                      />
                      {errors.lastname &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.lastname}
                          </span>
                        </div>
                      }
                    </div>
                    {/* Summary */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormContactSummaryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormContactTooltipSummary} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
                    </div>
                    {/* Type */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormContactTypeLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormContactTooltipType} ></Tooltip>
                      </div>
                      <Dropdown
                        selectedKey={type} disabled={this.isEditMode || !!this.props.preselectedType}
                        onChange={this.onChangeType}
                        options={typeChoices} />
                    </div>
                    {/* Item */}
                    {type !== Consts.CONTENT_TYPES.EUTELSAL_CONTACT &&
                      <div className={styles.field} >
                        <div className={styles.fieldLabel}>
                          <div className={styles.fieldLabelContainer}>
                            <Label required>{strings.FormContactItemLabel}</Label>
                          </div>
                          <Tooltip content={strings.FormContactTooltipItem} ></Tooltip>
                        </div>
                        <Combobox
                          data={itemChoices}
                          value={item}
                          disabled={this.state.isItemPreselected}
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
                    {/* Job Title */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormContactJobTitleLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormContactTooltipJobTitle} ></Tooltip>
                      </div>
                      <TextField
                        value={jobtitle}
                        onChange={this.onChangeJobTitle}
                      />
                    </div>
                    {/* Department */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormContactDepartmentLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormContactTooltipDepartment} ></Tooltip>
                      </div>
                      <TextField
                        value={department}
                        onChange={this.onChangeDepartment}
                      />
                    </div>
                    {/* Email */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormContactEmailLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormContactTooltipEmail} ></Tooltip>
                      </div>
                      <TextField
                        value={email}
                        onChange={this.onChangeEmail}
                      />
                      {errors.email &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.email}
                          </span>
                        </div>
                      }
                    </div>
                    {/* Phone number */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormContactPhoneNumberLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormContactTooltipPhoneNumber} ></Tooltip>
                      </div>
                      <PhoneInput
                        international={true}
                        value={phonenumber}
                        onChange={this.onChangePhoneNumber}
                      />
                      {errors.phonenumber &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.phonenumber}
                          </span>
                        </div>
                      }
                    </div>
                    {/* Comments */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormEutEntityCommentsLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormContactTooltipComments} ></Tooltip>
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

  private onChangeFirstName = (event, firstname: string) => {
    this.setState({ firstname });
  };

  private onChangeLastName = (event, lastname: string) => {
    this.setState({ lastname });
  };

  private onChangeSummary = (summary: string) => {
    this.setState({ summary });
    return summary;
  };

  private onChangeJobTitle = (event, jobtitle: string) => {
    this.setState({ jobtitle });
    return jobtitle;
  };

  private onChangeDepartment = (event, department: string) => {
    this.setState({ department });
    return department;
  };

  private onChangeEmail = (event, email: string) => {
    this.setState({ email });

    return email;
  };

  private onChangePhoneNumber = (value: string | undefined) => {
    this.setState({ phonenumber: value || '' });
  };

  private onChangeType = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ type: option.key.toString() });
    const itemChoices = this.getLookupItemsByType(option.text.toString());
    this.setState({ itemChoices, item: null });
  };

  private onChangeComments = (comments: string) => {
    this.setState({ comments });
    return comments;
  };

  private onChangeItem = (item: IDropdownOption) => {
    this.setState({ item });
  };

  private getLookupItemsByType = (contentTypeText: string) => {
    if (this.AllLookupItems.has(contentTypeText)) {
      const cachedItems = this.AllLookupItems.get(contentTypeText) || [];
      return cachedItems.slice().sort((a, b) => a.text?.localeCompare(b.text));
    }
  };

  private getDuplicateItems = async (): Promise<IBaseItem[]> => {
    const {
      pnpService,
      itemId
    } = this.props;
    const {
      title
    } = this.state;
    const duplicateItems: IBaseItem[] = [];
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields)();
    const match = items.find((i) => (i.Title?.toLowerCase() === UtilHelper.sanitizeSharePointTitle(title).toLowerCase()) && i.Id !== itemId);
    if (match)
      duplicateItems.push({ Id: match[Consts.FIELDS.COMMON.ID], Title: match[Consts.FIELDS.COMMON.TITLE] });
    return duplicateItems;
  };

  private getFormSummary = () => {
    const {
      firstname,
      lastname,
      summary,
      jobtitle,
      department,
      email,
      phonenumber,
      comments,
      typeChoices,
      type,
      item,
    } = this.state;
    const summaryTextValue = DomHelper.cleanRichHtml(summary);
    const commentsTextValue = DomHelper.cleanRichHtml(comments);
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormContactDatasheetNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormContactFirstNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {firstname}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormContactLastNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {lastname}
          </div>
        </div>

        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormContactSummaryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {summaryTextValue.length > 0 ? <RichText isEditMode={false} value={summaryTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormContactTypeLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {typeChoices.find((choice) => choice.key === type)?.text}
          </div>
        </div>
        {type !== Consts.CONTENT_TYPES.EUTELSAL_CONTACT &&
          <div className={styles.field}>
            <div className={styles.fieldLabel}>
              <div className={styles.fieldLabelContainer}>
                <Label>{strings.FormContactItemLabel}</Label>
              </div>
            </div>
            <div className={styles.fieldValue}>
              {item?.text}
            </div>
          </div>
        }
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormContactJobTitleLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {jobtitle}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormContactDepartmentLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {department}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormContactEmailLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {email}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormContactPhoneNumberLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {phonenumber}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormContactCommentsLabel}</Label>
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
      firstname,
      lastname,
      summary,
      comments,
      typeChoices,
      type,
      jobtitle,
      department,
      email,
      phonenumber,
      item
    } = this.state;
    const {
      pnpService,
    } = this.props;
    const itemData = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = `${firstname} ${lastname}`;
    itemData[Consts.FIELDS.CONTACT.FIRST_NAME] = firstname;
    itemData[Consts.FIELDS.CONTACT.LAST_NAME] = lastname;
    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = typeChoices.find((choice) => choice.key === type)?.key;
    itemData[Consts.FIELDS.CONTACT.SUMMARY] = summary;
    if (type === Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT || type === Consts.CONTENT_TYPES.ORGANIZATION_CONTACT || type === Consts.CONTENT_TYPES.REGULATOR_CONTACT) {
      itemData[Consts.FIELDS.CONTACT.AUTHORITY_ID] = item?.key;
    } else if (type === Consts.CONTENT_TYPES.LAW_FIRM_CONTACT || type === Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE_CONTACT) {
      itemData[Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER_ID] = item?.key;
    } else if (type === Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT) {
      itemData[Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER_ID] = item?.key;
    } else if (type === Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT) {
      itemData[Consts.FIELDS.CONTACT.TELEPORT_PARTNER_ID] = item?.key;
    }
    console.log('Item Key:', item?.key);
    itemData[Consts.FIELDS.CONTACT.FIRST_NAME] = firstname;
    itemData[Consts.FIELDS.CONTACT.LAST_NAME] = lastname;
    itemData[Consts.FIELDS.CONTACT.JOB_TITLE] = jobtitle;
    itemData[Consts.FIELDS.CONTACT.DEPARTMENT] = department;
    itemData[Consts.FIELDS.CONTACT.EMAIL] = email;
    itemData[Consts.FIELDS.CONTACT.PHONE_NUMBER] = phonenumber;
    itemData[Consts.FIELDS.CONTACT.COMMENTS] = comments;
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);
    return createdItem;
  };

  private processEdit = async () => {
    const {
      firstname,
      lastname,
      summary,
      comments,
      typeChoices,
      type,
      jobtitle,
      department,
      email,
      phonenumber,
      item
    } = this.state;
    const {
      pnpService,
      itemId,
    } = this.props;
    const itemData = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = `${firstname} ${lastname}`;
    itemData[Consts.FIELDS.CONTACT.FIRST_NAME] = firstname;
    itemData[Consts.FIELDS.CONTACT.LAST_NAME] = lastname;
    itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = typeChoices.find((choice) => choice.key === type)?.key;
    itemData[Consts.FIELDS.CONTACT.SUMMARY] = summary;
    if (type === Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT || type === Consts.CONTENT_TYPES.ORGANIZATION_CONTACT || type === Consts.CONTENT_TYPES.REGULATOR_CONTACT) {
      itemData[Consts.FIELDS.CONTACT.AUTHORITY_ID] = item?.key;
    } else if (type === Consts.CONTENT_TYPES.LAW_FIRM_CONTACT || type === Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE_CONTACT) {
      itemData[Consts.FIELDS.CONTACT.LEGAL_SVC_PROVIDER_ID] = item?.key;
    } else if (type === Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT) {
      itemData[Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER_ID] = item?.key;
    } else if (type === Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT) {
      itemData[Consts.FIELDS.CONTACT.TELEPORT_PARTNER_ID] = item?.key;
    }
    itemData[Consts.FIELDS.CONTACT.JOB_TITLE] = jobtitle;
    itemData[Consts.FIELDS.CONTACT.DEPARTMENT] = department;
    itemData[Consts.FIELDS.CONTACT.EMAIL] = email;
    itemData[Consts.FIELDS.CONTACT.PHONE_NUMBER] = phonenumber;
    itemData[Consts.FIELDS.CONTACT.COMMENTS] = comments;
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
  };
}
