import * as React from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import { IDistributionPartnerProps, IDistributionPartnerState } from './DistributionPartner.type';
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
import { UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';
export default class DistributionPartner extends React.Component<
  IDistributionPartnerProps,
  IDistributionPartnerState
> {
  private isEditMode: boolean;
  private config: IFormConfig;

  constructor(props: Readonly<IDistributionPartnerProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      name: '',
      summary: '',
      address: '',
      comments: '',
      status: '',
      statusChoices: [],
      errors: {},
      isFormReady: false,
      portal: '',
      eutelsatOwnersChoices: [],
      eutelsatOwners: []
    };
    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.DP);
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
      this.initEutelsatOwnersChoices()
    );
    await Promise.all(promises);
    if (this.isEditMode)
      await this.setInitialFormValues();
  };

  private initStatusChoices = async () => {
    const {
      pnpService
    } = this.props;
    let statusChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.DISTRIBUTION_PARTNER.STATUS)
      .select('Choices')();
    statusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    const status = statusChoices.length > 0 ? statusChoices[0].key as string : '';
    this.setState({
      statusChoices, status
    });
  };

  private initEutelsatOwnersChoices = async () => {
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
  };

  private setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.DP).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.DISTRIBUTION_PARTNER.SUMMARY,
      Consts.FIELDS.DISTRIBUTION_PARTNER.ADDRESS,
      Consts.FIELDS.DISTRIBUTION_PARTNER.STATUS,
      Consts.FIELDS.DISTRIBUTION_PARTNER.COMMENTS,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      Consts.FIELDS.DISTRIBUTION_PARTNER.PORTAL,
      `${Consts.FIELDS.DISTRIBUTION_PARTNER.OWNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.DISTRIBUTION_PARTNER.OWNER}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const expand = [Consts.FIELDS.DISTRIBUTION_PARTNER.OWNER];
    const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
    const initialValues = {
      name: item[Consts.FIELDS.COMMON.TITLE],
      summary: item[Consts.FIELDS.DISTRIBUTION_PARTNER.SUMMARY],
      address: item[Consts.FIELDS.DISTRIBUTION_PARTNER.ADDRESS],
      comments: item[Consts.FIELDS.DISTRIBUTION_PARTNER.COMMENTS],
      status: item[Consts.FIELDS.DISTRIBUTION_PARTNER.STATUS],
      eutelsatOwners: item[Consts.FIELDS.DISTRIBUTION_PARTNER.OWNER]?.map((eutowner) => ({
        key: eutowner[Consts.FIELDS.COMMON.ID],
        text: eutowner[Consts.FIELDS.COMMON.TITLE]
      })) || [],
      portal: item[Consts.FIELDS.DISTRIBUTION_PARTNER.PORTAL]?.Url?.trim() || '',
    };


    this.setState({
      ...initialValues
    });
  };

  private validateForm = (): boolean => {
    const errors: { [key: string]: string } = {};
    const {
      name,
      portal
    } = this.state;
    if (name?.trim() === '' || !name)
      errors.name = strings.FormEutEntityNameValidationError;

    if (portal?.trim() && !UtilHelper.validateURL(portal)) {
      errors.portal = strings.FormDistributionPartnerPortalValidationError;
    }
    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

  public render(): React.ReactElement<IDistributionPartnerProps> {
    const {
      name,
      summary,
      address,
      status,
      statusChoices,
      comments,
      errors,
      isFormReady,
      portal,
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
            getDuplicateItems={this.getDuplicatetems}
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
                          <Label required>{strings.FormDistributionPartnerNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormDistributionPartnerTooltipName} ></Tooltip>
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
                          <Label>{strings.FormDistributionPartnerSummaryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormDistributionPartnerTooltipSummary} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
                    </div>


                    {/* Address */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormDistributionPartnerAddressLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormDistributionPartnerTooltipAddress} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={address} onChange={this.onChangeAddress} />
                    </div>


                    {/* Eutelsat Owners*/}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormDistributionPartnerEutelsatOwnersLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormDistributionPartnerTooltipEutelsatOwners} ></Tooltip>
                      </div>
                      <MultiselectWrapper value={eutelsatOwners}
                        data={eutelsatOwnersChoices}
                        dataKey={(item: IDropdownOption) => item.key}
                        textField={(item: IDropdownOption) => item.text}
                        filter="contains"
                        onChange={this.onChangeEutelsatOwners} />
                    </div>
                    {/*Portal*/}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormDistributionPartnerPortalLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormDistributionPartnerTooltipPortal} ></Tooltip>
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
                    {/* Status */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormDistributionPartnerStatusLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormDistributionPartnerTooltipStatus} ></Tooltip>
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
                        <Tooltip content={strings.FormDistributionPartnerTooltipComments} ></Tooltip>
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

  private onChangeEutelsatOwners = (eutelsatOwners: IDropdownOption[]) => {
    this.setState({ eutelsatOwners });
  };

  private onChangeStatus = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({ status: option.key.toString() });
  };

  private onChangePortal = (event, portal: string) => {
    this.setState({ portal });
  };

  private onChangeComments = (comments: string) => {
    this.setState({ comments });
    return comments;
  };

  private getDuplicatetems = async (): Promise<IBaseItem[]> => {
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
      status,
      comments,
      portal,
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
              <Label>{strings.FormDistributionPartnerDatasheetNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormDistributionPartnerNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormDistributionPartnerSummaryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {summaryTextValue.length > 0 ? <RichText isEditMode={false} value={summaryTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormDistributionPartnerAddressLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {addressTextValue.length > 0 ? <RichText isEditMode={false} value={addressTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormDistributionPartnerEutelsatOwnersLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {eutelsatOwners.map((c) => c.text).join(", ")}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormDistributionPartnerPortalLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {portal}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormDistributionPartnerStatusLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {status}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormDistributionPartnerCommentsLabel}</Label>
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
      status,
      comments,
      portal,
      eutelsatOwners
    } = this.state;

    const {
      pnpService,
    } = this.props;
    const itemData = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = name;
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.SUMMARY] = summary;
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.ADDRESS] = address;
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.OWNERS_ID] = eutelsatOwners.map((c) => c.key);
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.STATUS] = status;
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.COMMENTS] = comments;
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.PORTAL] =
      { [Consts.FIELDS.COMMON.URL]: portal, [Consts.FIELDS.COMMON.DESCRIPTION]: "Link" };

    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);

    const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
    const documentSpaceUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.DOCUMENT_SPACE_URL);

    const parentFolderUrl = `${documentSpaceUrl}${Consts.DOCUMENT_SPACE.DISTRIBUTION_PARTNER_FOLDER_URL}`;
    const folderUrl = await createUniqueFolder(pnpService, parentFolderUrl, sanitizedFolderName);
    const updateData = {};
    updateData[Consts.FIELDS.COMMON.DOCUMENTS_SPACE] =
    {
      Url: folderUrl,
      Description: 'Link'
    },

      await pnpService.getListItems(listUrl).getById(createdItem.data.ID).update(updateData);
    return createdItem;

    //return createdItem;

  };

  private processEdit = async () => {
    const {
      name,
      summary,
      address,
      eutelsatOwners,
      status,
      comments,
      portal
    } = this.state;

    const {
      pnpService,
      itemId,
    } = this.props;



    const itemData = {};

    itemData[Consts.FIELDS.COMMON.TITLE] = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.SUMMARY] = summary;
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.ADDRESS] = address;
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.OWNERS_ID] =
      eutelsatOwners.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.COMMENTS] = comments;
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.STATUS] = status;
    itemData[Consts.FIELDS.DISTRIBUTION_PARTNER.PORTAL] =
      { [Consts.FIELDS.COMMON.URL]: portal, [Consts.FIELDS.COMMON.DESCRIPTION]: "Link" };
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
  };

}
