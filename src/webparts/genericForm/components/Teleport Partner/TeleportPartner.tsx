import * as React from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import { ITeleportPartnerProps, ITeleportPartnerState } from './TeleportPartner.type';
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
export default class TeleportPartner extends React.Component<
  ITeleportPartnerProps,
  ITeleportPartnerState
> {
  private isEditMode: boolean;
  private config: IFormConfig;

  constructor(props: Readonly<ITeleportPartnerProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      name: '',
      summary: '',
      address: '',
      comments: '',
      status: '',
      statusChoices: [],
      markets: [],
      marketsChoices: [],
      errors: {},
      isFormReady: false,
      portal: '',
      eutelsatOwnersChoices: [],
      eutelsatOwners: []
    };
    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.TP);
    this.validateForm = this.validateForm.bind(this);
    this.getDuplicatetems = this.getDuplicatetems.bind(this);
    this.getFormSummary = this.getFormSummary.bind(this);
    this.callback = this.callback.bind(this);
    this.processCreation = this.processCreation.bind(this);
    this.processEdit = this.processEdit.bind(this);
    this.initStatusChoices = this.initStatusChoices.bind(this);
    this.initMarketChoices = this.initMarketChoices.bind(this);
    this.initEutelsatOwnersChoices = this.initEutelsatOwnersChoices.bind(this);
    this.onChangeName = this.onChangeName.bind(this);
    this.onChangeStatus = this.onChangeStatus.bind(this);
    this.onChangeMarkets = this.onChangeMarkets.bind(this);
    this.onChangeSummary = this.onChangeSummary.bind(this);
    this.onChangeAddress = this.onChangeAddress.bind(this);
    this.onChangeComments = this.onChangeComments.bind(this);
    this.onChangeEutelsatOwners = this.onChangeEutelsatOwners.bind(this);
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

  private async init() {
    const promises = [];
    promises.push(
      this.initStatusChoices()
    );
    promises.push(
      this.initMarketChoices());
    promises.push(
      this.initEutelsatOwnersChoices()
    );
    await Promise.all(promises);
    if (this.isEditMode)
      await this.setInitialFormValues();
  }
  public async initStatusChoices() {
    const {
      pnpService
    } = this.props;
    let statusChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.TELEPORT_PARTNER.STATUS)
      .select('Choices')();
    statusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    const status = statusChoices.length > 0 ? statusChoices[0].key as string : '';
    this.setState({
      statusChoices, status
    });
  }

  public async initMarketChoices() {
    const {
      pnpService
    } = this.props;
    const marketsChoices: IDropdownOption[] = [];
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Market).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      marketsChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    const markets = [];
    this.setState({ marketsChoices, markets });
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
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.TP).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.TELEPORT_PARTNER.SUMMARY,
      Consts.FIELDS.TELEPORT_PARTNER.ADDRESS,
      Consts.FIELDS.TELEPORT_PARTNER.STATUS,
      Consts.FIELDS.TELEPORT_PARTNER.COMMENTS,
      Consts.FIELDS.COMMON.CONTENTTYPE_ID,
      Consts.FIELDS.TELEPORT_PARTNER.PORTAL,
      `${Consts.FIELDS.TELEPORT_PARTNER.MARKETS}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.TELEPORT_PARTNER.MARKETS}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.TELEPORT_PARTNER.OWNER}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.TELEPORT_PARTNER.OWNER}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const expand = [Consts.FIELDS.TELEPORT_PARTNER.MARKETS, Consts.FIELDS.TELEPORT_PARTNER.OWNER];
    const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
    const initialValues = {
      name: item[Consts.FIELDS.COMMON.TITLE],
      summary: item[Consts.FIELDS.TELEPORT_PARTNER.SUMMARY],
      address: item[Consts.FIELDS.TELEPORT_PARTNER.ADDRESS],
      comments: item[Consts.FIELDS.TELEPORT_PARTNER.COMMENTS],
      status: item[Consts.FIELDS.TELEPORT_PARTNER.STATUS],
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      markets: item[Consts.FIELDS.TELEPORT_PARTNER.MARKETS]?.map((market: Record<string, any>) => ({
        key: market[Consts.FIELDS.COMMON.ID],
        text: market[Consts.FIELDS.COMMON.TITLE]
      })) || [],
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      eutelsatOwners: item[Consts.FIELDS.TELEPORT_PARTNER.OWNER]?.map((etuowner: Record<string, any>) => ({
        key: etuowner[Consts.FIELDS.COMMON.ID],
        text: etuowner[Consts.FIELDS.COMMON.TITLE]
      })) || [],
      portal: item[Consts.FIELDS.TELEPORT_PARTNER.PORTAL]?.Url?.trim() || '',
    };
    console.log("after assigning owner", initialValues.eutelsatOwners);

    this.setState({
      ...initialValues
    });
    console.log("After setting state", initialValues);
  };
  private validateForm(): boolean {
    const errors: { [key: string]: string } = {};
    const {
      name,
      markets,
      portal
    } = this.state;
    if (name?.trim() === '' || !name)
      errors.name = strings.FormEutEntityNameValidationError;


    if (!Array.isArray(markets) || markets.length === 0) {
      errors.market = strings.FormTeleportPartnerMarketValidationError;
    }

    if (portal?.trim() && !UtilHelper.validateURL(portal)) {
      errors.portal = strings.FormTeleportPartnerPortalValidationError;
    }
    this.setState({ errors });
    return Object.keys(errors).length === 0;
  }

  public render(): React.ReactElement<ITeleportPartnerProps> {
    const {
      name,
      summary,
      address,
      markets,
      marketsChoices,
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
                          <Label required>{strings.FormTeleportPartnerNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormTeleportPartnerTooltipName} ></Tooltip>
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
                          <Label>{strings.FormTeleportPartnerSummaryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormTeleportPartnerTooltipSummary} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
                    </div>


                    {/* Address */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormTeleportPartnerAddressLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormTeleportPartnerTooltipAddress} ></Tooltip>
                      </div>
                      <RichText isEditMode={true} value={address} onChange={this.onChangeAddress} />
                    </div>


                    {/* markets */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormTeleportPartnerMarketLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormTeleportPartnerTooltipMarket} ></Tooltip>
                      </div>
                      <MultiselectWrapper data={marketsChoices}
                        dataKey={(item: IDropdownOption) => item.key}
                        textField={(item: IDropdownOption) => item.text}
                        value={markets}
                        filter="contains"
                        onChange={this.onChangeMarkets} />
                      {errors.market &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.market}
                          </span>
                        </div>
                      }
                    </div>

                    {/* Eutelsat Owners*/}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormTeleportPartnerEutelsatOwnersLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormTeleportPartnerTooltipEutelsatOwners} ></Tooltip>
                      </div>
                      <MultiselectWrapper
                        data={eutelsatOwnersChoices}
                        dataKey={(item: IDropdownOption) => item.key}
                        textField={(item: IDropdownOption) => item.text}
                        value={eutelsatOwners}
                        filter="contains"
                        onChange={this.onChangeEutelsatOwners} />
                    </div>

                    {/*Portal*/}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormTeleportPartnerPortalLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormTeleportPartnerTooltipPortal} ></Tooltip>
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
                          <Label>{strings.FormTeleportPartnerStatusLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormTeleportPartnerTooltipStatus} ></Tooltip>
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
                        <Tooltip content={strings.FormTeleportPartnerTooltipComments} ></Tooltip>
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

  private onChangeMarkets(markets: IDropdownOption[]) {
    this.setState({ markets });

  }
  private onChangeEutelsatOwners(eutelsatOwners: IDropdownOption[]) {
    this.setState({ eutelsatOwners });
  }
  private onChangeStatus(event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) {
    this.setState({ status: option.key.toString() });
  }
  private onChangePortal = (event, portal: string) => {
    this.setState({ portal });
  };
  private onChangeComments(comments: string) {
    this.setState({ comments });
    return comments;
  }

  private async getDuplicatetems(): Promise<IBaseItem[]> {
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
  }

  private getFormSummary() {
    const {
      name,
      summary,
      address,
      markets,
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
              <Label>{strings.FormTeleportPartnerDatasheetNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormTeleportPartnerNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormTeleportPartnerSummaryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {summaryTextValue.length > 0 ? <RichText isEditMode={false} value={summaryTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormTeleportPartnerAddressLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {addressTextValue.length > 0 ? <RichText isEditMode={false} value={addressTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormTeleportPartnerMarketLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {markets.map((c) => c.text).join(", ")}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormTeleportPartnerEutelsatOwnersLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {eutelsatOwners.map((c) => c.text).join(", ")}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormTeleportPartnerPortalLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {portal}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormTeleportPartnerStatusLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {status}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormTeleportPartnerCommentsLabel}</Label>
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
      markets,
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
    itemData[Consts.FIELDS.TELEPORT_PARTNER.SUMMARY] = summary;
    itemData[Consts.FIELDS.TELEPORT_PARTNER.ADDRESS] = address;
    itemData[Consts.FIELDS.TELEPORT_PARTNER.MARKETS_ID] = markets.map((c) => c.key);
    itemData[Consts.FIELDS.TELEPORT_PARTNER.OWNERS_ID] = eutelsatOwners.map((c) => c.key);
    itemData[Consts.FIELDS.TELEPORT_PARTNER.STATUS] = status;
    itemData[Consts.FIELDS.TELEPORT_PARTNER.COMMENTS] = comments;
    itemData[Consts.FIELDS.TELEPORT_PARTNER.PORTAL] =
      { [Consts.FIELDS.COMMON.URL]: portal, [Consts.FIELDS.COMMON.DESCRIPTION]: "Link" };
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);

    const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
    const documentSpaceUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.DOCUMENT_SPACE_URL);

    const parentFolderUrl = `${documentSpaceUrl}${Consts.DOCUMENT_SPACE.TELEPORT_PARTNER_FOLDER_URL}`;
    const folderUrl = await createUniqueFolder(pnpService, parentFolderUrl, sanitizedFolderName);
    const updateData = {};
    updateData[Consts.FIELDS.COMMON.DOCUMENTS_SPACE] =
    {
      Url: folderUrl,
      Description: 'Link'
    },

      await pnpService.getListItems(listUrl).getById(createdItem.data.ID).update(updateData);
    return createdItem;


  }

  private processEdit = async () => {
    const {
      name,
      summary,
      address,
      markets,
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
    itemData[Consts.FIELDS.TELEPORT_PARTNER.SUMMARY] = summary;
    itemData[Consts.FIELDS.TELEPORT_PARTNER.ADDRESS] = address;
    itemData[Consts.FIELDS.TELEPORT_PARTNER.MARKETS_ID] = markets.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.TELEPORT_PARTNER.OWNERS_ID] =
      eutelsatOwners.map((c) => (parseInt(c.key.toString())));
    itemData[Consts.FIELDS.TELEPORT_PARTNER.COMMENTS] = comments;
    itemData[Consts.FIELDS.TELEPORT_PARTNER.STATUS] = status;
    itemData[Consts.FIELDS.TELEPORT_PARTNER.PORTAL] =
      { [Consts.FIELDS.COMMON.URL]: portal, [Consts.FIELDS.COMMON.DESCRIPTION]: "Link" };
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
  };

}
