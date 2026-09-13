import * as React from 'react';
import { SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import { Dropdown, Fabric, Icon, IDropdownOption, Label, TextField, Toggle } from 'office-ui-fabric-react';
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
import { ModernTaxonomyPicker } from "@pnp/spfx-controls-react/lib/ModernTaxonomyPicker";
import { ITermInfo } from '@pnp/spfx-controls-react/node_modules/@pnp/sp/taxonomy';
import { Config } from '../../../../common/config/Config';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { ITooltipHostStyles, LayerHost, TooltipHost } from 'office-ui-fabric-react';
import { Combobox } from 'react-widgets/cjs';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import { Text } from '@microsoft/sp-core-library';
import { createUniqueFolder } from '../Form.utility';
import { TaxonomyService } from '../../../../common/services/TaxonomyService';
import { ITaxonomyTerm } from '../../../../common/models/SPEntities';
import { IMarketProps, IMarketState, ISanctionOption } from './Market.types';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import { UserRole } from '../../../../common/models/Enums';
import DomHelper from '../../../../common/helpers/DomHelper';

export default class Market extends React.Component<
  IMarketProps,
  IMarketState
> {
  private isEditMode: boolean;
  private config: IFormConfig;

  constructor(props: Readonly<IMarketProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      name: '',
      country: null,
      countryChoices: [],
      groupings: [],
      isSanctioned: false,
      sanctionCategory: null,
      sanctionCategoryChoices: [],
      maPriorities: [],
      maPrioritiesChoices: [],
      maCapacity: '',
      comments: '',
      errors: {},
      isFormReady: false
    };
    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.Market);
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
      this.initMaPrioritiesChoices());
    promises.push(
      this.initSanctionCategoryChoices());
    promises.push(
      this.initCountryChoices());
    await Promise.all(promises);
    if (this.isEditMode)
      await this.setInitialFormValues();
  };

  private initSanctionCategoryChoices = async () => {
    const {
      pnpService
    } = this.props;
    const sanctionCategoryChoices: ISanctionOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.SANCTION_CATEGORY_URL);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.SANCTION_CATEGORY.COLOR,
      Consts.FIELDS.SANCTION_CATEGORY.SUMMARY
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      const sanctionCategory: ISanctionOption = {
        key: item[Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.COMMON.TITLE],
        color: item[Consts.FIELDS.SANCTION_CATEGORY.COLOR],
        summary: item[Consts.FIELDS.SANCTION_CATEGORY.SUMMARY]
      };
      sanctionCategoryChoices.push(sanctionCategory);
    });

    if (!this.isEditMode) {
      const sanctionCategory = sanctionCategoryChoices.length > 0 ? sanctionCategoryChoices[0] as ISanctionOption : null;
      this.setState({ sanctionCategory, sanctionCategoryChoices });
    } else {
      this.setState({ sanctionCategoryChoices });
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

  public initMaPrioritiesChoices = async () => {
    const {
      pnpService
    } = this.props;
    let maPrioritiesChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.COUNTRY.MARKET_ACCESS_PRIORITY)
      .select('Choices')();
    maPrioritiesChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    if (!this.isEditMode) {
      const maPriorities = [];
      this.setState({
        maPriorities, maPrioritiesChoices
      });
    } else {
      this.setState({
        maPrioritiesChoices
      });
    }
  };

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const { sanctionCategoryChoices } = this.state;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Market).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);

    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.TITLE,
      `${Consts.FIELDS.MARKET.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.MARKET.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
      Consts.FIELDS.MARKET.COMMENTS,
      Consts.FIELDS.MARKET.MARKET_ACCESS_CAPACITY,
      Consts.FIELDS.MARKET.MARKET_ACCESS_PRIORITY,
      Consts.FIELDS.MARKET.SANCTION_CATEGORY_ID,
      Consts.FIELDS.MARKET.IS_SANCTIONED,
      Consts.FIELDS.MARKET.GROUPING
    ];
    const expand = [Consts.FIELDS.MARKET.COUNTRY];

    const item = await pnpService
      .getListByUrl(listUrl)
      .items.getById(itemId)
      .select(...selectedFields)
      .expand(...expand)();

    const taxonomyService = new TaxonomyService();
    await taxonomyService.init(pnpService);
    const groupings: ITaxonomyTerm[] = [];
    if (item[Consts.FIELDS.MARKET.GROUPING]) {
      await Promise.all(item[Consts.FIELDS.MARKET.GROUPING].map(async (g) => {
        const localizedLabelGrouping = await taxonomyService.getLocalizedTerm(
          g?.TermGuid,
          'en-us'
        );

        groupings.push({
          id: g?.TermGuid,
          labels: [{
            name: localizedLabelGrouping,
            languageTag: 'en-US',
            isDefault: true
          }
          ]
        });
      }));
    }


    const initialValues = {
      name: item[Consts.FIELDS.COMMON.TITLE],
      country: {
        key: item[Consts.FIELDS.MARKET.COUNTRY][Consts.FIELDS.COMMON.ID],
        text: item[Consts.FIELDS.MARKET.COUNTRY][Consts.FIELDS.COMMON.TITLE]
      },
      comments: item[Consts.FIELDS.MARKET.COMMENTS],
      maCapacity: item[Consts.FIELDS.MARKET.MARKET_ACCESS_CAPACITY],
      groupings,
      isSanctioned: item[Consts.FIELDS.MARKET.IS_SANCTIONED] || false,
      sanctionCategory: sanctionCategoryChoices.find((cat) => cat.key === item[Consts.FIELDS.MARKET.SANCTION_CATEGORY_ID]),
      maPriorities: item[Consts.FIELDS.MARKET.MARKET_ACCESS_PRIORITY]?.map((priority) => ({
        key: priority,
        text: priority
      })) || [],
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
      groupings
    } = this.state;
    if (name?.trim() === '' || !name)
      errors.name = strings.FormMarketNameValidationError;

    if (!country || !country.key || country.key === '')
      errors.country = strings.FormEutEntityCountryValidationError;

    if (!Array.isArray(groupings) || groupings.length === 0) {
      errors.groupings = strings.FormMarketGroupingValidationError;
    }

    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };


  private onChangeIsSanctioned = (event, checked: boolean) => {
    this.setState({ isSanctioned: checked });
  };

  private onChangeMaPriorities = (maPriorities: IDropdownOption[]) => {
    this.setState({ maPriorities });
  };
  public render(): React.ReactElement<IMarketProps> {
    const {
      name,
      country,
      countryChoices,
      isSanctioned,
      sanctionCategory,
      sanctionCategoryChoices,
      maPriorities,
      maPrioritiesChoices,
      maCapacity,
      comments,
      groupings,
      errors,
      isFormReady
    } = this.state;
    const swatch = (color: string) => ({
      width: 12,
      height: 12,
      borderRadius: 3,
      background: color,
      display: 'inline-block',
      verticalAlign: 'middle',
      marginRight: 5,
      border: '1px solid rgba(0,0,0,.1)',
    });

    const tooltipStyles: ITooltipHostStyles = { root: { display: 'block' } };
    const renderOption = (opt?: ISanctionOption): JSX.Element | null => {
      if (!opt) return null;
      const color = opt.color ?? '#888';
      const summary = opt.summary ?? '';
      // Tooltip for long summaries, while also showing an inline, muted line
      return (
        <TooltipHost content={<div style={{ color: '#666' }} dangerouslySetInnerHTML={{ __html: summary }}></div>}
          styles={tooltipStyles} calloutProps={{
            layerProps: { hostId: 'sanctionCategoryLayer' },
            styles: { root: { zIndex: 500002 } }
          }}>
          <div style={{ verticalAlign: "middle" }} title="">
            <span style={swatch(color)} />
            <span style={{ verticalAlign: "middle" }}>{opt.text}</span>
          </div>
        </TooltipHost>
      );
    };
    const renderTitle = (opts?: ISanctionOption[]): JSX.Element | null => {
      const opt = opts?.[0];
      if (!opt) return null;
      const color = opt.color ?? '#888';
      return (
        <div style={{ verticalAlign: "middle" }} title="">
          <span style={swatch(color)} />
          <span style={{ verticalAlign: "middle" }}>{opt.text}</span>
        </div>
      );
    };
    const renderSanctionCategoryTooltip = (): string => {
      const itemsHtml = sanctionCategoryChoices
        .map((item) =>
          Text.format(
            strings.FormMarketTooltipSanctionCategoryItemTemplate,
            item.color,
            item.text,
            item.summary
          )
        )
        .join('');

      return `<div style='font-size:14px'>${strings.FormMarketTooltipSanctionCategoryHeader}</div><div>&nbsp;</div><div>${itemsHtml}</div>`;
    };
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
                          <Label required>{strings.FormMarketNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormMarketTooltipName} ></Tooltip>
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
                    {/* Country */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormMarketCountryLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormMarketTooltipCountry} ></Tooltip>
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
                    {/* Taxonomy Picker */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormMarketGroupingLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormMarketTooltipGrouping} />
                      </div>
                      <ModernTaxonomyPicker
                        allowMultipleSelections={true}
                        termSetId={Config.TAXONOMY.GROUPING_TERMSETID}
                        panelTitle={strings.FormMarketGroupingLabel}
                        label=''
                        onChange={this.onChangeGrouping}
                        context={this.props.pnpService.getContext()}
                        placeHolder={strings.FormMarketGroupingPlaceHolder}
                        initialValues={groupings ? groupings : []} />
                    </div>
                    {errors.groupings &&
                      <div className={styles.errorContainer}>
                        <Icon iconName="Error" className={styles.errorIcon} />
                        <span className={styles.errorMessage}>
                          {errors.groupings}
                        </span>
                      </div>
                    }

                    {/* Is Sanctioned */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormMarketIsSanctionedLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormMarketTooltipIsSanctioned} ></Tooltip>
                      </div>
                      <Toggle
                        onText="Yes" offText="No"
                        checked={isSanctioned}
                        onChange={this.onChangeIsSanctioned}
                      />
                    </div>

                    {/* Sanction category */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormMarketSanctionCategoryLabel}</Label>
                        </div>
                        <Tooltip content={renderSanctionCategoryTooltip()} ></Tooltip>
                      </div>
                      <LayerHost
                        id='sanctionCategoryLayer'
                        style={{ position: 'relative', zIndex: 500000 }} // higher than page chrome
                      />
                      <Dropdown
                        options={sanctionCategoryChoices}
                        selectedKey={sanctionCategory?.key}
                        onChange={this.onChangeSanctionCategory}
                        onRenderOption={renderOption}
                        onRenderTitle={renderTitle}
                        calloutProps={{
                          layerProps: { hostId: 'sanctionCategoryLayer' },
                          styles: { root: { zIndex: 500001 } }
                        }}
                      />
                    </div>

                    {/* MA Priorities */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormMarketPriorityLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormMarketTooltipPriority} ></Tooltip>
                      </div>
                      <MultiselectWrapper
                        data={maPrioritiesChoices}
                        dataKey={(item: IDropdownOption) => item.key}
                        textField={(item: IDropdownOption) => item.text}
                        value={maPriorities}
                        filter="contains"
                        onChange={this.onChangeMaPriorities} />
                    </div>

                    {/* MA Capacity */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormMarketCapacityLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormMarketTooltipCapacity} ></Tooltip>
                      </div>
                      <TextField
                        value={maCapacity}
                        onChange={this.onChangeCapacity}
                        type="number"
                      />

                    </div>
                    {/* Comments */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormMarketCommentsLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormMarketTooltipComments} ></Tooltip>
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

  private onChangeCountry = (country: IDropdownOption) => {
    this.setState({ country });
  };

  private onChangeCapacity = (event, maCapacity: string) => {
    this.setState({ maCapacity });
  };

  private onChangeSanctionCategory = (event: React.FormEvent<HTMLDivElement>, option: ISanctionOption) => {
    this.setState({ sanctionCategory: option });
  };

  private onChangeGrouping = (info?: ITermInfo[]): void => {
    this.setState({ groupings: info });
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
      country,
      isSanctioned,
      sanctionCategory,
      maPriorities,
      comments,
      groupings,
      maCapacity
    } = this.state;
    const maPrioritiesValues = maPriorities.map((c) => c.text).join(", ");
    const groupingsValues = groupings?.map((e) => e?.labels?.[0]?.name).join(", ");
    const commentsTextValue = DomHelper.cleanRichHtml(comments);
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMarketDatasheetLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMarketNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {UtilHelper.sanitizeSharePointTitle(name)}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMarketCountryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {country?.text}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMarketGroupingLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {groupingsValues}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMarketIsSanctionedLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {isSanctioned ? 'Yes' : 'No'}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMarketSanctionCategoryLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {sanctionCategory?.text}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMarketPriorityLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {maPrioritiesValues}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMarketCapacityLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {maCapacity}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMarketCommentsLabel}</Label>
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
      country,
      groupings,
      isSanctioned,
      sanctionCategory,
      maPriorities,
      maCapacity,
      comments,
    } = this.state;

    const {
      pnpService,
    } = this.props;
    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const groupingHiddenField = await pnpService.getListByUrl(listUrl).fields.getByTitle(`${Consts.FIELDS.COUNTRY.GROUPING}_0`)();

    const itemData = {};
    const sanitizedTitle = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.COMMON.TITLE] = sanitizedTitle;
    itemData[Consts.FIELDS.MARKET.COUNTRY_ID] = parseInt(country.key.toString());
    itemData[Consts.FIELDS.MARKET.COMMENTS] = comments;
    itemData[Consts.FIELDS.MARKET.MARKET_ACCESS_CAPACITY] = maCapacity;
    itemData[Consts.FIELDS.MARKET.MARKET_ACCESS_PRIORITY] = maPriorities?.map((item) => item.key.toString());
    itemData[Consts.FIELDS.MARKET.SANCTION_CATEGORY_ID] = sanctionCategory?.key?.toString();
    itemData[Consts.FIELDS.MARKET.IS_SANCTIONED] = isSanctioned;
    if (Array.isArray(groupings) && groupings.length > 0) {
      itemData[groupingHiddenField.InternalName] = groupings
        .map((g) => `-1;#${g.labels[0].name}|${g.id}`)
        .join(";#");
    }

    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);
    const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
    const documentSpaceUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.DOCUMENT_SPACE_URL);
    const parentFolderUrl = `${documentSpaceUrl}${Consts.DOCUMENT_SPACE.MARKET_FOLDER_URL}`;
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
      country,
      groupings,
      isSanctioned,
      sanctionCategory,
      maPriorities,
      maCapacity,
      comments,
    } = this.state;

    const {
      pnpService,
      itemId,
    } = this.props;

    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const groupingHiddenField = await pnpService.getListByUrl(listUrl).fields.getByTitle(`${Consts.FIELDS.MARKET.GROUPING}_0`)();

    const itemData = {};
    const sanitizedTitle = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.COMMON.TITLE] = sanitizedTitle;
    itemData[Consts.FIELDS.MARKET.COMMENTS] = comments;
    itemData[Consts.FIELDS.MARKET.COUNTRY_ID] = parseInt(country.key.toString());
    itemData[Consts.FIELDS.MARKET.MARKET_ACCESS_CAPACITY] = maCapacity;
    itemData[Consts.FIELDS.MARKET.MARKET_ACCESS_PRIORITY] = maPriorities?.map((item) => item.key.toString());
    itemData[Consts.FIELDS.MARKET.IS_SANCTIONED] = isSanctioned;
    itemData[Consts.FIELDS.MARKET.SANCTION_CATEGORY_ID] = sanctionCategory?.key?.toString();
    if (Array.isArray(groupings) && groupings.length > 0) {
      itemData[groupingHiddenField.InternalName] = groupings
        .map((g) => `-1;#${g.labels[0].name}|${g.id}`)
        .join(";#");
    } else {
      itemData[groupingHiddenField.InternalName] = null;
    }

    await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
  };
}
