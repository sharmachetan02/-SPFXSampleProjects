import * as React from 'react';
import { SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { ICountryProps, ICountryyState } from './Country.types';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import { Fabric, Icon, Label, TextField } from 'office-ui-fabric-react';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { Form, FormOrigin, FormType, IFormConfig } from '../Form.types';
import strings from 'GenericFormWebPartStrings';
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
import { TaxonomyService } from '../../../../common/services/TaxonomyService';
import { ITaxonomyTerm } from '../../../../common/models/SPEntities';
import { UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';

export default class Country extends React.Component<
  ICountryProps,
  ICountryyState
> {
  private isEditMode: boolean;
  private config: IFormConfig;


  constructor(props: Readonly<ICountryProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      name: '',
      identifier: '',
      groupings: [],
      errors: {},
      isFormReady: false
    };
    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.Country);
  }

  private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
    const userRoles = this.props.userService.userContext?.userRoles ?? [];
    return rolesToCheck.some((role) => userRoles.includes(role));
  };
  public async componentDidMount() {
    if (this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin)) {
      if (this.isEditMode)
        await this.setInitialFormValues();

      this.setState({ isFormReady: true });

    }
  }

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Country).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.COUNTRY.IDENTIFIER,
      Consts.FIELDS.COUNTRY.GROUPING
    ];
    const item = await pnpService
      .getListByUrl(listUrl)
      .items.getById(itemId)
      .select(...selectedFields)();

    const taxonomyService = new TaxonomyService();
    await taxonomyService.init(pnpService);

    const groupings: ITaxonomyTerm[] = [];
    if (item[Consts.FIELDS.AUTHORITY.GROUPING]) {
      await Promise.all(item[Consts.FIELDS.AUTHORITY.GROUPING].map(async (g) => {
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
      identifier: item[Consts.FIELDS.COUNTRY.IDENTIFIER],
      groupings,
    };

    this.setState({
      ...initialValues
    });
  };

  private validateForm = (): boolean => {
    const errors: { [key: string]: string } = {};
    const {
      name,
      identifier,
      groupings
    } = this.state;
    if (name?.trim() === '' || !name)
      errors.name = strings.FormCountryNameValidationError;

    if (identifier?.trim() === '' || !identifier)
      errors.identifier = strings.FormCountryIdentifierValidationError;


    if (!Array.isArray(groupings) || groupings.length === 0) {
      errors.groupings = strings.FormCountryGroupingValidationError;
    }

    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

  public render(): React.ReactElement<ICountryProps> {
    const {
      name,
      identifier,
      groupings,
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
                          <Label required>{strings.FormCountryNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormCountryTooltipName} ></Tooltip>
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
                    {/* Identifier */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormCountryIdentifierLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormCountryTooltipIdentifier} ></Tooltip>
                      </div>
                      <TextField
                        value={identifier}
                        onChange={this.onChangeIdentifier}
                      />
                      {errors.identifier &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.identifier}
                          </span>
                        </div>
                      }
                    </div>
                    {/* Taxonomy Picker */}
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label required>{strings.FormCountryGroupingLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormCountryTooltipGrouping} />
                      </div>
                      <ModernTaxonomyPicker
                        allowMultipleSelections={true}
                        termSetId={Config.TAXONOMY.GROUPING_TERMSETID}
                        panelTitle={strings.FormCountryGroupingLabel}
                        label=''
                        onChange={this.onChangeGrouping}
                        context={this.props.pnpService.getContext()}
                        placeHolder={strings.FormCountryGroupingPlaceHolder}
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

  private onChangeIdentifier = (event, identifier: string) => {
    this.setState({ identifier });
  };

  private onChangeGrouping = (info?: ITermInfo[]): void => {
    this.setState({ groupings: info });
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
      identifier,
      groupings,
    } = this.state;
    const groupingsValues = groupings?.map((e) => e?.labels?.[0]?.name).join(", ");
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormCountryDatasheetLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {this.config.name}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormCountryNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {UtilHelper.sanitizeSharePointTitle(name)}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormCountryIdentifierLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {identifier}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormCountryGroupingLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {groupingsValues}
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
      identifier,
      groupings,
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
    itemData[Consts.FIELDS.COUNTRY.IDENTIFIER] = identifier;
    if (Array.isArray(groupings) && groupings.length > 0) {
      itemData[groupingHiddenField.InternalName] = groupings
        .map((g) => `-1;#${g.labels[0].name}|${g.id}`)
        .join(";#");
    }

    const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);
    return createdItem;
  };

  private processEdit = async () => {
    const {
      name,
      identifier,
      groupings,
    } = this.state;

    const {
      pnpService,
      itemId,
    } = this.props;

    const list = this.config.list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const groupingHiddenField = await pnpService.getListByUrl(listUrl).fields.getByTitle(`${Consts.FIELDS.COUNTRY.GROUPING}_0`)();

    const itemData = {};
    const sanitizedTitle = UtilHelper.sanitizeSharePointTitle(name);
    itemData[Consts.FIELDS.COMMON.TITLE] = sanitizedTitle;
    itemData[Consts.FIELDS.COUNTRY.IDENTIFIER] = identifier;
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
