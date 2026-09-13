import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import styles from '../Form.module.scss';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import FormWizard from '../Form Wizard/FormWizard';
import { Dropdown, Icon, IDropdownOption, Label, DatePicker } from '@fluentui/react';
import { Combobox } from 'react-widgets';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { Accordion, AccordionItem, AccordionHeader, AccordionPanel } from '../../../../common/components/Accordion/Accordion';
import { Form, FormOrigin, FormType, IFormConfig } from '../Form.types';
import strings from 'GenericFormWebPartStrings';
import { RichText } from '../../../../common/components/RichText/RichText';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { UtilHelper } from '../../../../common/helpers/Util';
import { Consts } from '../../../../common/consts/Consts';
import { IItemAddResult } from '@pnp/sp/items';

import 'react-widgets/styles.css';
import 'react-phone-number-input/style.css';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { MarketReadinessResponsibleParty, UserRole, Vertical } from '../../../../common/models/Enums';
import { datepickerstyles } from '../Form.utility';
import { IMarketReadinessProps, IMarketReadinessState } from './MarketReadiness.types';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';
export default class MarketReadiness extends React.Component<
    IMarketReadinessProps,
    IMarketReadinessState
> {
    private isEditMode: boolean;
    private config: IFormConfig;
    private AllLookupItems: Map<string, IDropdownOption[]> = new Map();
    constructor(props: Readonly<IMarketReadinessProps>) {
        super(props);
        this.isEditMode = this.props.itemId !== undefined;
        this.AllLookupItems = new Map();
        this.state = {
            name: '',
            market: null,
            marketChoices: [],
            responsiblePartyTypeChoices: [],
            responsiblePartyType: '',
            errors: {},
            isFormReady: false,
            responsiblePartyItem: null,
            responsiblePartyItemChoices: [],
            eutelsatOwnersChoices: [],
            eutelsatOwners: [],
            verticals: [],
            verticalsChoices: [],
            accordionOpenItems: [],
            verticalRAGStatus: '',
            verticalRAGStatusChoices: [],
            verticalEstimatedDate: null,
            verticalEffectiveDate: null,
            verticalComments: '',
            spaceRAGStatus: '',
            spaceRAGStatusChoices: [],
            spaceEstimatedDate: null,
            spaceEffectiveDate: null,
            spaceComments: '',
            landFixedRAGStatus: '',
            landFixedRAGStatusChoices: [],
            landFixedEstimatedDate: null,
            landFixedEffectiveDate: null,
            landFixedComments: '',
            landMobilityRAGStatus: '',
            landMobilityRAGStatusChoices: [],
            landMobilityEstimatedDate: null,
            landMobilityEffectiveDate: null,
            landMobilityComments: '',
            maritimeRAGStatus: '',
            maritimeRAGStatusChoices: [],
            maritimeEstimatedDate: null,
            maritimeEffectiveDate: null,
            maritimeComments: '',
            aviationRAGStatus: '',
            aviationRAGStatusChoices: [],
            aviationEstimatedDate: null,
            aviationEffectiveDate: null,
            aviationComments: ''
        };
        this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.MarketReadiness);
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
            this.initVerticalChoices()
        );
        promises.push(
            this.initEutelsatOwnersChoices()
        );
        promises.push(
            this.initMarketsChoices()
        );
        promises.push(
            this.initVerticalRAGStatusChoices()
        );
        promises.push(
            this.initSpaceRAGStatusChoices()
        );
        promises.push(
            this.initLandFixedRAGStatusChoices()
        );
        promises.push(
            this.initLandMobilityRAGStatusChoices()
        );
        promises.push(
            this.initMaritimeRAGStatusChoices()
        );
        promises.push(
            this.initAviationRAGStatusChoices()
        );
        await Promise.all(promises);

        if (this.isEditMode)
            await this.setInitialFormValues();
        this.setResponsiblePartyItems();
    };

    public initResponsiblePartyChoices = async () => {
        const {
            pnpService
        } = this.props;
        let responsiblePartyTypeChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MARKET_READINESS.RESPONSIBLE_PARTY)
            .select('Choices')();
        data.Choices.sort((a: string, b: string) => a?.localeCompare(b));
        responsiblePartyTypeChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
        const responsiblePartyType = responsiblePartyTypeChoices.length > 0 ? responsiblePartyTypeChoices[0].key as string : '';
        this.setState({
            responsiblePartyTypeChoices, responsiblePartyType
        });
    };

    public initVerticalChoices = async () => {
        const {
            pnpService
        } = this.props;
        let verticalsChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MARKET_READINESS.VERTICALS)
            .select('Choices')();
        data.Choices.sort((a: string, b: string) => a?.localeCompare(b));
        verticalsChoices = data.Choices.filter((choice: string) => choice !== Vertical.SNP).map((choice) => ({ key: choice, text: choice }));
        const verticals = [];
        this.setState({
            verticalsChoices, verticals
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

    private initMarketsChoices = async () => {
        if (!this.isEditMode) {
            const {
                pnpService,
                preselectedItemId
            } = this.props;
            const formConfigData = Form.config.find((f) => f.type.toLowerCase() === FormType.Market);

            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), formConfigData.list);
            const selectedFields: string[] = [
                Consts.FIELDS.COMMON.ID,
                Consts.FIELDS.COMMON.TITLE
            ];

            const marketChoices: IDropdownOption[] = [];
            const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)
                .filter(`Id eq ${preselectedItemId}`)
                .top(1)();
            items.forEach((item) => {
                marketChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
            });
            const market: IDropdownOption = marketChoices.length > 0 ? marketChoices[0] : null;
            this.setState({
                marketChoices, market
            });
        }
    };

    public initVerticalRAGStatusChoices = async () => {
        const {
            pnpService
        } = this.props;
        let verticalRAGStatusChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_RAGSTATUS)
            .select('Choices')();
        data.Choices.sort((a: string, b: string) => a?.localeCompare(b));
        verticalRAGStatusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
        // const verticalRAGStatus = verticalRAGStatusChoices.length > 0 ? verticalRAGStatusChoices[0].key as string : '';
        this.setState({
            verticalRAGStatusChoices, verticalRAGStatus: ''
        });
    };

    public initSpaceRAGStatusChoices = async () => {
        const {
            pnpService
        } = this.props;
        let spaceRAGStatusChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MARKET_READINESS.SPACE_NW_RAGSTATUS)
            .select('Choices')();
        data.Choices.sort((a: string, b: string) => a?.localeCompare(b));
        spaceRAGStatusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
        // const spaceRAGStatus = spaceRAGStatusChoices.length > 0 ? spaceRAGStatusChoices[0].key as string : '';
        this.setState({
            spaceRAGStatusChoices, spaceRAGStatus: ''
        });
    };

    public initLandFixedRAGStatusChoices = async () => {
        const {
            pnpService
        } = this.props;
        let landFixedRAGStatusChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MARKET_READINESS.LAND_FIXED_RAGSTATUS)
            .select('Choices')();
        data.Choices.sort((a: string, b: string) => a?.localeCompare(b));
        landFixedRAGStatusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
        // const landFixedRAGStatus = landFixedRAGStatusChoices.length > 0 ? landFixedRAGStatusChoices[0].key as string : '';
        this.setState({
            landFixedRAGStatusChoices, landFixedRAGStatus: ''
        });
    };

    public initLandMobilityRAGStatusChoices = async () => {
        const {
            pnpService
        } = this.props;
        let landMobilityRAGStatusChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MARKET_READINESS.LAND_FIXED_RAGSTATUS)
            .select('Choices')();
        data.Choices.sort((a: string, b: string) => a?.localeCompare(b));
        landMobilityRAGStatusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
        // const landMobilityRAGStatus = landMobilityRAGStatusChoices.length > 0 ? landMobilityRAGStatusChoices[0].key as string : '';
        this.setState({
            landMobilityRAGStatusChoices, landMobilityRAGStatus: ''
        });
    };

    public initMaritimeRAGStatusChoices = async () => {
        const {
            pnpService
        } = this.props;
        let maritimeRAGStatusChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MARKET_READINESS.MARITIME_RAGSTATUS)
            .select('Choices')();
        data.Choices.sort((a: string, b: string) => a?.localeCompare(b));
        maritimeRAGStatusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
        // const maritimeRAGStatus = maritimeRAGStatusChoices.length > 0 ? maritimeRAGStatusChoices[0].key as string : '';
        this.setState({
            maritimeRAGStatusChoices, maritimeRAGStatus: ''
        });
    };

    public initAviationRAGStatusChoices = async () => {
        const {
            pnpService
        } = this.props;
        let aviationRAGStatusChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MARKET_READINESS.AVIATION_RAGSTATUS)
            .select('Choices')();
        data.Choices.sort((a: string, b: string) => a?.localeCompare(b));
        aviationRAGStatusChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
        // const aviationRAGStatus = aviationRAGStatusChoices.length > 0 ? aviationRAGStatusChoices[0].key as string : '';
        this.setState({
            aviationRAGStatusChoices, aviationRAGStatus: ''
        });
    };

    private setResponsiblePartyItems = () => {
        const responsiblePartyText = this.state.responsiblePartyTypeChoices?.[0]?.text;
        const responsiblePartyItems = this.AllLookupItems.get(responsiblePartyText) || [];
        this.setState({ responsiblePartyItemChoices: responsiblePartyItems });
    };

    private validateForm = async (): Promise<boolean> => {
        const errors: { [key: string]: string } = {};
        const {
            responsiblePartyType,
            responsiblePartyItem,
            eutelsatOwners,
            verticals,
            market,
            verticalRAGStatus,
            /* verticalEstimatedDate,
             verticalEffectiveDate,
             verticalComments,*/
            spaceRAGStatus,
            /* spaceEstimatedDate,
             spaceEffectiveDate,
             spaceComments,*/
            landFixedRAGStatus,
            /*   landFixedEstimatedDate,
               landFixedEffectiveDate,
               landFixedComments,*/
            landMobilityRAGStatus,
            /*   landMobilityEstimatedDate,
               landMobilityEffectiveDate,
               landMobilityComments,*/
            maritimeRAGStatus,
            /*   maritimeEstimatedDate,
               maritimeEffectiveDate,
               maritimeComments, */
            aviationRAGStatus,
            /*  aviationEstimatedDate,
              aviationEffectiveDate,
              aviationComments  */
            accordionOpenItems
        } = this.state;
        const { pnpService, preselectedItemId } = this.props;

        // ✅ Base validations
        if (
            ![MarketReadinessResponsibleParty.Eutelsat, MarketReadinessResponsibleParty.Overall].includes(responsiblePartyType as MarketReadinessResponsibleParty) &&
            !responsiblePartyItem
        ) {
            errors.responsiblePartyItem = strings.FormMarketReadinessItemValidationError;
        }
        if (!Array.isArray(eutelsatOwners) || eutelsatOwners.length === 0) {
            errors.eutelsatOwners = strings.FormMarketReadinessEutelsatOwnersValidationError;
        }
        if (!Array.isArray(verticals) || verticals.length === 0) {
            errors.verticals = strings.FormMarketReadinessVerticalValidationError;
        }

        if (!errors.responsiblePartyType && preselectedItemId && this.isEditMode === false &&
            (responsiblePartyType === MarketReadinessResponsibleParty.Overall || responsiblePartyType === MarketReadinessResponsibleParty.Eutelsat)) {
            const list = this.config.list;
            //const exists = await pnpService.get(fileUrlDocument);
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);

            // Check if item exists
            const existingItems = await pnpService.getListItems(listUrl).
                filter(`${Consts.FIELDS.MARKET_READINESS.MARKET_ID} eq ${market?.key} and ${Consts.FIELDS.MARKET_READINESS.RESPONSIBLE_PARTY} eq '${responsiblePartyType}'`)();
            if (existingItems.length > 0) {
                errors.responsiblePartyType = responsiblePartyType + " " + strings.FormMarketReadinessOverAllandEtulsatValidationError;
            }

        }

        verticals.forEach((v) => {
            switch (v.key) {
                case Vertical.AllVerticals:
                    if (!verticalRAGStatus) {
                        errors.verticalRAGStatus = strings.FormCountryVerticalRAGStatusValidationError;
                        if (!accordionOpenItems.includes(v.key.toString())) {
                            accordionOpenItems.push(v.key.toString());
                        }
                    }
                    /*   if (!verticalEstimatedDate) errors.verticalEstimatedDate = strings.FormCountryVerticalEstimatedDateValidationError;
                       if (!verticalEffectiveDate) errors.verticalEffectiveDate = strings.FormCountryVerticalEffectiveDateValidationError;
                       if (!verticalComments) errors.verticalComments = strings.FormCountryVerticalCommentsValidationError;
                      */
                    break;

                case Vertical.SpaceNW:
                    if (!spaceRAGStatus) {
                        errors.spaceRAGStatus = strings.FormCountrySpaceRAGStatusValidationError;
                        if (!accordionOpenItems.includes(v.key.toString())) {
                            accordionOpenItems.push(v.key.toString());
                        }
                    }
                    break;

                case Vertical.LandFixed:
                    if (!landFixedRAGStatus) {
                        errors.landFixedRAGStatus = strings.FormCountryLandFixedRAGStatusValidationError;
                        if (!accordionOpenItems.includes(v.key.toString())) {
                            accordionOpenItems.push(v.key.toString());
                        }
                    }
                    break;

                case Vertical.LandMobility:
                    if (!landMobilityRAGStatus) {
                        errors.landMobilityRAGStatus = strings.FormCountryLandMobilityRAGStatusValidationError;
                        if (!accordionOpenItems.includes(v.key.toString())) {
                            accordionOpenItems.push(v.key.toString());
                        }
                    }
                    break;

                case Vertical.Maritime:
                    if (!maritimeRAGStatus) {
                        errors.maritimeRAGStatus = strings.FormCountryMaritimeRAGStatusValidationError;
                        if (!accordionOpenItems.includes(v.key.toString())) {
                            accordionOpenItems.push(v.key.toString());
                        }
                    }
                    break;

                case Vertical.Aviation:
                    if (!aviationRAGStatus) {
                        errors.aviationRAGStatus = strings.FormCountryAviationRAGStatusValidationError;
                        if (!accordionOpenItems.includes(v.key.toString())) {
                            accordionOpenItems.push(v.key.toString());
                        }
                    }
                    break;
            }
        });
        this.setState({ errors, accordionOpenItems });
        return Object.keys(errors).length === 0;
    };


    private getFormSummary = () => {
        const {
            market,
            responsiblePartyType,
            responsiblePartyItem,
            eutelsatOwners,
            verticals,
            verticalsChoices,
            verticalRAGStatus,
            verticalEstimatedDate,
            verticalEffectiveDate,
            verticalComments,
            spaceRAGStatus,
            spaceEstimatedDate,
            spaceEffectiveDate,
            spaceComments,
            landFixedRAGStatus,
            landFixedEstimatedDate,
            landFixedEffectiveDate,
            landFixedComments,
            landMobilityRAGStatus,
            landMobilityEstimatedDate,
            landMobilityEffectiveDate,
            landMobilityComments,
            maritimeRAGStatus,
            maritimeEstimatedDate,
            maritimeEffectiveDate,
            maritimeComments,
            aviationRAGStatus,
            aviationEstimatedDate,
            aviationEffectiveDate,
            aviationComments,
            accordionOpenItems,
        } = this.state;
        const maReadinessText = Consts.NAMING.MARKET_READINESS_SUFFIX;
        const verticalCommentsTextValue = DomHelper.cleanRichHtml(verticalComments);
        const spaceCommentsTextValue = DomHelper.cleanRichHtml(spaceComments);
        const landFixedCommentsTextValue = DomHelper.cleanRichHtml(landFixedComments);
        const landMobilityCommentsTextValue = DomHelper.cleanRichHtml(landMobilityComments);
        const maritimeCommentsTextValue = DomHelper.cleanRichHtml(maritimeComments);
        const aviationCommentsTextValue = DomHelper.cleanRichHtml(aviationComments);
        let displayText = "";
        if (responsiblePartyType === MarketReadinessResponsibleParty.Overall || responsiblePartyType === MarketReadinessResponsibleParty.Eutelsat) {
            displayText = `${market?.text} - ${responsiblePartyType} ${maReadinessText}`;
        } else if (responsiblePartyType === MarketReadinessResponsibleParty.TP) {
            displayText = `${market?.text} - ${responsiblePartyItem?.text} ${maReadinessText}`;
        } else if (responsiblePartyType === MarketReadinessResponsibleParty.DP) {
            displayText = `${market?.text} - ${responsiblePartyItem?.text} ${maReadinessText}`;
        }
        const html = verticalsChoices.map((item) =>
        (verticals.find((v) => v.key === item.key) &&
            <AccordionItem value={item.key.toString()} key={item.key} >
                <AccordionHeader>{item.text}</AccordionHeader>
                <AccordionPanel>
                    {
                        item.text === Vertical.AllVerticals && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryVerticalRAGStatusLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{verticalRAGStatus}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryVerticalEstimatedDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{verticalEstimatedDate ? this.formatDate(verticalEstimatedDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryVerticalEffectiveDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{verticalEffectiveDate ? this.formatDate(verticalEffectiveDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryVerticalCommentsLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{verticalCommentsTextValue.length > 0 ? <RichText isEditMode={false} value={verticalCommentsTextValue} /> : ''}</div>
                                </div>
                            </>
                        )
                    }
                    {
                        item.text === Vertical.SpaceNW && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountrySpaceRAGStatusLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{spaceRAGStatus}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountrySpaceEstimatedDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{spaceEstimatedDate ? this.formatDate(spaceEstimatedDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountrySpaceEffectiveDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{spaceEffectiveDate ? this.formatDate(spaceEffectiveDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountrySpaceCommentsLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{spaceCommentsTextValue.length > 0 ? <RichText isEditMode={false} value={spaceCommentsTextValue} /> : ''}</div>
                                </div>
                            </>
                        )
                    }
                    {
                        item.text === Vertical.LandFixed && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandFixedRAGStatusLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{landFixedRAGStatus}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandFixedEstimatedDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{landFixedEstimatedDate ? this.formatDate(landFixedEstimatedDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandFixedEffectiveDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{landFixedEffectiveDate ? this.formatDate(landFixedEffectiveDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandFixedCommentsLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{landFixedCommentsTextValue.length > 0 ? <RichText isEditMode={false} value={landFixedCommentsTextValue} /> : ''}</div>
                                </div>
                            </>
                        )
                    }
                    {
                        item.text === Vertical.LandMobility && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandMobilityRAGStatusLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{landMobilityRAGStatus}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandMobilityEstimatedDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{landMobilityEstimatedDate ? this.formatDate(landMobilityEstimatedDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandMobilityEffectiveDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{landMobilityEffectiveDate ? this.formatDate(landMobilityEffectiveDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandMobilityCommentsLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{landMobilityCommentsTextValue.length > 0 ? <RichText isEditMode={false} value={landMobilityCommentsTextValue} /> : ''}</div>
                                </div>
                            </>
                        )
                    }
                    {
                        item.text === Vertical.Maritime && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryMaritimeRAGStatusLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{maritimeRAGStatus}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryMaritimeEstimatedDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{maritimeEstimatedDate ? this.formatDate(maritimeEstimatedDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryMaritimeEffectiveDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{maritimeEffectiveDate ? this.formatDate(maritimeEffectiveDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryMaritimeCommentsLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{maritimeCommentsTextValue.length > 0 ? <RichText isEditMode={false} value={maritimeCommentsTextValue} /> : ''}</div>
                                </div>
                            </>
                        )
                    }
                    {
                        item.text === Vertical.Aviation && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryAviationRAGStatusLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{aviationRAGStatus}</div>

                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryAviationEstimatedDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{aviationEstimatedDate ? this.formatDate(aviationEstimatedDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryAviationEffectiveDateLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}> {aviationEffectiveDate ? this.formatDate(aviationEffectiveDate) : ''}</div>
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryAviationCommentsLabel}</Label>
                                        </div>
                                    </div>
                                    <div className={styles.fieldValue}>{aviationCommentsTextValue.length > 0 ? <RichText isEditMode={false} value={aviationCommentsTextValue} /> : ''}</div>
                                </div>
                            </>
                        )
                    }
                </AccordionPanel>
            </AccordionItem>
        )
        );

        return (
            <div className={styles.summary}>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormCountryDatasheetNameLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>{this.config.name}</div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMarketReadinessNameLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>{displayText}</div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormCountryLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {market?.text}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormCountryResponsiblePartyLabel}</Label>
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
                                <Label>{strings.FormCountryResponsibleItemLabel}</Label>
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
                            <Label>{strings.FormCountryEutelsatOwnersLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {eutelsatOwners.map((c) => c.text).join(", ")}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormCountryAllVerticalsLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {verticals.map((c) => c.text).join(", ")}
                    </div>
                </div>
                <div>
                    <Accordion
                        multiple
                        collapsible
                        openItems={accordionOpenItems}
                        onOpenChange={(_, d) => {
                            this.setState({ accordionOpenItems: d.openItems });
                        }}
                    >
                        {html}
                    </Accordion>
                </div>
            </div>
        );
    };

    public render(): React.ReactElement<IMarketReadinessProps> {
        const {
            isFormReady,
            market,
            marketChoices,
            responsiblePartyTypeChoices,
            responsiblePartyType,
            responsiblePartyItemChoices,
            responsiblePartyItem,
            errors,
            eutelsatOwnersChoices,
            eutelsatOwners,
            verticals,
            verticalsChoices,
            verticalRAGStatus,
            verticalRAGStatusChoices,
            verticalEffectiveDate,
            verticalEstimatedDate,
            verticalComments,
            spaceRAGStatus,
            spaceRAGStatusChoices,
            spaceEstimatedDate,
            spaceEffectiveDate,
            spaceComments,
            landFixedRAGStatus,
            landFixedRAGStatusChoices,
            landFixedEstimatedDate,
            landFixedEffectiveDate,
            landFixedComments,
            landMobilityRAGStatus,
            landMobilityRAGStatusChoices,
            landMobilityEstimatedDate,
            landMobilityEffectiveDate,
            landMobilityComments,
            maritimeRAGStatus,
            maritimeRAGStatusChoices,
            maritimeEstimatedDate,
            maritimeEffectiveDate,
            maritimeComments,
            aviationRAGStatus,
            aviationRAGStatusChoices,
            aviationEstimatedDate,
            aviationEffectiveDate,
            aviationComments,
            accordionOpenItems

        } = this.state;

        const html = verticalsChoices.map((item) =>
        (verticals.find((v) => v.key === item.key) &&
            <AccordionItem value={item.key.toString()} key={item.key}  >
                <AccordionHeader>{item.text}</AccordionHeader>
                <AccordionPanel>
                    {
                        item.text === Vertical.AllVerticals && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryVerticalRAGStatusLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipVerticalRAGStatus} ></Tooltip>
                                    </div>
                                    <Dropdown
                                        selectedKey={verticalRAGStatus}
                                        onChange={this.onChangeVerticalRAGStatus}
                                        options={verticalRAGStatusChoices} />
                                    {errors.verticalRAGStatus &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.verticalRAGStatus}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryVerticalEstimatedDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipVerticalEstimatedDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeVerticalEstimatedDate}
                                        value={verticalEstimatedDate}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                        styles={datepickerstyles}
                                    />
                                    {errors.verticalEstimatedDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.verticalEstimatedDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryVerticalEffectiveDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipVerticalEffectiveDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeVerticalEffectiveDate}
                                        value={verticalEffectiveDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.verticalEffectiveDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.verticalEffectiveDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryVerticalCommentsLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryVerticalCommentsLabel} ></Tooltip>
                                    </div>
                                    <RichText isEditMode={true} value={verticalComments} onChange={this.onChangeVerticalComments} />
                                    {errors.verticalComments &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.verticalComments}
                                            </span>
                                        </div>
                                    }
                                </div>
                            </>
                        )
                    }
                    {
                        item.text === Vertical.SpaceNW && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountrySpaceRAGStatusLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipSpaceRAGStatus} ></Tooltip>
                                    </div>
                                    <Dropdown
                                        selectedKey={spaceRAGStatus}
                                        onChange={this.onChangeSpaceRAGStatus}
                                        options={spaceRAGStatusChoices} />
                                    {errors.spaceRAGStatus &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.spaceRAGStatus}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountrySpaceEstimatedDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipSpaceEstimatedDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeSpaceEstimatedDate}
                                        value={spaceEstimatedDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.spaceEstimatedDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.spaceEstimatedDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountrySpaceEffectiveDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipSpaceEffectiveDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeSpaceEffectiveDate}
                                        value={spaceEffectiveDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.spaceEffectiveDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.spaceEffectiveDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountrySpaceCommentsLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountrySpaceCommentsLabel} ></Tooltip>
                                    </div>
                                    <RichText isEditMode={true} value={spaceComments} onChange={this.onChangeSpaceComments} />
                                    {errors.spaceComments &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.spaceComments}
                                            </span>
                                        </div>
                                    }
                                </div>
                            </>
                        )
                    }
                    {
                        item.text === Vertical.LandFixed && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandFixedRAGStatusLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipLandFixedRAGStatus} ></Tooltip>
                                    </div>
                                    <Dropdown
                                        selectedKey={landFixedRAGStatus}
                                        onChange={this.onChangeLandFixedRAGStatus}
                                        options={landFixedRAGStatusChoices} />
                                    {errors.landFixedRAGStatus &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.landFixedRAGStatus}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandFixedEstimatedDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipLandFixedEstimatedDate} ></Tooltip>

                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeLandFixedEstimatedDate}
                                        value={landFixedEstimatedDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.landFixedEstimatedDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.landFixedEstimatedDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandFixedEffectiveDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipLandFixedEffectiveDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeLandFixedEffectiveDate}
                                        value={landFixedEffectiveDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.landFixedEffectiveDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.landFixedEffectiveDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandFixedCommentsLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipLandFixedComments} ></Tooltip>
                                    </div>
                                    <RichText isEditMode={true} value={landFixedComments} onChange={this.onChangeLandFixedComments} />
                                    {errors.landFixedComments &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.landFixedComments}
                                            </span>
                                        </div>
                                    }
                                </div>
                            </>
                        )
                    }
                    {
                        item.text === Vertical.LandMobility && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandMobilityRAGStatusLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipLandMobilityRAGStatus} ></Tooltip>
                                    </div>
                                    <Dropdown
                                        selectedKey={landMobilityRAGStatus}
                                        onChange={this.onChangeLandMobilityRAGStatus}
                                        options={landMobilityRAGStatusChoices} />
                                    {errors.landMobilityRAGStatus &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.landMobilityRAGStatus}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandMobilityEstimatedDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipLandMobilityEstimatedDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeLandMobilityEstimatedDate}
                                        value={landMobilityEstimatedDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.landMobilityEstimatedDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.landMobilityEstimatedDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandMobilityEffectiveDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipLandMobilityEffectiveDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeLandMobilityEffectiveDate}
                                        value={landMobilityEffectiveDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.landMobilityEffectiveDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.landMobilityEffectiveDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryLandMobilityCommentsLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipLandMobilityComments} ></Tooltip>
                                    </div>
                                    <RichText isEditMode={true} value={landMobilityComments} onChange={this.onChangeLandMobilityComments} />
                                    {errors.landMobilityComments &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.landMobilityComments}
                                            </span>
                                        </div>
                                    }
                                </div>
                            </>
                        )
                    }
                    {
                        item.text === Vertical.Maritime && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryMaritimeRAGStatusLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipMaritimeRAGStatus} ></Tooltip>
                                    </div>
                                    <Dropdown
                                        selectedKey={maritimeRAGStatus}
                                        onChange={this.onChangeMaritimeRAGStatus}
                                        options={maritimeRAGStatusChoices} />
                                    {errors.maritimeRAGStatus &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.maritimeRAGStatus}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryMaritimeEstimatedDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipMaritimeEstimatedDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeMaritimeEstimatedDate}
                                        value={maritimeEstimatedDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.maritimeEstimatedDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.maritimeEstimatedDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryMaritimeEffectiveDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipMaritimeEffectiveDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeMaritimeEffectiveDate}
                                        value={maritimeEffectiveDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.maritimeEffectiveDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.maritimeEffectiveDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryMaritimeCommentsLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipMaritimeComments} ></Tooltip>
                                    </div>
                                    <RichText isEditMode={true} value={maritimeComments} onChange={this.onChangeMaritimeComments} />
                                    {errors.maritimeComments &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.maritimeComments}
                                            </span>
                                        </div>
                                    }
                                </div>
                            </>
                        )
                    }
                    {
                        item.text === Vertical.Aviation && (
                            <>
                                <div className={styles.field} >
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryAviationRAGStatusLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipAviationRAGStatus} ></Tooltip>
                                    </div>
                                    <Dropdown
                                        selectedKey={aviationRAGStatus}
                                        onChange={this.onChangeAviationRAGStatus}
                                        options={aviationRAGStatusChoices} />
                                    {errors.aviationRAGStatus &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.aviationRAGStatus}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryAviationEstimatedDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipAviationEstimatedDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeAviationEstimatedDate}
                                        value={aviationEstimatedDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.aviationEstimatedDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.aviationEstimatedDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryAviationEffectiveDateLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipAviationEffectiveDate} ></Tooltip>
                                    </div>
                                    <DatePicker
                                        onSelectDate={this.onChangeAviationEffectiveDate}
                                        value={aviationEffectiveDate}
                                        styles={datepickerstyles}
                                        formatDate={this.formatDate}
                                        allowTextInput
                                    />
                                    {errors.aviationEffectiveDate &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.aviationEffectiveDate}
                                            </span>
                                        </div>
                                    }
                                </div>
                                <div className={styles.field}>
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormCountryAviationCommentsLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormCountryTooltipAviationComments} ></Tooltip>
                                    </div>
                                    <RichText isEditMode={true} value={aviationComments} onChange={this.onChangeAviationComments} />
                                    {errors.aviationComments &&
                                        <div className={styles.errorContainer}>
                                            <Icon iconName="Error" className={styles.errorIcon} />
                                            <span className={styles.errorMessage}>
                                                {errors.aviationComments}
                                            </span>
                                        </div>
                                    }
                                </div>
                            </>
                        )
                    }

                </AccordionPanel>
            </AccordionItem>
        )
        );
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
                        getSummary={() => this.getFormSummary()}
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
                                    {/* Market */}
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMarketLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMarketTooltip} ></Tooltip>
                                        </div>
                                        <Combobox
                                            disabled={true}
                                            data={marketChoices}
                                            value={market}
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
                                                <Label>{strings.FormCountryResponsiblePartyLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormCountryTooltipResponsibleParty} ></Tooltip>
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
                                                onChange={this.onChangeResponsiblePartyItem}
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
                                            <Tooltip content={strings.FormCountryTooltipEutelsatOwners} ></Tooltip>
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
                                    {/* Verticals */}
                                    <div className={styles.field} >
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label required>{strings.FormCountryAllVerticalsLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormCountryTooltipAllVerticals} ></Tooltip>
                                        </div>
                                        <MultiselectWrapper
                                            value={verticals}
                                            data={verticalsChoices}
                                            dataKey={(item: IDropdownOption) => item.key}
                                            textField={(item: IDropdownOption) => item.text}
                                            filter="contains"
                                            onChange={this.onChangeVerticals} />
                                        {errors.verticals &&
                                            <div className={styles.errorContainer}>
                                                <Icon iconName="Error" className={styles.errorIcon} />
                                                <span className={styles.errorMessage}>
                                                    {errors.verticals}
                                                </span>
                                            </div>
                                        }

                                    </div>
                                    <Accordion
                                        multiple
                                        collapsible
                                        openItems={accordionOpenItems}
                                        onOpenChange={(_, d) => {
                                            this.setState({ accordionOpenItems: d.openItems });
                                        }}
                                    >

                                        {html}
                                    </Accordion>
                                </form>
                            }
                        </>
                    </FormWizard>}
                </div>
            </div >
        );
    }
    private onChangeVerticalEstimatedDate = (verticalEstimatedDate: Date) => {
        this.setState({ verticalEstimatedDate });
        return verticalEstimatedDate;
    };
    private onChangeVerticalEffectiveDate = (verticalEffectiveDate: Date) => {
        this.setState({ verticalEffectiveDate });
        return verticalEffectiveDate;
    };
    private onChangeVerticalComments = (verticalComments: string) => {
        this.setState({ verticalComments });
        return verticalComments;
    };
    private onChangeVerticalRAGStatus = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
        this.setState({ verticalRAGStatus: option.key.toString() });
    };
    private onChangeSpaceEstimatedDate = (spaceEstimatedDate: Date) => {
        this.setState({ spaceEstimatedDate });
        return spaceEstimatedDate;
    };
    private onChangeSpaceEffectiveDate = (spaceEffectiveDate: Date) => {
        this.setState({ spaceEffectiveDate });
        return spaceEffectiveDate;
    };
    private onChangeSpaceComments = (spaceComments: string) => {
        this.setState({ spaceComments });
        return spaceComments;
    };
    private onChangeSpaceRAGStatus = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
        this.setState({ spaceRAGStatus: option.key.toString() });
    };
    private onChangeLandFixedRAGStatus = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
        this.setState({ landFixedRAGStatus: option.key.toString() });
    };
    private onChangeLandFixedEstimatedDate = (landFixedEstimatedDate: Date) => {
        this.setState({ landFixedEstimatedDate });
        return landFixedEstimatedDate;
    };
    private onChangeLandFixedEffectiveDate = (landFixedEffectiveDate: Date) => {
        this.setState({ landFixedEffectiveDate });
        return landFixedEffectiveDate;
    };
    private onChangeLandFixedComments = (landFixedComments: string) => {
        this.setState({ landFixedComments });
        return landFixedComments;
    };
    private onChangeLandMobilityRAGStatus = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
        this.setState({ landMobilityRAGStatus: option.key.toString() });
    };
    private onChangeLandMobilityEstimatedDate = (landMobilityEstimatedDate: Date) => {
        this.setState({ landMobilityEstimatedDate });
        return landMobilityEstimatedDate;
    };
    private onChangeLandMobilityEffectiveDate = (landMobilityEffectiveDate: Date) => {
        this.setState({ landMobilityEffectiveDate });
        return landMobilityEffectiveDate;
    };
    private onChangeLandMobilityComments = (landMobilityComments: string) => {
        this.setState({ landMobilityComments });
        return landMobilityComments;
    };
    private onChangeMaritimeRAGStatus = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
        this.setState({ maritimeRAGStatus: option.key.toString() });
    };
    private onChangeMaritimeEstimatedDate = (maritimeEstimatedDate: Date) => {
        this.setState({ maritimeEstimatedDate });
        return maritimeEstimatedDate;
    };
    private onChangeMaritimeEffectiveDate = (maritimeEffectiveDate: Date) => {
        this.setState({ maritimeEffectiveDate });
        return maritimeEffectiveDate;
    };
    private onChangeMaritimeComments = (maritimeComments: string) => {
        this.setState({ maritimeComments });
        return maritimeComments;
    };
    private onChangeAviationRAGStatus = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
        this.setState({ aviationRAGStatus: option.key.toString() });
    };
    private onChangeAviationEstimatedDate = (aviationEstimatedDate: Date) => {
        this.setState({ aviationEstimatedDate });
        return aviationEstimatedDate;
    };
    private onChangeAviationEffectiveDate = (aviationEffectiveDate: Date) => {
        this.setState({ aviationEffectiveDate });
        return aviationEffectiveDate;
    };
    private onChangeAviationComments = (aviationComments: string) => {
        this.setState({ aviationComments });
        return aviationComments;
    };

    private onChangeEutelsatOwners = (eutelsatOwners: IDropdownOption[]) => {
        this.setState({ eutelsatOwners });
    };

    private onChangeResponsibleParty = async (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
        this.setState({ responsiblePartyType: option.key.toString() });
        if (option.key.toString() === MarketReadinessResponsibleParty.TP || option.key.toString() === MarketReadinessResponsibleParty.DP) {
            const typeChoices = this.getItemsByContentTypeId(option.text.toString());
            this.setState({ responsiblePartyItemChoices: typeChoices, responsiblePartyItem: null });
        }
    };

    private onChangeResponsiblePartyItem = (responsiblePartyItem: IDropdownOption) => {
        this.setState({ responsiblePartyItem });
    };

    private onChangeVerticals = (verticals: IDropdownOption[]) => {
        const selectedKeys = verticals.map((v) => v.key);
        const newState: Partial<IMarketReadinessState> = { verticals };

        if (!selectedKeys.includes(Vertical.AllVerticals)) {
            newState.verticalRAGStatus = '';
            newState.verticalEstimatedDate = null;
            newState.verticalEffectiveDate = null;
            newState.verticalComments = '';
        }
        if (!selectedKeys.includes(Vertical.SpaceNW)) {
            newState.spaceRAGStatus = '';
            newState.spaceEstimatedDate = null;
            newState.spaceEffectiveDate = null;
            newState.spaceComments = '';
        }
        if (!selectedKeys.includes(Vertical.LandFixed)) {
            newState.landFixedRAGStatus = '';
            newState.landFixedEstimatedDate = null;
            newState.landFixedEffectiveDate = null;
            newState.landFixedComments = '';
        }
        if (!selectedKeys.includes(Vertical.LandMobility)) {
            newState.landMobilityRAGStatus = '';
            newState.landMobilityEstimatedDate = null;
            newState.landMobilityEffectiveDate = null;
            newState.landMobilityComments = '';
        }
        if (!selectedKeys.includes(Vertical.Maritime)) {
            newState.maritimeRAGStatus = '';
            newState.maritimeEstimatedDate = null;
            newState.maritimeEffectiveDate = null;
            newState.maritimeComments = '';
        }
        if (!selectedKeys.includes(Vertical.Aviation)) {
            newState.aviationRAGStatus = '';
            newState.aviationEstimatedDate = null;
            newState.aviationEffectiveDate = null;
            newState.aviationComments = '';
        }
        this.setState(newState as IMarketReadinessState);
    };

    private processCreation = async (): Promise<IItemAddResult> => {
        const {
            market,
            responsiblePartyType,
            responsiblePartyItem,
            eutelsatOwners,
            verticals,
            verticalRAGStatus,
            verticalEstimatedDate,
            verticalEffectiveDate,
            verticalComments,
            spaceRAGStatus,
            spaceEstimatedDate,
            spaceEffectiveDate,
            spaceComments,
            landFixedRAGStatus,
            landFixedEstimatedDate,
            landFixedEffectiveDate,
            landFixedComments,
            landMobilityRAGStatus,
            landMobilityEstimatedDate,
            landMobilityEffectiveDate,
            landMobilityComments,
            maritimeRAGStatus,
            maritimeEstimatedDate,
            maritimeEffectiveDate,
            maritimeComments,
            aviationComments,
            aviationRAGStatus,
            aviationEstimatedDate,
            aviationEffectiveDate,
        } = this.state;

        const {
            pnpService,
        } = this.props;
        const itemData = {};
        const maReadinessText = Consts.NAMING.MARKET_READINESS_SUFFIX;
        let displayText = "";
        if (responsiblePartyType === MarketReadinessResponsibleParty.Overall || responsiblePartyType === MarketReadinessResponsibleParty.Eutelsat) {
            displayText = `${market?.text} - ${responsiblePartyType} ${maReadinessText}`;
        } else if (responsiblePartyType === MarketReadinessResponsibleParty.DP) {
            displayText = `${market?.text} - ${responsiblePartyItem?.text} ${maReadinessText}`;
        } else if (responsiblePartyType === MarketReadinessResponsibleParty.TP) {
            displayText = `${market?.text} - ${responsiblePartyItem?.text} ${maReadinessText}`;
        }
        itemData[Consts.FIELDS.COMMON.TITLE] = displayText;
        itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = Consts.CONTENT_TYPES.MA_READINESS_COUNTRY;
        itemData[Consts.FIELDS.MARKET_READINESS.MARKET_ID] = market?.key;
        itemData[Consts.FIELDS.MARKET_READINESS.RESPONSIBLE_PARTY] = responsiblePartyType;
        if (responsiblePartyType === MarketReadinessResponsibleParty.DP)
            itemData[Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER_ID] = responsiblePartyItem?.key;
        if (responsiblePartyType === MarketReadinessResponsibleParty.TP)
            itemData[Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER_ID] = responsiblePartyItem?.key;
        itemData[Consts.FIELDS.MARKET_READINESS.OWNERS_ID] = eutelsatOwners.map((c) => c.key);
        itemData[Consts.FIELDS.MARKET_READINESS.VERTICALS] = verticals.map((c) => c.key);

        // All Verticals
        itemData[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_RAGSTATUS] = verticalRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_ESTIMATED_DATE] = verticalEstimatedDate ? verticalEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_EFFECTIVE_DATE] = verticalEffectiveDate ? verticalEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_COMMENTS] = verticalComments;
        // Space & NW
        itemData[Consts.FIELDS.MARKET_READINESS.SPACE_NW_RAGSTATUS] = spaceRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.SPACE_NW_ESTIMATED_DATE] = spaceEstimatedDate ? spaceEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.SPACE_NW_EFFECTIVE_DATE] = spaceEffectiveDate ? spaceEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.SPACE_NW_COMMENTS] = spaceComments;
        // Land Fixed
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_RAGSTATUS] = landFixedRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_ESTIMATED_DATE] = landFixedEstimatedDate ? landFixedEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_EFFECTIVE_DATE] = landFixedEffectiveDate ? landFixedEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_COMMENTS] = landFixedComments;
        // Land Mobility
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_RAGSTATUS] = landMobilityRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_ESTIMATED_DATE] = landMobilityEstimatedDate ? landMobilityEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_EFFECTIVE_DATE] = landMobilityEffectiveDate ? landMobilityEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_COMMENTS] = landMobilityComments;
        // Maritime
        itemData[Consts.FIELDS.MARKET_READINESS.MARITIME_RAGSTATUS] = maritimeRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.MARITIME_ESTIMATED_DATE] = maritimeEstimatedDate ? maritimeEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.MARITIME_EFFECTIVE_DATE] = maritimeEffectiveDate ? maritimeEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.MARITIME_COMMENTS] = maritimeComments;
        // Aviation
        itemData[Consts.FIELDS.MARKET_READINESS.AVIATION_RAGSTATUS] = aviationRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.AVIATION_ESTIMATED_DATE] = aviationEstimatedDate ? aviationEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.AVIATION_EFFECTIVE_DATE] = aviationEffectiveDate ? aviationEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.AVIATION_COMMENTS] = aviationComments;


        const list = this.config.list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);
        return createdItem;
    };
    private processEdit = async () => {
        const {
            market,
            responsiblePartyType,
            responsiblePartyItem,
            eutelsatOwners,
            verticals,
            verticalRAGStatus,
            verticalEstimatedDate,
            verticalEffectiveDate,
            verticalComments,
            spaceRAGStatus,
            spaceEstimatedDate,
            spaceEffectiveDate,
            spaceComments,
            landFixedRAGStatus,
            landFixedEstimatedDate,
            landFixedEffectiveDate,
            landFixedComments,
            landMobilityRAGStatus,
            landMobilityEstimatedDate,
            landMobilityEffectiveDate,
            landMobilityComments,
            maritimeRAGStatus,
            maritimeEstimatedDate,
            maritimeEffectiveDate,
            maritimeComments,
            aviationComments,
            aviationRAGStatus,
            aviationEstimatedDate,
            aviationEffectiveDate,
        } = this.state;
        const {
            pnpService,
            itemId,
        } = this.props;
        const maReadinessText = Consts.NAMING.MARKET_READINESS_SUFFIX;
        let displayText = "";
        if (responsiblePartyType === MarketReadinessResponsibleParty.Overall || responsiblePartyType === MarketReadinessResponsibleParty.Eutelsat) {
            displayText = `${market?.text} - ${responsiblePartyType} ${maReadinessText}`;
        } else if (responsiblePartyType === MarketReadinessResponsibleParty.DP) {
            displayText = `${market?.text} - ${responsiblePartyItem?.text} ${maReadinessText}`;
        } else if (responsiblePartyType === MarketReadinessResponsibleParty.TP) {
            displayText = `${market?.text} - ${responsiblePartyItem?.text} ${maReadinessText}`;
        }
        const itemData = {};
        itemData[Consts.FIELDS.COMMON.TITLE] = displayText;
        itemData[Consts.FIELDS.MARKET_READINESS.MARKET_ID] = market?.key;
        itemData[Consts.FIELDS.MARKET_READINESS.RESPONSIBLE_PARTY] = responsiblePartyType;
        itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = Consts.CONTENT_TYPES.MA_READINESS_COUNTRY;
        if (responsiblePartyType === MarketReadinessResponsibleParty.TP) {
            itemData[Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER_ID] = responsiblePartyItem?.key;
        } else if (responsiblePartyType === MarketReadinessResponsibleParty.DP) {
            itemData[Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER_ID] = responsiblePartyItem?.key;
        }
        itemData[Consts.FIELDS.MARKET_READINESS.OWNERS_ID] =
            eutelsatOwners.map((c) => (parseInt(c.key.toString())));
        itemData[Consts.FIELDS.MARKET_READINESS.VERTICALS] = verticals.map((c) => c.key);
        // All Verticals
        itemData[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_RAGSTATUS] = verticalRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_ESTIMATED_DATE] = verticalEstimatedDate ? verticalEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_EFFECTIVE_DATE] = verticalEffectiveDate ? verticalEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_COMMENTS] = verticalComments;
        // Space & NW
        itemData[Consts.FIELDS.MARKET_READINESS.SPACE_NW_RAGSTATUS] = spaceRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.SPACE_NW_ESTIMATED_DATE] = spaceEstimatedDate ? spaceEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.SPACE_NW_EFFECTIVE_DATE] = spaceEffectiveDate ? spaceEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.SPACE_NW_COMMENTS] = spaceComments;
        // Land Fixed
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_RAGSTATUS] = landFixedRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_ESTIMATED_DATE] = landFixedEstimatedDate ? landFixedEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_EFFECTIVE_DATE] = landFixedEffectiveDate ? landFixedEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_COMMENTS] = landFixedComments;
        // Land Mobility
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_RAGSTATUS] = landMobilityRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_ESTIMATED_DATE] = landMobilityEstimatedDate ? landMobilityEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_EFFECTIVE_DATE] = landMobilityEffectiveDate ? landMobilityEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_COMMENTS] = landMobilityComments;
        // Maritime
        itemData[Consts.FIELDS.MARKET_READINESS.MARITIME_RAGSTATUS] = maritimeRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.MARITIME_ESTIMATED_DATE] = maritimeEstimatedDate ? maritimeEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.MARITIME_EFFECTIVE_DATE] = maritimeEffectiveDate ? maritimeEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.MARITIME_COMMENTS] = maritimeComments;
        // Aviation
        itemData[Consts.FIELDS.MARKET_READINESS.AVIATION_RAGSTATUS] = aviationRAGStatus;
        itemData[Consts.FIELDS.MARKET_READINESS.AVIATION_ESTIMATED_DATE] = aviationEstimatedDate ? aviationEstimatedDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.AVIATION_EFFECTIVE_DATE] = aviationEffectiveDate ? aviationEffectiveDate?.toISOString() : null;
        itemData[Consts.FIELDS.MARKET_READINESS.AVIATION_COMMENTS] = aviationComments;
        const list = this.config.list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        await pnpService.getListItems(listUrl).getById(itemId).update(itemData);

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
    private getItemsByContentTypeId = (contentTypeText: string) => {
        // Check if AllLookupItems already has data
        if (this.AllLookupItems.has(contentTypeText)) {
            const cachedItems = this.AllLookupItems.get(contentTypeText) || [];
            return cachedItems;
        }
    };

    public setInitialFormValues = async () => {
        const { pnpService, itemId } = this.props;
        const list = Form.config.find((f) => f.type.toLowerCase() === FormType.MarketReadiness).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.MARKET_READINESS.RESPONSIBLE_PARTY,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
            Consts.FIELDS.MARKET_READINESS.VERTICALS,
            Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_RAGSTATUS,
            Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_ESTIMATED_DATE,
            Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_EFFECTIVE_DATE,
            Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_COMMENTS,
            Consts.FIELDS.MARKET_READINESS.SPACE_NW_RAGSTATUS,
            Consts.FIELDS.MARKET_READINESS.SPACE_NW_ESTIMATED_DATE,
            Consts.FIELDS.MARKET_READINESS.SPACE_NW_EFFECTIVE_DATE,
            Consts.FIELDS.MARKET_READINESS.SPACE_NW_COMMENTS,
            Consts.FIELDS.MARKET_READINESS.LAND_FIXED_RAGSTATUS,
            Consts.FIELDS.MARKET_READINESS.LAND_FIXED_ESTIMATED_DATE,
            Consts.FIELDS.MARKET_READINESS.LAND_FIXED_EFFECTIVE_DATE,
            Consts.FIELDS.MARKET_READINESS.LAND_FIXED_COMMENTS,
            Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_RAGSTATUS,
            Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_ESTIMATED_DATE,
            Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_EFFECTIVE_DATE,
            Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_COMMENTS,
            Consts.FIELDS.MARKET_READINESS.MARITIME_RAGSTATUS,
            Consts.FIELDS.MARKET_READINESS.MARITIME_ESTIMATED_DATE,
            Consts.FIELDS.MARKET_READINESS.MARITIME_EFFECTIVE_DATE,
            Consts.FIELDS.MARKET_READINESS.MARITIME_COMMENTS,
            Consts.FIELDS.MARKET_READINESS.AVIATION_RAGSTATUS,
            Consts.FIELDS.MARKET_READINESS.AVIATION_ESTIMATED_DATE,
            Consts.FIELDS.MARKET_READINESS.AVIATION_EFFECTIVE_DATE,
            Consts.FIELDS.MARKET_READINESS.AVIATION_COMMENTS,
            `${Consts.FIELDS.MARKET_READINESS.MARKET}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MARKET_READINESS.MARKET}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MARKET_READINESS.OWNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MARKET_READINESS.OWNER}/${Consts.FIELDS.COMMON.TITLE}`
        ];
        const expand = [Consts.FIELDS.MARKET_READINESS.MARKET,
        Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER,
        Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER,
        Consts.FIELDS.MARKET_READINESS.OWNER];
        const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
        let market: IDropdownOption = null;
        market = {
            key: item[Consts.FIELDS.MARKET_READINESS.MARKET][Consts.FIELDS.COMMON.ID],
            text: item[Consts.FIELDS.MARKET_READINESS.MARKET][Consts.FIELDS.COMMON.TITLE],
        };

        let responsiblePartyItem: IDropdownOption | null = null;
        const responsiblePartyType: string = item[Consts.FIELDS.MARKET_READINESS.RESPONSIBLE_PARTY];
        if (responsiblePartyType === MarketReadinessResponsibleParty.TP) {
            responsiblePartyItem = {
                key: item[Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER][Consts.FIELDS.COMMON.ID],
                text: item[Consts.FIELDS.MARKET_READINESS.TELEPORT_PARTNER][Consts.FIELDS.COMMON.TITLE],
            };
        }
        if (responsiblePartyType === MarketReadinessResponsibleParty.DP) {
            responsiblePartyItem = {
                key: item[Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
                text: item[Consts.FIELDS.MARKET_READINESS.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE],
            };
        }
        const initialValues = {
            name: item[Consts.FIELDS.COMMON.TITLE],
            market,
            responsiblePartyType,
            responsiblePartyItem,
            eutelsatOwners: item[Consts.FIELDS.MARKET_READINESS.OWNER]?.map((etuowner) => ({
                key: etuowner[Consts.FIELDS.COMMON.ID],
                text: etuowner[Consts.FIELDS.COMMON.TITLE]
            })) || [],
            verticals: item[Consts.FIELDS.MARKET_READINESS.VERTICALS]?.map((vert) => ({
                key: vert,
                text: vert
            })) || [],
            verticalRAGStatus: item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_RAGSTATUS] || '',
            verticalEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_ESTIMATED_DATE]) : null,
            verticalEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_EFFECTIVE_DATE]) : null,
            verticalComments: item[Consts.FIELDS.MARKET_READINESS.ALL_VERTICALS_COMMENTS] || '',
            spaceRAGStatus: item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_RAGSTATUS] || '',
            spaceEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_ESTIMATED_DATE]) : null,
            spaceEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_EFFECTIVE_DATE]) : null,
            spaceComments: item[Consts.FIELDS.MARKET_READINESS.SPACE_NW_COMMENTS] || '',
            landFixedRAGStatus: item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_RAGSTATUS] || '',
            landFixedEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_ESTIMATED_DATE]) : null,
            landFixedEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_EFFECTIVE_DATE]) : null,
            landFixedComments: item[Consts.FIELDS.MARKET_READINESS.LAND_FIXED_COMMENTS] || '',
            landMobilityRAGStatus: item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_RAGSTATUS] || '',
            landMobilityEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_ESTIMATED_DATE]) : null,
            landMobilityEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_EFFECTIVE_DATE]) : null,
            landMobilityComments: item[Consts.FIELDS.MARKET_READINESS.LAND_MOBILITY_COMMENTS] || '',
            maritimeRAGStatus: item[Consts.FIELDS.MARKET_READINESS.MARITIME_RAGSTATUS] || '',
            maritimeEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.MARITIME_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.MARITIME_ESTIMATED_DATE]) : null,
            maritimeEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.MARITIME_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.MARITIME_EFFECTIVE_DATE]) : null,
            maritimeComments: item[Consts.FIELDS.MARKET_READINESS.MARITIME_COMMENTS] || '',
            aviationRAGStatus: item[Consts.FIELDS.MARKET_READINESS.AVIATION_RAGSTATUS] || '',
            aviationEstimatedDate: item[Consts.FIELDS.MARKET_READINESS.AVIATION_ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.AVIATION_ESTIMATED_DATE]) : null,
            aviationEffectiveDate: item[Consts.FIELDS.MARKET_READINESS.AVIATION_EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.MARKET_READINESS.AVIATION_EFFECTIVE_DATE]) : null,
            aviationComments: item[Consts.FIELDS.MARKET_READINESS.AVIATION_COMMENTS] || '',
            accordionOpenItems: []
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

    private formatDate = (date?: Date): string => {
        if (!date) return '';
        return UtilHelper.formatDate(date, 'LL', 'en-us');
    };
}
