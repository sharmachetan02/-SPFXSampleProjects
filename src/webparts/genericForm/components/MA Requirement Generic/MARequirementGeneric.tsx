import * as React from 'react';
import { IMARequirementGenericProps, IMARequirementGenericState, IReletedItem } from './MARequirementGeneric.types';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import { DefaultButton, DirectionalHint, Dropdown, Fabric, Icon, Label, Spinner, SpinnerSize, TextField } from '@fluentui/react';
import { IDropdownOption, } from '@fluentui/react';
import { DatePicker } from '@fluentui/react';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { Form, FormOrigin, FormType, IFormConfig } from '../Form.types';
import strings from 'GenericFormWebPartStrings';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { UtilHelper } from '../../../../common/helpers/Util';
import { Consts } from '../../../../common/consts/Consts';
import 'react-widgets/styles.css';
import { IItemAddResult } from '@pnp/sp/items';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { createUniqueFolder, datepickerstyles } from '../Form.utility';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import { RichText } from '../../../../common/components/RichText/RichText';
import { ICategorizedDropdownOption, MARequirementBasicData } from '../MA Requirement Basic/MARequirementBasic.types';
import MARequirementBasic from '../MA Requirement Basic/MARequirementBasic';
import { Accordion, AccordionHeader, AccordionItem, AccordionPanel } from '../../../../common/components/Accordion/Accordion';
import { ContactType, MARequirementResponsibleParty, MARequirementType, RelatedItemType, UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';
export default class MARequirementGeneric extends React.Component<IMARequirementGenericProps, IMARequirementGenericState> {
    private isEditMode: boolean;
    private config: IFormConfig;
    private AllResponsibleItems: Map<string, IDropdownOption[]> = new Map();
    private AllSynthesisStatusItems: Map<string, IDropdownOption[]> = new Map();
    private AllContactsItems: Map<string, IDropdownOption[]> = new Map();
    private AllReletedItems: Map<string, IDropdownOption[]> = new Map();
    constructor(props: Readonly<IMARequirementGenericProps>) {
        super(props);
        this.isEditMode = this.props.itemId !== undefined;
        this.state = {
            maRequirementBasicData: {
                maVertical: [],
                maVerticalChoices: [],
                snps: [],
                snpsChoices: [],
                geoLeo: '',
                geoLeoChoices: [],
                comments: '',
                name: '',
                type: '',
                typeChoices: [],
                requirementType: '',
                requirementTypeChoices: [],
                summary: '',
                markets: [],
                marketsChoices: [],
                responsibleParty: '',
                responsiblePartyChoices: [],
                responsiblePartyItem: null,
                responsiblePartyItemChoices: [],
                eutelsatOwners: [],
                eutelsatOwnersChoices: [],
                responsiblePartyContacts: [],
                responsiblePartyContactsChoices: [],
                synthesisStatus: '',
                synthesisStatusChoices: [],
                ragStatus: '',
                ragStatusChoices: [],
            },
            reletedItems: [],
            isFormReady: false,
            applicationDate: null,
            estimatedDate: null,
            effectiveDate: null,
            errors: {}
        };
        this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.MARequirementGeneric);
        this.AllResponsibleItems = new Map();
        this.AllSynthesisStatusItems = new Map();
        this.AllContactsItems = new Map();
        this.AllReletedItems = new Map();
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
        await Promise.all([
            this.fetchResposnibleItems(),
            this.fetchContactsItems(),
            this.fetchAllReletedItems()
        ]);
        promises.push(this.initTypeChoices());
        promises.push(
            this.initMarketChoices());
        promises.push(
            this.initRequirementTypeChoices());
        promises.push(
            this.initMAVerticalsChoices());
        promises.push(
            this.initResponsiblePartyChoices());
        promises.push(
            this.initSynthesisStatusChoices());
        promises.push(
            this.initRagStatusChoices());
        promises.push(
            this.initEutelsatOwnersChoices());
        promises.push(
            this.initSNPsChoices());
        promises.push(
            this.initGeoLeoChoices());
        await Promise.all(promises);
        if (this.isEditMode)
            await this.setInitialFormValues();
    }

    private fetchResposnibleItems = async () => {
        const { pnpService } = this.props;
        const selectedFields = [Consts.FIELDS.COMMON.ID, Consts.FIELDS.COMMON.TITLE];

        const configMap = [
            { key: MARequirementResponsibleParty.Eutelsat, type: FormType.EutelsatEntity, contentTypeId: null },
            { key: MARequirementResponsibleParty.TP, type: FormType.TP, contentTypeId: null },
            { key: MARequirementResponsibleParty.DP, type: FormType.DP, contentTypeId: null }
        ];

        const fetchPromises = configMap?.map(async (config) => {
            const list = Form.config.find((f) => f.type.toLowerCase() === config.type.toLowerCase())?.list;
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
            let query = pnpService.getListItems(listUrl).select(...selectedFields).expand(Consts.FIELDS.COMMON.CONTENTTYPE);
            if (config.contentTypeId) {
                query = query.filter(`startswith(${Consts.FIELDS.COMMON.CONTENTTYPE_ID}, '${config.contentTypeId}')`);
            }
            const items = await query.orderBy(Consts.FIELDS.COMMON.TITLE, true).top(5000)();
            const choices = items.map((item) => ({
                key: item[Consts.FIELDS.COMMON.ID] + "",
                text: item[Consts.FIELDS.COMMON.TITLE]
            }));
            return { key: config.key, choices };
        });

        const results = await Promise.all(fetchPromises);
        results.forEach(({ key, choices }) => {
            this.AllResponsibleItems.set(key, choices);
        });
    };
    private getConfigMap = (): { key: string; text: FormType; contentTypeId: string | null }[] => [
        { key: RelatedItemType.DistributionPartner, text: FormType.DP, contentTypeId: null },
        { key: RelatedItemType.SNP, text: FormType.SNP, contentTypeId: Consts.CONTENT_TYPES.SNP },
        { key: RelatedItemType.Administration, text: FormType.Authority, contentTypeId: Consts.CONTENT_TYPES.ADMINISTRATION },
        { key: RelatedItemType.Organization, text: FormType.Authority, contentTypeId: Consts.CONTENT_TYPES.ORGANIZATION },
        { key: RelatedItemType.Regulator, text: FormType.Authority, contentTypeId: Consts.CONTENT_TYPES.REGULATOR },
        { key: RelatedItemType.LawFirm, text: FormType.LegalServiceProvider, contentTypeId: Consts.CONTENT_TYPES.LAW_FIRM },
        { key: RelatedItemType.LegalRepresentative, text: FormType.LegalServiceProvider, contentTypeId: Consts.CONTENT_TYPES.LEGAL_REPRESENTATIVE },
        // { key: strings.FormMARequirementGenericTypeTeleportPartner, text: FormType.TP, contentTypeId: null }
    ];

    private createDefaultReletedItem = async (): Promise<IReletedItem> => {
        const releted = RelatedItemType.DistributionPartner;
        const relatedChoicesData: IDropdownOption[] = await this.getConfigMap();
        const reletedChoices = relatedChoicesData.map((option) => ({
            key: option.key,
            text: option.key + ""
        }));
        const reletedItemChoices: IDropdownOption[] = (this.AllReletedItems.get(releted));
        return {
            Id: undefined,
            relationShip: "",
            releted,
            reletedChoices,
            reletedItem: [],
            reletedItemChoices
        };
    };

    private fetchAllReletedItems = async () => {
        const { pnpService } = this.props;
        const selectedFields = [Consts.FIELDS.COMMON.ID, Consts.FIELDS.COMMON.TITLE];

        const configMap = await this.getConfigMap();

        const fetchPromises = configMap.map(async (config) => {
            const list = Form.config.find((f) => f.type.toLowerCase() === config.text.toLowerCase())?.list;
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
            let query = pnpService.getListItems(listUrl).select(...selectedFields).expand(Consts.FIELDS.COMMON.CONTENTTYPE);
            if (config.contentTypeId) {
                query = query.filter(`startswith(${Consts.FIELDS.COMMON.CONTENTTYPE_ID}, '${config.contentTypeId}')`);
            }
            const items = await query.orderBy(Consts.FIELDS.COMMON.TITLE, true).top(5000)();
            const choices = items?.map((item) => ({
                key: item[Consts.FIELDS.COMMON.ID],
                text: item[Consts.FIELDS.COMMON.TITLE]
            }));
            return { key: config.key, choices };
        });

        const results = await Promise.all(fetchPromises);
        results.forEach(({ key, choices }) => {
            this.AllReletedItems.set(key, choices);
        });

    };
    private fetchContactsItems = async () => {
        const { pnpService } = this.props;
        const selectedFields = [Consts.FIELDS.COMMON.ID, Consts.FIELDS.COMMON.TITLE, `${Consts.FIELDS.CONTACT.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
        `${Consts.FIELDS.CONTACT.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
        `${Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
        `${Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`

        ];

        const configMap = [
            { key: ContactType.TP, type: FormType.Contact, contentTypeId: Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT },
            { key: ContactType.DP, type: FormType.Contact, contentTypeId: Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT }
        ];

        const fetchPromises = configMap.map(async (config) => {
            const list = Form.config.find((f) => f.type.toLowerCase() === config.type.toLowerCase())?.list;
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
            const expand = [
                Consts.FIELDS.COMMON.CONTENTTYPE,
                Consts.FIELDS.CONTACT.TELEPORT_PARTNER,
                Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER
            ];
            let query = pnpService.getListItems(listUrl).select(...selectedFields).expand(...expand);
            if (config.contentTypeId) {
                query = query.filter(`startswith(${Consts.FIELDS.COMMON.CONTENTTYPE_ID}, '${config.contentTypeId}')`);
            }
            const items = await query.orderBy(Consts.FIELDS.COMMON.TITLE, true).top(5000)();
            const choices = items?.map((item) => ({
                key: item[Consts.FIELDS.COMMON.ID] + "",
                text: item[Consts.FIELDS.COMMON.TITLE],
                TPId: item[Consts.FIELDS.CONTACT.TELEPORT_PARTNER]?.[Consts.FIELDS.COMMON.ID] + "",
                DPId: item[Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.ID] + "",
            }));
            return { key: config.key, choices };
        });

        const results = await Promise.all(fetchPromises);
        results.forEach(({ key, choices }) => {
            this.AllContactsItems.set(key, choices);
        });
    };

    private initEditModeReletedItems = async (itemIds: number[]): Promise<IReletedItem[] | null> => {
        if (!this.isEditMode || !itemIds || itemIds.length <= 0) return null;

        const { pnpService } = this.props;

        const reletedItemsListUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.REQUIREMENT_RELTIONSHIP_URL);
        if (!reletedItemsListUrl) return null;

        //const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), reletedItemsListUrl);

        const selected: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.MA_REQUIREMENTRELATIONSHIPTYPE,
            `${Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.SNP}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.SNP}/${Consts.FIELDS.COMMON.TITLE}`,
        ];

        const expand: string[] = [
            Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.LEGALSERVICEPROVIDER,
            Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.AUTHORITY,
            Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.DISTRIBUTION_PARTNER,
            Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.SNP,
        ];


        try {
            const promises = itemIds?.map((id) =>
                pnpService.getListByUrl(reletedItemsListUrl).items.getById(id).select(...selected).expand(...expand)()
            );

            const items = await Promise.all(promises);

            const result: IReletedItem[] = await Promise.all(
                items?.map(async (item) => {
                    // Create and update default item inside loop
                    const defaultItem = await this.createDefaultReletedItem();
                    if (defaultItem) {

                        defaultItem.Id = item[Consts.FIELDS.COMMON.ID];
                        defaultItem.relationShip = item[Consts.FIELDS.COMMON.TITLE] ?? "";
                        defaultItem.releted = item[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.MA_REQUIREMENTRELATIONSHIPTYPE] ?? "";
                        //defaultItem.reletedItem = [];
                        const reletedItemChoicesData = this.AllReletedItems.get(defaultItem.releted) || [];
                        const reletedItemChoices = reletedItemChoicesData.map((option) => ({
                            key: option.key + "",
                            text: option.text + ""
                        }));
                        defaultItem.reletedItemChoices = reletedItemChoices;
                        switch (item[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.MA_REQUIREMENTRELATIONSHIPTYPE]) {
                            // Group 1: Law Firm & Legal Representative
                            case RelatedItemType.LawFirm:
                            case RelatedItemType.LegalRepresentative:
                                defaultItem.reletedItem = item[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.LEGALSERVICEPROVIDER]?.map((relatedItem) => ({
                                    key: relatedItem[Consts.FIELDS.COMMON.ID],
                                    text: relatedItem[Consts.FIELDS.COMMON.TITLE]
                                })) || [];
                                break;

                            // Group 2: Organization & Regulator
                            case RelatedItemType.Administration:
                            case RelatedItemType.Organization:
                            case RelatedItemType.Regulator:
                                defaultItem.reletedItem = item[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.AUTHORITY]?.map((relatedItem) => ({
                                    key: relatedItem[Consts.FIELDS.COMMON.ID],
                                    text: relatedItem[Consts.FIELDS.COMMON.TITLE]
                                })) || [];
                                break;
                            // Group 3: DP
                            case RelatedItemType.DistributionPartner:
                                defaultItem.reletedItem = item[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.DISTRIBUTION_PARTNER]?.map((relatedItem) => ({
                                    key: relatedItem[Consts.FIELDS.COMMON.ID],
                                    text: relatedItem[Consts.FIELDS.COMMON.TITLE]
                                })) || [];
                                break;
                            // Group 4: SNP
                            case RelatedItemType.SNP:
                                defaultItem.reletedItem = item[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.SNP]?.map((relatedItem) => ({
                                    key: relatedItem[Consts.FIELDS.COMMON.ID],
                                    text: relatedItem[Consts.FIELDS.COMMON.TITLE]
                                })) || [];
                                break;

                            // Group 5: Teleport Partner
                            case RelatedItemType.TeleportPartner:
                                defaultItem.reletedItem = item[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.TELEPORTPARTNER]?.map((relatedItem) => ({
                                    key: relatedItem[Consts.FIELDS.COMMON.ID],
                                    text: relatedItem[Consts.FIELDS.COMMON.TITLE]
                                })) || [];
                                break;

                            // Default: do nothing
                            default:
                                break;
                        }
                    }

                    return defaultItem;
                })
            );
            return result;
        } catch (error) {
            console.error("Error fetching related items:", error);
            return null;
        }
    };

    public setInitialFormValues = async () => {

        const { pnpService, itemId } = this.props;
        const list = Form.config.find((f) => f.type.toLowerCase() === FormType.MARequirementGeneric).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
            Consts.FIELDS.MAGENERIC.SUMMARY,
            `${Consts.FIELDS.MAGENERIC.MARKETS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.MARKETS}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.VERTICALS,
            `${Consts.FIELDS.MAGENERIC.SYNTHESIS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.SYNTHESIS}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.RAGSTATUS,
            `${Consts.FIELDS.MAGENERIC.SNPS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.SNPS}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY,
            `${Consts.FIELDS.MAGENERIC.EUTELSATENTITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.EUTELSATENTITY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.CONTACTS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.CONTACTS}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.EUTELSATOWNERS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.EUTELSATOWNERS}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.GEOLEO,
            Consts.FIELDS.MAGENERIC.COMMENTS,
            Consts.FIELDS.MAGENERIC.APPLICATIONDATE,
            Consts.FIELDS.MAGENERIC.ESTIMATEDDATE,
            Consts.FIELDS.MAGENERIC.EFFECTIVEDATE,
            `${Consts.FIELDS.MAGENERIC.RELATEDITEMS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.RELATEDITEMS}/${Consts.FIELDS.COMMON.TITLE}`,
        ];
        const expand = [Consts.FIELDS.MAGENERIC.MARKETS, Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE,
        Consts.FIELDS.MAGENERIC.SYNTHESIS, Consts.FIELDS.MAGENERIC.SNPS, Consts.FIELDS.MAGENERIC.EUTELSATENTITY, Consts.FIELDS.MAGENERIC.TELEPORTPARTNER,
        Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER, Consts.FIELDS.MAGENERIC.CONTACTS, Consts.FIELDS.MAGENERIC.EUTELSATOWNERS,
        Consts.FIELDS.MAGENERIC.RELATEDITEMS
        ];
        const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
        console.log(item);
        const filterTypeChoices = "" + this.state.maRequirementBasicData.typeChoices.find((type) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1)?.key;
        const maVertical = item[Consts.FIELDS.MAGENERIC.VERTICALS]?.map((v) => ({
            key: v,
            text: v
        }));
        const responsibleParty = item[Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY];
        let responsiblePartyItem = null;
        switch (responsibleParty) {
            case MARequirementResponsibleParty.Eutelsat:
                responsiblePartyItem = {
                    key: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID],
                    text: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.TITLE]
                };
                break;
            case MARequirementResponsibleParty.TP:
                responsiblePartyItem = {
                    key: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
                    text: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE]
                };
                break;
            case MARequirementResponsibleParty.DP:
                responsiblePartyItem = {
                    key: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
                    text: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE]
                };
                break;
            default:
                break;
        }

        const rawApplicationDate = item[Consts.FIELDS.MAGENERIC.APPLICATIONDATE];
        const applicationDate: Date | null = rawApplicationDate ? new Date(rawApplicationDate) : null;


        const rawEstimatedDate = item[Consts.FIELDS.MAGENERIC.ESTIMATEDDATE];
        const estimatedDate: Date | null = rawEstimatedDate ? new Date(rawEstimatedDate) : null;

        const rawEffectiveDate = item[Consts.FIELDS.MAGENERIC.EFFECTIVEDATE];
        const effectiveDate: Date | null = rawEffectiveDate ? new Date(rawEffectiveDate) : null;

        const reletedItemsFieldIds = item[`${Consts.FIELDS.MAGENERIC.RELATEDITEMS}`]?.map((f) => (f[Consts.FIELDS.COMMON.ID]));
        let reletedItems = [];
        if (reletedItemsFieldIds) {
            reletedItems = await this.initEditModeReletedItems(reletedItemsFieldIds);
        }

        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            name: item[Consts.FIELDS.COMMON.TITLE],
            type: filterTypeChoices,
            summary: item[Consts.FIELDS.MAGENERIC.SUMMARY],
            markets: item[Consts.FIELDS.MAGENERIC.MARKETS]?.map((markets) => ({
                key: markets[Consts.FIELDS.COMMON.ID],
                text: markets[Consts.FIELDS.COMMON.TITLE]
            })) || [],
            requirementType: item[Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE]?.[Consts.FIELDS.COMMON.ID] + "",
            synthesisStatus: item[Consts.FIELDS.MAGENERIC.SYNTHESIS]?.[Consts.FIELDS.COMMON.ID] + "",
            maVertical,
            ragStatus: item[Consts.FIELDS.MAGENERIC.RAGSTATUS],
            responsibleParty,
            responsiblePartyItem,
            responsiblePartyContacts: item[Consts.FIELDS.MAGENERIC.CONTACTS]?.map((contacts) => ({
                key: contacts[Consts.FIELDS.COMMON.ID],
                text: contacts[Consts.FIELDS.COMMON.TITLE]
            })) || [],
            eutelsatOwners: item[Consts.FIELDS.MAGENERIC.EUTELSATOWNERS]?.map((owners) => ({
                key: owners[Consts.FIELDS.COMMON.ID],
                text: owners[Consts.FIELDS.COMMON.TITLE]
            })) || [],
            geoLeo: item[Consts.FIELDS.MAGENERIC.GEOLEO],
            comments: item[Consts.FIELDS.MAGENERIC.COMMENTS],
            snps: item[Consts.FIELDS.MAGENERIC.SNPS]?.map((snp) => ({
                key: snp[Consts.FIELDS.COMMON.ID],
                text: snp[Consts.FIELDS.COMMON.TITLE]
            })) || [],
        };

        const initialValues = {
            maRequirementBasicData,
            applicationDate,
            estimatedDate,
            effectiveDate,
            reletedItems
        };
        this.setState({
            ...initialValues
        });


    };

    private initRequirementTypeChoices = async () => {
        const {
            pnpService
        } = this.props;

        const list = Form.config.find((f) => f.type.toLowerCase() === FormType.RequirementType).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.REQUIREMENT_TYPE.CATEGORY,
            Consts.FIELDS.REQUIREMENT_TYPE.VERTICAL
        ];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        const requirementTypeChoices: ICategorizedDropdownOption[] = [];
        items.forEach((item) => {
            requirementTypeChoices.push({
                key: item[Consts.FIELDS.COMMON.ID],
                text: item[Consts.FIELDS.COMMON.TITLE],
                maVerticals: item[Consts.FIELDS.REQUIREMENT_TYPE.VERTICAL],
                category: item[Consts.FIELDS.REQUIREMENT_TYPE.CATEGORY]
            });
        });

        const requirementType = '';
        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            requirementTypeChoices,
            requirementType
        };
        this.setState(() => ({ maRequirementBasicData }));

    };

    private initTypeChoices = async (): Promise<void> => {
        const typeChoices: IDropdownOption[] = [
            { key: Consts.CONTENT_TYPES.MA_RequirementGeneric, text: MARequirementType.Generic },
        ];
        const type: string = Consts.CONTENT_TYPES.MA_RequirementGeneric;

        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            typeChoices,
            type
        };
        this.setState(() => ({ maRequirementBasicData }));
    };

    private initMarketChoices = async (): Promise<void> => {
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
        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            marketsChoices,
            markets
        };
        this.setState(() => ({ maRequirementBasicData }));

    };

    private initSNPsChoices = async (): Promise<void> => {
        const {
            pnpService
        } = this.props;

        const snpsChoices: IDropdownOption[] = [];
        const list = Form.config.find((f) => f.type.toLowerCase() === FormType.SNP).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE
        ];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        items.forEach((item) => {
            snpsChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
        });
        const snps = [];
        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            snpsChoices,
            snps

        };
        this.setState(() => ({ maRequirementBasicData }));

    };

    public initResponsiblePartyChoices = async (): Promise<void> => {
        const {
            pnpService
        } = this.props;
        let responsiblePartyChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY)
            .select('Choices')();
        responsiblePartyChoices = data.Choices?.map((choice) => ({ key: choice, text: choice }));

        const responsibleParty = responsiblePartyChoices.length > 0 ? responsiblePartyChoices[0].key as string : '';
        const responsiblePartyItemChoices: IDropdownOption[] = this.AllResponsibleItems.get(responsibleParty);
        // If found, assign it to synthesisStatusChoices


        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            responsiblePartyChoices,
            responsibleParty,
            responsiblePartyItemChoices
        };
        this.setState(() => ({ maRequirementBasicData }));


    };

    private initSynthesisStatusChoices = async (): Promise<void> => {
        const {
            pnpService
        } = this.props;
        const synthesisStatusChoices: IDropdownOption[] = [];
        const list = Form.config.find((f) => f.type.toLowerCase() === FormType.Requirement_Status).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            `${Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE}/${Consts.FIELDS.COMMON.TITLE}`,

        ];
        const expand = [Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).expand(...expand).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        items.forEach((item) => {
            synthesisStatusChoices.push({ key: `${item[Consts.FIELDS.COMMON.ID]}`, text: item[Consts.FIELDS.COMMON.TITLE] });
            const key = item[Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE][Consts.FIELDS.COMMON.ID] + "";
            const value = {
                key: item[Consts.FIELDS.COMMON.ID] + "",
                text: item[Consts.FIELDS.COMMON.TITLE]
            };
            if (this.AllSynthesisStatusItems.has(key)) {
                this.AllSynthesisStatusItems.get(key)?.push(value);
            } else {
                this.AllSynthesisStatusItems.set(key, [value]);
            }
        });

        const synthesisStatus = '';
        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            synthesisStatusChoices,
            synthesisStatus
        };
        this.setState(() => ({ maRequirementBasicData }));

    };

    public initMAVerticalsChoices = async (): Promise<void> => {
        const {
            pnpService
        } = this.props;
        let maVerticalChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MAGENERIC.VERTICALS)
            .select('Choices')();
        maVerticalChoices = data.Choices?.map((choice) => ({ key: choice, text: choice }));
        const maVertical: IDropdownOption[] = [];


        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            maVerticalChoices,
            maVertical
        };
        this.setState({ maRequirementBasicData });
    };

    public initRagStatusChoices = async (): Promise<void> => {
        const {
            pnpService
        } = this.props;
        let ragStatusChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MAGENERIC.RAGSTATUS)
            .select('Choices')();
        ragStatusChoices = data.Choices?.map((choice) => ({ key: choice, text: choice }));
        const ragStatus = ragStatusChoices.length > 0 ? ragStatusChoices[0].key as string : '';

        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            ragStatusChoices,
            ragStatus
        };
        this.setState({ maRequirementBasicData });
    };

    private initEutelsatOwnersChoices = async (): Promise<void> => {
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
        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            eutelsatOwnersChoices,
            eutelsatOwners
        };
        this.setState(() => ({ maRequirementBasicData }));

    };

    public initGeoLeoChoices = async (): Promise<void> => {
        const {
            pnpService
        } = this.props;
        let geoLeoChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MAGENERIC.GEOLEO)
            .select('Choices')();
        geoLeoChoices = data.Choices?.map((choice) => ({ key: choice, text: choice }));
        const geoLeo = geoLeoChoices.length > 0 ? geoLeoChoices[0].key as string : '';
        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            geoLeoChoices,
            geoLeo
        };
        this.setState({ maRequirementBasicData });
    };


    private validateForm = (): boolean => {
        const errors: { [key: string]: string } = {};

        const {
            maRequirementBasicData,
            reletedItems
        } = this.state;
        const {
            name,
            markets,
            requirementType,
            maVertical,
            synthesisStatus,
            responsiblePartyItem,
            eutelsatOwners,


        } = maRequirementBasicData;
        if (name?.trim() === '' || !name)
            errors.name = strings.FormMARequirementErrorNameRequired;
        if (!markets || markets.length === 0) {
            errors['markets'] = strings.FormMARequirementErrorMarketsRequired;
        }
        if (!maVertical || maVertical.length === 0) {
            errors['maVertical'] = strings.FormMARequirementErrorMAVerticalsRequired;
        }
        if (!requirementType?.trim() || !requirementType) {
            errors['requirementType'] = strings.FormMARequirementErrorRequirementTypeRequired;
        }
        if (!synthesisStatus?.trim() || !synthesisStatus) {
            errors['synthesisStatus'] = strings.FormMARequirementErrorSynthesisRequired;
        }
        if (!responsiblePartyItem || !responsiblePartyItem?.key || responsiblePartyItem?.key === '') {
            errors['responsiblePartyItem'] = strings.FormMARequirementErrorResbonsiblePartyItempRequired;
        }
        if (!eutelsatOwners || eutelsatOwners.length === 0) {
            errors['eutelsatOwners'] = strings.FormMARequirementErrorEutelsatOwnersRequired;
        }
        if (reletedItems && reletedItems.length > 0) {
            reletedItems.forEach((item, idx) => {
                if (!item.relationShip || item.relationShip.trim() === '') {
                    errors[`reletedItems${[idx]}relationShip`] = strings.FormMARequirementErrorRelationshipRequired;
                }
                if (!item.reletedItem || item.reletedItem.length === 0) {
                    errors[`reletedItems${[idx]}reletedItem`] = strings.FormMARequirementErrorRelationshipItemRequired;
                }
            });
        }


        this.setState({ errors });
        return Object.keys(errors).length === 0;
    };

    public render(): React.ReactElement<IMARequirementGenericProps> {
        const {
            estimatedDate,
            applicationDate,
            effectiveDate,
            maRequirementBasicData,
            errors,
            isFormReady,
            reletedItems,

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
                                    <MARequirementBasic maRequirementBasicData={maRequirementBasicData} onChange={this.onChangeMARequiremntBasic} errors={errors}></MARequirementBasic>
                                    <div className="ms-Grid" >
                                        <div className="ms-Grid-row" >
                                            <div className="ms-Grid-col ms-sm6 ms-md6">
                                                {/* Date row */}
                                                {/* Application Date */}
                                                <div className={styles.field} >
                                                    <div className={styles.fieldLabel}>
                                                        <div className={styles.fieldLabelContainer}>
                                                            <Label >{strings.FormMARequirementGenericApplicationDateLabel}</Label>
                                                        </div>
                                                        <Tooltip content={strings.FormMARequirementGenericTooltipApplicationDate} ></Tooltip>
                                                    </div>
                                                    <div>
                                                        <DatePicker
                                                            onSelectDate={this.onchangeApplicationDate}
                                                            value={applicationDate}
                                                            formatDate={this.formatDate}
                                                            allowTextInput
                                                            styles={datepickerstyles}
                                                        />
                                                    </div>

                                                </div>
                                            </div>
                                            <div className="ms-Grid-col ms-sm6 ms-md6">
                                                {/* Estimated Date */}
                                                <div className={styles.field}>
                                                    <div className={styles.fieldLabel}>
                                                        <div className={styles.fieldLabelContainer}>
                                                            <Label >{strings.FormMARequirementGenericEstimatedDateLabel}</Label>
                                                        </div>
                                                        <Tooltip content={strings.FormMARequirementGenericTooltipEstimatedDate} ></Tooltip>
                                                    </div>
                                                    <DatePicker
                                                        onSelectDate={this.onchangeEstimatedDate}
                                                        value={estimatedDate}
                                                        formatDate={this.formatDate}
                                                        allowTextInput
                                                        styles={datepickerstyles}
                                                    />
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    {/* Effective Date */}
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMARequirementGenericEffectiveDateLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementGenericTooltipEffectiveDate} ></Tooltip>
                                        </div>
                                        <DatePicker
                                            onSelectDate={this.onchangeEffectiveDate}
                                            value={effectiveDate}
                                            formatDate={this.formatDate}
                                            allowTextInput
                                            styles={datepickerstyles}
                                        />
                                    </div>
                                    {/* Accordion wrapper */}
                                    <div className={styles.fieldLabel}>
                                        <div className={styles.fieldLabelContainer}>
                                            <Label>{strings.FormMARequirementGenericRelatedItemsLabel}</Label>
                                        </div>
                                        <Tooltip content={strings.FormMARequirementGenericTooltipRelatedItems} ></Tooltip>
                                    </div>
                                    <div className={styles.field}>
                                        <DefaultButton
                                            text="Add Releted Item"
                                            iconProps={{ iconName: 'Add' }}
                                            onClick={this.onAddItem} // your handler
                                        />
                                    </div>

                                    {reletedItems && reletedItems.length > 0 && (
                                        <Accordion
                                            multiple
                                            collapsible
                                            defaultOpenItems={Array.from({ length: reletedItems?.length }, (_, i) => i.toString())}
                                        >
                                            {reletedItems?.map((_, idx) => (
                                                <AccordionItem value={idx + ""} key={idx}>
                                                    <AccordionHeader>{_.relationShip && _.relationShip.trim() ? _.relationShip : `Item ${idx + 1}`}</AccordionHeader>
                                                    <AccordionPanel>
                                                        {/* RelationShip */}
                                                        <div className={styles.field}>
                                                            <div className={styles.fieldLabel}>
                                                                <div className={styles.fieldLabelContainer}>
                                                                    <Label required>{strings.FormMARequirementLabelRelationship}</Label>
                                                                </div>
                                                                <Tooltip content={strings.FormMARequirementTooltipRelationship} ></Tooltip>
                                                            </div>
                                                            <TextField
                                                                value={_.relationShip}
                                                                onChange={(e) => this.onChangeRelationship(e, idx)}
                                                            />
                                                            {errors[`reletedItems${[idx]}relationShip`] &&
                                                                <div className={styles.errorContainer}>
                                                                    <Icon iconName="Error" className={styles.errorIcon} />
                                                                    <span className={styles.errorMessage}>
                                                                        {errors[`reletedItems${[idx]}relationShip`]}
                                                                    </span>
                                                                </div>
                                                            }
                                                        </div>

                                                        {/* Item By */}
                                                        <div className={styles.field}>
                                                            <div className={styles.fieldLabel}>
                                                                <div className={styles.fieldLabelContainer}>
                                                                    <Label>{strings.FormMARequirementLabelItemType}</Label>
                                                                </div>
                                                                <Tooltip content={strings.FormMARequirementTooltipItemType} ></Tooltip>
                                                            </div>
                                                            <Dropdown
                                                                selectedKey={_.releted}
                                                                onChange={(e, o) => this.onChangeItem(e, o, idx)}
                                                                calloutProps={{
                                                                    directionalHint: DirectionalHint.bottomLeftEdge,
                                                                    directionalHintFixed: false,
                                                                }}
                                                                options={_.reletedChoices} />
                                                        </div>
                                                        {/* Item */}
                                                        <div className={styles.field} >
                                                            <div className={styles.fieldLabel}>
                                                                <div className={styles.fieldLabelContainer}>
                                                                    <Label required>{strings.FormMARequirementLabelItem}</Label>
                                                                </div>
                                                                <Tooltip content={strings.FormMARequirementTooltipItem} ></Tooltip>
                                                            </div>
                                                            <MultiselectWrapper
                                                                id={idx.toString()}
                                                                data={_.reletedItemChoices}
                                                                dataKey={(item: IDropdownOption) => item.key}
                                                                textField={(item: IDropdownOption) => item.text}
                                                                filter="contains"
                                                                value={_.reletedItemChoices.filter((item) => _.reletedItem.some((r) => r.key.toString() === item.key.toString()))}

                                                                onChange={(o) => this.onChangeByItem(o, idx)} />


                                                            {errors[`reletedItems${[idx]}reletedItem`] &&
                                                                <div className={styles.errorContainer}>
                                                                    <Icon iconName="Error" className={styles.errorIcon} />
                                                                    <span className={styles.errorMessage}>
                                                                        {errors[`reletedItems${[idx]}reletedItem`]}
                                                                    </span>
                                                                </div>
                                                            }
                                                        </div>

                                                        <div>
                                                            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                                                                <DefaultButton
                                                                    text="Delete"
                                                                    iconProps={{ iconName: 'Delete' }}
                                                                    styles={{
                                                                        flexContainer: {
                                                                            flexDirection: 'row-reverse' // moves icon to the right
                                                                        },
                                                                        icon: {
                                                                            marginLeft: 8 // spacing between text and icon
                                                                        }
                                                                    }}
                                                                    onClick={() => this.onDeleteItem(idx)}
                                                                />
                                                            </div>

                                                        </div>
                                                    </AccordionPanel>
                                                </AccordionItem>
                                            ))}

                                        </Accordion>
                                    )}
                                </Fabric>
                            }
                        </>
                    </FormWizard>}
                </div>
            </div >
        );
    }

    private onChangeMARequiremntBasic = (maRequirementBasicData: MARequirementBasicData) => {

        // console.log(maRequirementBasicData);


        // Example: Check if 'status' field changed
        if (maRequirementBasicData.requirementType !== this.state.maRequirementBasicData.requirementType) {
            const requirementType = maRequirementBasicData.requirementType;
            // Get array from the map using requirementType as key
            const synthesisArray: IDropdownOption[] = this.AllSynthesisStatusItems.get(requirementType);
            // If found, assign it to synthesisStatusChoices
            if (synthesisArray) {
                maRequirementBasicData.synthesisStatusChoices = synthesisArray;
            } else {
                maRequirementBasicData.synthesisStatusChoices = [];
            }
        }

        // Example: Check if 'status' field changed
        if (maRequirementBasicData.responsibleParty !== this.state.maRequirementBasicData.responsibleParty) {
            const responsibleParty = maRequirementBasicData.responsibleParty;
            // Get array from the map using requirementType as key
            const responsiblePartyItemChoices: IDropdownOption[] = this.AllResponsibleItems.get(responsibleParty);
            // If found, assign it to synthesisStatusChoices
            if (responsiblePartyItemChoices) {
                maRequirementBasicData.responsiblePartyItemChoices = responsiblePartyItemChoices;
            } else {
                maRequirementBasicData.responsiblePartyItemChoices = [];
            }
        }

        if (
            maRequirementBasicData?.responsiblePartyItem?.key !== this.state.maRequirementBasicData?.responsiblePartyItem?.key &&
            maRequirementBasicData?.responsiblePartyItem?.key !== MARequirementResponsibleParty.Eutelsat
        ) {
            const contactAllData = this.AllContactsItems.get(maRequirementBasicData?.responsibleParty);

            const filterKey = maRequirementBasicData?.responsiblePartyItem?.key;

            let filteredContacts: IDropdownOption[] = [];
            switch (maRequirementBasicData.responsibleParty) {
                case MARequirementResponsibleParty.Eutelsat:
                    break;
                case MARequirementResponsibleParty.TP:
                    filteredContacts = contactAllData?.filter((item) => item["TPId"] === filterKey);
                    break;
                case MARequirementResponsibleParty.DP:
                    filteredContacts = contactAllData?.filter((item) => item["DPId"] === filterKey);
                // Filter the contacts by TPId matching the key

            }
            maRequirementBasicData.responsiblePartyContactsChoices = filteredContacts;


        }

        this.setState({ maRequirementBasicData });
    };
    private onAddItem = async () => {
        const defaultReletedItem: IReletedItem = await this.createDefaultReletedItem();
        this.setState((prevState) => ({
            reletedItems: [
                ...prevState.reletedItems, defaultReletedItem
            ]
        })
        );
    };



    private onDeleteItem = (index: number) => {
        this.setState((prevState) => {
            if (index < 0 || index >= prevState.reletedItems.length) return null; // no change
            return {
                reletedItems: prevState.reletedItems.filter((_, i) => i !== index),
            };
        });
    };

    private formatDate = (date?: Date): string => {
        if (!date) return '';
        return UtilHelper.formatDate(date, 'LL', 'en-us');
    };
    private onchangeEstimatedDate = (estimatedDate: Date) => {
        this.setState({ estimatedDate });
    };
    private onchangeApplicationDate = (applicationDate: Date) => {
        this.setState({ applicationDate });
    };
    private onchangeEffectiveDate = (effectiveDate: Date) => {
        this.setState({ effectiveDate });
    };
    private onChangeRelationship = (event, id: number) => {
        const relationShip = event.target.value;
        const reletedItems = this.state.reletedItems;
        reletedItems[id].relationShip = relationShip;
        this.setState({ reletedItems });
    };

    private onChangeItem = async (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption, id: number) => {
        const reletedItemChoices: IDropdownOption[] = (this.AllReletedItems.get(option.key.toString()));
        const reletedItems = this.state.reletedItems;
        reletedItems[id].releted = option.text;
        reletedItems[id].reletedItemChoices = reletedItemChoices;
        this.setState({ reletedItems });
    };
    private onChangeByItem = (option, id: number) => {
        const reletedItems = this.state.reletedItems;
        reletedItems[id].reletedItem = option;
        this.setState({ reletedItems });
    };


    private getDuplicateItems = async (): Promise<IBaseItem[]> => {
        const duplicateItems = [];
        return duplicateItems;
    };

    private getFormSummary = () => {
        const {
            maRequirementBasicData,
            estimatedDate,
            applicationDate,
            effectiveDate,
            reletedItems

        } = this.state;

        const {
            name,
            typeChoices,
            type,
            requirementType,
            requirementTypeChoices,
            summary,
            markets,
            maVertical,
            synthesisStatus,
            synthesisStatusChoices,
            ragStatus,
            snps,
            responsibleParty,
            responsiblePartyItem,
            responsiblePartyContacts,
            eutelsatOwners,
            geoLeo,
            geoLeoChoices,
            comments

        } = maRequirementBasicData;
        const typeTextValue = typeChoices.find((choice) => choice.key === type)?.text;
        const requirementTypeTextValue = requirementTypeChoices.find((choice) => choice.key.toString() === requirementType)?.text;
        const marketsValues = markets?.map((c) => c.text).join(", ");
        const maVerticalValues = maVertical?.map((c) => c.text).join(", ");
        const synthesisTextValue = synthesisStatusChoices.find((choice) => choice.key === synthesisStatus)?.text;
        const snpsValues = snps?.map((c) => c.text).join(", ");
        const contactsValues = responsiblePartyContacts?.map((c) => c.text).join(", ");
        const eutelsatOwnersValues = eutelsatOwners?.map((c) => c.text).join(", ");
        const geoLeoTextValue = geoLeoChoices.find((choice) => choice.key === geoLeo)?.text;
        const summaryTextValue = DomHelper.cleanRichHtml(summary);
        const commentsTextValue = DomHelper.cleanRichHtml(comments);
        return (
            <div className={styles.summary}>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelName}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {name}
                    </div>
                </div>


                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelType}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {typeTextValue}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelSummary}</Label>
                        </div>
                    </div>

                    <div className={styles.fieldValue}>
                        {summaryTextValue.length > 0 ? <RichText isEditMode={false} value={summaryTextValue} /> : ''}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelMarkets}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {marketsValues}
                    </div>
                </div>

                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelRequirementType}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {requirementTypeTextValue}
                    </div>

                    <div className={styles.field}>
                        <div className={styles.fieldLabel}>
                            <div className={styles.fieldLabelContainer}>
                                <Label>{strings.FormMARequirementLabelMAVertical}</Label>
                            </div>
                        </div>
                        <div className={styles.fieldValue}>
                            {maVerticalValues}
                        </div>
                    </div>

                    <div className={styles.field}>
                        <div className={styles.fieldLabel}>
                            <div className={styles.fieldLabelContainer}>
                                <Label>{strings.FormMARequirementLabelSynthesis}</Label>
                            </div>
                        </div>
                        <div className={styles.fieldValue}>
                            {synthesisTextValue}
                        </div>
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelRagStatus}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {ragStatus}
                    </div>
                </div>
                {maVertical?.some((v) => v.key === "SNP") &&
                    (<div className={styles.field}>
                        <div className={styles.fieldLabel}>
                            <div className={styles.fieldLabelContainer}>
                                <Label>{strings.FormMARequirementLabelSNPS}</Label>
                            </div>
                        </div>
                        <div className={styles.fieldValue}>
                            {snpsValues}
                        </div>
                    </div>)
                }

                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelResponsibleParty}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {responsibleParty}
                    </div>
                </div>

                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelResponsiblePartyItem}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {responsiblePartyItem?.text}
                    </div>
                </div>
                {
                    responsibleParty !== MARequirementResponsibleParty.Eutelsat &&
                    (<div className={styles.field}>
                        <div className={styles.fieldLabel}>
                            <div className={styles.fieldLabelContainer}>
                                <Label>{strings.FormMARequirementLabelContacts}</Label>
                            </div>
                        </div>
                        <div className={styles.fieldValue}>
                            {contactsValues}
                        </div>
                    </div>)
                }


                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelEutelsatOwners}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {eutelsatOwnersValues}
                    </div>
                </div>

                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelGeoLeo}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {geoLeoTextValue}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelComments}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {commentsTextValue.length > 0 ? <RichText isEditMode={false} value={commentsTextValue} /> : ''}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementGenericApplicationDateLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>

                        {applicationDate ? this.formatDate(applicationDate) : ''}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementGenericEstimatedDateLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>

                        {estimatedDate ? this.formatDate(estimatedDate) : ''}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementGenericEffectiveDateLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>

                        {effectiveDate ? this.formatDate(effectiveDate) : ''}
                    </div>
                </div>

                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementGenericRelatedItemsLabel}</Label>
                        </div>
                    </div>
                    <Accordion
                        multiple
                        collapsible
                        defaultOpenItems={Array.from({ length: reletedItems?.length }, (_, i) => i.toString())}
                    >
                        {reletedItems?.map((_, idx) => (
                            <AccordionItem value={idx + ""} key={idx}>
                                <AccordionHeader>{_.relationShip && _.relationShip.trim() ? _.relationShip : `Item ${idx + 1}`}</AccordionHeader>
                                <AccordionPanel>

                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMARequirementLabelRelationship}</Label>
                                            </div>
                                        </div>
                                        <div className={styles.fieldValue}>

                                            {_.relationShip}
                                        </div>
                                    </div>
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMARequirementLabelItemType}</Label>
                                            </div>
                                        </div>
                                        <div className={styles.fieldValue}>

                                            {_.releted}
                                        </div>
                                    </div>
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMARequirementLabelItem}</Label>
                                            </div>
                                        </div>
                                        <div className={styles.fieldValue}>
                                            {_.reletedItem?.map((c) => c.text).join(", ")}
                                        </div>
                                    </div>

                                </AccordionPanel>
                            </AccordionItem>
                        ))}

                    </Accordion>
                </div >
            </div >

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
            maRequirementBasicData,
            applicationDate,
            estimatedDate,
            effectiveDate,
            reletedItems
        } = this.state;
        const {
            name,
            type,
            summary,
            markets,
            requirementType,
            maVertical,
            synthesisStatus,
            ragStatus,
            snps,
            responsibleParty,
            responsiblePartyItem,
            responsiblePartyContacts,
            eutelsatOwners,
            geoLeo,
            comments

        } = maRequirementBasicData;

        const {
            pnpService,
        } = this.props;

        const itemData = {};
        itemData[Consts.FIELDS.COMMON.TITLE] = name;
        itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = type;
        itemData[Consts.FIELDS.MAGENERIC.SUMMARY] = summary;
        itemData[Consts.FIELDS.MAGENERIC.MARKETS_ID] = markets.map((c) => (parseInt(c.key.toString())));
        itemData[Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE_ID] = requirementType;
        itemData[Consts.FIELDS.MAGENERIC.VERTICALS] = maVertical?.map((item) => item.key) || '';
        itemData[Consts.FIELDS.MAGENERIC.SYNTHESIS_ID] = synthesisStatus;
        itemData[Consts.FIELDS.MAGENERIC.RAGSTATUS] = ragStatus;
        itemData[Consts.FIELDS.MAGENERIC.SNPS_ID] = snps.map((c) => (parseInt(c.key.toString())));
        itemData[Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY] = responsibleParty;
        switch (responsibleParty) {
            case MARequirementResponsibleParty.Eutelsat:
                itemData[Consts.FIELDS.MAGENERIC.EUTELSATENTITY_ID] = responsiblePartyItem?.key;
                break;
            case MARequirementResponsibleParty.TP:
                itemData[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER_ID] = responsiblePartyItem?.key;
                break;
            case MARequirementResponsibleParty.DP:
                itemData[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER_ID] = responsiblePartyItem?.key;
                break;
            default:
                break;
        }
        itemData[Consts.FIELDS.MAGENERIC.CONTACTS_ID] = responsiblePartyContacts.map((c) => (parseInt(c.key.toString())));
        itemData[Consts.FIELDS.MAGENERIC.EUTELSATOWNERS_ID] = eutelsatOwners.map((c) => (parseInt(c.key.toString())));
        itemData[Consts.FIELDS.MAGENERIC.GEOLEO] = geoLeo;
        itemData[Consts.FIELDS.MAGENERIC.COMMENTS] = comments;
        //MA Generic start
        itemData[Consts.FIELDS.MAGENERIC.APPLICATIONDATE] = applicationDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.ESTIMATEDDATE] = estimatedDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.EFFECTIVEDATE] = effectiveDate?.toISOString();

        const reletedItemsListUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.REQUIREMENT_RELTIONSHIP_URL);
        const relatedItemIds: number[] = [];
        if (Array.isArray(reletedItems) && reletedItems.length > 0) {
            for (const item of reletedItems) {
                const relatedItemData = {};
                const itemsReletedItemId = item.reletedItem?.map((t) => parseInt(t.key.toString()));
                relatedItemData[Consts.FIELDS.COMMON.TITLE] = item.relationShip;
                relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.MA_REQUIREMENTRELATIONSHIPTYPE] = item.releted;

                switch (item.releted) {
                    // Group 1: Law Firm & Legal Representative
                    case RelatedItemType.LawFirm:
                    case RelatedItemType.LegalRepresentative:
                        relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.LEGALSERVICEPROVIDER_ID] = itemsReletedItemId;
                        break;

                    // Group 2: Organization & Regulator
                    case RelatedItemType.Administration:
                    case RelatedItemType.Organization:
                    case RelatedItemType.Regulator:
                        relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.AUTHORITY_ID] = itemsReletedItemId;
                        break;
                    // Group 3: DP
                    case RelatedItemType.DistributionPartner:
                        relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.DISTRIBUTION_PARTNER_ID] = itemsReletedItemId;

                        break;
                    // Group 4: SNP
                    case RelatedItemType.SNP:
                        relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.SNP_ID] = itemsReletedItemId;
                        break;

                    /* // Group 5: Teleport Partner
                     case strings.FormMARequirementGenericTypeTeleportPartner:
                         relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.TELEPORTPARTNER_ID] = itemsReletedItemId;
                         break;*/

                    // Default: do nothing
                    default:
                        break;
                }

                const createdItem: IItemAddResult = await pnpService
                    .getListByUrl(reletedItemsListUrl)
                    .items.add(relatedItemData);

                relatedItemIds.push(createdItem.data.ID);
            }
            // Store all created item IDs in the multi-lookup field
            itemData[Consts.FIELDS.MAGENERIC.RELATEDITEMS_ID] = relatedItemIds.map((id) => parseInt(id.toString()));
        }


        const list = this.config.list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);

        const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
        const documentSpaceUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.DOCUMENT_SPACE_URL);
        const parentFolderUrl = `${documentSpaceUrl}${Consts.DOCUMENT_SPACE.MAGENERIC_FOLDER_URL}`;
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
            maRequirementBasicData,
            applicationDate,
            estimatedDate,
            effectiveDate,
            reletedItems

        } = this.state;
        const {
            name,
            type,
            summary,
            markets,
            requirementType,
            maVertical,
            synthesisStatus,
            ragStatus,
            snps,
            responsibleParty,
            responsiblePartyItem,
            responsiblePartyContacts,
            eutelsatOwners,
            geoLeo,
            comments
        } = maRequirementBasicData;
        const {
            pnpService,
            itemId
        } = this.props;
        const itemData = {};
        itemData[Consts.FIELDS.COMMON.TITLE] = name;
        itemData[Consts.FIELDS.COMMON.CONTENTTYPE_ID] = type;
        itemData[Consts.FIELDS.MAGENERIC.SUMMARY] = summary;
        itemData[Consts.FIELDS.MAGENERIC.MARKETS_ID] = markets.map((c) => (parseInt(c.key.toString())));
        itemData[Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE_ID] = requirementType;
        itemData[Consts.FIELDS.MAGENERIC.VERTICALS] = maVertical?.map((item) => item.key) || '';
        itemData[Consts.FIELDS.MAGENERIC.SYNTHESIS_ID] = synthesisStatus;
        itemData[Consts.FIELDS.MAGENERIC.RAGSTATUS] = ragStatus;
        itemData[Consts.FIELDS.MAGENERIC.SNPS_ID] = snps.map((c) => (parseInt(c.key.toString())));
        itemData[Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY] = responsibleParty;
        switch (responsibleParty) {
            case MARequirementResponsibleParty.Eutelsat:
                itemData[Consts.FIELDS.MAGENERIC.EUTELSATENTITY_ID] = responsiblePartyItem?.key;
                break;
            case MARequirementResponsibleParty.TP:
                itemData[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER_ID] = responsiblePartyItem?.key;
                break;
            case MARequirementResponsibleParty.DP:
                itemData[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER_ID] = responsiblePartyItem?.key;
                break;
            default:
                break;
        }
        itemData[Consts.FIELDS.MAGENERIC.CONTACTS_ID] = responsiblePartyContacts.map((c) => (parseInt(c.key.toString())));
        itemData[Consts.FIELDS.MAGENERIC.EUTELSATOWNERS_ID] = eutelsatOwners.map((c) => (parseInt(c.key.toString())));
        itemData[Consts.FIELDS.MAGENERIC.GEOLEO] = geoLeo;
        itemData[Consts.FIELDS.MAGENERIC.COMMENTS] = comments;
        //MA Generic start
        itemData[Consts.FIELDS.MAGENERIC.APPLICATIONDATE] = applicationDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.ESTIMATEDDATE] = estimatedDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.EFFECTIVEDATE] = effectiveDate?.toISOString();
        //adding items into releteditems
        const reletedItemsListUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.REQUIREMENT_RELTIONSHIP_URL);
        const relatedItemIds: number[] = [];
        if (Array.isArray(reletedItems) && reletedItems.length > 0) {
            for (const item of reletedItems) {
                const relatedItemData = {};
                const itemsReletedItemId = item.reletedItem?.map((t) => parseInt(t.key.toString()));
                relatedItemData[Consts.FIELDS.COMMON.TITLE] = item.relationShip;
                relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.MA_REQUIREMENTRELATIONSHIPTYPE] = item.releted;
                switch (item.releted) {
                    // Group 1: Law Firm & Legal Representative
                    case RelatedItemType.LawFirm:
                    case RelatedItemType.LegalRepresentative:
                        relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.LEGALSERVICEPROVIDER_ID] = itemsReletedItemId;
                        break;
                    // Group 2: Organization & Regulator
                    case RelatedItemType.Administration:
                    case RelatedItemType.Organization:
                    case RelatedItemType.Regulator:
                        relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.AUTHORITY_ID] = itemsReletedItemId;
                        break;
                    // Group 3: DP
                    case RelatedItemType.DistributionPartner:
                        relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.DISTRIBUTION_PARTNER_ID] = itemsReletedItemId;
                        break;
                    // Group 4: SNP
                    case RelatedItemType.SNP:
                        relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.SNP_ID] = itemsReletedItemId;
                        break;

                    /* // Group 5: Teleport Partner
                     case strings.FormMARequirementGenericTypeTeleportPartner:
                         relatedItemData[Consts.FIELDS.REQUIREMENTRELATIONSHIP_TYPE_FIELDS.TELEPORTPARTNER_ID] = itemsReletedItemId;
                         break;*/

                    // Default: do nothing
                    default:
                        break;
                }
                if (item.Id) {
                    await pnpService
                        .getListByUrl(reletedItemsListUrl)
                        .items.getById(parseInt(item.Id)).update(relatedItemData);
                    relatedItemIds.push(parseInt(item.Id));
                } else {
                    const createdItem: IItemAddResult = await pnpService
                        .getListByUrl(reletedItemsListUrl)
                        .items.add(relatedItemData);
                    relatedItemIds.push(createdItem?.data?.ID);
                }
            }
            // Store all created item IDs in the multi-lookup field
            itemData[Consts.FIELDS.MAGENERIC.RELATEDITEMS_ID] = relatedItemIds;
        }
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
    };
}
