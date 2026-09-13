import * as React from 'react';
import { ILicenseProps, ILicenseState } from './MALicense.types';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import { DatePicker, Dropdown, Icon, IDropdownOption, Label, Spinner, SpinnerSize, TextField, Toggle } from '@fluentui/react';
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
import { RichText } from '../../../../common/components/RichText/RichText';
import RecurrencePatternPicker from '../../../../common/components/Recurrence Pattern Picker/RecurrencePatternPicker';
import { RecurrenceData, RecurrenceFrequency } from '../../../../common/components/Recurrence Pattern Picker/RecurrencePatternPicker.types';
import { Combobox } from 'react-widgets/cjs';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import { ICategorizedDropdownOption, MARequirementBasicData } from '../MA Requirement Basic/MARequirementBasic.types';
import { AuthorityType, ContactType, LicenceRenewalTerm, MARequirementResponsibleParty, MARequirementType, UserRole } from '../../../../common/models/Enums';
import MARequirementBasic from '../MA Requirement Basic/MARequirementBasic';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';
export default class MALicense extends React.Component<ILicenseProps, ILicenseState> {
    private isEditMode: boolean;
    private config: IFormConfig;
    private AllResponsibleItems: Map<string, IDropdownOption[]> = new Map();
    private AllRequestedPartyItems: Map<string, IDropdownOption[]> = new Map();
    private AllContactsItems: Map<string, IDropdownOption[]> = new Map();
    private AllSynthesisStatusItems: Map<string, IDropdownOption[]> = new Map();
    constructor(props: Readonly<ILicenseProps>) {
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
                responsiblePartyContactsChoices: [],
                responsiblePartyContacts: [],
                eutelsatOwners: [],
                eutelsatOwnersChoices: [],
                synthesisStatus: '',
                synthesisStatusChoices: [],
                ragStatus: '',
                ragStatusChoices: [],
            },
            recurrencePattern: {
                frequency: '',
                weekDay: null,
                month: null,
                day: null,
                interval: null,
                endDate: null,
            },
            isFormReady: false,
            errors: {},
            recurrencPatternId: 0,
            applicationDate: null,
            estimatedDate: null,
            effectiveDate: null,
            expirationDate: null,
            requestedPartyChoices: [],
            requestedParty: '',
            requestedPartyItem: null,
            requestedPartyItemChoices: [],
            requestedPartyContacts: [],
            requestedPartyContactsChoices: [],
            reference: '',
            trialDemo: false,
            renewalTerm: '',
            renewalTermChoices: [],
        };
        this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.MALicense);
        this.AllResponsibleItems = new Map();
        this.AllSynthesisStatusItems = new Map();
        this.AllContactsItems = new Map();
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
            this.fetchRequestedPartyItems(),
            this.fetchResposniblePartyItems(),
            this.fetchContactsItems(),
        ]);
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
            this.initSNPsChoices());
        promises.push(
            this.initGeoLeoChoices());
        promises.push(
            this.initRenewalTermChoices());
        await Promise.all(promises);
        this.initDatasheetTypeChoices();
        this.initRequestedPartyChoices();
        this.initRequestedPartyItemChoices();
        this.initEutelsatOwnersChoices();
        this.initResponsiblePartyItemChoices();

        if (this.isEditMode)
            await this.setInitialFormValues();
    }

    private fetchResposniblePartyItems = async () => {
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

    public fetchRequestedPartyItems = async () => {
        const { pnpService } = this.props;
        const typeChoices: IDropdownOption[] = [
            { key: Consts.CONTENT_TYPES.ADMINISTRATION, text: AuthorityType.Administration },
            { key: Consts.CONTENT_TYPES.ORGANIZATION, text: AuthorityType.Organization },
            { key: Consts.CONTENT_TYPES.REGULATOR, text: AuthorityType.Regulator }
        ];
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID
        ];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.AUTHORITY_URL);
        const items = await pnpService.getListItems(listUrl).
            select(...selectedFields)
            .orderBy(Consts.FIELDS.COMMON.TITLE)();
        const requestedPartyItems = items.map((item) => {
            const type = typeChoices
                .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.key) !== -1);
            return {
                key: item[Consts.FIELDS.COMMON.ID] + "",
                text: item[Consts.FIELDS.COMMON.TITLE],
                type
            };
        });
        typeChoices.map((type) => {
            const choices = requestedPartyItems.filter((item) => item.type === type);
            this.AllRequestedPartyItems.set(type.text, choices);
        });
    };

    private fetchContactsItems = async () => {
        const { pnpService } = this.props;
        const selectedFields = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
            `${Consts.FIELDS.CONTACT.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.CONTACT.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.CONTACT.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.CONTACT.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`
        ];
        const configMap = [
            { key: ContactType.TP, contentTypeId: Consts.CONTENT_TYPES.TELEPORT_PARTNER_CONTACT },
            { key: ContactType.DP, contentTypeId: Consts.CONTENT_TYPES.DISTRIBUTION_PARTNER_CONTACT },
            { key: ContactType.Administration, contentTypeId: Consts.CONTENT_TYPES.ADMINISTRATION_CONTACT },
            { key: ContactType.Organization, contentTypeId: Consts.CONTENT_TYPES.ORGANIZATION_CONTACT },
            { key: ContactType.Regulator, contentTypeId: Consts.CONTENT_TYPES.REGULATOR_CONTACT },
            { key: ContactType.Eutelsat, contentTypeId: Consts.CONTENT_TYPES.EUTELSAL_CONTACT },
        ];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.CONTACT_URL);
        const expand = [
            Consts.FIELDS.COMMON.CONTENTTYPE,
            Consts.FIELDS.CONTACT.TELEPORT_PARTNER,
            Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER,
            Consts.FIELDS.CONTACT.AUTHORITY,
        ];
        const items = await pnpService.getListByUrl(listUrl)
            .items
            .select(...selectedFields)
            .expand(...expand)
            .orderBy(Consts.FIELDS.COMMON.TITLE, true)
            .top(5000)();
        const contactchoices = items?.map((item) => {
            const type = configMap
                .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.contentTypeId) !== -1)?.key;
            return {
                key: item[Consts.FIELDS.COMMON.ID] + "",
                text: item[Consts.FIELDS.COMMON.TITLE],
                type,
                TPId: item[Consts.FIELDS.CONTACT.TELEPORT_PARTNER]?.[Consts.FIELDS.COMMON.ID] + "",
                DPId: item[Consts.FIELDS.CONTACT.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.ID] + "",
                AuthorityId: item[Consts.FIELDS.CONTACT.AUTHORITY]?.[Consts.FIELDS.COMMON.ID] + "",
            };
        });

        configMap.map((type) => {
            const choices = contactchoices.filter((item) => item.type === type.key);
            this.AllContactsItems.set(type.key, choices);
        });

    };

    private initEditModeRecurrence = async (itemId: number): Promise<RecurrenceData | null> => {
        if (!this.isEditMode || !itemId) return null;
        const { pnpService } = this.props;
        const list = Consts.LISTS.RECURRENCE_PATTERN_URL;
        if (!list) return null;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.RECURRENCE_PATTERN.FREQUENCY,
            Consts.FIELDS.RECURRENCE_PATTERN.WEEKDAY,
            Consts.FIELDS.RECURRENCE_PATTERN.MONTH,
            Consts.FIELDS.RECURRENCE_PATTERN.DAY,
            Consts.FIELDS.RECURRENCE_PATTERN.INTERVAL,
            Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE
        ];
        try {
            const item = await pnpService.getListItems(listUrl)
                .getById(itemId)
                .select(...selectedFields)();
            const recurrencePattern: RecurrenceData = {
                frequency: item[Consts.FIELDS.RECURRENCE_PATTERN.FREQUENCY],
                weekDay: item[Consts.FIELDS.RECURRENCE_PATTERN.WEEKDAY] ?? null,
                month: item[Consts.FIELDS.RECURRENCE_PATTERN.MONTH] ?? null,
                day: item[Consts.FIELDS.RECURRENCE_PATTERN.DAY] ?? null,
                interval: item[Consts.FIELDS.RECURRENCE_PATTERN.INTERVAL] ?? null,
                endDate: item[Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE] ? new Date(item[Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE]) : null
            };
            return recurrencePattern;
        } catch (error) {
            console.error("Error fetching recurrence pattern:", error);
            return null;
        }
    };

    public setInitialFormValues = async () => {
        const { pnpService, itemId } = this.props;
        const list = Form.config.find((f) => f.type.toLowerCase() === FormType.MALicense).list;
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
            `${Consts.FIELDS.MAGENERIC.RECURRENCE_PATTERN}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.RECURRENCE_PATTERN}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_TYPE,
            `${Consts.FIELDS.MAGENERIC.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.REFERENCE,
            Consts.FIELDS.MAGENERIC.TRIALDEMO,
            Consts.FIELDS.MAGENERIC.RENEWALTERM,
            Consts.FIELDS.MAGENERIC.EXPIRATIONDATE

        ];
        const expand = [Consts.FIELDS.MAGENERIC.MARKETS, Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE,
        Consts.FIELDS.MAGENERIC.SYNTHESIS, Consts.FIELDS.MAGENERIC.SNPS, Consts.FIELDS.MAGENERIC.EUTELSATENTITY, Consts.FIELDS.MAGENERIC.TELEPORTPARTNER,
        Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER, Consts.FIELDS.MAGENERIC.CONTACTS, Consts.FIELDS.MAGENERIC.EUTELSATOWNERS, Consts.FIELDS.MAGENERIC.RECURRENCE_PATTERN,
        Consts.FIELDS.MAGENERIC.AUTHORITY, Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT
        ];
        const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
        const filterTypeChoices = "" + this.state.maRequirementBasicData.typeChoices.find((type) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID].indexOf(type.key) !== -1)?.key;
        const maVertical = item[Consts.FIELDS.MAGENERIC.VERTICALS]?.map((v) => ({
            key: v,
            text: v
        }));
        const responsibleParty = item[Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY];
        let responsiblePartyItem = null;
        let filteredContacts: IDropdownOption[] = [];
        const contactAllData = this.AllContactsItems.get(responsibleParty);
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
                filteredContacts = contactAllData?.filter((item) => item["TPId"] === responsiblePartyItem.key.toString());
                break;
            case MARequirementResponsibleParty.DP:
                responsiblePartyItem = {
                    key: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
                    text: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE]
                };
                filteredContacts = contactAllData?.filter((item) => item["DPId"] === responsiblePartyItem.key.toString());
                break;
            default:
                break;
        }

        const responsiblePartyItemChoices: IDropdownOption[] = this.AllResponsibleItems.get(responsibleParty);
        const snps = item[Consts.FIELDS.MAGENERIC.SNPS]?.map((snp) => ({
            key: snp[Consts.FIELDS.COMMON.ID],
            text: snp[Consts.FIELDS.COMMON.TITLE]
        }));
        const requirementType = item[Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE]?.[Consts.FIELDS.COMMON.ID] + "";
        let synthesisStatusChoices: IDropdownOption[] = [];
        const synthesisArray: IDropdownOption[] = this.AllSynthesisStatusItems.get(requirementType);
        if (synthesisArray) {
            synthesisStatusChoices = synthesisArray;
        }
        const rawApplicationDate = item[Consts.FIELDS.MAGENERIC.APPLICATIONDATE];
        const applicationDate: Date | null = rawApplicationDate ? new Date(rawApplicationDate) : null;
        const rawEstimatedDate = item[Consts.FIELDS.MAGENERIC.ESTIMATEDDATE];
        const estimatedDate: Date | null = rawEstimatedDate ? new Date(rawEstimatedDate) : null;
        const rawEffectiveDate = item[Consts.FIELDS.MAGENERIC.EFFECTIVEDATE];
        const effectiveDate: Date | null = rawEffectiveDate ? new Date(rawEffectiveDate) : null;
        const recurrencePatternFieldId = item[`${Consts.FIELDS.MAGENERIC.RECURRENCE_PATTERN}`]?.[Consts.FIELDS.COMMON.ID];
        const recurrencePattern = await this.initEditModeRecurrence(recurrencePatternFieldId);
        let recurrencPatternId;
        if (item[Consts.FIELDS.MAGENERIC.RECURRENCE_PATTERN]) {
            recurrencPatternId = item[Consts.FIELDS.MAGENERIC.RECURRENCE_PATTERN][Consts.FIELDS.COMMON.ID];
        }
        const requestedParty = this.state.requestedPartyChoices.find(
            (choice) => choice.text === item[Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_TYPE]
        )?.text as string || '';
        // Read values from authority lookup field
        const authorityLookup = item[Consts.FIELDS.MAGENERIC.AUTHORITY];
        const authorityValue = authorityLookup ? {
            key: authorityLookup[Consts.FIELDS.COMMON.ID],
            text: authorityLookup[Consts.FIELDS.COMMON.TITLE]
        } : null;
        const requestedPartyContacts =
            item[Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT] ? item[Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT].map((contact) => ({
                key: contact[Consts.FIELDS.COMMON.ID],
                text: contact[Consts.FIELDS.COMMON.TITLE]
            })) : [];
        const requestedPartyContactsChoices = this.AllContactsItems.get(requestedParty).filter((contact) => contact['AuthorityId'] === authorityValue.key.toString());
        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            name: item[Consts.FIELDS.COMMON.TITLE],
            type: filterTypeChoices,
            summary: item[Consts.FIELDS.MAGENERIC.SUMMARY],
            markets: item[Consts.FIELDS.MAGENERIC.MARKETS]?.map((market) => ({
                key: market[Consts.FIELDS.COMMON.ID],
                text: market[Consts.FIELDS.COMMON.TITLE]
            })) || [],
            requirementType,
            synthesisStatus: item[Consts.FIELDS.MAGENERIC.SYNTHESIS]?.[Consts.FIELDS.COMMON.ID] + "",
            synthesisStatusChoices,
            maVertical,
            ragStatus: item[Consts.FIELDS.MAGENERIC.RAGSTATUS],
            responsibleParty,
            responsiblePartyItem,
            responsiblePartyItemChoices,
            responsiblePartyContacts: item[Consts.FIELDS.MAGENERIC.CONTACTS]?.map((contacts) => ({
                key: contacts[Consts.FIELDS.COMMON.ID],
                text: contacts[Consts.FIELDS.COMMON.TITLE]
            })) || [],
            responsiblePartyContactsChoices: filteredContacts,
            eutelsatOwners: item[Consts.FIELDS.MAGENERIC.EUTELSATOWNERS]?.map((owners) => ({
                key: owners[Consts.FIELDS.COMMON.ID],
                text: owners[Consts.FIELDS.COMMON.TITLE]
            })) || [],
            geoLeo: item[Consts.FIELDS.MAGENERIC.GEOLEO],
            comments: item[Consts.FIELDS.MAGENERIC.COMMENTS],
            snps
        };
        const reference = item[Consts.FIELDS.MAGENERIC.REFERENCE];
        const trialDemo = item[Consts.FIELDS.MAGENERIC.TRIALDEMO];
        const renewalTerm = item[Consts.FIELDS.MAGENERIC.RENEWALTERM];
        const rawExpirationDate = item[Consts.FIELDS.MAGENERIC.EXPIRATIONDATE];
        const expirationDate: Date | null = rawExpirationDate ? new Date(rawExpirationDate) : null;
        const initialValues = {
            maRequirementBasicData,
            applicationDate,
            estimatedDate,
            effectiveDate,
            recurrencePattern,
            requestedParty,
            requestedPartyItem: authorityValue,
            requestedPartyContacts,
            requestedPartyContactsChoices,
            reference,
            trialDemo,
            renewalTerm,
            expirationDate,
            recurrencPatternId
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

    private initDatasheetTypeChoices = async (): Promise<void> => {
        const typeChoices: IDropdownOption[] = [
            { key: Consts.CONTENT_TYPES.MA_RequirementCredential, text: MARequirementType.Credential },
        ];
        const type: string = Consts.CONTENT_TYPES.MA_RequirementCredential;
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
        const eutelsatOwnersChoices: IDropdownOption[] = [];
        this.AllContactsItems.get(ContactType.Eutelsat).forEach((item) => {
            eutelsatOwnersChoices.push({ key: item.key, text: item.text });
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

    public initResponsiblePartyChoices = async (): Promise<void> => {
        const {
            pnpService
        } = this.props;
        let responsiblePartyChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY)
            .select('Choices')();
        responsiblePartyChoices = data.Choices?.map((choice) => ({ key: choice, text: choice }));
        if (!this.isEditMode) {
            const responsibleParty = responsiblePartyChoices.length > 0 ? responsiblePartyChoices[0].key as string : '';
            const maRequirementBasicData = {
                ...this.state.maRequirementBasicData,
                responsiblePartyChoices,
                responsibleParty
            };
            this.setState(() => ({ maRequirementBasicData }));
        } else {
            const maRequirementBasicData = {
                ...this.state.maRequirementBasicData,
                responsiblePartyChoices
            };
            this.setState(() => ({ maRequirementBasicData }));
        }
    };


    public initResponsiblePartyItemChoices = () => {
        const {
            responsibleParty
        } = this.state.maRequirementBasicData;
        const responsiblePartyItemChoices: IDropdownOption[] = this.AllResponsibleItems.get(responsibleParty);
        const maRequirementBasicData = {
            ...this.state.maRequirementBasicData,
            responsiblePartyItemChoices
        };
        this.setState(() => ({ maRequirementBasicData }));
    };

    private initRequestedPartyItemChoices = () => {
        const {
            requestedParty
        } = this.state;
        const requestedPartyItemChoices = this.AllRequestedPartyItems.get(requestedParty)?.map((item) => ({
            key: item.key,
            text: item.text
        })) || [];
        this.setState({ requestedPartyItemChoices });
    };

    private initRequestedPartyChoices = async () => {
        const requestedPartyChoices: IDropdownOption[] = [
            { key: Consts.CONTENT_TYPES.REGULATOR, text: ContactType.Regulator },
            { key: Consts.CONTENT_TYPES.ADMINISTRATION, text: ContactType.Administration },
            { key: Consts.CONTENT_TYPES.ORGANIZATION, text: ContactType.Organization },
        ];
        if (!this.isEditMode) {
            const requestedParty = requestedPartyChoices.length > 0 ? requestedPartyChoices[0].text as string : '';
            this.setState({
                requestedPartyChoices, requestedParty
            });
        } else {
            this.setState({
                requestedPartyChoices
            });
        }
    };

    public initRenewalTermChoices = async (): Promise<void> => {
        const {
            pnpService
        } = this.props;
        let renewalTermChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MAGENERIC.RENEWALTERM)
            .select('Choices')();
        renewalTermChoices = data.Choices?.map((choice) => ({ key: choice, text: choice }));
        const renewalTerm = renewalTermChoices.length > 0 ? renewalTermChoices[0].key as string : '';
        this.setState({
            renewalTermChoices,
            renewalTerm
        });
    };

    private validateForm = (): boolean => {
        const errors: { [key: string]: string } = {};
        const {
            maRequirementBasicData,
            requestedPartyItem
        } = this.state;
        const {
            name,
            markets,
            requirementType,
            maVertical,
            synthesisStatus,
            responsiblePartyItem,
            eutelsatOwners
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
        if (!requestedPartyItem || !requestedPartyItem?.key || requestedPartyItem?.key === '') {
            errors['requestedPartyItem'] = strings.FormMARequirementErrorRequestedPartyItemRequired;
        }
        this.setState({ errors });
        return Object.keys(errors).length === 0;
    };

    public render(): React.ReactElement<ILicenseProps> {
        const {
            estimatedDate,
            applicationDate,
            effectiveDate,
            maRequirementBasicData,
            errors,
            isFormReady,
            recurrencePattern,
            requestedParty,
            requestedPartyChoices,
            requestedPartyContacts,
            requestedPartyContactsChoices,
            requestedPartyItem,
            requestedPartyItemChoices,
            reference,
            trialDemo,
            renewalTerm,
            renewalTermChoices,
            expirationDate
        } = this.state;
        const requestedPartyKey = requestedPartyChoices.find(
            (choice) => choice.text === requestedParty
        )?.key;
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
                                <form>
                                    <MARequirementBasic
                                        maRequirementBasicData={maRequirementBasicData} onChange={this.onChangeMARequiremntBasic} errors={errors}></MARequirementBasic>
                                    {/* Authority Type */}
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMARequirementRequestedPartyTypeLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementTooltipRequestedPartyType} ></Tooltip>
                                        </div>
                                        <Dropdown
                                            selectedKey={requestedPartyKey} disabled={this.isEditMode}
                                            onChange={this.onChangeRequestedParty}
                                            options={requestedPartyChoices as IDropdownOption[]} />
                                    </div>

                                    {/* Authority */}

                                    <div className={styles.field} >
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label required>{strings.FormMARequirementRequestedPartyItemLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementTooltipRequestedPartyItem} ></Tooltip>
                                        </div>
                                        <Combobox
                                            data={requestedPartyItemChoices}
                                            value={requestedPartyItem}
                                            disabled={this.isEditMode}
                                            textField="text"
                                            filter="contains"
                                            onChange={this.onChangeRequestedPartyItem}
                                            selectIcon={
                                                <span className="ms-Dropdown-caretDownWrapper">
                                                  <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
                                                </span>
                                              }
                                        />
                                        {errors.requestedPartyItem &&
                                            <div className={styles.errorContainer}>
                                                <Icon iconName="Error" className={styles.errorIcon} />
                                                <span className={styles.errorMessage}>
                                                    {errors.requestedPartyItem}
                                                </span>
                                            </div>
                                        }
                                    </div>
                                    {/*Authority Contacts */}
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMARequirementRequestedPartyContactsLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementTooltipRequestedPartyContacts} ></Tooltip>
                                        </div>
                                        <MultiselectWrapper
                                            value={requestedPartyContacts}
                                            data={requestedPartyContactsChoices}
                                            dataKey={(item: IDropdownOption) => item.key}
                                            textField={(item: IDropdownOption) => item.text}
                                            filter="contains"
                                            onChange={this.onChangeRequestedPartyContacts} />
                                        {errors.requestedPartyContacts &&
                                            <div className={styles.errorContainer}>
                                                <Icon iconName="Error" className={styles.errorIcon} />
                                                <span className={styles.errorMessage}>
                                                    {errors.requestedPartyContacts}
                                                </span>
                                            </div>
                                        }
                                    </div>
                                    {/* Reference */}
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label >{strings.FormMARequirementReferenceLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementTooltipReference} ></Tooltip>
                                        </div>
                                        <TextField
                                            value={reference}
                                            onChange={this.onChangeReference}
                                        />

                                    </div>
                                    {/* Date row */}
                                    {/* Application Date */}
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label >{strings.FormMARequirementApplicationDateLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementTooltipApplicationDate} ></Tooltip>
                                        </div>

                                        <DatePicker
                                            onSelectDate={this.onchangeApplicationDate}
                                            value={applicationDate}
                                            formatDate={this.formatDate}
                                            allowTextInput
                                            styles={datepickerstyles}
                                        />
                                    </div>

                                    {/* Estimated Date */}
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label >{strings.FormMARequirementEstimatedDateLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementTooltipEstimatedDate} ></Tooltip>
                                        </div>
                                        <DatePicker
                                            onSelectDate={this.onchangeEstimatedDate}
                                            value={estimatedDate}
                                            formatDate={this.formatDate}
                                            allowTextInput
                                            styles={datepickerstyles}
                                        />
                                    </div>

                                    {/* Trial/Demo */}
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMARequirementTrialDemoLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementTooltipTrialDemo} ></Tooltip>
                                        </div>
                                        <Toggle
                                            checked={trialDemo}
                                            onChange={this.onChangetrialDemo}
                                        />
                                    </div>
                                    {/* Renewal Term */}
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMARequirementRenewalTermLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementTooltipRenewalTerm} ></Tooltip>
                                        </div>
                                        <Dropdown
                                            selectedKey={renewalTerm}
                                            onChange={this.onChangeRenewalTerm}
                                            options={renewalTermChoices as IDropdownOption[]} />
                                    </div>
                                    {/* Date row */}
                                    {/* Effective Date */}
                                    <div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMARequirementEffectiveDateLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementTooltipEffectiveDate} ></Tooltip>
                                        </div>
                                        <DatePicker
                                            onSelectDate={this.onchangeEffectiveDate}
                                            value={effectiveDate}
                                            formatDate={this.formatDate}
                                            allowTextInput
                                            styles={datepickerstyles}
                                        />
                                    </div>


                                    {/* Expiration Date */}
                                    {renewalTerm === LicenceRenewalTerm.Fixed && (<div className={styles.field}>
                                        <div className={styles.fieldLabel}>
                                            <div className={styles.fieldLabelContainer}>
                                                <Label>{strings.FormMARequirementExpirationDateLabel}</Label>
                                            </div>
                                            <Tooltip content={strings.FormMARequirementTooltipExpirationDate} ></Tooltip>
                                        </div>
                                        <DatePicker
                                            onSelectDate={this.onchangeExpirationDate}
                                            value={expirationDate}
                                            formatDate={this.formatDate}
                                            allowTextInput
                                            styles={datepickerstyles}
                                        />
                                    </div>)}

                                    {/* Recurance */}
                                    {renewalTerm === LicenceRenewalTerm.Recurrent && (
                                        <RecurrencePatternPicker
                                            recurrenceData={recurrencePattern}
                                            onChange={this.onChangeRecurrencePattern}
                                            availableFrequencies={[RecurrenceFrequency.Monthly, RecurrenceFrequency.Annual, RecurrenceFrequency.MultiAnnual]}
                                        >
                                        </RecurrencePatternPicker>
                                    )}
                                </form>
                            }
                        </>
                    </FormWizard>}
                </div>
            </div >
        );
    }


    private onChangeRecurrencePattern = (recurrencePattern: RecurrenceData) => {
        this.setState({ recurrencePattern });
    };

    private onChangeMARequiremntBasic = (maRequirementBasicData: MARequirementBasicData) => {

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
            }
            maRequirementBasicData.responsiblePartyContactsChoices = filteredContacts;
        }
        this.setState({ maRequirementBasicData });
    };

    private onChangeReference = (event, reference: string) => {
        this.setState({ reference });
    };

    private onChangetrialDemo = (event, trialDemo: boolean) => {
        this.setState({ trialDemo });
    };

    private onChangeRenewalTerm = (event, option: IDropdownOption): void => {
        if (option) {
            this.setState({ renewalTerm: option.text });
            if (option.text !== LicenceRenewalTerm.Fixed) {
                this.setState({ expirationDate: null });
            }
        }
    };

    private formatDate = (date?: Date): string => {
        if (!date) return '';
        return UtilHelper.formatDate(date, 'LL', 'en-us');
    };

    private formatRecurrencePattern = (recurrencePattern: RecurrenceData): string => {
        const frequency = recurrencePattern.frequency || '';
        const weekDay = recurrencePattern.weekDay || '';
        const month = recurrencePattern.month || '';
        const day = recurrencePattern.day || 1;
        const interval = recurrencePattern.interval || 1;
        const endDate = recurrencePattern.endDate || null;

        let result = '';

        switch (frequency) {
            case 'Weekly':
                result = `${weekDay} of every ${interval} week${interval > 1 ? 's' : ''}`;
                break;
            case 'Monthly':
                result = `Day ${day} of every ${interval} month${interval > 1 ? 's' : ''}`;
                break;
            case 'Quarterly':
                result = `Day ${day} of every 3 months`;
                break;
            case 'Half Annual':
                result = `Day ${day} of every 6 months`;
                break;
            case 'Annual':
                result = `Every ${month} ${day} of every 1 year`;
                break;
            case 'Multi Annual':
                result = `Every ${month} ${day} of every ${interval} year${interval > 1 ? 's' : ''}`;
                break;
            default:
                result = frequency;
        }

        if (frequency && endDate) {
            const date = UtilHelper.formatDate(endDate, 'LL', 'en-us');
            result += ` and expired on ${date} `;
        }

        return result;
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

    private onchangeExpirationDate = (expirationDate: Date) => {
        this.setState({ expirationDate });
    };

    private getRequestedPartyItemsByType = (contentTypeText: string) => {
        if (this.AllRequestedPartyItems.has(contentTypeText)) {
            const cachedItems = this.AllRequestedPartyItems.get(contentTypeText) || [];
            return cachedItems;
        }
    };

    private onChangeRequestedParty = async (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
        const requestedParty = option.text.toString();
        this.setState({ requestedParty });
        const typeChoices = this.getRequestedPartyItemsByType(option.text.toString());
        this.setState({ requestedPartyItemChoices: typeChoices, requestedPartyItem: null, requestedPartyContacts: [] });
    };

    private onChangeRequestedPartyContacts = (requestedPartyContacts: IDropdownOption[]): void => {
        this.setState({ requestedPartyContacts });
    };

    private onChangeRequestedPartyItem = async (option: IDropdownOption) => {
        const {
            requestedParty,
        } = this.state;
        const contacts = this.AllContactsItems.get(requestedParty).filter((contact) => contact['AuthorityId'] === option.key.toString());

        this.setState({ requestedPartyItem: option, requestedPartyContactsChoices: contacts, requestedPartyContacts: [] });
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
            recurrencePattern,
            reference,
            renewalTerm,
            requestedParty,
            requestedPartyContacts,
            requestedPartyItem,
            expirationDate,
            trialDemo,
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
        const requirementTypeTextValue = requirementTypeChoices.find((choice) => choice.key === +requirementType)?.text;
        const marketsValues = markets?.map((c) => c.text).join(", ");
        const maVerticalValues = maVertical?.map((c) => c.text).join(", ");
        const synthesisTextValue = synthesisStatusChoices.find((choice) => choice.key === synthesisStatus)?.text;
        const snpsValues = snps?.map((c) => c.text).join(", ");
        const contactsValues = responsiblePartyContacts?.map((c) => c.text).join(", ");
        const eutelsatOwnersValues = eutelsatOwners?.map((c) => c.text).join(", ");
        const geoLeoTextValue = geoLeoChoices.find((choice) => choice.key === geoLeo)?.text;
        const requestedPartyContactsValues = requestedPartyContacts?.map((c) => c.text).join(", ");
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
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelSNPS}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {snpsValues}
                    </div>
                </div>
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

                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementLabelContacts}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {contactsValues}
                    </div>
                </div>

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
                            <Label>{strings.FormMARequirementRequestedPartyTypeLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {requestedParty}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementRequestedPartyItemLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {requestedParty && requestedPartyItem?.text}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementRequestedPartyContactsLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {requestedPartyContactsValues}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementReferenceLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {reference}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementTrialDemoLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {trialDemo + ""}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementRenewalTermLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {renewalTerm}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementApplicationDateLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>

                        {applicationDate ? this.formatDate(applicationDate) : ''}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementEstimatedDateLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>

                        {estimatedDate ? this.formatDate(estimatedDate) : ''}
                    </div>
                </div>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementEffectiveDateLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {effectiveDate ? this.formatDate(effectiveDate) : ''}
                    </div>
                </div>
                {renewalTerm === LicenceRenewalTerm.Fixed && (<div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMARequirementExpirationDateLabel}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {expirationDate ? this.formatDate(expirationDate) : ''}
                    </div>
                </div>)}
                {renewalTerm === LicenceRenewalTerm.Recurrent && recurrencePattern && (<>
                    <div className={styles.field}>
                        <div className={styles.fieldLabel}>
                            <div className={styles.fieldLabelContainer}>
                                <Label>Repeat Pattern</Label>
                            </div>
                        </div>
                        <div className={styles.fieldValue}>
                            {this.formatRecurrencePattern(recurrencePattern)}
                        </div>
                    </div>

                    {recurrencePattern.month && <div className={styles.field}>
                        <div className={styles.fieldLabel}>
                            <div className={styles.fieldLabelContainer}>
                                <Label>Month:</Label>
                            </div>
                        </div>
                        <div className={styles.fieldValue}>
                            {recurrencePattern.month ?? ''}
                        </div>
                    </div>}

                    {recurrencePattern.day && <div className={styles.field}>
                        <div className={styles.fieldLabel}>
                            <div className={styles.fieldLabelContainer}>
                                <Label>Day:</Label>
                            </div>
                        </div>
                        <div className={styles.fieldValue}>
                            {recurrencePattern.day ?? ''}
                        </div>
                    </div>}
                </>)}
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
            recurrencePattern,
            reference,
            renewalTerm,
            requestedParty,
            requestedPartyContacts,
            requestedPartyItem,
            expirationDate,
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
        if (maVertical.length > 0) {
            itemData[Consts.FIELDS.MAGENERIC.VERTICALS] = maVertical?.map((item) => item.key) || '';
        }
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
        //License
        itemData[Consts.FIELDS.MAGENERIC.APPLICATIONDATE] = applicationDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.ESTIMATEDDATE] = estimatedDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.EFFECTIVEDATE] = effectiveDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.REFERENCE] = reference;
        itemData[Consts.FIELDS.MAGENERIC.RENEWALTERM] = renewalTerm;
        itemData[Consts.FIELDS.MAGENERIC.EXPIRATIONDATE] = expirationDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_TYPE] = requestedParty;
        itemData[Consts.FIELDS.MAGENERIC.AUTHORITY_ID] = requestedPartyItem?.key;
        itemData[Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT_ID] = requestedPartyContacts.map((c) => (parseInt(c.key.toString())));
        //adding item into recurrence list
        const recurrenceListUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.RECURRENCE_PATTERN_URL);
        let recurrenceItemId;
        if (renewalTerm === LicenceRenewalTerm.Recurrent) {
            const recurrenceItemData = {
                [Consts.FIELDS.COMMON.TITLE]: 'Recurrence Entry',
                [Consts.FIELDS.RECURRENCE_PATTERN.FREQUENCY]: recurrencePattern.frequency,
                [Consts.FIELDS.RECURRENCE_PATTERN.WEEKDAY]: recurrencePattern.weekDay,
                [Consts.FIELDS.RECURRENCE_PATTERN.MONTH]: recurrencePattern.month,
                [Consts.FIELDS.RECURRENCE_PATTERN.DAY]: recurrencePattern.day,
                [Consts.FIELDS.RECURRENCE_PATTERN.INTERVAL]: recurrencePattern.interval,
                [Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE]: recurrencePattern.endDate?.toISOString(),
            };
            const createdRecurrenceItem: IItemAddResult = await pnpService
                .getListByUrl(recurrenceListUrl)
                .items.add(recurrenceItemData);

            recurrenceItemId = createdRecurrenceItem.data.ID;
            itemData[Consts.FIELDS.MAGENERIC.RECURRENCE_PATTERN_ID] = recurrenceItemId;
        }
        const list = this.config.list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const createdItem: IItemAddResult = await pnpService.getListItems(listUrl).add(itemData);
        if (recurrenceItemId) {
            await pnpService
                .getListByUrl(recurrenceListUrl)
                .items.getById(recurrenceItemId)
                .update({
                    Title: `${this.config.type}-${createdItem.data.ID}`,
                });
        }
        const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
        const documentSpaceUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.DOCUMENT_SPACE_URL);
        const parentFolderUrl = `${documentSpaceUrl}${Consts.DOCUMENT_SPACE.LICENCE_REQUIREMENTS_FOLDER_URL}`;
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
            maRequirementBasicData,
            applicationDate,
            estimatedDate,
            effectiveDate,
            recurrencePattern,
            reference,
            renewalTerm,
            requestedParty,
            requestedPartyContacts,
            requestedPartyItem,
            expirationDate,
            recurrencPatternId,
            trialDemo
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
        if (maVertical.length > 0) {
            itemData[Consts.FIELDS.MAGENERIC.VERTICALS] = maVertical?.map((item) => item.key) || '';
        }
        itemData[Consts.FIELDS.MAGENERIC.SYNTHESIS_ID] = synthesisStatus;
        itemData[Consts.FIELDS.MAGENERIC.RAGSTATUS] = ragStatus;
        itemData[Consts.FIELDS.MAGENERIC.SNPS_ID] = snps.map((c) => (parseInt(c.key.toString())));
        itemData[Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY] = responsibleParty;
        itemData[Consts.FIELDS.MAGENERIC.TRIALDEMO] = trialDemo;
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
        itemData[Consts.FIELDS.MAGENERIC.APPLICATIONDATE] = applicationDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.ESTIMATEDDATE] = estimatedDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.EFFECTIVEDATE] = effectiveDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.REFERENCE] = reference;
        itemData[Consts.FIELDS.MAGENERIC.RENEWALTERM] = renewalTerm;
        itemData[Consts.FIELDS.MAGENERIC.EXPIRATIONDATE] = expirationDate?.toISOString();
        itemData[Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_TYPE] = requestedParty;
        itemData[Consts.FIELDS.MAGENERIC.AUTHORITY_ID] = requestedPartyItem?.key;
        itemData[Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT_ID] = requestedPartyContacts.map((c) => (parseInt(c.key.toString())));
        // Recurrence update
        const recurrenceListUrl = UrlHelper.getListWebRelativeUrl(
            pnpService.getUrl(),
            Consts.LISTS.RECURRENCE_PATTERN_URL
        );
        const recurrenceItemData = {
            [Consts.FIELDS.COMMON.TITLE]: 'Recurrence Entry',
            [Consts.FIELDS.RECURRENCE_PATTERN.FREQUENCY]: recurrencePattern?.frequency,
            [Consts.FIELDS.RECURRENCE_PATTERN.WEEKDAY]: recurrencePattern?.weekDay,
            [Consts.FIELDS.RECURRENCE_PATTERN.MONTH]: recurrencePattern?.month,
            [Consts.FIELDS.RECURRENCE_PATTERN.DAY]: recurrencePattern?.day,
            [Consts.FIELDS.RECURRENCE_PATTERN.INTERVAL]: recurrencePattern?.interval,
            [Consts.FIELDS.RECURRENCE_PATTERN.ENDDATE]: recurrencePattern?.endDate?.toISOString(),
        };
        if (renewalTerm === LicenceRenewalTerm.Recurrent && recurrencPatternId) {
            await pnpService
                .getListByUrl(recurrenceListUrl)
                .items.getById(recurrencPatternId)
                .update(recurrenceItemData);
        } else if (renewalTerm === LicenceRenewalTerm.Recurrent && !recurrencPatternId) {
            const createdRecurrenceItem: IItemAddResult = await pnpService
                .getListByUrl(recurrenceListUrl)
                .items.add(recurrenceItemData);
            const recurrenceItemId = createdRecurrenceItem.data.ID;
            itemData[Consts.FIELDS.FEE.RECURRENCEPATTERN_ID] = recurrenceItemId;
        }
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
    };
}




