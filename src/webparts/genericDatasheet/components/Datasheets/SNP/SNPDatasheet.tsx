/* eslint-disable max-len */
import * as React from 'react';
import { ISNPDatasheetProps, ISNPDatasheetState } from './SNP.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import * as commonStrings from 'AppCustomizerStrings';
import {
    ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, IDropdownOption, ITooltipHostStyles, Label,
    LayerHost, Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost
} from '@fluentui/react';
import SNP from '../../../../genericForm/components/SNP/SNP';
import { Datasheet, DatasheetType, IDatasheetConfig } from '../../Datasheet.types';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { IBaseItem, IContact, IContentType, IFEE, IItemOption, IMARequirement, ISNP, ISNPReadiness } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { renderArrayPills, renderObjectPills, StatusPill } from '../../Datasheet.utility';
import { RichText } from '../../../../../common/components/RichText/RichText';
import { ViewType } from '../../../../genericView/components/View.types';
import MARequirementView from '../../../../genericView/components/MARequirement View/MARequirementView';
import { SortDirection } from '@pnp/sp/search';
import { FeeChargedBy, FeeType, MarketReadinessResponsibleParty, SNPType, UserRole } from '../../../../../common/models/Enums';
import SNPReadinessView from '../../Datasheet Views/SNPReadinessDatasheetView';
import LibraryView from '../../../../genericView/components/Library View/LibraryView';
import FeeDatasheetView from '../../Datasheet Views/FeeDatasheetView';
import { FormType, IMessageBanner } from '../../../../genericForm/components/Form.types';
import AccessDeniedMessage from '../../../../../common/components/Access Denied/AccessDenied';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';
import DomHelper from '../../../../../common/helpers/DomHelper';

export default class SNPDatasheet extends React.Component<
    ISNPDatasheetProps,
    ISNPDatasheetState
> {
    private config: IDatasheetConfig;
    private snpTypeChoices: { key: string; text: string }[] = [
        { key: Consts.CONTENT_TYPES.STARGATE, text: SNPType.Stargate },
        { key: Consts.CONTENT_TYPES.EUTELSAT_SNP, text: SNPType.EutelsatOwned },
        { key: Consts.CONTENT_TYPES.PARTNER_SNP, text: SNPType.Other }
    ];
    private feeTypeChoices: { key: string; text: string }[] = [
        { key: Consts.CONTENT_TYPES.SNP_FEE, text: FeeType.SNP },
        { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_FEE, text: FeeType.MATypeRequirement }
    ];
    private ChargeByItemChoices: IItemOption[] = [];

    private DPandTPDataInfo: IItemOption[] = [];
    constructor(props: Readonly<ISNPDatasheetProps>) {
        super(props);
        this.state = {
            isDatasheetReady: false,
            showEditFormDialog: false,
            showDeleteDialog: false,
            item: null,
            feeViewItems: null,
            maRequirements: [],
            snpReadinessItems: [],
            errorMessage: null,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'error',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        };
        this.config = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.SNP);
    }
    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };
    public async componentDidMount() {
        await this.init();
        this.setState({ isDatasheetReady: true });
    }
    public initItem = async () => {
        const { pnpService, itemId } = this.props;
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.COMMON.CREATION_TIME,
            Consts.FIELDS.COMMON.MODIFICATION_TIME,
            Consts.FIELDS.SNP.SUMMARY,
            Consts.FIELDS.SNP.CITY,
            Consts.FIELDS.SNP.COMMENTS,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
            Consts.FIELDS.SNP.STATUS,
            Consts.FIELDS.COMMON.DOCUMENTS_SPACE,
            `${Consts.FIELDS.SNP.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.SNP.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.SNP.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.SNP.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.SNP.EUTELSATENTITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.SNP.EUTELSATENTITY}/${Consts.FIELDS.COMMON.TITLE}`

        ];
        const expand = [Consts.FIELDS.SNP.COUNTRY, Consts.FIELDS.SNP.TELEPORTPARTNER, Consts.FIELDS.SNP.EUTELSATENTITY];
        const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.SNP).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        try {
            const item = await this.props.pnpService.getListItems(listUrl).
                getById(itemId).
                select(...selectedFields)
                .expand(...expand)();
            const ContentType: IContentType = this.snpTypeChoices
                .map((type) => ({ Id: type.key + "", Name: type.text }))
                .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1) || { Id: '', Name: '' };
            let TeleportPartner;
            if (item[Consts.FIELDS.SNP.TELEPORTPARTNER]) {
                TeleportPartner = {
                    Id: item[Consts.FIELDS.SNP.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.SNP.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE]
                };
            } else {
                TeleportPartner = null;
            }
            let Country;
            if (item[Consts.FIELDS.SNP.COUNTRY]) {
                Country = {
                    Id: item[Consts.FIELDS.SNP.COUNTRY][Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.SNP.COUNTRY][Consts.FIELDS.COMMON.TITLE]
                };
            } else {
                Country = null;
            }
            let EutEntity;
            if (item[Consts.FIELDS.SNP.EUTELSATENTITY]) {
                EutEntity = {
                    Id: item[Consts.FIELDS.SNP.EUTELSATENTITY][Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.SNP.EUTELSATENTITY][Consts.FIELDS.COMMON.TITLE]
                };
            } else {
                EutEntity = null;
            }
            const snp: ISNP = {
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                Created: item[Consts.FIELDS.COMMON.CREATION_TIME],
                Modified: item[Consts.FIELDS.COMMON.MODIFICATION_TIME],
                ContentType,
                Summary: item[Consts.FIELDS.SNP.SUMMARY],
                City: item[Consts.FIELDS.SNP.CITY],
                Country,
                TeleportPartner,
                Status: item[Consts.FIELDS.SNP.STATUS],
                EutEntity,
                Comments: item[Consts.FIELDS.SNP.COMMENTS],
                DocumentSpace: item[Consts.FIELDS.COMMON.DOCUMENTS_SPACE]?.Url
            };
            return snp;
        } catch (error) {
            if (error instanceof Error && error.message.includes(strings.DataSheetItemDoesNotExist)) {
                const errorMessage = strings.DataSheetItemDoesNotExist;
                this.setState({ errorMessage });
                return null;

            }
            throw error;
        }
    };
    public initFeeViewItems = async () => {
        const { pnpService, itemId } = this.props;
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.FEE.SUMMARY,
            Consts.FIELDS.FEE.STATUS,
            Consts.FIELDS.FEE.COMMENTS,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
            Consts.FIELDS.FEE.CATEGORY,
            Consts.FIELDS.FEE.SUMMARY,
            Consts.FIELDS.FEE.VATRATE,
            Consts.FIELDS.FEE.DUEDATE,
            Consts.FIELDS.FEE.VATFREECOST,
            Consts.FIELDS.FEE.VAT,
            Consts.FIELDS.FEE.VATRATE,
            Consts.FIELDS.FEE.COSTTYPE,
            Consts.FIELDS.FEE.PONONPO,
            Consts.FIELDS.FEE.CHARGEDBY,
            `${Consts.FIELDS.FEE.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.CURRENCY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.CURRENCY}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.CURRENCY}/${Consts.FIELDS.FEE.CURRENCY_ISOCODE}`,
            `${Consts.FIELDS.FEE.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.LEGALSERVICEPROVIDER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.MA_REQUIREMENT}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.MA_REQUIREMENT}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.RECURRENCEPATTERN}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.RECURRENCEPATTERN}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.SNP}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.SNP}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.FEE.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.FEE.TELEPORTPARTNER}/${Consts.FIELDS.COMMON.TITLE}`

        ];
        const expand = [Consts.FIELDS.FEE.AUTHORITY, Consts.FIELDS.FEE.CURRENCY, Consts.FIELDS.FEE.LEGALSERVICEPROVIDER,
        Consts.FIELDS.FEE.MA_REQUIREMENT, Consts.FIELDS.FEE.RECURRENCEPATTERN, Consts.FIELDS.FEE.SNP, Consts.FIELDS.FEE.TELEPORTPARTNER];
        const list = Consts.LISTS.FEE_URL;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const items = await pnpService.getListItems(listUrl).
            select(...selectedFields)
            .expand(...expand)
            .filter(`${Consts.FIELDS.FEE.SNP}/Id eq ${itemId}`)
            .orderBy(Consts.FIELDS.COMMON.TITLE)();
        const viewItems: IFEE[] = [];
        items.map((item) => {
            const rawDueDate = item[Consts.FIELDS.FEE.DUEDATE];
            const DueDate: Date | null = rawDueDate ? new Date(rawDueDate) : null;
            const ContentType: IContentType = this.feeTypeChoices
                .map((type) => ({ Id: type.key + "", Name: type.text }))
                .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1) || { Id: '', Name: '' };
            let itemValue = null;

            if (ContentType.Id === Consts.CONTENT_TYPES.SNP_FEE && item[Consts.FIELDS.FEE.SNP] && item[Consts.FIELDS.FEE.SNP][Consts.FIELDS.COMMON.ID]) {
                itemValue = {
                    id: item[Consts.FIELDS.FEE.SNP]?.[Consts.FIELDS.COMMON.ID],
                    key: item[Consts.FIELDS.FEE.SNP]?.[Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.SNP,
                    text: item[Consts.FIELDS.FEE.SNP]?.[Consts.FIELDS.COMMON.TITLE],
                    type: DatasheetType.SNP
                };
            } else if (ContentType.Id === Consts.CONTENT_TYPES.MA_REQUIREMENT_FEE && item[Consts.FIELDS.FEE.MA_REQUIREMENT] && item[Consts.FIELDS.FEE.MA_REQUIREMENT][Consts.FIELDS.COMMON.ID]) {
                itemValue = {
                    id: item[Consts.FIELDS.FEE.MA_REQUIREMENT]?.[Consts.FIELDS.COMMON.ID],
                    key: item[Consts.FIELDS.FEE.MA_REQUIREMENT]?.[Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.MARequirement,
                    text: item[Consts.FIELDS.FEE.MA_REQUIREMENT]?.[Consts.FIELDS.COMMON.TITLE],
                    type: DatasheetType.MARequirement
                };
            }

            let chargeByItem = null;
            switch (item[Consts.FIELDS.FEE.CHARGEDBY]) {
                case FeeChargedBy.LawFirm:
                case FeeChargedBy.LegalRepresentative:
                    chargeByItem = {
                        id: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.ID] +
                            "-" + DatasheetType.LegalSvcProvider,
                        text: item[Consts.FIELDS.FEE.LEGALSERVICEPROVIDER][Consts.FIELDS.COMMON.TITLE],
                        type: DatasheetType.LegalSvcProvider
                    };
                    this.ChargeByItemChoices.push(chargeByItem);
                    break;
                case FeeChargedBy.Administration:
                case FeeChargedBy.Organization:
                case FeeChargedBy.Regulator:
                    chargeByItem = {
                        id: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.Authority,
                        text: item[Consts.FIELDS.FEE.AUTHORITY][Consts.FIELDS.COMMON.TITLE],
                        type: DatasheetType.TeleportPartner,
                    };
                    this.ChargeByItemChoices.push(chargeByItem);
                    break;
                case FeeChargedBy.TeleportPartner:
                    chargeByItem = {
                        id: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
                        key: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.TeleportPartner,
                        text: item[Consts.FIELDS.FEE.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
                        type: DatasheetType.TeleportPartner,
                    };
                    this.ChargeByItemChoices.push(chargeByItem);
                    break;
                default:
                    break;
            }

            const vatFreeCost = parseInt(item[Consts.FIELDS.FEE.VATFREECOST]);
            const vat = parseInt(item[Consts.FIELDS.FEE.VAT]);
            const total = (!isNaN(vatFreeCost) ? vatFreeCost : 0) + (!isNaN(vat) ? vat : 0);
            const currencyCode = item[Consts.FIELDS.FEE.CURRENCY]?.[Consts.FIELDS.FEE.CURRENCY_ISOCODE] ?? '';
            viewItems.push({
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                Category: item[Consts.FIELDS.FEE.CATEGORY],
                ContentType,
                Item: itemValue,
                Summary: item[Consts.FIELDS.FEE.SUMMARY],
                VatFreeCost: item[Consts.FIELDS.FEE.VATFREECOST],
                Vat: item[Consts.FIELDS.FEE.VAT],
                Status: item[Consts.FIELDS.FEE.STATUS],
                CostTotal: total ? total.toString() + " " + currencyCode : '0',
                DueDate,
                ChargeByItem: chargeByItem,
            });
        });
        return viewItems;
    };
    public initMARequirementItems = async () => {
        const { pnpService, itemId } = this.props;
        const typeChoices: IDropdownOption[] = [
            { key: Consts.CONTENT_TYPES.MA_RequirementCredential, text: commonStrings.MARequirementCredentialsContentTypeName },
            { key: Consts.CONTENT_TYPES.MA_RequirementGeneric, text: commonStrings.MARequirementGenericContentTypeName },
            { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_NOL, text: commonStrings.MARequirementNolContentTypeName },
            { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_LEGAL, text: commonStrings.MARequirementLegalContentTypeName },
        ];
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
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
            `${Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT}/${Consts.FIELDS.COMMON.TITLE}`,
            Consts.FIELDS.MAGENERIC.GEOLEO,
            Consts.FIELDS.MAGENERIC.COMMENTS,
            Consts.FIELDS.MAGENERIC.APPLICATIONDATE,
            Consts.FIELDS.MAGENERIC.ESTIMATEDDATE,
            Consts.FIELDS.MAGENERIC.EFFECTIVEDATE,
            Consts.FIELDS.MAGENERIC.EXPIRATIONDATE,
            `${Consts.FIELDS.MAGENERIC.RELATEDITEMS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.RELATEDITEMS}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MAGENERIC.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MAGENERIC.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
        ];
        const expand = [Consts.FIELDS.MAGENERIC.MARKETS, Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE,
        Consts.FIELDS.MAGENERIC.SYNTHESIS, Consts.FIELDS.MAGENERIC.SNPS, Consts.FIELDS.MAGENERIC.EUTELSATENTITY, Consts.FIELDS.MAGENERIC.TELEPORTPARTNER,
        Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER, Consts.FIELDS.MAGENERIC.CONTACTS, Consts.FIELDS.MAGENERIC.EUTELSATOWNERS,
        Consts.FIELDS.MAGENERIC.RELATEDITEMS, Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT, Consts.FIELDS.MAGENERIC.AUTHORITY
        ];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.MA_REQUIREMENT_URL);
        const items = await pnpService.getListItems(listUrl).
            filter(`${Consts.FIELDS.MAGENERIC.SNPS_ID} eq ${itemId}`)
            .select(...selectedFields)
            .expand(...expand)
            .orderBy(Consts.FIELDS.COMMON.TITLE)();
        const viewItems: IMARequirement[] = [];
        items.map((item) => {
            const rawApplicationDate = item[Consts.FIELDS.MAGENERIC.APPLICATIONDATE];
            const ApplicationDate: Date | null = rawApplicationDate ? new Date(rawApplicationDate) : null;
            const rawEffectiveDate = item[Consts.FIELDS.MAGENERIC.EFFECTIVEDATE];
            const EffectiveDate: Date | null = rawEffectiveDate ? new Date(rawEffectiveDate) : null;
            const rawExpirationDate = item[Consts.FIELDS.MAGENERIC.EXPIRATIONDATE];
            const ExpirationDate: Date | null = rawExpirationDate ? new Date(rawExpirationDate) : null;
            const responsibleParty = item[Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY];
            let responsiblePartyItem = null;
            switch (responsibleParty) {
                case commonStrings.MARequirementEutelsatTypeName:
                    if (item?.[Consts.FIELDS.MAGENERIC.EUTELSATENTITY]?.[Consts.FIELDS.COMMON.ID] && item?.[Consts.FIELDS.MAGENERIC.EUTELSATENTITY]?.[Consts.FIELDS.COMMON.TITLE]) {
                        responsiblePartyItem = {
                            id: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID] + "-" + ViewType.EutelsatEntity,
                            text: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.TITLE],
                            type: ViewType.EutelsatEntity,
                        };
                    }
                    break;
                case commonStrings.MARequirementTeleportPartnerTypeName:
                    if (item?.[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER]?.[Consts.FIELDS.COMMON.ID] &&
                        item?.[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER]?.[Consts.FIELDS.COMMON.TITLE]) {
                        responsiblePartyItem = {
                            id: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + ViewType.TeleportPartner,
                            text: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
                            type: ViewType.TeleportPartner,
                        };
                    }
                    break;
                case commonStrings.MARequirementDistributionPartnerTypeName:
                    if (item?.[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.ID] && item?.[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.TITLE]) {
                        responsiblePartyItem = {
                            id: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID] + "-" + ViewType.DistributionPartner,
                            text: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE],
                            type: ViewType.DistributionPartner,
                        };
                    }
                    break;
                default:
                    break;
            }
            const Type: IContentType = typeChoices
                .map((type) => ({ Id: type.key + "", Name: type.text }))
                .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1) || { Id: '', Name: '' };
            const authority = item?.[Consts.FIELDS.MAGENERIC.AUTHORITY];
            const RequestedPartyItem = authority ? {
                Id: authority[Consts.FIELDS.COMMON.ID],
                Title: authority[Consts.FIELDS.COMMON.TITLE]
            } : null;
            viewItems.push({
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                MaVertical: item[Consts.FIELDS.MAGENERIC.VERTICALS]?.map((v) => ({
                    key: v,
                    text: v
                })),
                Type,
                Markets: item[Consts.FIELDS.MAGENERIC.MARKETS]?.map((market) => ({
                    Id: market[Consts.FIELDS.COMMON.ID],
                    Title: market[Consts.FIELDS.COMMON.TITLE]
                })) || [],
                Summary: item[Consts.FIELDS.PAYMENT.SUMMARY],
                ResponsibleParty: responsibleParty,
                ResponsiblePartyItem: responsiblePartyItem,
                RequestedPartyItem,
                ApplicationDate,
                ExpirationDate,
                EffectiveDate,
                RagStatus: item[Consts.FIELDS.MAGENERIC.RAGSTATUS],
                Synthesis: (item?.[Consts.FIELDS.MAGENERIC.SYNTHESIS]?.[Consts.FIELDS.COMMON.ID] &&
                    item?.[Consts.FIELDS.MAGENERIC.SYNTHESIS]?.[Consts.FIELDS.COMMON.TITLE]) ? {
                    Id: item[Consts.FIELDS.MAGENERIC.SYNTHESIS][Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.MAGENERIC.SYNTHESIS][Consts.FIELDS.COMMON.TITLE]
                } : null,
            });
        });
        return viewItems;
    };
    public initSNPReadinessItems = async (): Promise<ISNPReadiness[]> => {
        const { pnpService, itemId } = this.props;
        const list = Consts.LISTS.MA_READINESS_URL;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.SNP_MA_READINESS.RESPONSIBLE_PARTY,
            Consts.FIELDS.SNP_MA_READINESS.RAGSTATUS,
            Consts.FIELDS.SNP_MA_READINESS.ESTIMATED_DATE,
            Consts.FIELDS.SNP_MA_READINESS.EFFECTIVE_DATE,
            Consts.FIELDS.SNP_MA_READINESS.COMMENTS,
            Consts.FIELDS.COMMON.CONTENTTYPE_ID,
            `${Consts.FIELDS.SNP_MA_READINESS.SNP}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.SNP_MA_READINESS.SNP}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.SNP_MA_READINESS.OWNER}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.SNP_MA_READINESS.OWNER}/${Consts.FIELDS.COMMON.TITLE}`
        ];
        const expand = [Consts.FIELDS.SNP_MA_READINESS.SNP, Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER, Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER, Consts.FIELDS.SNP_MA_READINESS.OWNER];

        const items = await pnpService.getListItems(listUrl).
            filter(`${Consts.FIELDS.SNP_MA_READINESS.SNP_ID} eq ${itemId}`).
            select(...selectedFields).
            expand(...expand)();




        return items.map((item): ISNPReadiness => {
            let ResponsiblePartyItem: IBaseItem | null = null;
            const ResponsiblePartyType: string = item[Consts.FIELDS.SNP_MA_READINESS.RESPONSIBLE_PARTY];






            if (ResponsiblePartyType === MarketReadinessResponsibleParty.TP) {
                ResponsiblePartyItem = {
                    Id: item[Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER][Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.SNP_MA_READINESS.TELEPORT_PARTNER][Consts.FIELDS.COMMON.TITLE],
                };
                const exists = this.DPandTPDataInfo.some(
                    (d) => d.id === String(ResponsiblePartyItem.Id) && d.key === ResponsiblePartyItem.Title && d.type === DatasheetType.TeleportPartner
                );

                if (!exists) {
                    this.DPandTPDataInfo.push({
                        id: String(ResponsiblePartyItem.Id),
                        key: ResponsiblePartyItem.Title,
                        text: ResponsiblePartyItem.Title,
                        type: DatasheetType.TeleportPartner
                    });

                }
            }
            if (ResponsiblePartyType === MarketReadinessResponsibleParty.DP) {
                ResponsiblePartyItem = {
                    Id: item[Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.SNP_MA_READINESS.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE],
                };
                const exists = this.DPandTPDataInfo.some(
                    (d) => d.id === String(ResponsiblePartyItem.Id) && d.key === ResponsiblePartyItem.Title && d.type === DatasheetType.DistributionPartner
                );

                if (!exists) {
                    this.DPandTPDataInfo.push({
                        id: String(ResponsiblePartyItem.Id),
                        key: ResponsiblePartyItem.Title,
                        text: ResponsiblePartyItem.Title,
                        type: DatasheetType.DistributionPartner
                    });
                }
            }
            const snpReadiness: ISNPReadiness = {
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                Snp: {
                    Id: item[Consts.FIELDS.SNP_MA_READINESS.SNP][Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.SNP_MA_READINESS.SNP][Consts.FIELDS.COMMON.TITLE],
                },
                ResponsiblePartyType,
                ResponsiblePartyItem,
                EutelsatOwners: item[Consts.FIELDS.SNP_MA_READINESS.OWNER]?.map((etuowner) => ({
                    Id: etuowner[Consts.FIELDS.COMMON.ID],
                    Title: etuowner[Consts.FIELDS.COMMON.TITLE]
                })) || [],
                RAGStatus: item[Consts.FIELDS.SNP_MA_READINESS.RAGSTATUS] || '',
                EstimatedDate: item[Consts.FIELDS.SNP_MA_READINESS.ESTIMATED_DATE] ? new Date(item[Consts.FIELDS.SNP_MA_READINESS.ESTIMATED_DATE]) : null,
                EffectiveDate: item[Consts.FIELDS.SNP_MA_READINESS.EFFECTIVE_DATE] ? new Date(item[Consts.FIELDS.SNP_MA_READINESS.EFFECTIVE_DATE]) : null,
                Comments: item[Consts.FIELDS.SNP_MA_READINESS.COMMENTS] || '',

            };
            return snpReadiness;
        });

    };

    private init = async () => {
        let item: ISNP;
        let feeItems: IFEE[] = [];
        let maRequirements: IMARequirement[] = [];
        let snpReadinessItems: ISNPReadiness[] = [];
        const promises: Promise<void>[] = [];
        if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner, UserRole.Admin)) {
            promises.push(this.initFeeViewItems().then((opts) => { feeItems = opts as IFEE[]; }));
        }
        promises.push(this.initItem().then((opts) => { item = opts as ISNP; }));
        promises.push(this.initMARequirementItems().then((opts) => { maRequirements = opts as IMARequirement[]; }));
        promises.push(this.initSNPReadinessItems().then((opts) => { snpReadinessItems = opts as ISNPReadiness[]; }));
        await Promise.all(promises);
        this.setState({
            item,
            feeViewItems: feeItems,
            maRequirements,
            snpReadinessItems
        });
    };

    private renderDatasheetInfo = (): JSX.Element | null => {
        const leftColumns = ['Summary', 'ContentType', 'TeleportPartner', 'EutEntity', 'Country'];
        const rightColumns = ['City', 'Status', 'Comments'];

        return (
            <div className={`${styles.datasheetInfoContainer}`}>
                <div className={styles.datasheetInfoContainerLeft}>
                    {leftColumns.map((key) =>
                        (this.renderDatasheetInfoRow(key))
                    )}
                </div>
                <div className={styles.datasheetInfoContainerRight}>
                    {rightColumns.map((key) =>
                        (this.renderDatasheetInfoRow(key))
                    )}
                </div>
            </div>
        );
    };

    private renderDatasheetInfoRow = (key: string) => {
        const {
            item
        } = this.state;
        // Hide fields based on ContentType
        const contentType = item?.ContentType?.Name;
        if (
            (contentType === SNPType.Stargate && (key === 'TeleportPartner' || key === 'EutEntity')) ||
            (contentType === SNPType.EutelsatOwned && key === 'TeleportPartner') ||
            (contentType === SNPType.Other && key === 'EutEntity')
        ) {
            return null;
        }
        const labelKey = `SNPDataSheetInfoLabel${key}`;
        return (
            <div className="ms-Grid">
                <div className={`ms-Grid-row ${styles.datasheetInfoRow}`}>
                    <div className="ms-Grid-col ms-sm12 ms-md12">
                        <div className="ms-Grid-col ms-sm3 ms-md3">
                            <div className={styles.datasheetInfoLabel}>
                                <Label>{`${strings[labelKey]}:`}</Label>
                            </div>
                        </div>
                        <div className="ms-Grid-col ms-sm9 ms-md9">
                            <div>
                                {this.renderDatasheetInfoValue(item, key)}
                            </div>
                        </div>
                    </div>
                </div>
            </div>);
    };


    private renderDatasheetInfoValue = (item, fieldName) => {
        const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        const itemDatasheetUrl = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(this.config.type)}&itemId=${item.Id}`;
        const value = fieldName ? item[fieldName] : undefined;

        if (fieldName === 'ResponsiblePartyItem') {
            const options = (this.DPandTPDataInfo ?? []) as IItemOption[];
            return renderObjectPills(
                options,
                (option) => {
                    const viewType = option.type; // Use directly, no mapping needed
                    if (!viewType || !option.id) return null;
                    return (
                        <span key={option.key ?? option.id} className={styles.termPill}>
                            <span className={styles.termPillText}>
                                <a
                                    data-interception='off'
                                    rel="noreferrer"
                                    title={String(option.text ?? '')}
                                    onClick={(e) => {
                                        e.preventDefault();
                                        const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${viewType}&itemId=${option.id}`;
                                        UrlHelper.navigate(url, false);
                                    }}
                                >
                                    {String(option.text ?? '')}
                                </a>
                            </span>
                        </span>
                    );
                }
            );
            // return (
            //     <div className={styles.termPills}>
            //         {options.map((t) => {
            //             const viewType = t.type; // Use directly, no mapping needed
            //             if (!viewType || !t.id) return null;

            //             return (
            //                 <span key={t.key ?? t.id} className={styles.termPill}>
            //                     <span className={styles.termPillText}>
            //                         <a
            //                             data-interception='off'
            //                             href={`${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${viewType}&itemId=${t.id}`}
            //                             rel="noreferrer"
            //                             title={String(t.text ?? '')}
            //                         >
            //                             {String(t.text ?? '')}
            //                         </a>
            //                     </span>
            //                 </span>
            //             );
            //         })
            //         }
            //     </div>
            // );
        }
        if (value === null || value === undefined || (Array.isArray(value) && value.length === 0)) {
            return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
        }
        if (fieldName === 'Title') {
            return (
                <a
                    data-interception="on"
                    rel="noreferrer"
                    onClick={(e) => {
                        e.preventDefault();
                        UrlHelper.navigate(itemDatasheetUrl, false);
                    }}
                >
                    {String(value)}
                </a>
            );
        }
        if (fieldName === 'ContentType') {
            return (
                <span className={styles.detailsListValue}>{String(value?.Name)}</span>
            );
        }
        if (fieldName === 'Comments' || fieldName === 'Address' || fieldName === 'Summary') {
            //return <RichText isEditMode={false} value={item[fieldName]} />;
            const cleanValue = DomHelper.cleanRichHtml(item[fieldName]);
            if (!cleanValue) return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
            return <RichText isEditMode={false} value={cleanValue} />;
        }
        if (fieldName === 'EutEntity') {
            const eutEntityListType = DatasheetType.EutelsatEntity;
            return (
                <div className={styles.termPills}>
                    <span key={value.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>

                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(eutEntityListType)}&itemId=${value?.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{value?.Title}</a>
                        </span>
                    </span>
                </div>
            );
        }
        if (fieldName === 'TeleportPartner') {
            const teleportPartnerListType = DatasheetType.TeleportPartner;
            return (
                <div className={styles.termPills}>
                    <span key={value.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>

                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(teleportPartnerListType)}&itemId=${value?.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{value?.Title}</a>
                        </span>
                    </span>
                </div>
            );
        }
        if (fieldName === 'Country') {
            const countryListType = DatasheetType.Country;

            return (
                <div className={styles.termPills}>

                    <span key={value.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(countryListType)}&itemId=${value?.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{value?.Title}</a>
                        </span>
                    </span>

                </div>
            );
        }
        if (fieldName === 'TeleportPartner') {
            const teleportPartnerListType = DatasheetType.TeleportPartner;
            return (
                <div className={styles.termPills}>

                    <span key={value.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(teleportPartnerListType)}&itemId=${value?.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{value?.Title}</a>
                        </span>
                    </span>

                </div>
            );
        }
        if (fieldName === 'EutelsatOwners') {
            return renderObjectPills(
                value as IContact[],
                (owner) => (
                    <span key={owner.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Contact)}&itemId=${owner.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{owner.Title}</a>
                        </span>
                    </span>
                )
            );
        }
        if (fieldName === 'Comments') {
            return <RichText isEditMode={false} value={value} />;
        }

        if (fieldName === 'Status' || fieldName === 'RAGStatus') {
            return StatusPill(String(value));
        }
        if (fieldName === 'City') {
            return (
                <span className={styles.detailsListValue}>{String(value)}</span>
            );
        }
        if (typeof value === 'boolean') {
            const iconName = value ? 'AcceptMedium' : 'Cancel';
            const iconClass = value ? styles.iconGreen : styles.iconRed;
            return <div className={styles.detailsListValue}><Icon iconName={iconName} className={iconClass} /></div>;
        }
        if (value instanceof Date) {
            return <span className={styles.detailsListValue}>{UtilHelper.formatDate(value, 'LL', 'en-us')}</span>;
        }
        if (typeof value === 'string' || typeof value === 'number') {
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }
        if (Array.isArray(value) && value?.every((v) => typeof v === 'string' || typeof v === 'number')) {
            return renderArrayPills(value as (string | number)[]);
        }

    };

    public render(): React.ReactElement<ISNPDatasheetProps> {
        const {
            showEditFormDialog,
            showDeleteDialog,
            item,
            feeViewItems,
            maRequirements,
            errorMessage
        } = this.state;
        const { pnpService, userService, itemId } = this.props;
        const created = UtilHelper.formatDate(item?.Created, 'LL', 'en-us');
        const modified = UtilHelper.formatDate(item?.Modified, 'LL', 'en-us');
        if (!this.hasAnyRole(UserRole.Contributor, UserRole.Owner, UserRole.Visitor,
            UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Admin)) {
            return (
                <AccessDeniedMessage message={strings.FormAccessDeniedMessage}> </AccessDeniedMessage>
            );
        }
        if (errorMessage) {
            if (errorMessage.includes(strings.DataSheetItemDoesNotExist)) {
                return (
                    <div className={styles.noDataMessage}>
                        <Icon iconName="Remove" className={styles.mutedIcon} />
                        <span>{strings.DataSheetNoRecordFound || 'No record found'}</span>
                        <Icon iconName="Remove" className={styles.mutedIcon} />
                    </div>
                );
            }
        }
        return (
            <>
                {
                    this.state && this.state.isDatasheetReady ?

                        <div>
                            <div className={styles.datasheetHeader}>
                                <div className={styles.datasheetTypeContainer}>
                                    < h5 className={styles.datasheetType} >
                                        {strings.SNPDataSheetTitle}
                                    </h5 >
                                </div>
                                {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.datasheetCommandsContainer}>
                                    {<div className={styles.datasheetCommand} title={strings.DataSheetEditButton}>
                                        <a href="#" onClick={this.openEditForm} >
                                            <Icon iconName="Edit" />
                                        </a>
                                    </div>
                                    }
                                    {showEditFormDialog &&
                                        this.renderEditFormDialog()
                                    }
                                    {<div className={styles.datasheetCommand} title={strings.DataSheetDeleteButton}>
                                        <a href="#" className="" onClick={this.openDeleteDialog}>
                                            <Icon iconName="Delete" />
                                        </a>
                                    </div>}
                                    {showDeleteDialog && this.renderDeleteDialog()}
                                </div>}
                            </div>
                            <div className={styles.datasheetTitleContainer}>
                                <Icon iconName='DateTime' className={styles.dateIcon} />
                                <span className={styles.dateText} >
                                    {`${strings.DataSheetLabelCreatedOn} ${created} / ${strings.DataSheetLabelUpdatedOn} ${modified}`}
                                </span>
                                <div className={styles.titleText}>
                                    <h1>{item?.Title}</h1>
                                </div>
                            </div>
                            <div className={styles.tabsContainer}>
                                <Pivot linkSize="large" >
                                    <PivotItem headerText="Info" itemIcon="Info" >
                                        <div className={styles.section}>
                                            <Icon iconName="TaskManager" className="sectionIcon" />
                                            <h3 >{strings.DataSheetSectionProperties}</h3>
                                        </div>
                                        <LayerHost
                                            id='sanctionCategoryLayer'
                                            style={{ position: 'relative', zIndex: 500000 }} // higher than page chrome
                                        />
                                        {this.renderDatasheetInfo()}
                                        <div className={styles.section}>
                                            <Icon iconName="Library" className="sectionIcon" />
                                            <h3 >{strings.DataSheetSectionDocs}</h3>
                                        </div>
                                        <LibraryView
                                            folderPath={item.DocumentSpace}
                                            pnpService={pnpService}
                                            userService={userService}
                                            pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}>
                                        </LibraryView>
                                    </PivotItem>
                                    {this.hasAnyRole(UserRole.Owner, UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Admin) && <PivotItem headerText="Fees" itemIcon="Bank" >
                                        <FeeDatasheetView
                                            items={feeViewItems || []}
                                            pnpService={pnpService}
                                            userService={userService}
                                            refresh={() => this.init()}
                                            relatedItemId={item?.Id}
                                            relatedItemType={FormType.SNP}
                                        />
                                    </PivotItem>}
                                    <PivotItem headerText="Market Readiness" itemIcon="Rocket" >
                                        <SNPReadinessView
                                            items={this.state.snpReadinessItems}
                                            renderField={(item, key) => this.renderDatasheetInfoValue(item, key)}
                                            pnpService={pnpService}
                                            userService={userService}
                                            refresh={this.init}
                                            snpItemId={itemId}

                                        />
                                    </PivotItem>
                                    <PivotItem headerText="MA Requirements" itemIcon="Bullseye" >
                                        <MARequirementView pnpService={pnpService} userService={userService}
                                            columns={['Title', 'Countries', 'Summary', 'Type', 'RequestedPartyItem', 'ResponsiblePartyItem', 'RagStatus', 'Synthesis', 'ApplicationDate', 'EffectiveDate', 'ExpirationDate']}
                                            filters={['Search', 'Type', 'Countries', 'RequestedPartyItem', 'ResponsiblePartyItem', 'RagStatus', 'Synthesis']}
                                            pagination={{ isEnabled: false }}
                                            grouping={{ isEnabled: true, groupBy: 'ResponsibleParty' }}
                                            sort={{
                                                Property: 'ResponsibleParty',
                                                Direction: SortDirection.Ascending
                                            }}
                                            viewItems={maRequirements}
                                        ></MARequirementView>

                                    </PivotItem>
                                </Pivot>
                            </div>
                        </div > :
                        <div className="ms-Grid-col ms-sm12 ms-md12">
                            <div className={styles.spinnerContainer}>
                                <Spinner label={'Loading...'} />
                            </div>
                        </div>
                }
            </>




        );
    }
    // private renderSnpReadiness = (): JSX.Element | null => {
    //     const {
    //         snpReadinessItems
    //     } = this.state;
    //     const itemsHtml = snpReadinessItems.map((readinessItem) => {
    //         let responsibleTitle = "";
    //         let responsibleIcon = "";
    //         switch (readinessItem.ResponsiblePartyType) {
    //             case MarketReadinessResponsibleParty.Overall:
    //                 responsibleTitle = `${readinessItem.ResponsiblePartyType}`;
    //                 responsibleIcon = 'Country';
    //                 break;
    //             case MarketReadinessResponsibleParty.Eutelsat:
    //                 responsibleIcon = 'Satellite';
    //                 responsibleTitle = `${readinessItem.ResponsiblePartyType}`;
    //                 break;
    //             case MarketReadinessResponsibleParty.DP:
    //                 responsibleIcon = 'Agreement';
    //                 responsibleTitle = `${readinessItem.ResponsiblePartyType} - ${readinessItem.ResponsiblePartyItem?.Title}`;
    //                 break;
    //             case MarketReadinessResponsibleParty.TP:
    //                 responsibleIcon = 'Dish';
    //                 responsibleTitle = `${readinessItem.ResponsiblePartyType} - ${readinessItem.ResponsiblePartyItem?.Title}`;
    //                 break;
    //         }

    //         return (
    //             <AccordionItem value={responsibleTitle} key={responsibleTitle} >
    //                 <AccordionHeader icon={responsibleIcon}>{responsibleTitle}</AccordionHeader>
    //                 <AccordionPanel>
    //                     <div>
    //                         <Accordion
    //                             multiple
    //                             collapsible
    //                         >
    //                             {this.renderSnpReadinessInfo(readinessItem)}
    //                         </Accordion>
    //                     </div>

    //                 </AccordionPanel>
    //             </AccordionItem>
    //         );
    //     });
    //     return (
    //         <div className={styles.readinessContainer}>
    //             <div className={styles.readinessNewButton}>
    //                 <ActionButton iconProps={{ iconName: 'Add' }}>
    //                     {strings.SnpReadinessLabelNewButton}
    //                 </ActionButton>
    //             </div>
    //             <Accordion multiple collapsible>
    //                 {itemsHtml}
    //             </Accordion>
    //         </div>
    //     );
    // };

    // private renderSnpReadinessInfo = (readinessItem: ISNPReadiness): JSX.Element | null => {
    //     const leftColumns = ['RAGStatus', 'EstimatedDate', 'EffectiveDate', 'EutelsatOwners'];
    //     const rightColumns = ['Comments'];
    //     return (
    //         <div className={`${styles.datasheetInfoContainer}`}>
    //             <div className={styles.datasheetInfoContainerLeft}>
    //                 {leftColumns.map((key) =>
    //                     (this.renderSnpReadinessInfoRow(readinessItem, key))
    //                 )}
    //             </div>
    //             <div className={styles.datasheetInfoContainerRight}>
    //                 {rightColumns.map((key) =>
    //                     (this.renderSnpReadinessInfoRow(readinessItem, key))
    //                 )}
    //             </div>
    //         </div>
    //     );
    // };

    // private renderSnpReadinessInfoRow = (readinessItem: ISNPReadiness, key: string) => {
    //     const labelKey = `SnpReadinessLabel${key}`;
    //     return (
    //         <div className="ms-Grid">
    //             <div className={`ms-Grid-row ${styles.datasheetInfoRow}`}>
    //                 <div className="ms-Grid-col ms-sm12 ms-md12">
    //                     <div className="ms-Grid-col ms-sm3 ms-md3">
    //                         <div className={styles.datasheetInfoLabel}>
    //                             <Label>{`${strings[labelKey]}:`}</Label>
    //                         </div>
    //                     </div>
    //                     <div className="ms-Grid-col ms-sm9 ms-md9">
    //                         <div>
    //                             {this.renderDatasheetInfoValue(readinessItem, key)}
    //                         </div>
    //                     </div>
    //                 </div>
    //             </div>
    //         </div>);
    // };

    private openEditForm = () => {
        this.setState({ showEditFormDialog: true });
    };
    private closeEditForm = async (refreshData = false) => {
        this.setState({ showEditFormDialog: false });
        if (refreshData) {
            await this.init();
        }
    };
    private openDeleteDialog = () => {
        this.setState({ showDeleteDialog: true });
    };
    private closeDeleteDialog = () => {
        this.setState({ showDeleteDialog: false });
    };
    private deleteDataSheet = async () => {
        try {
            this.setState({ isFormProcessing: true });
            const { pnpService, itemId } = this.props;
            const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.SNP).list;
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
            await pnpService.getListItems(listUrl).getById(itemId).recycle();
            const messageBanner: IMessageBanner = {
                message: strings.DeleteSuccessMessage,
                type: 'success',
                visible: true
            };
            this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
            setTimeout(() => { window.location.href = Config.SITE.ROOT_WEB_URL; }, 2000);
        } catch (error) {
            console.error(strings.DeleteErrorMessage, error);
            const messageBanner: IMessageBanner = {
                message: strings.DeleteErrorMessage,
                type: 'error',
                visible: true
            };
            this.setState({ isFormProcessing: false, messageBanner });
        }
    };
    private renderEditFormDialog = () => {

        const {
            showEditFormDialog,
        } = this.state;
        const {
            pnpService,
            userService,
            itemId
        } = this.props;

        const editFormControl = (<SNP pnpService={pnpService} userService={userService} itemId={itemId} callback={async () => { await this.closeEditForm(true); }}></SNP>);
        const editFormTitle = strings.SNPDataSheetEditFormTitle;

        return (<Modal
            isOpen={showEditFormDialog}
            onDismiss={() => this.closeEditForm(true)}
            containerClassName={styles.modalContainer}
            isBlocking={true}
        >
            <div>
                <div className={styles.modalHeader}>
                    {editFormTitle}
                    <ActionButton
                        iconProps={{ iconName: 'Cancel' }}
                        onClick={() => this.closeEditForm()}
                        className={styles.iconButton}
                    >{strings.DataSheetCancelButton}</ActionButton>
                </div>
                <div className={styles.modalBody}>
                    {editFormControl}
                </div>
            </div>
        </Modal>);
    };
    private renderDeleteDialog = () =>
    (<Dialog
        hidden={false}
        onDismiss={this.closeDeleteDialog}
        dialogContentProps={{
            type: DialogType.normal,
            title: strings.DataSheetDeleteModalTitle
        }}
        modalProps={{
            isBlocking: true
        }}
        styles={{
            main: {
                selectors: {
                    '@media (min-width: 480px)': {
                        maxWidth: '410px !important', // Override media query
                    },
                },

            }
        }}
    >
        <div>
            {this.state.messageBanner.visible && (
                <div className={styles.topMessageBanner}>
                    <MessageBanner
                        message={this.state.messageBanner.message}
                        type={this.state.messageBanner.type}
                        visible={this.state.messageBanner.visible}
                    />
                </div>
            )}
            {this.state.messageBanner.type !== 'success' && (
                <span>{strings.DataSheetDeleteModalMessage}</span>
            )}
            <DialogFooter>
                <PrimaryButton
                    disabled={this.state.isConfirmButtonDisabled}
                    onClick={this.deleteDataSheet}
                //text={strings.DataSheetConfirmButton}
                >
                    {this.state.isFormProcessing ? (
                        <Stack horizontal verticalAlign="center">
                            <Spinner size={SpinnerSize.xSmall} labelPosition="right" />
                            <span style={{ marginLeft: 8 }}>Processing</span>
                        </Stack>
                    ) : (
                        strings.DataSheetConfirmButton
                    )}
                </PrimaryButton>
                <DefaultButton
                    onClick={this.closeDeleteDialog}
                    text={strings.DataSheetBackButton}
                />
            </DialogFooter>
        </div>
    </Dialog>);
}
