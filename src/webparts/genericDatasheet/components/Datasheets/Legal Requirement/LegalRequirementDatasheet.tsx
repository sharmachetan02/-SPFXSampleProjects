/* eslint-disable max-len */
import * as React from 'react';
import { ILegalRequirementDatasheetProps, ILegalRequirementDatasheetState } from './LegalRequirement.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import {
    ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, IDropdownOption, ITooltipHostStyles, Label,
    LayerHost, Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost
} from '@fluentui/react';
import { Datasheet, DatasheetType, IDatasheetConfig } from '../../Datasheet.types';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { IContact, IContentType, IFEE, IItemOption, IMARequirement, IMarket, ISNP } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { renderArrayPills, renderObjectPills, StatusPill } from '../../Datasheet.utility';
import { RichText } from '../../../../../common/components/RichText/RichText';
import LibraryView from '../../../../genericView/components/Library View/LibraryView';
import FeeDatasheetView from '../../Datasheet Views/FeeDatasheetView';
import { FormType, IMessageBanner } from '../../../../genericForm/components/Form.types';
import LegalRequirement from '../../../../genericForm/components/MA Requirement Legal/LegalRequirement';
import { FeeChargedBy, FeeType, MARequirementResponsibleParty, MARequirementType, UserRole } from '../../../../../common/models/Enums';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';
import AccessDeniedMessage from '../../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../../common/helpers/DomHelper';
//import { CallOutButton } from '../../../../../common/components/CallOutButton/CallOutButton';

export default class LegalRequirementDatasheet extends React.Component<
    ILegalRequirementDatasheetProps,
    ILegalRequirementDatasheetState
> {
    private config: IDatasheetConfig;
    private maReqTypeChoices: { key: string; text: string }[] = [
        { key: Consts.CONTENT_TYPES.MA_RequirementCredential, text: MARequirementType.Credential },
        { key: Consts.CONTENT_TYPES.MA_RequirementGeneric, text: MARequirementType.Generic },
        { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_NOL, text: MARequirementType.NOL },
        { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_LEGAL, text: MARequirementType.Legal },
    ];
    private feeTypeChoices: { key: string; text: string }[] = [
        { key: Consts.CONTENT_TYPES.SNP_FEE, text: FeeType.SNP },
        { key: Consts.CONTENT_TYPES.MA_REQUIREMENT_FEE, text: FeeType.MATypeRequirement }
    ];
    private ChargeByItemChoices: IItemOption[] = [];
    private responsibleItemChoices: IItemOption[] = [];
    private requestedItemChoices: IDropdownOption[] = [];
    constructor(props: Readonly<ILegalRequirementDatasheetProps>) {
        super(props);
        this.state = {
            isDatasheetReady: false,
            showEditFormDialog: false,
            showDeleteDialog: false,
            item: null,
            errorMessage: null,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'error',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        };
        this.config = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.MARequirementLegal);
        //console.log("Teleport partner view config:", this.config);
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
        const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.MARequirementLegal).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        try {
            const selectedFields: string[] = [
                Consts.FIELDS.COMMON.TITLE,
                Consts.FIELDS.COMMON.ID,
                Consts.FIELDS.COMMON.CONTENTTYPE_ID,
                Consts.FIELDS.COMMON.CREATION_TIME,
                Consts.FIELDS.COMMON.MODIFICATION_TIME,
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
                `${Consts.FIELDS.MAGENERIC.AUTHORITY}/${Consts.FIELDS.COMMON.ID}`,
                `${Consts.FIELDS.MAGENERIC.AUTHORITY}/${Consts.FIELDS.COMMON.TITLE}`,
                `${Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT}/${Consts.FIELDS.COMMON.ID}`,
                `${Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT}/${Consts.FIELDS.COMMON.TITLE}`,
                Consts.FIELDS.MAGENERIC.APPLICATIONDATE,
                Consts.FIELDS.MAGENERIC.ESTIMATEDDATE,
                Consts.FIELDS.MAGENERIC.EFFECTIVEDATE,
                Consts.FIELDS.MA_REQUIREMENT_NOL.LETTER1_SENTDATE,
                Consts.FIELDS.MA_REQUIREMENT_NOL.LETTER2_SENTDATE,
                Consts.FIELDS.MA_REQUIREMENT_NOL.LETTER1_MILESTONE,
                Consts.FIELDS.MA_REQUIREMENT_NOL.LETTER2_MILESTONE,
                Consts.FIELDS.MA_REQUIREMENT_NOL.RESPONSE_DATE,
                Consts.FIELDS.MA_REQUIREMENT_NOL.RESPONSE_SUMMARY,
                Consts.FIELDS.COMMON.DOCUMENTS_SPACE,
                `${Consts.FIELDS.MAGENERIC.RELATEDITEMS}/${Consts.FIELDS.COMMON.ID}`,
                `${Consts.FIELDS.MAGENERIC.RELATEDITEMS}/${Consts.FIELDS.COMMON.TITLE}`,
                `${Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY_CONTACT}/${Consts.FIELDS.COMMON.ID}`,
                `${Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY_CONTACT}/${Consts.FIELDS.COMMON.TITLE}`,
                Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_TYPE,

            ];
            const expand = [Consts.FIELDS.MAGENERIC.MARKETS, Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE,
            Consts.FIELDS.MAGENERIC.SYNTHESIS, Consts.FIELDS.MAGENERIC.SNPS, Consts.FIELDS.MAGENERIC.EUTELSATENTITY, Consts.FIELDS.MAGENERIC.TELEPORTPARTNER,
            Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER, Consts.FIELDS.MAGENERIC.CONTACTS, Consts.FIELDS.MAGENERIC.EUTELSATOWNERS,
            Consts.FIELDS.MAGENERIC.RELATEDITEMS, Consts.FIELDS.MAGENERIC.AUTHORITY, Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT
            ];
            const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
            const rawApplicationDate = item[Consts.FIELDS.MAGENERIC.APPLICATIONDATE];
            const ApplicationDate: Date | null = rawApplicationDate ? new Date(rawApplicationDate) : null;
            const rawEffectiveDate = item[Consts.FIELDS.MAGENERIC.EFFECTIVEDATE];
            const EffectiveDate: Date | null = rawEffectiveDate ? new Date(rawEffectiveDate) : null;
            const rawExpirationDate = item[Consts.FIELDS.MAGENERIC.EXPIRATIONDATE];
            const ExpirationDate: Date | null = rawExpirationDate ? new Date(rawExpirationDate) : null;
            const rawEstimatedDate = item[Consts.FIELDS.MAGENERIC.ESTIMATEDDATE];
            const EstimatedDate: Date | null = rawEstimatedDate ? new Date(rawEstimatedDate) : null;
            const rawLetter1SentDate = item[Consts.FIELDS.MA_REQUIREMENT_NOL.LETTER1_SENTDATE];
            const Letter1SentDate: Date | null = rawLetter1SentDate ? new Date(rawLetter1SentDate) : null;
            const rawLetter2SentDate = item[Consts.FIELDS.MA_REQUIREMENT_NOL.LETTER2_SENTDATE];
            const Letter2SentDate: Date | null = rawLetter2SentDate ? new Date(rawLetter2SentDate) : null;
            const rawLetter1Milestone = item[Consts.FIELDS.MA_REQUIREMENT_NOL.LETTER1_MILESTONE];
            const Letter1Milestone: Date | null = rawLetter1Milestone ? new Date(rawLetter1Milestone) : null;
            const rawLetter2Milestone = item[Consts.FIELDS.MA_REQUIREMENT_NOL.LETTER2_MILESTONE];
            const Letter2Milestone: Date | null = rawLetter2Milestone ? new Date(rawLetter2Milestone) : null;
            const rawResponseDate = item[Consts.FIELDS.MA_REQUIREMENT_NOL.RESPONSE_DATE];
            const ResponseDate: Date | null = rawResponseDate ? new Date(rawResponseDate) : null;
            const responsibleParty = item[Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY];
            let responsiblePartyItem = null;
            switch (responsibleParty) {
                case MARequirementResponsibleParty.Eutelsat:
                    if (item?.[Consts.FIELDS.MAGENERIC.EUTELSATENTITY]?.[Consts.FIELDS.COMMON.ID] && item?.[Consts.FIELDS.MAGENERIC.EUTELSATENTITY]?.[Consts.FIELDS.COMMON.TITLE]) {
                        responsiblePartyItem = {
                            id: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.EutelsatEntity,
                            text: item[Consts.FIELDS.MAGENERIC.EUTELSATENTITY][Consts.FIELDS.COMMON.TITLE],
                            type: DatasheetType.EutelsatEntity,
                        };
                        this.responsibleItemChoices.push(responsiblePartyItem);
                    }
                    break;
                case MARequirementResponsibleParty.TP:
                    if (item?.[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER]?.[Consts.FIELDS.COMMON.ID] &&
                        item?.[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER]?.[Consts.FIELDS.COMMON.TITLE]) {
                        responsiblePartyItem = {
                            id: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.TeleportPartner,
                            text: item[Consts.FIELDS.MAGENERIC.TELEPORTPARTNER][Consts.FIELDS.COMMON.TITLE],
                            type: DatasheetType.TeleportPartner,
                        };
                        this.responsibleItemChoices.push(responsiblePartyItem);
                    }
                    break;
                case MARequirementResponsibleParty.DP:
                    if (item?.[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.ID] && item?.[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER]?.[Consts.FIELDS.COMMON.TITLE]) {
                        responsiblePartyItem = {
                            id: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID],
                            key: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.ID] + "-" + DatasheetType.DistributionPartner,
                            text: item[Consts.FIELDS.MAGENERIC.DISTRIBUTION_PARTNER][Consts.FIELDS.COMMON.TITLE],
                            type: DatasheetType.DistributionPartner,
                        };
                        this.responsibleItemChoices.push(responsiblePartyItem);
                    }
                    break;
                default:
                    break;
            }
            const Type: IContentType = this.maReqTypeChoices
                .map((type) => ({ Id: type.key + "", Name: type.text }))
                .find((ct) => item[Consts.FIELDS.COMMON.CONTENTTYPE_ID]?.indexOf(ct.Id) !== -1) || { Id: '', Name: '' };
            const authority = item?.[Consts.FIELDS.MAGENERIC.AUTHORITY];
            const RequestedPartyItem = authority ? {
                Id: authority[Consts.FIELDS.COMMON.ID],
                Title: authority[Consts.FIELDS.COMMON.TITLE]
            } : null;
            if (RequestedPartyItem) {
                this.requestedItemChoices.push({
                    key: RequestedPartyItem.Id,
                    text: RequestedPartyItem.Title
                });
            }
            const maRequirementItem: IMARequirement = {
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
                EutelsatOwner: item[Consts.FIELDS.MAGENERIC.EUTELSATOWNERS]?.map((eutOwner) => ({
                    Id: eutOwner[Consts.FIELDS.COMMON.ID],
                    Title: eutOwner[Consts.FIELDS.COMMON.TITLE]
                })) || [],
                RequestedPartyContacts: item[Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_CONTACT]?.map((reqContact) => ({
                    Id: reqContact[Consts.FIELDS.COMMON.ID],
                    Title: reqContact[Consts.FIELDS.COMMON.TITLE]
                })) || [],
                ResponsiblePartyContacts: item[Consts.FIELDS.MAGENERIC.RESPONSIBLEPARTY_CONTACT]?.map((reqContact) => ({
                    Id: reqContact[Consts.FIELDS.COMMON.ID],
                    Title: reqContact[Consts.FIELDS.COMMON.TITLE]
                })) || [],
                Snps: item[Consts.FIELDS.MAGENERIC.SNPS]?.map((snp) => ({
                    Id: snp[Consts.FIELDS.COMMON.ID],
                    Title: snp[Consts.FIELDS.COMMON.TITLE]
                })) || [],
                Summary: item[Consts.FIELDS.PAYMENT.SUMMARY],
                ResponsiblePartyItem: responsiblePartyItem,
                RequestedPartyType: item[Consts.FIELDS.MAGENERIC.REQUESTEDPARTY_TYPE],
                RequestedPartyItem,
                ApplicationDate,
                ExpirationDate,
                EffectiveDate,
                EstimatedDate,
                Letter1SentDate,
                Letter2SentDate,
                Letter1Milestone,
                Letter2Milestone,
                ResponseDate,
                RequirementType: item[Consts.FIELDS.MAGENERIC.REQUIREMENT_TYPE][Consts.FIELDS.COMMON.TITLE],
                ResponseSummary: item[Consts.FIELDS.MA_REQUIREMENT_NOL.RESPONSE_SUMMARY],
                RagStatus: item[Consts.FIELDS.MAGENERIC.RAGSTATUS],
                GeoLeo: item[Consts.FIELDS.MAGENERIC.GEOLEO],
                Synthesis: (item?.[Consts.FIELDS.MAGENERIC.SYNTHESIS]?.[Consts.FIELDS.COMMON.ID] &&
                    item?.[Consts.FIELDS.MAGENERIC.SYNTHESIS]?.[Consts.FIELDS.COMMON.TITLE]) ? {
                    Id: item[Consts.FIELDS.MAGENERIC.SYNTHESIS][Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.MAGENERIC.SYNTHESIS][Consts.FIELDS.COMMON.TITLE]
                } : null,
                Created: new Date(item[Consts.FIELDS.COMMON.CREATION_TIME]),
                Modified: new Date(item[Consts.FIELDS.COMMON.MODIFICATION_TIME]),
                DocumentSpace: item[Consts.FIELDS.COMMON.DOCUMENTS_SPACE]?.Url,
                ResponsibleParty: responsibleParty,
            };
            return maRequirementItem;

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
            .filter(`${Consts.FIELDS.FEE.MA_REQUIREMENT}/Id eq ${itemId}`)
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
                        type: DatasheetType.Authority,
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
                ChargeByItem: chargeByItem
            });
        });
        return viewItems;
    };
    private init = async () => {
        let item: IMARequirement;
        let feeItems: IFEE[] = [];
        const promises: Promise<void>[] = [];
        if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner, UserRole.Admin)) {
            promises.push(this.initFeeViewItems().then((opts) => { feeItems = opts as IFEE[]; }));
        }
        promises.push(this.initItem().then((opts) => { item = opts as IMARequirement; }));
        //promises.push(this.initSanctionCategoriesChoices().then((opts) => { sanctions = opts as ISanctionOption[]; }));
        await Promise.all(promises);

        /*const sanctionCategory = sanctions.find((cat) => cat.key.toString() === item.SanctionCategory.Id.toString());
        item.SanctionCategory = {
          Id: item.SanctionCategory.Id,
          Title: sanctionCategory.text,
          Color: sanctionCategory.color,
          Summary: sanctionCategory.summary
        };*/
        this.setState({
            item,
            feeViewItems: feeItems
        });
        //console.log('item', item);
    };

    private renderDatasheetInfo = (): JSX.Element | null => {
        const leftColumns = ['Summary', 'Markets', 'ResponsibleParty', 'ResponsiblePartyItem', 'ResponsiblePartyContacts', 'EutelsatOwner', 'RequestedPartyType', 'RequestedPartyItem', 'RequestedPartyContacts'];
        const rightColumns = ['RequirementType', 'ApplicationDate', 'EffectiveDate', 'EstimatedDate', 'RagStatus', 'Synthesis', 'MaVertical', 'Snps', 'GeoLeo', 'Comments'];

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
        if ((key === 'ResponsiblePartyContacts' || key === 'Holder Contact(s)') && item?.ResponsibleParty === MARequirementResponsibleParty.Eutelsat) {
            return null;
        }
        // Snps field will be visible only if Verticals array contains SNP
        if (key === 'Snps' && !item?.MaVertical?.some((v) => v.text === 'SNP')) {
            return null;
        }
        const labelKey = `LegalRequirementDataSheetInfoLabel${key}`;
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
        //const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        //const value = fieldName ? item[fieldName] : undefined;
        /*const swatchStyle = (color?: string) => ({
          width: 12,
          height: 12,
          borderRadius: 3,
          background: color || '#888',
          display: 'inline-block',
          verticalAlign: 'middle' as const,
          marginRight: 6,
          border: '1px solid rgba(0,0,0,.1)',
        });*/

        const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        const itemDatasheetUrl = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(this.config.type)}&itemId=${item.Id}`;
        //if (!column) return <span />;
        //const fieldName = column.fieldName as keyof ILegalRequirement | undefined;
        const value = fieldName ? item[fieldName] : undefined;
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
        if (fieldName === 'Markets') {
            return renderObjectPills(
                value as IMarket[],
                (market) => (
                    <span key={market.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Market)}&itemId=${market.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{market.Title}</a>
                        </span>
                    </span>
                )
            );
        }
        if (fieldName === 'Snps') {
            return renderObjectPills(
                value as ISNP[],
                (snp) => (
                    <span key={snp.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.SNP)}&itemId=${snp.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{snp.Title}</a>
                        </span>
                    </span>
                )
            );
        }
        if (fieldName === 'Type') {
            return (
                <div className={styles.tooltipHostContainer} title="">

                    <span className={styles.tooltipHostText}> {value?.Name}</span>
                </div>
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
        if (fieldName === 'RequirementType') {
            return (
                <div className={styles.termPills}>
                    <span key={(value)?.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            {value}
                        </span>
                    </span>

                </div>
            );
        }
        if (fieldName === 'RequestedPartyItem') {
            const eutEntityListType = DatasheetType.Authority;
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
        if (fieldName === 'ResponsiblePartyContacts') {
            return renderObjectPills(
                value as IContact[],
                (contact) => (
                    <span key={contact.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a data-interception="on" href={`${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Contact)}&itemId=${contact.Id}`}
                                rel="noreferrer" title={contact.Title} >{contact.Title}</a>
                        </span>
                    </span>
                )
            );
        }
        if (fieldName === 'ResponsiblePartyItem') {
            const eutEntityListType = value?.type;
            return (
                <div className={styles.termPills}>
                    <span key={value.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>

                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(eutEntityListType)}&itemId=${value?.id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{value?.text}</a>
                        </span>
                    </span>
                </div>
            );
        }

        if (fieldName === 'EutelsatOwner') {
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

        if (fieldName === 'RequestedPartyContacts') {
            return renderObjectPills(
                value as IContact[],
                (contact) => (
                    <span key={contact.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a
                                data-interception="on"
                                rel="noreferrer"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Contact)}&itemId=${contact.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                            >{contact.Title}</a>
                        </span>
                    </span>
                )
            );
        }
        if (fieldName === 'Synthesis') {
            return (
                <div className={styles.termPills}>
                    <span key={(value)?.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            {(value)?.Title}
                        </span>
                    </span>

                </div>
            );
        }
        if (fieldName === 'MaVertical') {
            return renderObjectPills(
                value as Array<{ key: string; text: string }>,
                (item) => (
                    <span className={styles.termPill}>
                        <span className={styles.termPillText}>
                            {item?.text}
                        </span>
                    </span>
                )
            );
        }
        /* if (fieldName === 'ContentType') {
             const ct = value as IDropdownOption;
             return <span className={styles.detailsListValue}>{String(ct?.['Name'])}</span>;
         }*/

        /*if (fieldName === 'Summary') {
          return (
            <CallOutButton icon={'View'} name={'Preview'} >
              <RichText isEditMode={false} value={item[fieldName]} />
            </CallOutButton>
          );
        }*/
        if (fieldName === 'Comments' || fieldName === 'Address' || fieldName === 'Summary') {
            return <RichText isEditMode={false} value={item[fieldName]} />;
        }
        if (fieldName === 'RagStatus') {
            return StatusPill(String(value));
        }
        if (fieldName === 'Portal') {
            return (
                <a data-interception="off" href={String(value)} target="_blank" rel="noreferrer">{"Link"}</a>
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

    public render(): React.ReactElement<ILegalRequirementDatasheetProps> {
        const {
            showEditFormDialog,
            showDeleteDialog,
            item,
            feeViewItems,
            errorMessage
        } = this.state;
        const { pnpService, userService } = this.props;
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
                                        {strings.LegalRequirementDataSheetTitle}
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
                                    {this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner, UserRole.Admin) && <PivotItem headerText="Fees" itemIcon="Bank" >
                                        <FeeDatasheetView
                                            items={feeViewItems || []}
                                            pnpService={pnpService}
                                            userService={userService}
                                            refresh={() => this.init()}
                                            relatedItemId={item?.Id}
                                            relatedItemType={FormType.MARequirementLegal}
                                        />
                                    </PivotItem>}
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
            const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.MARequirementLegal).list;
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
            await pnpService.getListItems(listUrl).getById(itemId).recycle();
            const messageBanner: IMessageBanner = {
                message: strings.DeleteSuccessMessage,
                type: 'success',
                visible: true
            };
            this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
            //setTimeout(() => { window.location.href = Config.SITE.ROOT_WEB_URL; }, 2000);
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

        const editFormControl = (<LegalRequirement pnpService={pnpService} userService={userService} itemId={itemId} callback={async () => { await this.closeEditForm(true); }}></LegalRequirement>);
        const editFormTitle = strings.LegalRequirementDataSheetEditFormTitle;

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
