import { MessageType } from '../../../common/components/Message Banner/MessageBanner';
import { Consts } from '../../../common/consts/Consts';

export enum FormType {
  EutelsatEntity = 'eutelsat_entity',
  Country = 'country',
  Market = 'market',
  LegalServiceProvider = 'legal_svc_provider',
  Authority = 'authority',
  Contact = 'contact',
  SNP = 'snp',
  TP = 'tp',
  DP = 'dp',
  Test = 'test',
  Fee = 'fee',
  Payment = 'payment',
  MARequirementGeneric = 'ma_requirement_generic',
  MARequirementLegal = 'ma_requirement_legal',
  MARequirementNol = 'ma_requirement_nol',
  MALicense = 'ma_requirement_credential',
  RequirementType = 'requirement_type',
  Requirement_Status = 'requirement_status',
  SNPMarketReadiness = 'snp_ma_readiness',
  MarketReadiness = 'country_ma_readiness',
  MADocument = 'ma_document',
  MAFolder = 'ma_folder',
}

export interface IFormConfig {
  name: string;
  type: FormType;
  list: string;
}

export interface IMessageBanner {
  message: string;
  type: MessageType;
  visible: boolean;
}

export class Form {
  public static config: IFormConfig[] = [
    {
      name: 'Eutelsat Entity',
      type: FormType.EutelsatEntity,
      list: Consts.LISTS.EUTELSAT_ENTITY_URL
    },
    {
      name: 'Country',
      type: FormType.Country,
      list: Consts.LISTS.COUNTRY_URL
    },
    {
      name: 'Market',
      type: FormType.Market,
      list: Consts.LISTS.MARKET_URL
    },
    {
      name: 'Legal Service Provider',
      type: FormType.LegalServiceProvider,
      list: Consts.LISTS.LEGAL_SERVICE_PROVIDER_URL

    },
    {
      name: 'Authority',
      type: FormType.Authority,
      list: Consts.LISTS.AUTHORITY_URL
    },
    {
      name: 'Contact',
      type: FormType.Contact,
      list: Consts.LISTS.CONTACT_URL
    },
    {
      name: 'SNP',
      type: FormType.SNP,
      list: Consts.LISTS.SNP_URL
    },
    {
      name: 'TP',
      type: FormType.TP,
      list: Consts.LISTS.TP_URL
    },
    {
      name: 'DP',
      type: FormType.DP,
      list: Consts.LISTS.DISTRIBUTION_PARTNER_URL
    },
    {
      name: 'Fee',
      type: FormType.Fee,
      list: Consts.LISTS.FEE_URL
    },
    {
      name: 'Payment',
      type: FormType.Payment,
      list: Consts.LISTS.PAYMENT_URL
    },
    {
      name: 'MA Requirement Generic',
      type: FormType.MARequirementGeneric,
      list: Consts.LISTS.MA_REQUIREMENT_URL
    },
    {
      name: 'MA Requirement Legal',
      type: FormType.MARequirementLegal,
      list: Consts.LISTS.MA_REQUIREMENT_URL
    },
    {
      name: 'MA Requirement NOL',
      type: FormType.MARequirementNol,
      list: Consts.LISTS.MA_REQUIREMENT_URL
    },
    {
      name: 'MA Requirement Credential',
      type: FormType.MALicense,
      list: Consts.LISTS.MA_REQUIREMENT_URL
    },
    {
      name: 'RequirementType',
      type: FormType.RequirementType,
      list: Consts.LISTS.REQUIREMENT_TYPE_URL
    },
    {
      name: 'RequirementStatus',
      type: FormType.Requirement_Status,
      list: Consts.LISTS.REQUIREMENT_STATUS_URL
    },
    {
      name: 'SNP Market Readiness',
      type: FormType.SNPMarketReadiness,
      list: Consts.LISTS.MA_READINESS_URL
    },
    {
      name: 'Market Readiness',
      type: FormType.MarketReadiness,
      list: Consts.LISTS.MA_READINESS_URL
    },
    {
      name: 'MA Document',
      type: FormType.MADocument,
      list: Consts.LISTS.MA_DOCUMENT_URL
    },
    {
      name: 'MA Folder',
      type: FormType.MAFolder,
      list: Consts.LISTS.DOCUMENT_SPACE_URL
    },
    {
      name: 'Test',
      type: FormType.Test,
      list: Consts.LISTS.EUTELSAT_ENTITY_URL
    }
  ];
}

export enum FormOrigin {
  Menu,
  Datasheet
}

