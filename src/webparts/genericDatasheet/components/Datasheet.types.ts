import { PNPService } from '../../../common/services/PNPService';
import { Consts } from '../../../common/consts/Consts';
import { IDropdownOption } from '@fluentui/react';
import { UserService } from '../../../common/services/UserContextService';

export interface IGenericDatasheetProps {
  pnpService: PNPService;
  userService: UserService;
}

export enum DatasheetType {
  Country = 'country',
  Market = 'market',
  EutelsatEntity = 'eutelsat_entity',
  LegalSvcProvider = 'legal_svc_provider',
  TeleportPartner = 'tp',
  Contact = 'contact',
  Authority = 'authority',
  DistributionPartner = 'dp',
  // Nol = 'nol',
  MARequirement = 'ma_requirement',
  // LegalRequirement = 'legal_requirement',
  SNP = 'snp',
  // MALicenseRequirment = 'ma_license_requirement',
  // MAGenericRequirment = 'ma_generic_requirement',
  MARequirementGeneric = 'ma_requirement_generic',
  MARequirementLegal = 'ma_requirement_legal',
  MARequirementNol = 'ma_requirement_nol',
  MARequirementLicense = 'ma_requirement_credential',
  Payment = "payment",
  Fee = "fee"
}

export interface IDatasheetConfig {
  name: string;
  type: DatasheetType;
  list: string;
}

export class Datasheet {
  public static config: IDatasheetConfig[] = [
    {
      name: 'Countries Datasheet',
      type: DatasheetType.Country,
      list: Consts.LISTS.COUNTRY_URL
    },
    {
      name: 'Markets Datasheet',
      type: DatasheetType.Market,
      list: Consts.LISTS.MARKET_URL
    },
    {
      name: 'Eutelsat Entities Datasheet',
      type: DatasheetType.EutelsatEntity,
      list: Consts.LISTS.EUTELSAT_ENTITY_URL
    },
    {
      name: 'Legal SVC Datasheet',
      type: DatasheetType.LegalSvcProvider,
      list: Consts.LISTS.LEGAL_SERVICE_PROVIDER_URL
    },
    {
      name: 'Teleport Partners Datasheet',
      type: DatasheetType.TeleportPartner,
      list: Consts.LISTS.TELEPORT_PARTNER_URL
    },
    {
      name: 'Contact Datasheet',
      type: DatasheetType.Contact,
      list: Consts.LISTS.CONTACT_URL
    },
    {
      name: 'Authority Datasheet',
      type: DatasheetType.Authority,
      list: Consts.LISTS.AUTHORITY_URL
    },
    {
      name: 'DP Datasheet',
      type: DatasheetType.DistributionPartner,
      list: Consts.LISTS.DISTRIBUTION_PARTNER_URL
    },
    {
      name: 'Nol Datasheet',
      type: DatasheetType.MARequirementNol,
      list: Consts.LISTS.MA_REQUIREMENT_URL
    },
    {
      name: 'MA License Datasheet',
      type: DatasheetType.MARequirementLicense,
      list: Consts.LISTS.MA_REQUIREMENT_URL
    },
    {
      name: 'FEE Datasheet',
      type: DatasheetType.Fee,
      list: Consts.LISTS.FEE_URL
    },
    {
      name: 'MA Requirements Datasheet',
      type: DatasheetType.MARequirement,
      list: Consts.LISTS.MA_REQUIREMENT_URL
    },
    {
      name: 'SNP Datasheet',
      type: DatasheetType.SNP,
      list: Consts.LISTS.SNP_URL
    },
    {
      name: 'MA Generic Datasheet',
      type: DatasheetType.MARequirementGeneric,
      list: Consts.LISTS.MA_REQUIREMENT_URL
    },
    {
      name: 'Payment Datasheet',
      type: DatasheetType.Payment,
      list: Consts.LISTS.PAYMENT_URL
    },
    {
      name: 'Fee Datasheet',
      type: DatasheetType.Fee,
      list: Consts.LISTS.FEE_URL
    },
    {
      name: 'Legal Requirement Datasheet',
      type: DatasheetType.MARequirementLegal,
      list: Consts.LISTS.MA_REQUIREMENT_URL
    },
  ];
}

export interface ISanctionOption extends IDropdownOption {
  color: string;
  summary: string;
}
