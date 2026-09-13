/* eslint-disable no-useless-escape */
interface IDOM {
  LAPTOP_SCREEN_OFFSET: number;
  MEDIUM_TABLET_OFFSET: number;
  TABLET_OFFSET: number;
  DESKTOP_OFFSET: number;
}
interface IREGEX {
  YOUTUBE_URL: string;
  STREAM_URL: string;
  STREAM_SP_URL: string;
}

interface ICOMMON_FIELDS {
  ID: string;
  TITLE: string;
  MODIFICATION_TIME: string;
  CREATION_TIME: string;
  FILE: string;
  FILE_LEAFREF: string;
  FILE_DIR_REF: string;
  FILE_REF: string;
  FILE_TYPE: string;
  ISFOLDER: string;
  LIST_ITEM_ALL_FIELDS: string;
  DOCUMENTS_SPACE: string;
  CONTENTTYPE_ID: string;
  CONTENTTYPE: string;
  URL: string;
  DESCRIPTION: string;

}

interface IEUT_ENTITY_FIELDS {
  STATUS: string;
  COMMENTS: string;
  COUNTRY: string;
  COUNTRY_ID: string;
  ADDRESS: string;
  SUMMARY: string;
}
interface IAUTHORITY_FIELDS {
  Id: any;
  TaxCatchAll: any;
  Term: any;
  COMMENTS: string;
  COUNTRY: string;
  COUNTRY_ID: string;
  ADDRESS: string;
  SUMMARY: string;
  GROUPING: string;
  ETU_OWNERS: string;
  ETU_OWNERS_ID: string;
  PORTAL: string;
}
interface ISANCTION_CATEGORY_FIELDS {
  SUMMARY: string;
  COLOR: string;
}
interface ICOUNTRY_FIELDS {
  IDENTIFIER: string;
  GROUPING: string;
  IS_SANCTIONED: string;
  SANCTION_CATEGORY: string;
  SANCTION_CATEGORY_ID: string;
  MARKET_ACCESS_CAPACITY: string;
  MARKET_ACCESS_PRIORITY: string;
  COMMENTS: string;
}
interface IMARKET_FIELDS {
  COUNTRY: string;
  COUNTRY_ID: string;
  GROUPING: string;
  IS_SANCTIONED: string;
  SANCTION_CATEGORY: string;
  SANCTION_CATEGORY_ID: string;
  MARKET_ACCESS_CAPACITY: string;
  MARKET_ACCESS_PRIORITY: string;
  COMMENTS: string;
}
interface ILEGAL_SVC_PROVIDER_FIELDS {
  STATUS: string;
  COMMENTS: string;
  COUNTRY_ID: string;
  COUNTRY: string;
  ADDRESS: string;
  SUMMARY: string;
  OWNERS_ID: string;
  OWNER: string;
}
interface ISNP_FIELDS {
  COMMENTS: string;
  COUNTRY: string;
  COUNTRY_ID: string;
  CITY: string;
  SUMMARY: string;
  EUTELSATENTITY: string;
  EUTELSATENTITY_ID: string;
  TELEPORTPARTNER: string;
  TELEPORTPARTNER_ID: string;
  STATUS: string;
}
interface ITELEPORT_PARTNER_FIELDS {
  STATUS: string;
  COMMENTS: string;
  MARKETS: string;
  MARKETS_ID: string;
  ADDRESS: string;
  SUMMARY: string;
  PORTAL: string;
  OWNERS_ID: string;
  OWNER: string;
}
interface IDISTRIBUTION_PARTNER_FIELDS {
  STATUS: string;
  COMMENTS: string;
  ADDRESS: string;
  SUMMARY: string;
  PORTAL: string;
  OWNERS_ID: string;
  OWNER: string;
}
interface ICONTACT_FIELDS {
  FIRST_NAME: string;
  LAST_NAME: string;
  SUMMARY: string;
  JOB_TITLE: string;
  DEPARTMENT: string;
  EMAIL: string;
  PHONE_NUMBER: string;
  COMMENTS: string;
  AUTHORITY: string;
  AUTHORITY_ID: string;
  LEGAL_SVC_PROVIDER: string;
  LEGAL_SVC_PROVIDER_ID: string;
  DISTRIBUTION_PARTNER: string;
  DISTRIBUTION_PARTNER_ID: string;
  TELEPORT_PARTNER: string;
  TELEPORT_PARTNER_ID: string;
}
interface IFEE_FIELDS {
  CATEGORY: string;
  SUMMARY: string;
  SNP: string;
  SNP_ID: string;
  MA_REQUIREMENT: string;
  MA_REQUIREMENT_ID: string;
  CHARGEDBY: string;
  CURRENCY: string;
  CURRENCY_ID: string;
  CURRENCY_ISOCODE: string;
  VATFREECOST: string;
  VAT: string;
  VATRATE: string;
  COSTTYPE: string;
  PONONPO: string;
  STATUS: string;
  COMMENTS: string;
  AUTHORITY: string;
  AUTHORITY_ID: string;
  LEGALSERVICEPROVIDER: string;
  LEGALSERVICEPROVIDER_ID: string;
  TELEPORTPARTNER: string;
  TELEPORTPARTNER_ID: string;
  DUEDATE: string;
  REPEATINGFEE: string;
  RECURRENCEPATTERN: string;
  RECURRENCEPATTERN_ID: string;
  FEETYPE: string;
  VATFREE: string;
}
interface IPAYMENT_FIELDS {
  SUMMARY: string;
  FEES: string;
  FEES_ID: string;
  BENEFICIARY: string;
  CONTACTS: string;
  CONTACTS_ID: string;
  CURRENCY: string;
  CURRENCY_ID: string;
  CURRENCY_ISOCODE: string;
  VATFREECOST: string;
  VAT: string;
  VATRATE: string;
  STATUS: string;
  COMMENTS: string;
  AUTHORITY: string;
  AUTHORITY_ID: string;
  LEGALSERVICEPROVIDER: string;
  LEGALSERVICEPROVIDER_ID: string;
  TELEPORTPARTNER: string;
  TELEPORTPARTNER_ID: string;
  DUEDATE: string;
  PAYMENTPAID: string;
}
interface INAMING {
  MA_SNP_READINESS_SUFFIX;
  MARKET_READINESS_SUFFIX;
}
interface IRECURRENCE_PATTERN_FIELDS {
  DAY: string;
  MONTH: string;
  ENDDATE: string;
  FREQUENCY: string;
  INTERVAL: string;
  WEEKDAY: string;
}
interface IREQUIREMENT_TYPE_FIELDS {
  CATEGORY: string;
  VERTICAL: string;
}
interface IREQUIREMENTRELATIONSHIP_TYPE_FIELDS {
  MA_REQUIREMENTRELATIONSHIPTYPE: string;
  TELEPORTPARTNER: string;
  TELEPORTPARTNER_ID: string;
  DISTRIBUTION_PARTNER: string;
  DISTRIBUTION_PARTNER_ID: string;
  LEGALSERVICEPROVIDER: string;
  LEGALSERVICEPROVIDER_ID: string;
  AUTHORITY: string;
  AUTHORITY_ID: string;
  SNP: string;
  SNP_ID: string;
}
interface IMAGENERIC_FIELDS {
  NAME: string;
  TYPE: string;
  SUMMARY: string;
  MARKETS: string;
  MARKETS_ID: string;
  RESPONSIBLEPARTY: string;
  RESPONSIBLE: string;
  EUTELSATOWNERS: string;
  EUTELSATOWNERS_ID: string;
  CONTACTS: string;
  CONTACTS_ID: string;
  SYNTHESIS: string;
  SYNTHESIS_ID: string;
  RAGSTATUS: string;
  VERTICALS: string;
  SNPS: string;
  SNPS_ID: string;
  GEOLEO: string;
  COMMENTS: string;
  TELEPORTPARTNER: string;
  TELEPORTPARTNER_ID: string;
  DISTRIBUTION_PARTNER: string;
  DISTRIBUTION_PARTNER_ID: string;
  REQUIREMENT_TYPE: string;
  REQUIREMENT_TYPE_ID: string;
  EUTELSATENTITY: string;
  EUTELSATENTITY_ID: string;
  EFFECTIVEDATE: string;
  ESTIMATEDDATE: string;
  APPLICATIONDATE: string;
  RELATEDITEMS: string;
  RELATEDITEMS_ID: string;
  REQUESTEDPARTY_TYPE: string;
  REQUESTEDPARTY_CONTACT: string;
  REQUESTEDPARTY_CONTACT_ID: string;
  AUTHORITY: string;
  AUTHORITY_ID: string;
  EXPIRATIONDATE: string;
  TRIALDEMO: string;
  RENEWALTERM: string;
  REFERENCE: string;
  RECURRENCE_PATTERN: string;
  RECURRENCE_PATTERN_ID: string;
  RESPONSIBLEPARTY_CONTACT: string;
  RESPONSIBLEPARTY_CONTACT_ID: string;
}
interface IMA_REQUIREMENT_LEGAL_FIELDS {
  REQUESTEDPARTY_TYPE: string;
  AUTHORITY: string;
  AUTHORITY_ID: string;
  REQUESTEDPARTY_CONTACT: string;
  REQUESTEDPARTY_CONTACT_ID: string;
}
interface IMA_REQUIREMENT_NOL_FIELDS {
  REQUESTEDPARTY_TYPE: string;
  AUTHORITY: string;
  AUTHORITY_ID: string;
  REQUESTEDPARTY_CONTACT: string;
  REQUESTEDPARTY_CONTACT_ID: string;
  LETTER1_SENTDATE: string;
  LETTER2_SENTDATE: string;
  LETTER1_MILESTONE: string;
  LETTER2_MILESTONE: string;
  RESPONSE_DATE: string;
  RESPONSE_SUMMARY: string;
  DEMO_TRIAL: string;
  REFERENCE: string;
  COMMENTS: string;

}
interface ISNP_MA_READINESS_FIELDS {
  SNP: string;
  SNP_ID: string;
  RESPONSIBLE_PARTY: string;
  OWNERS_ID: string;
  OWNER: string;
  DISTRIBUTION_PARTNER: string;
  DISTRIBUTION_PARTNER_ID: string;
  TELEPORT_PARTNER: string
  TELEPORT_PARTNER_ID: string
  RAGSTATUS: string;
  ESTIMATED_DATE: string;
  EFFECTIVE_DATE: string;

  COMMENTS: string;
}
interface ISPUser {
  IS_SITE_ADMIN: string;
  GROUPS: string;
}
interface ISPGroup {
  LOGIN_NAME: string;
}
interface IRenditions {
  RENDITION1: string;
  RENDITION2: string;
  RENDITION3: string;
  RENDITION4: string;
  RENDITION5: string;
  RENDITION6: string;
  RENDITION7: string;
  RENDITION8: string;
}
interface IFile {
  SERVER_RELATIVE_URL: string;
  RENDITIONS: IRenditions;
  PROPS: string;
}
interface IImage {
  HEIGHT: string;
  WIDTH: string;
}
interface ICONTENT_TYPES {
  LAW_FIRM: string;
  LEGAL_REPRESENTATIVE: string;
  EUTELSAL_CONTACT: string;
  LAW_FIRM_CONTACT: string;
  LEGAL_REPRESENTATIVE_CONTACT: string;
  ADMINISTRATION_CONTACT: string;
  ORGANIZATION_CONTACT: string;
  REGULATOR_CONTACT: string;
  TELEPORT_PARTNER_CONTACT: string;
  DISTRIBUTION_PARTNER_CONTACT: string;
  ADMINISTRATION: string;
  REGULATOR: string;
  ORGANIZATION: string;
  SNP: string;
  EUTELSAT_SNP: string;
  PARTNER_SNP: string;
  STARGATE: string;
  TELEPORT_PARTNER: string;
  DISTRIBUTION_PARTNER: string;
  SNP_FEE: string;
  MA_REQUIREMENT_FEE: string;
  MA_RequirementGeneric: string;
  MA_READINESS_SNP: string;
  MA_RequirementCredential: string;
  MA_REQUIREMENT_LEGAL: string;
  MA_REQUIREMENT_NOL: string;
  MA_READINESS_COUNTRY: string;
}
interface IMARKET_READINESS_FIELDS {
  MARKET: string;
  MARKET_ID: string;
  RESPONSIBLE_PARTY: string;
  OWNERS_ID: string;
  OWNER: string;
  DISTRIBUTION_PARTNER: string;
  DISTRIBUTION_PARTNER_ID: string;
  TELEPORT_PARTNER: string
  TELEPORT_PARTNER_ID: string;
  VERTICALS: string;
  ALL_VERTICALS_RAGSTATUS: string;
  ALL_VERTICALS_ESTIMATED_DATE: string;
  ALL_VERTICALS_EFFECTIVE_DATE: string;
  ALL_VERTICALS_COMMENTS: string;
  AVIATION_RAGSTATUS: string;
  AVIATION_ESTIMATED_DATE: string;
  AVIATION_EFFECTIVE_DATE: string;
  AVIATION_COMMENTS: string;
  SPACE_NW_RAGSTATUS: string;
  SPACE_NW_ESTIMATED_DATE: string;
  SPACE_NW_EFFECTIVE_DATE: string;
  SPACE_NW_COMMENTS: string;
  LAND_FIXED_RAGSTATUS: string;
  LAND_FIXED_EFFECTIVE_DATE: string;
  LAND_FIXED_ESTIMATED_DATE: string;
  LAND_FIXED_COMMENTS: string;
  LAND_MOBILITY_RAGSTATUS: string;
  LAND_MOBILITY_ESTIMATED_DATE: string;
  LAND_MOBILITY_EFFECTIVE_DATE: string;
  LAND_MOBILITY_COMMENTS: string;
  MARITIME_RAGSTATUS: string;
  MARITIME_ESTIMATED_DATE: string;
  MARITIME_EFFECTIVE_DATE: string;
  MARITIME_COMMENTS: string;

}
interface IMADOCUMENT_FIELDS {
  SUMMARY: string;
  CLASSIFICATION: string;
  FILE_PATH: string;
  FILE_TYPE: string;
  FILE_TYPE_ID: string;
  VERSION: string;
  KEYWORDS: string;
  KEYWORDS_ID: string;
  IS_ACTIVE: string;
  CATEGORY_ID: string;
  CATEGORY: string;
  UNIQUEID: string;
}
interface IFILETYPE_FIELDS {
  CATEGORY: string;
}
interface ITAXONOMY_FIELDS {
  TERM_ID: string;
  TERM_LABEL_FR: string;
  TERM_LABEL_EN: string;
}
interface IFIELDS {
  COMMON: ICOMMON_FIELDS;
  SP_USER: ISPUser;
  SP_GROUP: ISPGroup;
  FILE: IFile;
  IMAGE: IImage;
  EUT_ENTITY: IEUT_ENTITY_FIELDS;
  LEGAL_SVC_PROVIDER: ILEGAL_SVC_PROVIDER_FIELDS;
  AUTHORITY: IAUTHORITY_FIELDS;
  TAXONOMY: ITAXONOMY_FIELDS;
  SANCTION_CATEGORY: ISANCTION_CATEGORY_FIELDS;
  COUNTRY: ICOUNTRY_FIELDS;
  MARKET: IMARKET_FIELDS;
  SNP: ISNP_FIELDS;
  TELEPORT_PARTNER: ITELEPORT_PARTNER_FIELDS;
  DISTRIBUTION_PARTNER: IDISTRIBUTION_PARTNER_FIELDS;
  FEE: IFEE_FIELDS;
  PAYMENT: IPAYMENT_FIELDS;
  RECURRENCE_PATTERN: IRECURRENCE_PATTERN_FIELDS;
  REQUIREMENT_TYPE: IREQUIREMENT_TYPE_FIELDS;
  MAGENERIC: IMAGENERIC_FIELDS;
  MA_REQUIREMENT_LEGAL: IMA_REQUIREMENT_LEGAL_FIELDS;
  MA_REQUIREMENT_NOL: IMA_REQUIREMENT_NOL_FIELDS;
  CONTACT: ICONTACT_FIELDS;
  REQUIREMENTRELATIONSHIP_TYPE_FIELDS: IREQUIREMENTRELATIONSHIP_TYPE_FIELDS;
  SNP_MA_READINESS: ISNP_MA_READINESS_FIELDS;
  MARKET_READINESS: IMARKET_READINESS_FIELDS;
  MADOCUMENT: IMADOCUMENT_FIELDS;
  FILETYPE: IFILETYPE_FIELDS;
}
interface ILISTS {
  AUTHORITY_URL: string;
  LEGAL_SERVICE_PROVIDER_URL: string;
  COUNTRY_URL: string;
  MARKET_URL: string;
  SANCTION_CATEGORY_URL: string;
  EUTELSAT_ENTITY_URL: string;
  CONTACT_URL: string;
  SNP_URL: string;
  TP_URL: string;
  TELEPORT_PARTNER_URL: string;
  DISTRIBUTION_PARTNER_URL: string;
  FEE_URL: string;
  PAYMENT_URL: string;
  RECURRENCE_PATTERN_URL: string;
  MA_REQUIREMENT_URL: string;
  DOCUMENT_SPACE_URL: string;
  CURRENCY_URL: string;
  TAXONOMY_HIDDEN: string;
  REQUIREMENT_TYPE_URL: string;
  REQUIREMENT_STATUS_URL: string;
  REQUIREMENT_RELTIONSHIP_URL: string;
  MA_READINESS_URL: string;
  MA_DOCUMENT_URL: string;
  FILETYPE_URL: string;
  KEYWORD_URL: string;

}
interface IPAGES {
  DATASHEET_URL: string;
}

interface IGROUPS {
  FinanceContributors: string;
  FinanceVisitors: string;
  Contributors: string;
  Owners: string;
  Visitors: string;
}

interface IOFFICE {
  OFFICE_SIGNOUT_URL: string;
  OFFICE_SIGNIN_URL: string;
  OFFICE_ACCOUNT_URL: string;
  OFFICE_DELVE_PROFILE_URL: string;
  OFFICE_DELVE_PROFILE_IMAGE_BASE_URL: string;
}

interface IDOCUMENT_SPACE {
  EUT_ENTITY_FOLDER_URL: string;
  MARKET_FOLDER_URL: string;
  AUTHORITY_ADMINISTRATION_FOLDER_URL: string;
  AUTHORITY_ORGANIZATION_FOLDER_URL: string;
  AUTHORITY_REGULATOR_FOLDER_URL: string;
  LEGAL_LAW_FIRMS_FOLDER_URL: string;
  LEGAL_REPRESENTATIVES_FOLDER_URL: string;
  SNP_STARGATE_FOLDER_URL: string;
  SNP_EUTELSAT_SNP_FOLDER_URL: string;
  SNP_PARTNER_SNP_FOLDER_URL: string;
  TELEPORT_PARTNER_FOLDER_URL: string;
  DISTRIBUTION_PARTNER_FOLDER_URL: string;
  FEE_SNP_FOLDER_URL: string;
  FEE_MA_REQUIREMENT_fOLDER_URL: string;
  PAYMENT_FOLDER_URL: string;
  MAGENERIC_FOLDER_URL: string;
  LEGAL_REQUIREMENTS_FOLDER_URL: string;
  NOL_REQUIREMENTS_FOLDER_URL: string;
  LICENCE_REQUIREMENTS_FOLDER_URL: string;
}

interface IQUERY {
  PARAM_FORM_TYPE: string;
  PARAM_VIEW_TYPE: string;
  PARAM_DATASHEET_TYPE: string;
  PARAM_DATASHEET_ITEM_ID: string;
}
export class Consts {
  public static DOM: IDOM = {
    'LAPTOP_SCREEN_OFFSET': 960,
    'MEDIUM_TABLET_OFFSET': 800,
    'TABLET_OFFSET': 640,
    'DESKTOP_OFFSET': 1030
  };
  public static DOCUMENT_SPACE: IDOCUMENT_SPACE = {
    'EUT_ENTITY_FOLDER_URL': '/Eutelsat Entities',
    'MARKET_FOLDER_URL': '/Markets',
    'AUTHORITY_ADMINISTRATION_FOLDER_URL': '/Authorities/Administrations',
    'AUTHORITY_ORGANIZATION_FOLDER_URL': '/Authorities/Organizations',
    'AUTHORITY_REGULATOR_FOLDER_URL': '/Authorities/Regulators',
    'LEGAL_LAW_FIRMS_FOLDER_URL': '/Legal Service Providers/Law Firms',
    'LEGAL_REPRESENTATIVES_FOLDER_URL': '/Legal Service Providers/Legal Representatives',
    'SNP_EUTELSAT_SNP_FOLDER_URL': '/SNPs/Eutelsat',
    'SNP_PARTNER_SNP_FOLDER_URL': '/SNPs/Partners',
    'SNP_STARGATE_FOLDER_URL': '/SNPs/Stargates',
    'TELEPORT_PARTNER_FOLDER_URL': '/Teleport Partners',
    'DISTRIBUTION_PARTNER_FOLDER_URL': '/Distribution Partners',
    'FEE_SNP_FOLDER_URL': '/Fees/SNPs',
    'FEE_MA_REQUIREMENT_fOLDER_URL': '/Fees/MA Requirements',
    'PAYMENT_FOLDER_URL': '/Payments',
    'MAGENERIC_FOLDER_URL': '/MA Requirements/Generic',
    'LEGAL_REQUIREMENTS_FOLDER_URL': '/MA Requirements/Legal',
    'NOL_REQUIREMENTS_FOLDER_URL': '/MA Requirements/Nol',
    'LICENCE_REQUIREMENTS_FOLDER_URL': '/MA Requirements/Credential'

  };
  public static REGEX: IREGEX = {
    YOUTUBE_URL: '^(https?\:\/\/)?((www\.)?youtube\.com|youtu\.be)\/.+$',
    STREAM_URL: '^(https?\:\/\/)?(web\.microsoftstream\.com\/video)\/.+$',
    STREAM_SP_URL: '^(https?\:\/\/)?(eutelsatgroup\.sharepoint\.com\/)(.*)(\/stream.+$)'
  };
  public static NAMING: INAMING = {
    MA_SNP_READINESS_SUFFIX: 'Market Readiness',
    MARKET_READINESS_SUFFIX: 'Market Readiness',

  };
  public static FIELDS: IFIELDS = {
    COMMON: {
      'ID': 'Id',
      'TITLE': 'Title',
      'MODIFICATION_TIME': 'Modified',
      'CREATION_TIME': 'Created',
      'FILE': 'File',
      'DOCUMENTS_SPACE': 'MA_DocumentsSpace',
      'CONTENTTYPE_ID': 'ContentTypeId',
      'CONTENTTYPE': 'ContentType',
      'URL': 'Url',
      'DESCRIPTION': 'Description',
      'FILE_LEAFREF': 'FileLeafRef',
      'FILE_REF': 'FileRef',
      'ISFOLDER': 'FSObjType',
      'LIST_ITEM_ALL_FIELDS': 'ListItemAllFields',
      'FILE_TYPE': 'File_x0020_Type',
      'FILE_DIR_REF': 'FileDirRef'
    },
    EUT_ENTITY: {
      'STATUS': 'MA_EutelsatEntityStatus',
      'COMMENTS': 'MA_Comments',
      'COUNTRY_ID': 'MA_CountryId',
      'COUNTRY': 'MA_Country',
      'ADDRESS': 'MA_Address',
      'SUMMARY': 'MA_Summary'
    },
    AUTHORITY: {
      'COMMENTS': 'MA_Comments',
      'COUNTRY_ID': 'MA_CountryId',
      'COUNTRY': 'MA_Country',
      'ADDRESS': 'MA_Address',
      'SUMMARY': 'MA_Summary',
      'GROUPING': 'MA_Grouping',
      'ETU_OWNERS': 'MA_Owner',
      'ETU_OWNERS_ID': 'MA_OwnerId',
      'PORTAL': 'MA_Portal',
      'TaxCatchAll': 'TaxCatchAll',
      'Term': 'Term',
      'Id': 'Id'
    },
    LEGAL_SVC_PROVIDER: {
      'STATUS': 'MA_LegalSvcProviderStatus',
      'COMMENTS': 'MA_Comments',
      'COUNTRY_ID': 'MA_CountryId',
      'COUNTRY': 'MA_Country',
      'ADDRESS': 'MA_Address',
      'SUMMARY': 'MA_Summary',
      'OWNERS_ID': 'MA_OwnerId',
      'OWNER': 'MA_Owner'
    },
    SNP: {
      'COMMENTS': 'MA_Comments',
      'COUNTRY_ID': 'MA_CountryId',
      'COUNTRY': 'MA_Country',
      'SUMMARY': 'MA_Summary',
      'EUTELSATENTITY': 'MA_EutelsatEntity',
      'EUTELSATENTITY_ID': 'MA_EutelsatEntityId',
      'TELEPORTPARTNER': 'MA_TeleportPartner',
      'TELEPORTPARTNER_ID': 'MA_TeleportPartnerId',
      'STATUS': 'MA_SNPStatus',
      'CITY': 'MA_City'
    },
    TELEPORT_PARTNER: {
      'STATUS': 'MA_TeleportPartnerStatus',
      'COMMENTS': 'MA_Comments',
      'MARKETS_ID': 'MA_MarketId',
      'MARKETS': 'MA_Market',
      'ADDRESS': 'MA_Address',
      'SUMMARY': 'MA_Summary',
      'OWNERS_ID': 'MA_OwnerId',
      'OWNER': 'MA_Owner',
      'PORTAL': 'MA_Portal'
    },
    DISTRIBUTION_PARTNER: {
      'STATUS': 'MA_DistributionPartnerStatus',
      'COMMENTS': 'MA_Comments',
      'ADDRESS': 'MA_Address',
      'SUMMARY': 'MA_Summary',
      'OWNERS_ID': 'MA_OwnerId',
      'OWNER': 'MA_Owner',
      'PORTAL': 'MA_Portal'
    },
    FEE: {
      'CATEGORY': 'MA_FeeType',
      'SUMMARY': 'MA_Summary',
      'CHARGEDBY': 'MA_FeeChargedBy',
      'CURRENCY': 'MA_Currency',
      'CURRENCY_ID': 'MA_CurrencyId',
      'CURRENCY_ISOCODE': 'MA_CurrencyIsoCode',
      'VATFREECOST': 'MA_VatFree',
      'VAT': 'MA_Vat',
      'VATRATE': 'MA_VatRate',
      'COSTTYPE': 'MA_CostClassification',
      'PONONPO': 'MA_PoType',
      'STATUS': 'MA_FeeStatus',
      'COMMENTS': 'MA_Comments',
      'MA_REQUIREMENT': 'MA_Requirement',
      'MA_REQUIREMENT_ID': 'MA_RequirementId',
      'SNP': 'MA_SNP',
      'SNP_ID': 'MA_SNPId',
      'AUTHORITY': 'MA_Authority',
      'AUTHORITY_ID': 'MA_AuthorityId',
      'LEGALSERVICEPROVIDER': 'MA_LegalSvcProvider',
      'LEGALSERVICEPROVIDER_ID': 'MA_LegalSvcProviderId',
      'TELEPORTPARTNER': 'MA_TeleportPartner',
      'TELEPORTPARTNER_ID': 'MA_TeleportPartnerId',
      'DUEDATE': 'MA_DueDate',
      'REPEATINGFEE': 'MA_IsRepeatingFee',
      'RECURRENCEPATTERN': 'MA_RecurrencePattern',
      'RECURRENCEPATTERN_ID': 'MA_RecurrencePatternId',
      'FEETYPE': 'MA_FeeType',
      'VATFREE': 'MA_VatFree',
    },
    PAYMENT: {
      'FEES': 'MA_Fee',
      'FEES_ID': 'MA_FeeId',
      'SUMMARY': 'MA_Summary',
      'BENEFICIARY': 'MA_FeeChargedBy',
      'CURRENCY': 'MA_Currency',
      'CURRENCY_ID': 'MA_CurrencyId',
      'CURRENCY_ISOCODE': 'MA_CurrencyIsoCode',
      'VATFREECOST': 'MA_VatFree',
      'VAT': 'MA_Vat',
      'VATRATE': 'MA_VatRate',
      'CONTACTS': 'MA_Contact',
      'CONTACTS_ID': 'MA_ContactId',
      'DUEDATE': 'MA_DueDate',
      'STATUS': 'MA_PaymentStatus',
      'PAYMENTPAID': 'MA_PaymentDate',
      'COMMENTS': 'MA_Comments',
      'AUTHORITY': 'MA_Authority',
      'AUTHORITY_ID': 'MA_AuthorityId',
      'LEGALSERVICEPROVIDER': 'MA_LegalSvcProvider',
      'LEGALSERVICEPROVIDER_ID': 'MA_LegalSvcProviderId',
      'TELEPORTPARTNER': 'MA_TeleportPartner',
      'TELEPORTPARTNER_ID': 'MA_TeleportPartnerId'
    },
    REQUIREMENTRELATIONSHIP_TYPE_FIELDS: {
      'MA_REQUIREMENTRELATIONSHIPTYPE': 'MA_RequirementRelationshipType',
      'TELEPORTPARTNER': 'MA_TeleportPartner',
      'TELEPORTPARTNER_ID': 'MA_TeleportPartnerId',
      'DISTRIBUTION_PARTNER': 'MA_DistributionPartner',
      'DISTRIBUTION_PARTNER_ID': 'MA_DistributionPartnerId',
      'LEGALSERVICEPROVIDER': 'MA_LegalSvcProvider',
      'LEGALSERVICEPROVIDER_ID': 'MA_LegalSvcProviderId',
      'AUTHORITY': 'MA_Authority',
      'AUTHORITY_ID': 'MA_AuthorityId',
      'SNP': 'MA_SNP',
      'SNP_ID': 'MA_SNPId',
    },
    MAGENERIC: {
      'NAME': 'MA_Name',
      'TYPE': 'MA_Type',
      'SUMMARY': 'MA_Summary',
      'MARKETS_ID': 'MA_MarketId',
      'MARKETS': 'MA_Market',
      'RESPONSIBLEPARTY': 'MA_RequirementResponsible',
      'RESPONSIBLE': 'MA_REQUIREMENT_RESPONSIBLE',
      'EUTELSATOWNERS': 'MA_Owner',
      'EUTELSATOWNERS_ID': 'MA_OwnerId',
      'CONTACTS': 'MA_ResponsiblePartyContact',
      'CONTACTS_ID': 'MA_ResponsiblePartyContactId',
      'SYNTHESIS': 'MA_RequirementStatus',
      'SYNTHESIS_ID': 'MA_RequirementStatusId',
      'RAGSTATUS': 'MA_RAGStatus',
      'VERTICALS': 'MA_RequirementVertical',
      'SNPS': 'MA_SNP',
      'SNPS_ID': 'MA_SNPId',
      'GEOLEO': 'MA_SatelliteOrbit',
      'COMMENTS': 'MA_Comments',
      'TELEPORTPARTNER': 'MA_TeleportPartner',
      'TELEPORTPARTNER_ID': 'MA_TeleportPartnerId',
      'DISTRIBUTION_PARTNER': 'MA_DistributionPartner',
      'DISTRIBUTION_PARTNER_ID': 'MA_DistributionPartnerId',
      'EUTELSATENTITY': 'MA_EutelsatEntity',
      'EUTELSATENTITY_ID': 'MA_EutelsatEntityId',
      'REQUIREMENT_TYPE': 'MA_RequirementType',
      'REQUIREMENT_TYPE_ID': 'MA_RequirementTypeId',
      'EFFECTIVEDATE': 'MA_RequirementEffectiveDate',
      'ESTIMATEDDATE': 'MA_RequirementEstimatedDate',
      'APPLICATIONDATE': 'MA_RequirementApplicationDate',
      'RELATEDITEMS': 'MA_RequirementRelationship',
      'RELATEDITEMS_ID': 'MA_RequirementRelationshipId',
      'REQUESTEDPARTY_TYPE': 'MA_RequirementRequestedPartyType',
      'REQUESTEDPARTY_CONTACT': 'MA_RequestedPartyContact',
      'REQUESTEDPARTY_CONTACT_ID': 'MA_RequestedPartyContactId',
      'AUTHORITY': 'MA_Authority',
      'AUTHORITY_ID': 'MA_AuthorityId',
      'EXPIRATIONDATE': 'MA_RequirementExpirationDate',
      'TRIALDEMO': 'MA_IsTrial',
      'RENEWALTERM': 'MA_RenewalTerm',
      'REFERENCE': 'MA_Reference',
      'RECURRENCE_PATTERN': 'MA_RecurrencePattern',
      'RECURRENCE_PATTERN_ID': 'MA_RecurrencePatternId',
      'RESPONSIBLEPARTY_CONTACT': 'MA_ResponsiblePartyContact',
      'RESPONSIBLEPARTY_CONTACT_ID': 'MA_ResponsiblePartyContactId'
    },
    MA_REQUIREMENT_LEGAL: {
      'AUTHORITY': 'MA_Authority',
      'AUTHORITY_ID': 'MA_AuthorityId',
      'REQUESTEDPARTY_TYPE': 'MA_RequirementRequestedPartyType',
      'REQUESTEDPARTY_CONTACT': 'MA_RequestedPartyContact',
      'REQUESTEDPARTY_CONTACT_ID': 'MA_RequestedPartyContactId'
    },
    MA_REQUIREMENT_NOL: {
      'AUTHORITY': 'MA_Authority',
      'AUTHORITY_ID': 'MA_AuthorityId',
      'REQUESTEDPARTY_TYPE': 'MA_RequirementRequestedPartyType',
      'REQUESTEDPARTY_CONTACT': 'MA_RequestedPartyContact',
      'REQUESTEDPARTY_CONTACT_ID': 'MA_RequestedPartyContactId',
      'LETTER1_SENTDATE': 'MA_Letter1SentDate',
      'LETTER2_SENTDATE': 'MA_Letter2SentDate',
      'LETTER1_MILESTONE': 'MA_Letter1Milestone',
      'LETTER2_MILESTONE': 'MA_Letter2Milestone',
      'RESPONSE_DATE': 'MA_ResponseDate',
      'RESPONSE_SUMMARY': 'MA_ResponseSummary',
      'DEMO_TRIAL': "MA_IsTrial",
      'REFERENCE': "MA_Reference",
      'COMMENTS': "MA_Comments"

    },
    REQUIREMENT_TYPE: {
      'CATEGORY': 'MA_RequirementCategory',
      'VERTICAL': 'MA_RequirementVertical'
    },
    RECURRENCE_PATTERN: {
      'DAY': "MA_RecurrenceDay",
      'ENDDATE': "MA_RecurrenceEndDate",
      'FREQUENCY': "MA_RecurrenceFrequency",
      'INTERVAL': "MA_RecurrenceInterval",
      'WEEKDAY': "MA_RecurrenceWeekDay",
      'MONTH': "MA_RecurrenceMonth"

    },
    CONTACT: {
      'FIRST_NAME': 'MA_FirstName',
      'LAST_NAME': 'MA_LastName',
      'SUMMARY': 'MA_Summary',
      'JOB_TITLE': 'MA_JobTitle',
      'DEPARTMENT': 'MA_Department',
      'EMAIL': 'MA_Email',
      'PHONE_NUMBER': 'MA_PhoneNumber',
      'COMMENTS': 'MA_Comments',
      'AUTHORITY': 'MA_Authority',
      'AUTHORITY_ID': 'MA_AuthorityId',
      'LEGAL_SVC_PROVIDER': 'MA_LegalSvcProvider',
      'LEGAL_SVC_PROVIDER_ID': 'MA_LegalSvcProviderId',
      'DISTRIBUTION_PARTNER': 'MA_DistributionPartner',
      'DISTRIBUTION_PARTNER_ID': 'MA_DistributionPartnerId',
      'TELEPORT_PARTNER': 'MA_TeleportPartner',
      'TELEPORT_PARTNER_ID': 'MA_TeleportPartnerId'
    },
    SNP_MA_READINESS: {
      'RESPONSIBLE_PARTY': 'MA_ReadinessResponsible',
      'DISTRIBUTION_PARTNER': 'MA_DistributionPartner',
      'DISTRIBUTION_PARTNER_ID': 'MA_DistributionPartnerId',
      'TELEPORT_PARTNER': 'MA_TeleportPartner',
      'TELEPORT_PARTNER_ID': 'MA_TeleportPartnerId',
      'OWNERS_ID': 'MA_OwnerId',
      'OWNER': 'MA_Owner',
      'SNP': 'MA_SNP',
      'SNP_ID': "MA_SNPId",
      'RAGSTATUS': 'MA_SNPRAGStatus',
      'ESTIMATED_DATE': 'MA_SNPEstimatedDate',
      'EFFECTIVE_DATE': 'MA_SNPEffectiveDate',
      'COMMENTS': 'MA_SNPComments'
    },
    MARKET_READINESS: {
      'RESPONSIBLE_PARTY': 'MA_ReadinessResponsible',
      'MARKET_ID': 'MA_MarketId',
      'MARKET': 'MA_Market',
      'DISTRIBUTION_PARTNER': 'MA_DistributionPartner',
      'DISTRIBUTION_PARTNER_ID': 'MA_DistributionPartnerId',
      'TELEPORT_PARTNER': 'MA_TeleportPartner',
      'TELEPORT_PARTNER_ID': 'MA_TeleportPartnerId',
      'OWNERS_ID': 'MA_OwnerId',
      'OWNER': 'MA_Owner',
      'VERTICALS': 'MA_RequirementVertical',
      'ALL_VERTICALS_RAGSTATUS': 'MA_ReadinessRAGStatus',
      'ALL_VERTICALS_ESTIMATED_DATE': 'MA_ReadinessEstimatedDate',
      'ALL_VERTICALS_EFFECTIVE_DATE': 'MA_ReadinessEffectiveDate',
      'ALL_VERTICALS_COMMENTS': 'MA_ReadinessComments',
      'AVIATION_RAGSTATUS': 'MA_AviationRAGStatus',
      'AVIATION_ESTIMATED_DATE': 'MA_AviationEstimatedDate',
      'AVIATION_EFFECTIVE_DATE': 'MA_AviationEffectiveDate',
      'AVIATION_COMMENTS': 'MA_AviationComments',
      'SPACE_NW_RAGSTATUS': 'MA_SpaceRAGStatus',
      'SPACE_NW_ESTIMATED_DATE': 'MA_SpaceEstimatedDate',
      'SPACE_NW_EFFECTIVE_DATE': 'MA_SpaceEffectiveDate',
      'SPACE_NW_COMMENTS': 'MA_SpaceComments',
      'LAND_FIXED_RAGSTATUS': 'MA_LandFixedRAGStatus',
      'LAND_FIXED_ESTIMATED_DATE': 'MA_LandFixedEstimatedDate',
      'LAND_FIXED_EFFECTIVE_DATE': 'MA_LandFixedEffectiveDate',
      'LAND_FIXED_COMMENTS': 'MA_LandFixedComments',
      'LAND_MOBILITY_RAGSTATUS': 'MA_LandMobilityRAGStatus',
      'LAND_MOBILITY_ESTIMATED_DATE': 'MA_LandMobilityEstimatedDate',
      'LAND_MOBILITY_EFFECTIVE_DATE': 'MA_LandMobilityEffectiveDate',
      'LAND_MOBILITY_COMMENTS': 'MA_LandMobilityComments',
      'MARITIME_RAGSTATUS': 'MA_MaritimeRAGStatus',
      'MARITIME_ESTIMATED_DATE': 'MA_MaritimeEstimatedDate',
      'MARITIME_EFFECTIVE_DATE': 'MA_MaritimeEffectiveDate',
      'MARITIME_COMMENTS': 'MA_MaritimeComments'
    },
    MADOCUMENT: {
      'CLASSIFICATION': 'MA_DocumentClassification',
      'FILE_PATH': 'FileDirRef',
      'SUMMARY': 'MA_Summary',
      'FILE_TYPE': 'MA_FileType',
      'FILE_TYPE_ID': 'MA_FileTypeId',
      'VERSION': 'MA_DocumentVersion',
      'KEYWORDS': 'MA_Keyword',
      'KEYWORDS_ID': 'MA_KeywordId',
      'IS_ACTIVE': 'MA_IsActive',
      'CATEGORY': 'MA_DocumentCategory',
      'CATEGORY_ID': 'MA_DocumentCategoryId',
      'UNIQUEID': 'UniqueId'
    },
    FILETYPE: {
      'CATEGORY': 'MA_DocumentCategory'
    },
    SP_USER: {
      'IS_SITE_ADMIN': 'IsSiteAdmin',
      'GROUPS': 'Groups'
    },
    SP_GROUP: {
      'LOGIN_NAME': 'LoginName',
    },
    TAXONOMY: {
      TERM_ID: 'IdForTerm',
      TERM_LABEL_FR: 'Term1036',
      TERM_LABEL_EN: 'Term1033'
    },
    SANCTION_CATEGORY: {
      SUMMARY: 'MA_Summary',
      COLOR: 'MA_Color'
    },
    COUNTRY: {
      IDENTIFIER: 'MA_Identifier',
      GROUPING: 'MA_Grouping',
      IS_SANCTIONED: 'MA_IsSanctioned',
      SANCTION_CATEGORY: 'MA_SanctionCategory',
      SANCTION_CATEGORY_ID: 'MA_SanctionCategoryId',
      MARKET_ACCESS_CAPACITY: 'MA_Capacity',
      MARKET_ACCESS_PRIORITY: 'MA_Priority',
      COMMENTS: 'MA_Comments'
    },
    MARKET: {
      COUNTRY: 'MA_Country',
      COUNTRY_ID: 'MA_CountryId',
      GROUPING: 'MA_Grouping',
      IS_SANCTIONED: 'MA_IsSanctioned',
      SANCTION_CATEGORY: 'MA_SanctionCategory',
      SANCTION_CATEGORY_ID: 'MA_SanctionCategoryId',
      MARKET_ACCESS_CAPACITY: 'MA_Capacity',
      MARKET_ACCESS_PRIORITY: 'MA_Priority',
      COMMENTS: 'MA_Comments'
    },
    FILE: {
      'SERVER_RELATIVE_URL': 'ServerRelativeUrl',
      'PROPS': 'Properties',
      'RENDITIONS': {
        'RENDITION1': 'PublishingImageRendition1',
        'RENDITION2': 'PublishingImageRendition2',
        'RENDITION3': 'PublishingImageRendition3',
        'RENDITION4': 'PublishingImageRendition4',
        'RENDITION5': 'PublishingImageRendition5',
        'RENDITION6': 'PublishingImageRendition6',
        'RENDITION7': 'PublishingImageRendition7',
        'RENDITION8': 'PublishingImageRendition8'
      }
    },
    IMAGE: {
      'HEIGHT': 'ImageHeight',
      'WIDTH': 'ImageWidth'
    }
  };
  public static LISTS: ILISTS = {
    'AUTHORITY_URL': 'Lists/authority',
    'LEGAL_SERVICE_PROVIDER_URL': 'Lists/legal_svc_provider',
    'COUNTRY_URL': 'Lists/country',
    'MARKET_URL': 'Lists/market',
    'EUTELSAT_ENTITY_URL': 'Lists/eutelsat_entity',
    'DOCUMENT_SPACE_URL': 'documents_space',
    'TAXONOMY_HIDDEN': 'Lists/TaxonomyHiddenList',
    'CONTACT_URL': 'Lists/contact',
    'SANCTION_CATEGORY_URL': 'Lists/sanction_category',
    'SNP_URL': 'Lists/snp',
    'TP_URL': 'Lists/teleport_partner',
    'TELEPORT_PARTNER_URL': 'Lists/teleport_partner',
    'DISTRIBUTION_PARTNER_URL': 'Lists/distribution_partner',
    'FEE_URL': 'Lists/fee',
    'PAYMENT_URL': 'Lists/payment',
    'RECURRENCE_PATTERN_URL': 'Lists/recurrence_pattern',
    'MA_REQUIREMENT_URL': 'Lists/requirement',
    'CURRENCY_URL': 'Lists/currency',
    'REQUIREMENT_TYPE_URL': 'Lists/requirement_type',
    'REQUIREMENT_STATUS_URL': 'Lists/requirement_status',
    'REQUIREMENT_RELTIONSHIP_URL': 'Lists/requirement_relationship',
    'MA_READINESS_URL': 'Lists/ma_readiness',
    'MA_DOCUMENT_URL': 'documents_space',
    'FILETYPE_URL': 'Lists/file_type',
    'KEYWORD_URL': 'Lists/keyword',
  };
  public static Pages: IPAGES = {
    'DATASHEET_URL': 'SitePages/Datasheet.aspx',
  };
  public static Groups: IGROUPS = {
    FinanceContributors: 'Finance Contributors',
    FinanceVisitors: 'Finance Visitors',
    Contributors: 'MADB Contributors',
    Owners: 'MADB Owners',
    Visitors: 'MADB Visitors'
  };

  public static OFFICE: IOFFICE = {
    'OFFICE_SIGNOUT_URL': '/_layouts/15/SignOut.aspx',
    'OFFICE_SIGNIN_URL': 'https://www.office.com/login?prompt=select_account&ru=%2Flaunch%2Fsharepoint',
    'OFFICE_ACCOUNT_URL': 'https://myaccount.microsoft.com/?ref=MeControl',
    'OFFICE_DELVE_PROFILE_URL': 'https://eutelsatgroup-my.sharepoint.com/person.aspx',
    'OFFICE_DELVE_PROFILE_IMAGE_BASE_URL': 'https://delve.office.com/mt/v3/people/profileimage?size={0}&userId={1}'
  };

  public static QUERY: IQUERY = {
    'PARAM_FORM_TYPE': 'type',
    'PARAM_VIEW_TYPE': 'view',
    'PARAM_DATASHEET_TYPE': 'config',
    'PARAM_DATASHEET_ITEM_ID': 'itemId',
  };
  public static CONTENT_TYPES: ICONTENT_TYPES = {
    'LAW_FIRM': '0x010002E1FD95664234498E6FD54EB499E33A01',
    'LEGAL_REPRESENTATIVE': '0x010002E1FD95664234498E6FD54EB499E33A02',
    'ADMINISTRATION': '0x0100DD016807C9F8854299FFFF08E0D9D2C502',
    'ORGANIZATION': '0x0100DD016807C9F8854299FFFF08E0D9D2C503',
    'REGULATOR': '0x0100DD016807C9F8854299FFFF08E0D9D2C501',
    'EUTELSAL_CONTACT': '0x01009683AA37F0D20E469880B9832370E82605',
    'LAW_FIRM_CONTACT': '0x01009683AA37F0D20E469880B9832370E8260102',
    'LEGAL_REPRESENTATIVE_CONTACT': '0x01009683AA37F0D20E469880B9832370E8260101',
    'ADMINISTRATION_CONTACT': '0x01009683AA37F0D20E469880B9832370E8260202',
    'ORGANIZATION_CONTACT': '0x01009683AA37F0D20E469880B9832370E8260203',
    'REGULATOR_CONTACT': '0x01009683AA37F0D20E469880B9832370E8260201',
    'TELEPORT_PARTNER_CONTACT': '0x01009683AA37F0D20E469880B9832370E82603',
    'DISTRIBUTION_PARTNER_CONTACT': '0x01009683AA37F0D20E469880B9832370E82604',
    'SNP': '0x01007DBD40A3236C4447848111DE54A6D863',
    'EUTELSAT_SNP': '0x01007DBD40A3236C4447848111DE54A6D86302',
    'PARTNER_SNP': '0x01007DBD40A3236C4447848111DE54A6D86303',
    'STARGATE': '0x01007DBD40A3236C4447848111DE54A6D86301',
    'TELEPORT_PARTNER': '0x0100FAB3C8B22A402D4296663A098E9A9FF5',
    'DISTRIBUTION_PARTNER': '0x010025FCA5695FFF2D4CA3BD04B70D983932',
    'SNP_FEE': '0x010092D96D11717F014497021E5516FB818201',
    'MA_REQUIREMENT_FEE': '0x010092D96D11717F014497021E5516FB818202',
    'MA_RequirementGeneric': '0x01009E58F9850C7D4144AD7FB9331AD6A24801',
    'MA_READINESS_SNP': '0x0100518D373795382B4BAC75508AC0E73D3302',
    'MA_RequirementCredential': '0x01009E58F9850C7D4144AD7FB9331AD6A24804',
    'MA_REQUIREMENT_LEGAL': '0x01009E58F9850C7D4144AD7FB9331AD6A24802',
    'MA_REQUIREMENT_NOL': '0x01009E58F9850C7D4144AD7FB9331AD6A24803',
    'MA_READINESS_COUNTRY': '0x0100518D373795382B4BAC75508AC0E73D3301',

  };
}
