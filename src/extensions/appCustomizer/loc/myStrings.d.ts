declare interface IAppCustomizerStrings {
  Title: string;
  LawFirmContentTypeName: string;
  LegalRepresentativeContentTypeName: string;
  OrganizationContentTypeName: string;
  RegulatorContentTypeName: string;
  AdministrationContentTypeName: string;
  SnpStargateContentTypeName: string;
  SnpEutelsatContentTypeName: string;
  SnpPartnerContentTypeName: string;
  MARequirementEutelsatTypeName: string;
  MARequirementTeleportPartnerTypeName: string;
  MARequirementDistributionPartnerTypeName: string;
  MARequirementCredentialsContentTypeName: string;
  MARequirementGenericContentTypeName: string;
  MARequirementNolContentTypeName: string;
  MARequirementLegalContentTypeName: string;
}

declare module 'AppCustomizerStrings' {
  const strings: IAppCustomizerStrings;
  export = strings;
}
