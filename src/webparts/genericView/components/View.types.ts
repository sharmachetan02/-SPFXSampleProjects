import { PNPService } from '../../../common/services/PNPService';
import { Consts } from '../../../common/consts/Consts';
import { IFilter } from '../../../common/models/IFilter';
import { IColumn } from 'office-ui-fabric-react';
import { UserService } from '../../../common/services/UserContextService';

export interface IGenericViewProps {
  pnpService: PNPService;
  userService: UserService;
}

export enum ViewType {
  Country = 'country',
  Market = 'market',
  EutelsatEntity = 'eutelsat_entity',
  LegalSvcProvider = 'legal_svc_provider',
  TeleportPartner = 'tp',
  Contact = 'contact',
  Authority = 'authority',
  DistributionPartner = 'dp',
  Payment = 'payment',
  LegalRepresentative = 'legal_representative',
  MARequirement = 'ma_requirement',
  SNP = 'snp',
  FEE = 'fee',
  Library = 'library',
  FileType = 'file_type',
  Keyword = 'keyword'
}

export interface IViewConfig {
  name: string;
  type: ViewType;
  list: string;
}

export class View {
  public static config: IViewConfig[] = [
    {
      name: 'Countries View',
      type: ViewType.Country,
      list: Consts.LISTS.COUNTRY_URL
    },
    {
      name: 'Markets View',
      type: ViewType.Market,
      list: Consts.LISTS.MARKET_URL
    },
    {
      name: 'Eutelsat Entities View',
      type: ViewType.EutelsatEntity,
      list: Consts.LISTS.EUTELSAT_ENTITY_URL
    },
    {
      name: 'Legal Svc Provider View',
      type: ViewType.LegalSvcProvider,
      list: Consts.LISTS.LEGAL_SERVICE_PROVIDER_URL
    },
    {
      name: 'Teleport Partners View',
      type: ViewType.TeleportPartner,
      list: Consts.LISTS.TELEPORT_PARTNER_URL
    },
    {
      name: 'Contacts View',
      type: ViewType.Contact,
      list: Consts.LISTS.CONTACT_URL
    },
    {
      name: 'Authorities View',
      type: ViewType.Authority,
      list: Consts.LISTS.AUTHORITY_URL
    },
    {
      name: 'Distribution Partners View',
      type: ViewType.DistributionPartner,
      list: Consts.LISTS.DISTRIBUTION_PARTNER_URL
    },
    {
      name: 'Legal Representatives View',
      type: ViewType.LegalRepresentative,
      list: Consts.LISTS.LEGAL_SERVICE_PROVIDER_URL
    },
    {
      name: 'MA Requirements View',
      type: ViewType.MARequirement,
      list: Consts.LISTS.MA_REQUIREMENT_URL
    },
    {
      name: 'SNPs View',
      type: ViewType.SNP,
      list: Consts.LISTS.SNP_URL
    },
    {
      name: 'Fees View',
      type: ViewType.FEE,
      list: Consts.LISTS.FEE_URL
    },
    {
      name: 'Payments View',
      type: ViewType.Payment,
      list: Consts.LISTS.PAYMENT_URL
    },
    {
      name: 'Library View',
      type: ViewType.Library,
      list: Consts.LISTS.DOCUMENT_SPACE_URL
    },
    {
      name: 'File Types View',
      type: ViewType.FileType,
      list: Consts.LISTS.FILETYPE_URL
    },
    {
      name: 'Keyword View',
      type: ViewType.Keyword,
      list: Consts.LISTS.KEYWORD_URL
    }
  ];
}

export interface IViewBaseConfig {
  listUrl: string;
  filter: string;
  sort: string;
  select: string;
  top: number;
  expand: string;
  filters: IFilter[];
  columns: IColumn[];
  onRender?(filter);
}
