import { PNPService } from '../../../common/services/PNPService';
import { ITermInfo } from "@pnp/sp/taxonomy";

export interface INavTilesProps {
  pnpService: PNPService;
}

export interface INavTilesState {
  navItems: ITermInfo[];
}

