import { PNPService } from "../../../../common/services/PNPService";
import { IContextualMenuItem } from "@fluentui/react";
import { UserService } from "../../../../common/services/UserContextService";

export interface ISiteHeaderProps {
    pnpService: PNPService;
    userService: UserService
     onNavigate?: (url: string) => void;
}
export interface ISiteHeaderState {
    menuItems: IContextualMenuItem[];
    //userGroups: any[];
}
export interface ITaxonomyTermState {
    id: string;
    labels: {
        name: string;
        languageTag: string;
        isDefault: boolean;
    }[];
    customProperties?: { NavURL?: string, ParentID?: string };
    customSortOrder?: number;
}
