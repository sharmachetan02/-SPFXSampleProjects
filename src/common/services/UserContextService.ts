import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { Consts } from '../consts/Consts';
import { PNPService } from './PNPService';
import { UserRole } from '../models/Enums';



export interface ISpGroup {
  Id: number;
  Title: string;
  LoginName: string;
  /* Description?: string;
     AllowMembersEditMembership: boolean;
     AllowRequestToJoinLeave: boolean;
     AutoAcceptRequestToJoinLeave: boolean;
     IsHiddenInUI: boolean;
     OnlyAllowMembersViewMembership: boolean;
     OwnerTitle?: string;
     PrincipalType?: number;
     RequestToJoinLeaveEmailSetting?: string | null;
     // OData metadata
     'odata.type'?: string;
     'odata.id'?: string;
     'odata.editLink'?: string;*/
}

export interface IUserContext {
  userRoles: UserRole[];
}


export interface IExtendedSiteUserInfo extends ISiteUserInfo {
  Groups?: Array<ISpGroup>;
}


export class UserService {
  private _pnpService: PNPService;
  private _userContext: IUserContext | null = null;


  constructor(pnpService: PNPService) {
    this._pnpService = pnpService;
  }
  public get userContext(): IUserContext | null {
    return this._userContext;
  }
  public async setup() {
    await this.initUserRoles();
  }

  private checkGroup(tokens: string[], groupName: string): boolean {
    return tokens.some((token) => token.trim().toLowerCase() === groupName.trim().toLowerCase());
  }

  private async initUserRoles(): Promise<void> {
    const currentUser: IExtendedSiteUserInfo = await this._pnpService.getWeb().currentUser.expand('Groups')();
    const groupsPayload = currentUser.Groups ?? [];

    const normalizedGroupNames = groupsPayload
      .map((g) => (g?.LoginName ?? g?.Title ?? '').trim().toLowerCase())
      .filter(Boolean);

    const roles: UserRole[] = [];

    if (this.checkGroup(normalizedGroupNames, Consts.Groups.FinanceContributors)) {
      roles.push(UserRole.FinanceContributor);
    }
    if (this.checkGroup(normalizedGroupNames, Consts.Groups.FinanceVisitors)) {
      roles.push(UserRole.FinanceVisitor);
    }
    if (this.checkGroup(normalizedGroupNames, Consts.Groups.Contributors)) {
      roles.push(UserRole.Contributor);
    }
    if (this.checkGroup(normalizedGroupNames, Consts.Groups.Owners)) {
      roles.push(UserRole.Owner);
    }
    if (this.checkGroup(normalizedGroupNames, Consts.Groups.Visitors)) {
      roles.push(UserRole.Visitor);
    }

    // check if user is Site Collection Admin
    const isSiteAdmin = currentUser.IsSiteAdmin; // Property available in ISiteUserInfo
    if (isSiteAdmin) {
      roles.push(UserRole.Admin);
    }


    this._userContext = { userRoles: roles };
  }

}
