/* eslint-disable @typescript-eslint/no-empty-function */
import * as React from 'react';
import { ISiteHeaderProps, ISiteHeaderState } from './AppHeader.types';
import styles from '../../AppCustomizer.module.scss';
import { TaxonomyHelper } from '../../../../common/helpers/TaxonomyHelper';
import { Config } from '../../../../common/config/Config';
import { CommandBar, IContextualMenuItem } from '@fluentui/react';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { UserRole } from '../../../../common/models/Enums';
export default class SiteHeader
  extends React.Component<ISiteHeaderProps, ISiteHeaderState> {
  constructor(props) {
    super(props);
    this.state = {
      menuItems: []
    };
  }

  private async init(): Promise<void> {
    this.initSiteNavigationItems();
  }

  public initSiteNavigationItems = async (): Promise<void> => {
    const {
      pnpService
    } = this.props;
    const terms: ITermInfo[] = await TaxonomyHelper.getTermSetTerms(pnpService, Config.TAXONOMY.NAVIGATION_TERMSETID);
    const menuItems: IContextualMenuItem[] = terms
      .map(
        (term) => {
          const userRole = (term.localProperties[0]?.properties.find((p) => p.key === 'AssignedUserRoles')?.value || '').split(';').map((r) => r.trim()) as UserRole[];
          return this.hasAnyRole(...userRole) ?
            this.CreateMenuItems(term) : null;
        }
      )
      .filter((item): item is IContextualMenuItem => item !== null);
    this.setState({ menuItems });
  };
  private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
    const userRoles = this.props.userService.userContext?.userRoles ?? [];
    return rolesToCheck.some((role) => userRoles.includes(role));
  };
  public async componentDidMount(): Promise<void> {
    await Promise.all([
      this.init(),
    ]);
  }
  public render(): React.ReactElement<ISiteHeaderProps> {
    const { menuItems } = this.state;
    return (
      <div className={styles.appHeader} style={{ display: 'none' }}>
        <div className={styles.topBar}>
          <div className={styles.appInfo}>
            <div className={styles.appInfoContainer}>
              <div className={styles.appInfoLogo}>
                <div>
                  <div className={styles.appInfoLogoOuter}>
                    <div className={styles.appInfoLogoInner} >
                      <a href={Config.SITE.ROOT_WEB_URL} data-interception="off">
                        <div className={styles.appInfoLogoContainer}>
                          <img src={require('../../../../common/assets/images/EutelSatLogo.png')} />
                        </div>
                      </a>
                    </div>
                  </div>
                </div>
              </div>
              <div className={styles.appInfoTitle}>
                <span>
                  {Config.SITE.SITE_TITLE}</span>
              </div>
            </div>
          </div>
          <div className={styles.appNavigation}>
            <CommandBar
              items={menuItems}
            />
          </div>
        </div>
      </div>
    );
  }

  private CreateMenuItems(menuItem: ITermInfo): IContextualMenuItem | null {
    const isLink = menuItem.localProperties[0]?.properties.find((p) => p.key === 'NavURL') !== undefined;
    const navUrl = isLink ?
      menuItem.localProperties[0].properties.find((p) => p.key === 'NavURL')?.value : null;

    const userRoles = (menuItem.localProperties[0]?.properties.find((p) => p.key === 'AssignedUserRoles')?.value || '').split(';').map((r) => r.trim()) as UserRole[];
    if (!this.hasAnyRole(...userRoles)) {
      return null;
    }
    const childMenuItems = menuItem.children?.length > 0 ? menuItem.children
      .map((i) => this.CreateMenuItems(i as ITermInfo))
      .filter((item): item is IContextualMenuItem => item !== null) : [];
    return {
      id: menuItem.id,
      key: menuItem.id,
      name: menuItem.labels[0]?.name || '',
      text: menuItem.labels[0]?.name || '',
      onClick: navUrl ? (e: React.MouseEvent<HTMLElement>) => {
        e.preventDefault();
        if (this.props.onNavigate) {
          this.props.onNavigate(navUrl);
        }
      } : undefined,
      subMenuProps: childMenuItems?.length > 0 ?
        {
          items: childMenuItems
        } : null,
      isSubMenu: !isLink,
    };
  }
}
