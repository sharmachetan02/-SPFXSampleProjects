import * as React from 'react';
import styles from './NavTiles.module.scss';
import { INavTilesProps, INavTilesState } from './INavTilesProps.types';
import { TaxonomyHelper } from '../../../common/helpers/TaxonomyHelper';
import { Config } from '../../../common/config/Config';
import { ITermInfo } from '@pnp/sp/taxonomy';
import SvgIcon from '../../../common/components/svgIcon/SvgIcon';
import { Icon } from '../../../common/components/svgIcon/Interfaces';
import strings from 'NavTilesWebPartStrings';
import { Icon as FluentIcon } from '@fluentui/react';
import { UrlHelper } from '../../../common/helpers/UrlHelper';
export default class NavTiles extends React.Component<INavTilesProps, INavTilesState> {
  constructor(props: INavTilesProps) {
    super(props);
    this.state = {
      navItems: []
    };
  }

  public async componentDidMount(): Promise<void> {
    await Promise.all([
      this.init(),
    ]);
  }

  private async init(): Promise<void> {
    const promises: Promise<void>[] = [];
    promises.push(this.initSiteNavigationTerms());
    await Promise.all(promises);
  }

  public initSiteNavigationTerms = async (): Promise<void> => {
    const {
      pnpService
    } = this.props;
    const subTerms = await TaxonomyHelper.getSubTerms(pnpService, Config.TAXONOMY.NAVIGATION_TERMSETID, Config.TAXONOMY.NAVIGATION_CONTENT_TERMID);
    const navItems: ITermInfo[] = [];
    subTerms.forEach((term) => {
      navItems.push({
        ...term
      });
    });
    this.setState({ navItems: subTerms });
  };

  public render(): React.ReactElement<INavTilesProps> {
    const { navItems } = this.state;
    return (
      <>
        <div className={styles.navTiles}>
          <div className={styles.section}>
            <FluentIcon iconName="OpenFolderHorizontal" className="sectionIcon" />
            <h3 >{strings.NavTilesTitle}</h3>
          </div>
          <div className={styles.navTilesOuter}>
            <div className={styles.navTilesInner}>
              {navItems.map((term) => (
                <div key={term.id}>
                  <a data-interception="on"
                    key={term.id}
                    onClick={(e) => {
                      e.preventDefault();
                      const url = term.localProperties[0]?.properties.find((p) => p.key === 'NavURL')?.value;
                      UrlHelper.navigate(url, false);
                    }}
                    className={styles.tile}
                    rel='noreferrer'
                  >
                    <SvgIcon
                      icon={
                        Icon[term.labels?.[0]?.name?.replace(/\s+/g, '')] ??
                        term.labels?.[0]?.name
                      }
                    />
                    <span>{term.labels?.[0]?.name}</span>
                  </a>
                </div>
              ))}
            </div>
          </div >
        </div>
      </>
    );
  }
}
