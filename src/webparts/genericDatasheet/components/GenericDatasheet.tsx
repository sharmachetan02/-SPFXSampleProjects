/* eslint-disable max-len */
import * as React from 'react';
import { IGenericDatasheetProps, DatasheetType } from './Datasheet.types';
import { IDatasheetConfig, Datasheet } from './Datasheet.types';
import { UrlHelper } from '../../../common/helpers/UrlHelper';
import { Consts } from '../../../common/consts/Consts';
import styles from './Datasheet.module.scss';
import CountryDatasheet from './Datasheets/Country/CountryDatasheet';
import SNP from './Datasheets/SNP/SNPDatasheet';
import NolDatasheet from './Datasheets/NOL/NolDatasheet';
import MAGenericRequirementDatasheet from './Datasheets/MAGeneric/MAGenericDatasheet';
import MALicenseRequirementDatasheet from './Datasheets/MALicense/MALicenseDatasheet';
//import MAGenericRequirementDatasheet from '../MAGeneric/MAGenericDatasheet';

import PaymentDatasheet from './Datasheets/Payment/PaymentDatasheet';

import FeeDatasheet from './Datasheets/Fee Datasheet/FeeDatasheet';
import LegalRequirementDatasheet from './Datasheets/Legal Requirement/LegalRequirementDatasheet';
import AuthorityDatasheet from './Datasheets/Authority/AuthorityDatasheet';
import LegalSvcProviderDatasheet from './Datasheets/Legal Svc Provider Datasheet/LegalSvcProviderDatasheet';
import DPDatasheet from './Datasheets/DP/DPDatasheet';
import Contact from './Datasheets/Contact/ContactDatasheet';
import EutelsatEntityDatasheet from './Datasheets/Eutelsat Entity/EutelsatEntity';
import TPDatasheet from './Datasheets/TP/TPDatasheet';
import MarketDatasheet from './Datasheets/Market/MarketDatasheet';

export default class GenericDatasheet extends React.Component<IGenericDatasheetProps> {
    private DatasheetConfig: IDatasheetConfig;
    private ItemId: number;

    constructor(props: Readonly<IGenericDatasheetProps>) {
        super(props);
        this.updateDatasheetConfig();
    }

    private updateDatasheetConfig = (): void => {
        const datasheetType = UrlHelper.getQueryStringParam(Consts.QUERY.PARAM_DATASHEET_TYPE, window.location.href);
        this.ItemId = +UrlHelper.getQueryStringParam(Consts.QUERY.PARAM_DATASHEET_ITEM_ID, window.location.href);
        if (datasheetType) {
            this.DatasheetConfig = Datasheet.config.find((f) => f.type.toLowerCase() === datasheetType);
        }
    };

    public componentDidMount(): void {
        UrlHelper.onNavigationChange((path) => {
            this.updateDatasheetConfig();
            this._onUrlChange(path || window.location.search);
        });
    }
    private _onUrlChange(newpath = window.location.search): void {

        this.forceUpdate();
    }
    public render(): React.ReactElement<IGenericDatasheetProps> {
        return (
            <>
                {
                    !this.DatasheetConfig &&
                    <div className={styles.noView}>
                        <span className={styles.noViewText}>{'Datasheet Not Found'}</span>
                    </div>
                }
                {this.DatasheetConfig &&
                    <div className={styles.datasheet}>
                        <div className="ms-Grid">
                            <div className="ms-Grid-row">
                                <div className={`ms-Grid-col ms-sm12 ms-md12`}>
                                    <div className="ms-Grid">
                                        {this.selectDatasheet()}
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                }
            </>
        );
    }

    private selectDatasheet = () => {
        const {
            pnpService,
            userService
        } = this.props;
        const {
            ItemId
        } = this;
        if (this.DatasheetConfig !== undefined && this.DatasheetConfig !== null && this.DatasheetConfig.type !== null) {
            switch (this.DatasheetConfig.type) {
                case DatasheetType.Country:
                    return (<CountryDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.Market:
                    return (<MarketDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.SNP:
                    return (<SNP pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.MARequirementNol:
                    return (<NolDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.MARequirementGeneric:
                    return (<MAGenericRequirementDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.MARequirementLicense:
                    return (<MALicenseRequirementDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.Payment:
                    return (<PaymentDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.Fee:
                    return (<FeeDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.MARequirementLegal:
                    return (<LegalRequirementDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.Authority:
                    return (<AuthorityDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.LegalSvcProvider:
                    return (<LegalSvcProviderDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.DistributionPartner:
                    return (<DPDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.Contact:
                    return (<Contact pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.EutelsatEntity:
                    return (<EutelsatEntityDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);
                case DatasheetType.TeleportPartner:
                    return (<TPDatasheet pnpService={pnpService} userService={userService} itemId={ItemId} />);

                default:
                    return (<div>No Datasheet found</div>);
            }
        } else {
            return (<div>No Datasheet found</div>);
        }
    };
}
