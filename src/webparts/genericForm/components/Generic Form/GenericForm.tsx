
import * as React from 'react';
import styles from '../Form.module.scss';
import { IFormConfig, Form, FormType, FormOrigin } from '../Form.types';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { Consts } from '../../../../common/consts/Consts';
import * as strings from 'GenericFormWebPartStrings';
import 'office-ui-fabric-core/dist/css/fabric.min.css';
import { IGenericFormProps } from './GenericForm.types';
import EutelsatEntity from '../Eutelsat Entity/EutelsatEntity';
import LegalSvcProvider from '../Legal Svc Provider/LegalSvcProvider';
import Authority from '../Authority/Authority';
import Country from '../Country/Country';
import SNP from '../SNP/SNP';
import TeleportPartner from '../Teleport Partner/TeleportPartner';
import DistributionPartner from '../Distribution Partner/DistributionPartner';
import Contact from '../Contact/Contact';
import TestComponent from '../TestComponent/TestComponent';
import Fee from '../Fee/Fee';
import MARequirementGeneric from '../MA Requirement Generic/MARequirementGeneric';
import Payment from '../Payment/Payment';
import Nol from '../MA Requirement NOL/Nol';
import LegalRequirement from '../MA Requirement Legal/LegalRequirement';
import MALicense from '../MA Requirement License/MALicense';
import Market from '../Market/Market';

export default class GenericForm extends React.Component<IGenericFormProps> {

  private formConfig: IFormConfig;

  constructor(props: Readonly<IGenericFormProps>) {
    super(props);
    this.updateFormConfig();
  }
  private updateFormConfig = (): void => {
    const formType = UrlHelper.getQueryStringParam(Consts.QUERY.PARAM_FORM_TYPE, window.location.href);
    this.formConfig = Form.config.find((f) => f.type.toLowerCase() === formType);
  };
  public componentDidMount(): void {
    //empty
    UrlHelper.onNavigationChange((path) => {
      this.updateFormConfig();
      this._onUrlChange(path || window.location.search);
    });
  }
  private _onUrlChange(newpath = window.location.search): void {

    this.forceUpdate();
  }
  private reloadPage(redirectUrl) {
    if (redirectUrl) window.location.href = redirectUrl;
  }
  public render(): React.ReactElement<IGenericFormProps> {
    const titleText = this.formConfig?.name;
    return (
      <>
        {
          !this.formConfig &&
          <div className={styles.noForm}>
            <span className={styles.noFormText}>{strings.FormTypeNotfound}</span>
          </div>
        }
        {!!this.formConfig &&
          <div className={styles.genericForm}>
            <div className="ms-Grid">
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm3 ms-md3"></div>
                <div className={`ms-Grid-col ms-sm6 ms-md6 ${styles.formContainer}`}>
                  <div className="ms-Grid">
                    <div className="ms-Grid-row">
                      <div className="ms-Grid-col ms-sm12 ms-md12">
                        <div className={styles.formTitle}>
                          <span className={styles.formTitleText}>{titleText}</span>
                        </div>
                      </div>
                    </div>
                    {this.selectForm()}
                  </div>
                </div>
                <div className="ms-Grid-col ms-sm3 ms-md3"></div>
              </div>
            </div>
          </div>
        }
      </>
    );
  }
  private selectForm() {
    const {
      pnpService,
      userService
    } = this.props;
    switch (this.formConfig.type) {
      case FormType.EutelsatEntity:
        return (<EutelsatEntity pnpService={pnpService} userService={userService} callback={this.reloadPage}></EutelsatEntity>);

      case FormType.LegalServiceProvider:
        return (<LegalSvcProvider pnpService={pnpService} userService={userService} callback={this.reloadPage}></LegalSvcProvider>);
      case FormType.Authority:
        return (<Authority pnpService={pnpService} userService={userService} callback={this.reloadPage}></Authority>);
      case FormType.Country:
        return (<Country pnpService={pnpService} userService={userService} callback={this.reloadPage}></Country>);
      case FormType.Market:
        return (<Market pnpService={pnpService} userService={userService} callback={this.reloadPage}></Market>);
      case FormType.SNP:
        return (<SNP pnpService={pnpService} userService={userService} callback={this.reloadPage} ></SNP>);
      case FormType.TP:
        return (<TeleportPartner pnpService={pnpService} userService={userService} callback={this.reloadPage}></TeleportPartner>);
      case FormType.DP:
        return (<DistributionPartner pnpService={pnpService} userService={userService} callback={this.reloadPage}></DistributionPartner>);
      case FormType.Contact:
        return (<Contact pnpService={pnpService} userService={userService} callback={this.reloadPage} formOrigin={FormOrigin.Menu}></Contact>);
      case FormType.Fee:
        return (<Fee pnpService={pnpService} userService={userService} callback={this.reloadPage} ></Fee>);
      case FormType.Payment:
        return (<Payment pnpService={pnpService} userService={userService} formOrigin={FormOrigin.Menu}></Payment>);
      case FormType.MARequirementGeneric:
        return (<MARequirementGeneric pnpService={pnpService} userService={userService} callback={this.reloadPage}></MARequirementGeneric>);
      case FormType.MARequirementLegal:
        return (<LegalRequirement pnpService={pnpService} userService={userService}></LegalRequirement>);
      case FormType.MARequirementNol:
        return (<Nol pnpService={pnpService} userService={userService}></Nol>);
      case FormType.MALicense:
        return (<MALicense pnpService={pnpService} userService={userService} ></MALicense>);
      case FormType.Test:
        return (<TestComponent pnpService={pnpService} userService={userService} callback={this.reloadPage}></TestComponent>);

    }
  }
}
