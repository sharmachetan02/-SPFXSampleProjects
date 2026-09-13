import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps, IHelloWorldState } from './IHelloWorldProps.types';
import { escape } from '@microsoft/sp-lodash-subset';
import { ActionButton, Modal, PrimaryButton } from 'office-ui-fabric-react';
import { Form, FormOrigin, FormType } from '../../genericForm/components/Form.types';
import EutelsatEntity from '../../genericForm/components/Eutelsat Entity/EutelsatEntity';
import Authority from '../../genericForm/components/Authority/Authority';
import LegalSvcProvider from '../../genericForm/components/Legal Svc Provider/LegalSvcProvider';
import Country from '../../genericForm/components/Country/Country';
import SNP from '../../genericForm/components/SNP/SNP';
import TeleportPartner from '../../genericForm/components/Teleport Partner/TeleportPartner';
import DistributionPartner from '../../genericForm/components/Distribution Partner/DistributionPartner';
import TestComponent from '../../genericForm/components/TestComponent/TestComponent';
import Fee from '../../genericForm/components/Fee/Fee';
import MARequirementGeneric from '../../genericForm/components/MA Requirement Generic/MARequirementGeneric';
import Payment from '../../genericForm/components/Payment/Payment';
import SNPMarketReadiness from '../../genericForm/components/SNP Readiness/SNPReadiness';
import { UrlHelper } from '../../../common/helpers/UrlHelper';
import { Consts } from '../../../common/consts/Consts';
import MarketReadiness from '../../genericForm/components/Market Readiness/MarketReadiness';
import LegalRequirement from '../../genericForm/components/MA Requirement Legal/LegalRequirement';
import Nol from '../../genericForm/components/MA Requirement NOL/Nol';
import MALicense from '../../genericForm/components/MA Requirement License/MALicense';
import MADocument from '../../genericForm/components/MADocument/MADocument';
import MAFolder from '../../genericForm/components/MA Folder/MAFolder';

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
  constructor(props: IHelloWorldProps) {
    super(props);
    this.state = {
      showFormDialog: false,
      itemId: undefined,
      itemType: undefined,
      preselectedItemId: undefined,
      preselectedItemType: undefined
    };
    this.closeForm = this.closeForm.bind(this);
    this.renderFormDialog = this.renderFormDialog.bind(this);
    this.reloadPage = this.reloadPage.bind(this);
  }
  public render(): React.ReactElement<IHelloWorldProps> {
    const {
      description,
      environmentMessage,
      userDisplayName
    } = this.props;
    const {
      showFormDialog
    } = this.state;
    return (
      <>
        {!showFormDialog &&
          <section className={`${styles.helloWorld}`}>
            <div className={styles.welcome}>
              <h2>Well done, {escape(userDisplayName)}!</h2>
              <div>{environmentMessage}</div>
              <div>Web part property value: <strong>{escape(description)}</strong></div>
            </div>
            <div>
              <h3>Welcome to SharePoint Framework!</h3>
              <h4>Learn more about SPFx development:</h4>
              <ul className={styles.links}>
                <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
                <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
              </ul>
            </div>
          </section>
        }
        {

          <div style={{ marginTop: 10 }}>
            {this.renderFeeForm()}
            <PrimaryButton
              text="EUT Entity Edit Form"
              onClick={this.openEUTEntityEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />
            <PrimaryButton
              text="Authority Edit Form"
              onClick={this.openAuthorityEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="Legal Svc Provider Edit Form"
              onClick={this.openEUTLegalSvcProviderEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="Country Edit Form"
              onClick={this.openCountryEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="SNP Edit Form"
              onClick={this.openSNPEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="Teleport Partner Edit Form"
              onClick={this.openTeleportPartnerEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="Distribution Partner Edit Form"
              onClick={this.openDistributionPartnerEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="Fee New Form"
              onClick={this.openFeeNewForm}
              style={{ marginRight: 10 }} />

            <PrimaryButton
              text="Fee Edit Form"
              disabled={false}
              onClick={this.openFeeEditForm}
              style={{ marginRight: 10 }} />
            <PrimaryButton
              text="Payment Edit Form"
              onClick={this.openPaymentEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />
            <PrimaryButton
              text="MA Requirement Generic Edit Form"
              disabled={false}
              onClick={this.openMARequirementGenericEditForm}
              style={{ marginRight: 10 }} />

            <PrimaryButton
              text="TestComponent Edit Form"
              onClick={this.openTestComponentEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="TestComponent New Form"
              onClick={this.openTestComponentNewForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="SNP Market Readiness Edit Form"
              disabled={false}
              onClick={this.openSNPMarketReadinessEditForm}
              style={{ marginRight: 10, marginTop: 10 }} />

            <PrimaryButton
              text="SNP Market Readiness New Form"
              onClick={this.openSNPMarketReadinessNewForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="COUNTRY Market Readiness Edit Form"
              disabled={false}
              onClick={this.openMarketReadinessEditForm}
              style={{ marginRight: 10, marginTop: 10 }} />

            <PrimaryButton
              text="COUNTRY Market Readiness New Form"
              onClick={this.openMarketReadinessNewForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="MA Requirement Legal Edit Form"
              onClick={this.openLegalRequirementEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />

            <PrimaryButton
              text="MA NOL Requirement Edit Form"
              onClick={this.openNolRequirementEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />
            <PrimaryButton
              text="MALicense Edit Form"
              disabled={false}
              onClick={this.openMALicenseFormEditForm}
              style={{ marginRight: 10 }} />

            {<PrimaryButton
              text="MA Document New Form"
              onClick={this.openMADocumentNewForm}
              style={{ marginRight: 10, marginBottom: 10 }} />
            }
            {<PrimaryButton
              text="MA Document Edit Form"
              onClick={this.openMADocumentEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />
            }
            <PrimaryButton
              text="MA Folder New Form"
              onClick={this.openMAFolderNewForm}
              style={{ marginRight: 10, marginBottom: 10 }} />
            <PrimaryButton
              text="MA Folder Edit Form"
              onClick={this.openMAFolderEditForm}
              style={{ marginRight: 10, marginBottom: 10 }} />
          </div>
        }
        {showFormDialog &&
          this.renderFormDialog()
        }
      </>
    );
  }
  private closeForm = () => {
    this.setState({ showFormDialog: false });
  };
  private reloadPage() {
    this.closeForm();
    // if (redirectUrl) window.location.href = redirectUrl;
  }
  private openEUTEntityEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 1, itemType: FormType.EutelsatEntity });
  };
  private openCountryEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 16, itemType: FormType.Country });
  };
  private openEUTLegalSvcProviderEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 13, itemType: FormType.LegalServiceProvider });
  };
  private openAuthorityEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 16, itemType: FormType.Authority });
  };
  private openPaymentEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 5, itemType: FormType.Payment });
  };
  private openSNPEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 13, itemType: FormType.SNP });
  };
  private openTeleportPartnerEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 8, itemType: FormType.TP });
  };
  private openDistributionPartnerEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 6, itemType: FormType.DP });
  };
  private openSNPMarketReadinessEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 14, itemType: FormType.SNPMarketReadiness });
  };
  private openSNPMarketReadinessNewForm = () => {
    this.setState({
      showFormDialog: true,
      itemId: undefined,
      preselectedItemId: 1,
      itemType: FormType.SNPMarketReadiness
    });
  };
  private openMarketReadinessEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 18, itemType: FormType.MarketReadiness });
  };
  private openMarketReadinessNewForm = () => {
    this.setState({
      showFormDialog: true,
      itemId: undefined,
      preselectedItemId: 1,
      itemType: FormType.MarketReadiness
    });
  };
  private openFeeNewForm = () => {
    this.setState({
      showFormDialog: true,
      itemId: undefined,
      preselectedItemId: 9,
      preselectedItemType: FormType.SNP,
      itemType: FormType.Fee
    });
  };
  private openFeeEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 24, itemType: FormType.Fee });
  };
  private openMARequirementGenericEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 40, itemType: FormType.MARequirementGeneric });
  };
  private openLegalRequirementEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 21, itemType: FormType.MARequirementLegal });
  };
  private openNolRequirementEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 46, itemType: FormType.MARequirementNol });
  };
  private openMALicenseFormEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 20, itemType: FormType.MALicense });
  };
  private openMADocumentNewForm = () => {
    this.setState({ showFormDialog: true, itemId: undefined, itemType: FormType.MADocument });
  };
  private openMADocumentEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 236, itemType: FormType.MADocument });
  };
  private openMAFolderEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 238, itemType: FormType.MAFolder });
  };
  private openMAFolderNewForm = () => {
    this.setState({ showFormDialog: true, itemId: undefined, itemType: FormType.MAFolder });
  };
  private openTestComponentEditForm = () => {
    this.setState({ showFormDialog: true, itemId: 1, itemType: FormType.Test });
  };
  private openTestComponentNewForm = () => {
    this.setState({
      showFormDialog: true,
      itemId: undefined,
      preselectedItemId: 1,
      preselectedItemType: FormType.Test,
      itemType: FormType.Test
    });
  };
  private renderFeeForm() {
    const {
      pnpService,
      userService
    } = this.props;
    const ItemId: number = +UrlHelper.getQueryStringParam(Consts.QUERY.PARAM_DATASHEET_ITEM_ID, window.location.href);

    return <Fee pnpService={pnpService} userService={userService} itemId={ItemId} ></Fee>;
  }
  private renderFormDialog() {
    const {
      pnpService,
      userService
    } = this.props;
    const {
      showFormDialog,
      itemType,
      itemId,
      preselectedItemType,
      preselectedItemId
    } = this.state;
    let formControl, formTitle = null;
    const formConfig = Form.config.find((f) => f.type.toLowerCase() === itemType);
    switch (itemType) {
      case FormType.EutelsatEntity:
        formControl = (<EutelsatEntity pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></EutelsatEntity>);
        formTitle = `Edit ${formConfig.name}`; break;

      case FormType.Authority:
        formControl = (<Authority pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></Authority>);
        formTitle = `Edit ${formConfig.name}`; break;

      case FormType.LegalServiceProvider:
        formControl = (<LegalSvcProvider pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></LegalSvcProvider>);
        formTitle = `Edit ${formConfig.name}`;
        break;

      case FormType.Country:
        formControl = (<Country pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></Country>);
        formTitle = `Edit ${formConfig.name}`;
        break;

      case FormType.SNP:
        formControl = (<SNP pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></SNP>);
        formTitle = `Edit ${formConfig.name}`;
        break;

      case FormType.TP:
        formControl = (<TeleportPartner pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></TeleportPartner>);
        formTitle = `Edit ${formConfig.name}`;
        break;

      case FormType.DP:
        formControl = (<DistributionPartner pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></DistributionPartner>);
        formTitle = `Edit ${formConfig.name}`;
        break;

      case FormType.Fee:
        if (itemId) {
          formControl = (<Fee pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></Fee>);
          formTitle = `Edit ${formConfig.name}`;
        } else {
          formControl = (<Fee pnpService={pnpService} userService={userService} preselectedItemType={preselectedItemType}
            preselectedItemId={preselectedItemId} callback={this.reloadPage}></Fee>);
          formTitle = `New ${formConfig.name}`;
        }
        break;

      case FormType.SNPMarketReadiness:
        if (itemId) {
          formControl = (<SNPMarketReadiness pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></SNPMarketReadiness>);
          formTitle = `Edit ${formConfig.name}`;
        } else {
          formControl = (<SNPMarketReadiness pnpService={pnpService} userService={userService}
            preselectedItemId={preselectedItemId} callback={this.reloadPage}></SNPMarketReadiness>);
          formTitle = `New ${formConfig.name}`;
        }
        break;
      case FormType.MarketReadiness:
        if (itemId) {
          formControl = (<MarketReadiness pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></MarketReadiness>);
          formTitle = `Edit ${formConfig.name}`;
        } else {
          formControl = (<MarketReadiness pnpService={pnpService} userService={userService}
            preselectedItemId={preselectedItemId} callback={this.reloadPage}></MarketReadiness>);
          formTitle = `New ${formConfig.name}`;
        }
        break;
      case FormType.MARequirementGeneric:
        formControl = (<MARequirementGeneric pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></MARequirementGeneric>);
        formTitle = `Edit ${formConfig.name}`;
        break;

      case FormType.Payment:
        formControl = (<Payment pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage} formOrigin={FormOrigin.Datasheet}></Payment>);
        formTitle = `Edit ${formConfig.name}`;
        break;
      case FormType.MARequirementLegal:
        formControl = (<LegalRequirement pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></LegalRequirement>);
        formTitle = `Edit ${formConfig.name}`;
        break;
      case FormType.MARequirementNol:
        formControl = (<Nol pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></Nol>);
        formTitle = `Edit ${formConfig.name}`;
        break;
      case FormType.MALicense:
        formControl = (<MALicense pnpService={pnpService} userService={userService} itemId={35} callback={this.reloadPage}></MALicense>);
        formTitle = `Edit ${formConfig.name}`;
        break;
      case FormType.MADocument:
        if (itemId) {
          formControl = (<MADocument pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></MADocument>);
          formTitle = `Edit ${formConfig.name}`;
        } else {
          formControl = (<MADocument pnpService={pnpService} userService={userService} filepath={"/Authority/Administration"} callback={this.reloadPage}></MADocument>);
          formTitle = `New ${formConfig.name}`;
        }
        break;
      case FormType.MAFolder:
        if (itemId) {
          formControl = (<MAFolder pnpService={pnpService} userService={userService} itemId={itemId} currentFolderPath='Nol/Soma Test' callback={this.reloadPage}></MAFolder>);
          formTitle = `Edit ${formConfig.name}`;
        } else {
          formControl = (<MAFolder pnpService={pnpService} userService={userService}
            currentFolderPath='Nol/Soma Test' callback={this.reloadPage}></MAFolder>);
          formTitle = `New ${formConfig.name}`;
        }
        break;
      case FormType.Test:
        if (itemId) {
          formControl = (<TestComponent pnpService={pnpService} userService={userService} itemId={itemId} callback={this.reloadPage}></TestComponent>);
          formTitle = `Edit ${formConfig.name}`;
        } else {
          formControl = (<TestComponent pnpService={pnpService} userService={userService} itemId={itemId} preselectedItemType={preselectedItemType}
            preselectedItemId={preselectedItemId} callback={this.reloadPage}></TestComponent>);
          formTitle = `New ${formConfig.name}`;
        }
        break;
    }
    return (<Modal
      isOpen={showFormDialog}
      onDismiss={this.closeForm}
      containerClassName="modalContainer"
      isBlocking={true} >
      <div className="document-form">
        <div className="modalHeader">
          {formTitle}
          <ActionButton
            iconProps={{ iconName: 'Cancel' }}
            onClick={this.closeForm}
            className="iconButton"
          >Close</ActionButton>
        </div>
        {/* <hr /> */}
        <div className="modalBody">
          {formControl}
        </div>
      </div>
    </Modal>);
  }
}
