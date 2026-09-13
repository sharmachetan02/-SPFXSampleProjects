/* eslint-disable max-len */
import * as React from 'react';
import { IGenericViewProps, ViewType } from '../View.types';
import { IViewConfig, View } from '../View.types';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { Consts } from '../../../../common/consts/Consts';
import styles from '../View.module.scss';
import { Icon } from '@fluentui/react';
import CountryView from '../Country View/CountryView';
import EutelsatEntityView from '../Eutelsat Entity View/EutelsatEntityView';
import ContactView from '../Contact View/ContactView';
import LegalSvcProviderView from '../Legal Svc Provider View/LegalSvcProviderView';
import DPView from '../DP View/DPView';
import TPView from '../TP View/TPView';
import AuthorityView from '../Authority View/AuthorityView';
import SNPView from '../SNP View/SNPView';
import FeeView from '../Fees View/FeesView';
import PaymentView from '../Payment View/PaymentView';
import MARequirementView from '../MARequirement View/MARequirementView';
import MarketView from '../Market View/MarketView';
// import { SortDirection } from '@pnp/sp/search';

export default class GenericView extends React.Component<IGenericViewProps> {
    private viewConfig: IViewConfig;
    constructor(props: Readonly<IGenericViewProps>) {
        super(props);

        this.updateViewConfig();
    }

    private updateViewConfig =  (): void => {
        const viewType = UrlHelper.getQueryStringParam(Consts.QUERY.PARAM_VIEW_TYPE, window.location.href);
        if (viewType) {
            this.viewConfig = View.config.find((f) => f.type.toLowerCase() === viewType);
        }
    };
    public componentDidMount(): void {
        //empty
        UrlHelper.onNavigationChange((path) => {
            this.updateViewConfig();
            this._onUrlChange(path || window.location.search);
        });
    }
    private _onUrlChange(newpath = window.location.search): void {

        this.forceUpdate();
    }
    public render(): React.ReactElement<IGenericViewProps> {
        const titleText = this.viewConfig?.name;
        return (
            <>
                {
                    !this.viewConfig &&
                    <div className={styles.noView}>
                        <span className={styles.noViewText}>{'View Not Found'}</span>
                    </div>
                }
                {this.viewConfig &&
                    <div className={styles.viewContainer}>
                        <div className="ms-Grid">
                            <div className="ms-Grid-row">
                                <div className={`ms-Grid-col ms-sm12 ms-md12`}>
                                    <div className="ms-Grid">
                                        <div className="ms-Grid-row">
                                            <div className="ms-Grid-col ms-sm12 ms-md12">
                                                <div className={styles.section}>
                                                    <Icon iconName="ViewList" className="sectionIcon" />
                                                    <h3 >{titleText}</h3>
                                                </div>
                                            </div>
                                        </div>
                                        {this.selectView()}
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                }
            </>
        );
    }

    private selectView() {
        const {
            pnpService,
            userService
        } = this.props;
        if (this.viewConfig !== undefined && this.viewConfig !== null && this.viewConfig.type !== null) {
            switch (this.viewConfig.type) {
                case ViewType.Country:
                    // return (<CountryView pnpService={pnpService}
                    //     filters={['Search', 'Priority', 'Grouping', 'SanctionCategory']}
                    //     columns={['Title', 'Priority', 'Grouping', 'SanctionCategory']}
                    //     sort={{ Property: 'SanctionCategory', Direction: SortDirection.Ascending }}
                    //     viewItems={[
                    //         { "Id": 1, "Title": "France", "Identifier": "FR", "Capacity": "120 Gbps", "Grouping": [{ "Label": "APAC", "TermGuid": "2a0b2a0b-8f16-4c9d-9b06-6a5d1f3b5a01" }, { "Label": "NA", "TermGuid": "9b6e0a5f-6c24-4c2b-8f2a-3b7a1c9d4e02" }], "SanctionCategory": { "Id": 1, "Title": "Category A", "Summary": "No active sanctions", "Color": "#4caf50" }, "MaPriority": ["EU", "0.0"], "IsSanctioned": false, "Comments": "Key EU market" },
                    //         { "Id": 2, "Title": "Iran", "Identifier": "IR", "Capacity": "60 Gbps", "Grouping": [{ "Label": "APAC", "TermGuid": "c1d2e3f4-5678-49ab-9cde-102938475601" }], "SanctionCategory": { "Id": 5, "Title": "Category E", "Summary": "Comprehensive international sanctions apply", "Color": "#d32f2f" }, "MaPriority": ["EU"], "IsSanctioned": true, "Comments": "High regulatory risk" },
                    //         { "Id": 3, "Title": "Tunisia", "Identifier": "TN", "Capacity": "25 Gbps", "Grouping": [{ "Label": "APAC", "TermGuid": "0f1e2d3c-4b5a-6978-90ab-cdef12345603" }, { "Label": "NA", "TermGuid": "7c6b5a49-3827-4f16-9a0b-1c2d3e4f5604" }], "SanctionCategory": { "Id": 2, "Title": "Category B", "Summary": "Standard compliance checks required", "Color": "#ffb300" }, "MaPriority": ["0.0"], "IsSanctioned": false, "Comments": "Growing demand; moderate compliance" },
                    //         { "Id": 4, "Title": "United States", "Identifier": "US", "Capacity": "300 Gbps", "Grouping": [{ "Label": "EMEA", "TermGuid": "11112222-3333-4444-8888-aaaaaaaa0004" }], "SanctionCategory": { "Id": 1, "Title": "Category A", "Summary": "No country-level sanctions; export controls may apply", "Color": "#4caf50" }, "MaPriority": ["EU", "0.0"], "IsSanctioned": false, "Comments": "Large market; strict export controls (EAR/ITAR)" },
                    //         { "Id": 5, "Title": "Nigeria", "Identifier": "NG", "Capacity": "45 Gbps", "Grouping": [{ "Label": "EMEA", "TermGuid": "5e4d3c2b-1a09-48fe-9abc-def012345605" }, { "Label": "LATAM", "TermGuid": "abcde123-4567-89ab-cdef-012345678906" }], "SanctionCategory": { "Id": 3, "Title": "Category C", "Summary": "Heightened due diligence", "Color": "#fb8c00" }, "MaPriority": ["EU"], "IsSanctioned": false, "Comments": "Operational constraints in some regions" }
                    //     ]}>
                    // </CountryView>);
                    return (<CountryView
                        pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}></CountryView>);
                case ViewType.EutelsatEntity:
                    /*return (<EutelsatEntityView pnpService={pnpService}
                        filters={['Search', 'Title', 'Country', 'Status']}
                        columns={['Title', 'Country', 'Status', 'Summary']}
                        viewItems={[
                            {
                                "Id": 1, "Title": "Eutelsat France",
                                "Country": { "Title": "France", "Id": 1 },
                                "Status": "Active", "Summary": "French subsidiary of Eutelsat focusing on European markets."
                            },
                            {
                                "Id": 2, "Title": "Eutelsat Germany",
                                "Country": { "Title": "Germany", "Id": 3 },
                                "Status": "Inactive", "Summary": "German branch specializing in satellite communications."
                            },
                            {
                                "Id": 3, "Title": "Eutelsat Italy", "Country": { "Title": "Italy", "Id": 5 },
                                "Status": "Revoked", "Summary": "Italian entity formerly providing satellite services in the Mediterranean region."
                            }
                        ]}>
                    </EutelsatEntityView>);*/
                    return (<EutelsatEntityView
                        pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }} ></EutelsatEntityView>);
                case ViewType.Contact:
                    /*return (<ContactView pnpService={pnpService}
                        filters={['Search', 'ContentType', 'Item', 'JobTitle', 'Email', 'PhoneNumber']}
                        columns={['Title', 'ContentType', 'JobTitle', 'Email', 'PhoneNumber', 'Item', 'Summary']}
                        viewItems={[
                            {
                                "Id": 1, "Title": "John Doe",
                                "ContentType": { "Name": "Law Firm", "Id": "" },
                                "JobTitle": "Manager", "Email": "asdasd@gmail.com", "PhoneNumber": "123-456-7890",
                                'Item': { "key": "Item1-DP", "text": "Item1-DP", "id": "1", "type": "LegalSvcProvider" },
                                "Summary": "Experienced manager in the telecommunications sector."
                            },
                            {
                                "Id": 2, "Title": "Jane Smith", "ContentType": { "Name": "Legal Representative", "Id": "0x01009683AA37F0D20E469880B9832370E8260101" },
                                "JobTitle": "Director", "Email": "ad@gmail.com", "PhoneNumber": "987-654-3210",
                                'Item': { "key": "Item1-TP", "text": "Item1-TP", "id": "1", "type": "LegalSvcProvider" },
                                "Summary": "Director with extensive regulatory experience."
                            },
                        ]}
                    ></ContactView>);*/
                    return (<ContactView pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}
                    ></ContactView>);
                case ViewType.LegalSvcProvider:
                    /*return (<LegalSvcProviderView pnpService={pnpService}
                        filters={['Search', 'Title', 'ContentType', 'Country', 'Status']}
                        columns={['Title', 'Country', 'ContentType', 'Summary', 'Status']}
                        viewItems={
                            [
                                {
                                    "Id": 1, "Title": "Global Legal Advisors",
                                    "Countries": [{ "Title": "France", "Id": 1 }, { "Title": "United States", "Id": 4 }],
                                    "ContentType": { "Name": "Law Firm", "Id": "0x010002E1FD95664234498E6FD54EB499E33A01" },
                                    "Status": "Active",
                                    "Summary": "Leading international law firm with expertise in telecommunications regulations."
                                },
                                {
                                    "Id": 2, "Title": "Regulatory Experts Inc.",
                                    "Countries": [{ "Title": "Germany", "Id": 3 }, { "Title": "United Kingdom", "Id": 2 }],
                                    "ContentType": { "Name": "Legal Representative", "Id": "0x010002E1FD95664234498E6FD54EB499E33A02" },
                                    "Status": "Active",
                                    "Summary": "Specialized consultancy focusing on regulatory compliance in sanctioned regions."
                                },
                                {
                                    "Id": 3, "Title": "Compliant Counsel LLP",
                                    "Countries": [{ "Title": "Italy", "Id": 5 }],
                                    "ContentType": { "Name": "Law Firm", "Id": "0x010002E1FD95664234498E6FD54EB499E33A01" },
                                    "Status": "Revoked",
                                    "Summary": "Experienced in handling complex sanction scenarios for multinational clients."
                                }
                            ]} >
                    </ LegalSvcProviderView >);*/
                    return (<LegalSvcProviderView userService={userService}
                        pnpService={pnpService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}
                    ></LegalSvcProviderView>);
                case ViewType.DistributionPartner:
                    /*return (<DPView pnpService={pnpService}
                        filters={['Search', 'Country', 'Status']}
                        columns={['Title', 'Country', 'Status', 'Summary']}
                        viewItems={[
                            {
                                "Id": 1, "Title": "Teleport Solutions Ltd.",
                                "Status": "Active", "Summary": "Leading teleport services provider with global reach."
                            },
                            {
                                "Id": 2, "Title": "Global Teleport Services",
                                "Status": "Inactive", "Summary": "Comprehensive teleport solutions for diverse markets."
                            },
                            {
                                "Id": 3, "Title": "SatCom Teleport Inc.",
                                "Status": "Revoked", "Summary": "Specialized in satellite communication and teleport services."
                            }
                        ]}></DPView>);*/
                    return (<DPView
                        pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}
                    ></DPView>);
                case ViewType.TeleportPartner:
                    /*return (<TPView pnpService={pnpService}
                        filters={['Search', 'Country', 'Status']}
                        columns={['Title', 'Country', 'Status', 'Summary']}
                        viewItems={[
                            {
                                "Id": 1, "Title": "Teleport Solutions Ltd.",
                                "Countries": [{ "Title": "France", "Id": 1 }],
                                "Status": "Active", "Summary": "Leading teleport services provider with global reach."
                            },
                            {
                                "Id": 2, "Title": "Global Teleport Services",
                                "Countries": [{ "Title": "Germany", "Id": 3 }],
                                "Status": "Inactive", "Summary": "Comprehensive teleport solutions for diverse markets."
                            },
                            {
                                "Id": 3, "Title": "SatCom Teleport Inc.", "Countries": [{ "Title": "Italy", "Id": 5 }],
                                "Status": "Revoked", "Summary": "Specialized in satellite communication and teleport services."
                            }
                        ]}></TPView>);*/
                    return (<TPView pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}
                    ></TPView>);
                case ViewType.Authority:
                    /*  return (<AuthorityView pnpService={pnpService}
                          columns={['Title', 'Country', 'AuthorityType', 'Contact', 'Address', 'Summary']}
                          filters={['Search', 'Countries']}
                          viewItems={[
                              { "Id": 1, "Title": "FDA", "Countries": [{ "Title": "France", "Id": 1 }], "Owners": [{ Id: 1, Title: "contact@fda.gov" }], "Address": "10903 New Hampshire Ave, Silver Spring, MD 20993", "Summary": "test", },
                              { "Id": 2, "Title": "EMA", "Countries": [{ "Title": "France", "Id": 1 }], "Owners": [{ Id: 1, Title: "contact@fda.gov" }], "Address": "Domenico Scarlattilaan 6, Amsterdam", "Summary": "test", },
                              { "Id": 3, "Title": "MHRA", "Countries": [{ "Title": "France", "Id": 1 }], "Owners": [{ Id: 1, Title: "contact@fda.gov" }], "Address": "10 South Colonnade, London", "Summary": "test", },
                              { "Id": 4, "Title": "PMDA", "Countries": [{ "Title": "France", "Id": 1 }], "Owners": [{ Id: 1, Title: "contact@fda.gov" }], "Address": "3-3-2 Kasumigaseki, Tokyo", "Summary": "test", },
                              { "Id": 5, "Title": "TGA", "Countries": [{ "Title": "France", "Id": 1 }], "Owners": [{ Id: 1, Title: "contact@fda.gov" }], "Address": "136 Narrabundah Lane, Symonston", "Summary": "test", }
                          ]}>
                      </AuthorityView>);*/
                    return (<AuthorityView
                        pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}
                    ></AuthorityView>);
                case ViewType.SNP:

                    /*return (<SNPView pnpService={pnpService}
                        columns={['Title', 'Country', 'ContentType', 'City', 'TeleportPartner', 'Status']}
                        filters={['Search', 'Country', 'ContentType', 'TeleportPartner', 'Status']}
                        viewItems={[
                            { "Id": 1, "Title": "FDA", "Country": { "Title": "France", "Id": 1 }, ContentType: { Name: "Fee", Id: '0x01007DBD40A3236C4447848111DE54A6D86301' }, "TeleportPartner": { Id: 1, Title: "Tp-1" }, "Status": "Active", },
                            { "Id": 2, "Title": "EMA", "Country": { "Title": "France", "Id": 1 }, ContentType: { Name: "Fee", Id: '0x01007DBD40A3236C4447848111DE54A6D86302' }, "TeleportPartner": { Id: 2, Title: "Tp-2" }, "Summary": "test", },
                            { "Id": 3, "Title": "MHRA", "Country": { "Title": "France", "Id": 1 }, ContentType: { Name: "Fee", Id: '0x01007DBD40A3236C4447848111DE54A6D86302' }, "TeleportPartner": { Id: 4, Title: "Tp-2" }, "Summary": "test", },
                            { "Id": 4, "Title": "PMDA", "Country": { "Title": "France", "Id": 1 }, ContentType: { Name: "Fee", Id: '0x01007DBD40A3236C4447848111DE54A6D86302' }, "TeleportPartner": { Id: 5, Title: "Tp-6" }, "Summary": "test", },
                            { "Id": 5, "Title": "TGA", "Country": { "Title": "France", "Id": 1 }, ContentType: { Name: "Fee", Id: '0x01007DBD40A3236C4447848111DE54A6D86301' }, "TeleportPartner": { Id: 10, Title: "Tp-5" }, "Summary": "test", }
                        ]}></SNPView>);*/
                    return (<SNPView pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}
                    ></SNPView>);
                case ViewType.FEE:
                    /*  return (<FeeView pnpService={pnpService}
                          columns={['Title', 'ContentType', 'Item', 'CostTotal', 'Summary', 'Status', 'DueDate', 'ChargeByItem']}
                          filters={['Search', 'ContentType', 'Status', 'Item', 'ChargeByItem']}
                          viewItems={[
                              {
                                  Id: 1,
                                  Title: 'Setup Fee',
                                  Category: 'Implementation',
                                  Summary: 'One-time setup and configuration',
                                  ContentType: { Name: 'Fee' },
                                  Item: {
                                      key: '101', text: 'PO-1001', type: ViewType.Authority, id: "1"
                                  },
                                  ChargedBy: 'Vendor A',
                                  ChargeByItem: { Id: 201, Title: 'Service Line' },
                                  Currency: { Id: 1, Title: 'USD', },
                                  VatFreeCost: '0.00',
                                  Vat: '20.00',
                                  VatRate: '20%',
                                  CostType: 'One-time',
                                  PoNonPO: 'PO',
                                  Status: 'Open',
                                  DueDate: new Date('2025-01-15'),
                                  Recurrence: 'None',
                                  RecurranceFee: false,
                                  recurrencePattern: null,
                                  RecurrencPatternId: null,
                                  Comments: 'Charged once at project start',
                                  CostTotal: '120.00'
                              },
                              {
                                  Id: 2,
                                  Title: 'Monthly Support',
                                  Category: 'Support',
                                  Summary: 'Ongoing monthly support fee',
                                  ContentType: { Name: "Fee", Id: '0x010092D96D11717F014497021E5516FB818201' },
                                  Item: {
                                      key: '102', text: 'PO-1001', type: ViewType.Payment, id: "2"
                                  },
                                  ChargedBy: 'Vendor B',
                                  ChargeByItem: {
                                      key: '102', text: 'PO-1001', type: ViewType.Payment, id: "2"
                                  },
                                  Currency: { Id: 2, Title: 'Euro' },
                                  VatFreeCost: '0.00',
                                  Vat: '10.00',
                                  VatRate: '10%',
                                  CostType: 'Recurring',
                                  PoNonPO: 'PO',
                                  Status: 'Active',
                                  DueDate: new Date('2025-02-01'),
                                  Recurrence: 'Monthly',
                                  RecurranceFee: true,
                                  // recurrencePattern: { frequency: 'Monthly', interval: 1 },
                                  RecurrencPatternId: 1,
                                  Comments: 'Billed on the 1st of each month',
                                  CostTotal: '110.00'
                              },
                              {
                                  Id: 3,
                                  Title: 'License Fee',
                                  Category: 'License',
                                  Summary: 'Annual software license',
                                  ContentType: { Name: 'Fee' },
                                  Item: {
                                      key: '102', text: 'PO-1001', type: ViewType.Payment, id: "2"
                                  },
                                  ChargedBy: 'Vendor C',
                                  ChargeByItem: {
                                      key: '102', text: 'PO-1001', type: ViewType.Payment, id: "2"
                                  },
                                  Currency: { Id: 3, Title: 'British Pound' },
                                  VatFreeCost: '0.00',
                                  Vat: '40.00',
                                  VatRate: '20%',
                                  CostType: 'Recurring',
                                  PoNonPO: 'PO',
                                  Status: 'Pending',
                                  DueDate: new Date('2025-06-30'),
                                  Recurrence: 'Yearly',
                                  RecurranceFee: true,
                                  //  recurrencePattern: { frequency: 'Yearly', interval: 1 },
                                  RecurrencPatternId: 2,
                                  Comments: 'Renewal due mid-year',
                                  CostTotal: '240.00'
                              },
                              {
                                  Id: 4,
                                  Title: 'Travel Expense',
                                  Category: 'Expense',
                                  Summary: 'On-site visit travel reimbursement',
                                  ContentType: { Name: 'Fee' },
                                  Item: {
                                      key: '102', text: 'PO-1001', type: ViewType.Payment, id: "2"
                                  },
                                  ChargedBy: 'Consultant D',
                                  ChargeByItem: {
                                      key: '102', text: 'PO-1001', type: ViewType.Payment, id: "2"
                                  },
                                  Currency: { Id: 5, Title: 'US Dollar' },
                                  VatFreeCost: '0.00',
                                  Vat: '0.00',
                                  VatRate: '0%',
                                  CostType: 'One-time',
                                  PoNonPO: 'Non-PO',
                                  Status: 'Paid',
                                  DueDate: new Date('2024-12-10'),
                                  Recurrence: 'None',
                                  RecurranceFee: false,
                                  recurrencePattern: null,
                                  RecurrencPatternId: null,
                                  Comments: 'Reimbursed after trip',
                                  CostTotal: '350.00'
                              },
                              {
                                  Id: 5,
                                  Title: 'Consultancy Retainer',
                                  Category: 'Consulting',
                                  Summary: 'Quarterly retainer for expert consultancy',
                                  ContentType: { Name: 'Fee' },
                                  Item: {
                                      key: '102', text: 'PO-1001', type: ViewType.Payment, id: "2"
                                  },
                                  ChargedBy: 'Consultancy E',
                                  ChargeByItem: {
                                      key: '102', text: 'PO-1001', type: ViewType.Payment, id: "2"
                                  },
                                  Currency: { Id: 7, Title: 'US Dollar' },
                                  VatFreeCost: '0.00',
                                  Vat: '75.00',
                                  VatRate: '15%',
                                  CostType: 'Recurring',
                                  PoNonPO: 'PO',
                                  Status: 'Active',
                                  DueDate: new Date('2025-03-01'),
                                  Recurrence: 'Quarterly',
                                  RecurranceFee: true,
                                  // recurrencePattern: { frequency: 'Monthly', interval: 3 },
                                  RecurrencPatternId: 3,
                                  Comments: 'Quarterly invoicing',
                                  CostTotal: '575.00'
                              }
                          ]}></FeeView>);*/
                    return (<FeeView pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}
                    ></FeeView>);
                case ViewType.Payment:
                    return (<PaymentView pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}
                    ></PaymentView>);
                case ViewType.MARequirement:
                    return (<MARequirementView pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}
                    ></MARequirementView>);
                case ViewType.Market:
                    return (<MarketView pnpService={pnpService} userService={userService}
                        pagination={{ isEnabled: true, showTopPagination: true, showBottomPagination: true }}
                    ></MarketView>);
                default:
                    return (<div>No types found</div>);
            }
        } else {
            return (<div>No types found</div>);
        }
    }
}
