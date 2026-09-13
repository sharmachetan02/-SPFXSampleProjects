import * as React from 'react';
import { IContact, IMarketReadiness, IItemOption, ISanctionCategory, Iterm } from '../../../../common/models/IBusiness';
import { Accordion, AccordionItem, AccordionHeader, AccordionPanel } from '../../../../common/components/Accordion/Accordion';
import * as strings from 'GenericDatasheetStrings';
import { Label, ActionButton, Modal, Icon, TooltipHost, ITooltipHostStyles, DialogFooter, PrimaryButton, DefaultButton, Dialog, DialogType, Stack, Spinner, SpinnerSize } from '@fluentui/react';
import styles from '../Datasheet.module.scss';
import { Form, FormType, IMessageBanner } from '../../../genericForm/components/Form.types';
import { IMarketReadinessDatasheetViewProps, IMarketReadinessDatasheetViewState } from './DatasheetViews.types';
import MarketReadiness from '../../../genericForm/components/Market Readiness/MarketReadiness';
import { MarketReadinessResponsibleParty, UserRole, Vertical } from '../../../../common/models/Enums';
import { renderArrayPills, renderObjectPills, StatusPill } from '../Datasheet.utility';
import { UtilHelper } from '../../../../common/helpers/Util';
import { RichText } from '../../../../common/components/RichText/RichText';
import { Config } from '../../../../common/config/Config';
import { Consts } from '../../../../common/consts/Consts';
import { DatasheetType } from '../Datasheet.types';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { MessageBanner } from '../../../../common/components/Message Banner/MessageBanner';
import DomHelper from '../../../../common/helpers/DomHelper';
export default class MarketReadinessDatasheetView extends React.Component<IMarketReadinessDatasheetViewProps, IMarketReadinessDatasheetViewState> {

    private DPandTPDataInfo: IItemOption[] = [];
    constructor(props: IMarketReadinessDatasheetViewProps) {
        super(props);
        this.state = {
            showFormDialog: false,
            itemId: undefined,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'error',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        };
    }
    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };
    private renderMarketReadinessInfoRow = (readinessItem: IMarketReadiness, key: string) => {
        const labelKey = `MarketReadinessLabel${key}`;
        return (
            <div className="ms-Grid">
                <div className={`ms-Grid-row ${styles.datasheetInfoRow}`}>
                    <div className="ms-Grid-col ms-sm12 ms-md12">
                        <div className="ms-Grid-col ms-sm3 ms-md3">
                            <div className={styles.datasheetInfoLabel}>
                                <Label>{`${strings[labelKey]}:`}</Label>
                            </div>
                        </div>
                        <div className="ms-Grid-col ms-sm9 ms-md9">
                            <div>
                                {this.renderDatasheetInfoValue(readinessItem, key)}
                            </div>
                        </div>
                    </div>
                </div>
            </div>);
    };

    private renderDatasheetInfoValue = (item, fieldName) => {
        const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        const value = fieldName ? item[fieldName] : undefined;
        const swatchStyle = (color?: string) => ({
            width: 12,
            height: 12,
            borderRadius: 3,
            background: color || '#888',
            display: 'inline-block',
            verticalAlign: 'middle' as const,
            marginRight: 6,
            border: '1px solid rgba(0,0,0,.1)',
        });
        if (fieldName === 'ResponsiblePartyItem') {
            const options = (this.DPandTPDataInfo ?? []) as IItemOption[];

            return renderObjectPills(
                options,
                (t) => {
                    const viewType = t.type;
                    if (!viewType || !t.id) return null;

                    return (
                        <span key={t.key ?? t.id} className={styles.termPill}>
                            <span className={styles.termPillText}>
                                <a
                                    data-interception='off'
                                    rel="noreferrer"
                                    title={String(t.text ?? '')}
                                    onClick={(e) => {
                                        e.preventDefault();
                                        const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${viewType}&itemId=${t.id}`;
                                        UrlHelper.navigate(url, false);
                                    }}
                                >
                                    {String(t.text ?? '')}
                                </a>
                            </span>
                        </span>
                    );
                }
            );
        }
        if (value === null || value === undefined || (Array.isArray(value) && value.length === 0)) {
            return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
        }
        if (fieldName === 'Grouping') {
            return (
                <div className={styles.termPills}>
                    {(value as Iterm[]).map((t) => (
                        <span key={t.TermGuid} className={styles.termPill}>
                            <span className={styles.termPillText}>{t.Label}</span>
                        </span>
                    ))}
                </div>
            );
        }
        if (fieldName === 'SanctionCategory') {
            const sc = value as ISanctionCategory;
            return (
                <TooltipHost content={<div className={styles.tooltipHost} dangerouslySetInnerHTML={{ __html: sc.Summary }}></div>}
                    styles={tooltipStyles} calloutProps={{
                        layerProps: { hostId: 'sanctionCategoryLayer' },
                        styles: { root: { zIndex: 500002 } }
                    }}>
                    <div className={styles.tooltipHostContainer} title="">
                        <span style={swatchStyle(sc.Color)} />
                        <span className={styles.tooltipHostText}> {sc.Title}</span>
                    </div>
                </TooltipHost>
            );
        }
        if (fieldName === 'EutelsatOwners') {
            return renderObjectPills(
                value as IContact[],
                (contact) => (
                    <span key={contact.Id} className={styles.termPill}>
                        <span className={styles.termPillText}>
                            <a data-interception="on"
                                onClick={(e) => {
                                    e.preventDefault();
                                    const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Contact)}&itemId=${contact.Id}`;
                                    UrlHelper.navigate(url, false);
                                }}
                                rel="noreferrer" title={contact.Title}>{contact.Title}</a>
                        </span>
                    </span>
                )
            );
        }


        if ([
            'Comments',
            'AllVerticalComments',
            'SpaceComments',
            'LandFixedComments',
            'LandMobilityComments',
            'MaritimeComments',
            'AviationComments'
        ].includes(fieldName)) {
            //return <RichText isEditMode={false} value={value} />;
            const cleanValue = DomHelper.cleanRichHtml(item[fieldName]);
            if (!cleanValue) return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
            return <RichText isEditMode={false} value={cleanValue} />;
        }
        if ([
            'AllVerticalRAGStatus',
            'SpaceRAGStatus',
            'LandFixedRAGStatus',
            'LandMobilityRAGStatus',
            'MaritimeRAGStatus',
            'AviationRAGStatus'
        ].includes(fieldName)) {
            return StatusPill(String(value));
        }
        if (typeof value === 'boolean') {
            const iconName = value ? 'AcceptMedium' : 'Cancel';
            const iconClass = value ? styles.iconGreen : styles.iconRed;
            return <div className={styles.detailsListValue}><Icon iconName={iconName} className={iconClass} /></div>;
        }
        if (value instanceof Date) {
            return <span className={styles.detailsListValue}>{UtilHelper.formatDate(value, 'LL', 'en-us')}</span>;
        }
        if (typeof value === 'string' || typeof value === 'number') {
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }
        if (Array.isArray(value) && value?.every((v) => typeof v === 'string' || typeof v === 'number')) {
            return renderArrayPills(value as (string | number)[]);
        }
    };

    private renderMarketReadinessInfo = (readinessItem: IMarketReadiness, vertical: string): JSX.Element | null => {
        let leftColumns, rightColumns = [];
        const allVerticalsLeftColumns = ['AllVerticalRAGStatus', 'AllVerticalEstimatedDate', 'AllVerticalEffectiveDate', 'EutelsatOwners'];
        let allVerticalsRightColumns = ['AllVerticalComments'];
        if (MarketReadinessResponsibleParty.Overall === readinessItem.ResponsiblePartyType) {
            allVerticalsRightColumns = ['AllVerticalComments', 'ResponsiblePartyItem'];
        }
        const spaceLeftColumns = ['SpaceRAGStatus', 'SpaceEstimatedDate', 'SpaceEffectiveDate'];
        const spaceRightColumns = ['SpaceComments'];
        const landFixedLeftColumns = ['LandFixedRAGStatus', 'LandFixedEstimatedDate', 'LandFixedEffectiveDate'];
        const landFixedRightColumns = ['LandFixedComments'];
        const landMobilityLeftColumns = ['LandMobilityRAGStatus', 'LandMobilityEstimatedDate', 'LandMobilityEffectiveDate'];
        const landMobilityRightColumns = ['LandMobilityComments'];
        const maritimeLeftColumns = ['MaritimeRAGStatus', 'MaritimeEstimatedDate', 'MaritimeEffectiveDate'];
        const maritimeRightColumns = ['MaritimeComments'];
        const aviationLeftColumns = ['AviationRAGStatus', 'AviationEstimatedDate', 'AviationEffectiveDate'];
        const aviationRightColumns = ['AviationComments'];
        switch (vertical) {
            case Vertical.AllVerticals:
                leftColumns = allVerticalsLeftColumns;
                rightColumns = allVerticalsRightColumns;
                break;

            case Vertical.SpaceNW:
                leftColumns = spaceLeftColumns;
                rightColumns = spaceRightColumns;
                break;

            case Vertical.LandFixed:
                leftColumns = landFixedLeftColumns;
                rightColumns = landFixedRightColumns;
                break;

            case Vertical.LandMobility:
                leftColumns = landMobilityLeftColumns;
                rightColumns = landMobilityRightColumns;
                break;

            case Vertical.Maritime:
                leftColumns = maritimeLeftColumns;
                rightColumns = maritimeRightColumns;
                break;

            case Vertical.Aviation:
                leftColumns = aviationLeftColumns;
                rightColumns = aviationRightColumns;
                break;
        }
        return (
            <div className={`${styles.datasheetInfoContainer}`}>
                <div className={styles.datasheetInfoContainerLeft}>
                    {leftColumns.map((key) =>
                        (this.renderMarketReadinessInfoRow(readinessItem, key))
                    )}
                </div>
                <div className={styles.datasheetInfoContainerRight}>
                    {rightColumns.map((key) =>
                        (this.renderMarketReadinessInfoRow(readinessItem, key))
                    )}
                </div>
            </div>
        );
    };

    private renderMarketReadiness = (): JSX.Element | null => {
        const marketreadinessItems = this.props.items;

        if (!marketreadinessItems || marketreadinessItems.length === 0) {
            return (
                <div className={styles.noDataMessage}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span>{strings.DataSheetNoRecordFound}</span>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                </div>
            );
        }

        marketreadinessItems.forEach((readinessItem) => {
            const partyItem = readinessItem.ResponsiblePartyItem;
            if (!partyItem) return;

            let type = '';
            if (readinessItem.ResponsiblePartyType === MarketReadinessResponsibleParty.DP) {
                type = DatasheetType.DistributionPartner;
            } else if (readinessItem.ResponsiblePartyType === MarketReadinessResponsibleParty.TP) {
                type = DatasheetType.TeleportPartner;
            } else {
                type = null;
            }


            if (!type) return;

            const exists = this.DPandTPDataInfo.some(
                (d) => d.id === String(partyItem.Id) && d.key === partyItem.Title && d.type === type
            );

            if (!exists) {
                this.DPandTPDataInfo.push({
                    id: String(partyItem.Id),
                    key: partyItem.Title,
                    text: partyItem.Title,
                    type
                });
            }
        });
        const sortedMarketReadinessItems = marketreadinessItems.sort((a, b) => {
            const order = [MarketReadinessResponsibleParty.Overall, MarketReadinessResponsibleParty.Eutelsat, MarketReadinessResponsibleParty.DP, MarketReadinessResponsibleParty.TP];
            return order.indexOf(a.ResponsiblePartyType as MarketReadinessResponsibleParty) - order.indexOf(b.ResponsiblePartyType as MarketReadinessResponsibleParty);
        });
        const itemsHtml = sortedMarketReadinessItems.map((readinessItem) => {
            //const itemsHtml = countryreadinessItems.map((readinessItem) => {
            let responsibleTitle = "";
            let responsibleIcon = "";

            //console.log("readinessItem", readinessItem);
            switch (readinessItem.ResponsiblePartyType) {
                case MarketReadinessResponsibleParty.Overall:
                    responsibleTitle = `${readinessItem.ResponsiblePartyType}`;
                    responsibleIcon = 'Country';
                    break;
                case MarketReadinessResponsibleParty.Eutelsat:
                    responsibleIcon = 'Satellite';
                    responsibleTitle = `${readinessItem.ResponsiblePartyType}`;
                    break;
                case MarketReadinessResponsibleParty.DP:
                    responsibleIcon = 'Agreement';
                    responsibleTitle = `${readinessItem.ResponsiblePartyType} - ${readinessItem.ResponsiblePartyItem?.Title}`;
                    break;
                case MarketReadinessResponsibleParty.TP:
                    responsibleIcon = 'Dish';
                    responsibleTitle = `${readinessItem.ResponsiblePartyType} - ${readinessItem.ResponsiblePartyItem?.Title}`;
                    break;
            }
            const verticalsHtml = Object.values(Vertical)
                .filter((v) => v !== Vertical.SNP)
                .map((vertical) => {
                    let icon: string;
                    switch (vertical) {
                        case Vertical.AllVerticals:
                            icon = 'Binoculars';
                            break;
                        case Vertical.SpaceNW:
                            icon = 'Rocket';
                            break;
                        case Vertical.LandFixed:
                            icon = 'Wifi';
                            break;
                        case Vertical.LandMobility:
                            icon = 'Truck';
                            break;
                        case Vertical.Maritime:
                            icon = 'Ship';
                            break;
                        case Vertical.Aviation:
                            icon = 'Plane';
                            break;
                    }
                    return (readinessItem.Verticals.find((v) => v.Title === vertical) &&
                        <AccordionItem value={vertical} key={vertical} >
                            <AccordionHeader icon={icon}>{vertical}</AccordionHeader>
                            <AccordionPanel>
                                {
                                    this.renderMarketReadinessInfo(readinessItem, vertical)
                                }
                            </AccordionPanel>
                        </AccordionItem>
                    );
                }
                );
            return (
                <>
                    <AccordionItem value={responsibleTitle + readinessItem.Id} key={responsibleTitle} >
                        <AccordionHeader icon={responsibleIcon}>{responsibleTitle}</AccordionHeader>
                        <AccordionPanel>
                            {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.readinessCommandsBar}>
                                <ActionButton iconProps={{ iconName: 'Edit' }} onClick={() => this.openSNPMarketReadinessEditForm(readinessItem.Id)}>
                                    {strings.SnpReadinessLabelEditButton}
                                </ActionButton>
                                <ActionButton iconProps={{ iconName: 'Delete' }} onClick={() => this.openDeleteDialog(readinessItem)}>
                                    {strings.SnpReadinessLabelDeleteButton}
                                </ActionButton>
                            </div>}
                            <div className={styles.itemAccordionContainer}>
                                <Accordion
                                    multiple
                                    collapsible
                                >
                                    {verticalsHtml}
                                </Accordion>
                            </div>
                        </AccordionPanel>
                    </AccordionItem>
                </>
            );
        });
        return (
            <>

                <div className={styles.readinessContainer}>
                    <Accordion multiple collapsible>
                        {itemsHtml}
                    </Accordion>
                </div>
            </>
        );
    };

    public render(): React.ReactElement<IMarketReadinessDatasheetViewProps> {
        const {
            showFormDialog,
            showDeleteDialog
        } = this.state;
        return (
            <>
                <div className={styles.readinessContainer}>
                    {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.readinessNewButton}>
                        <ActionButton iconProps={{ iconName: 'Add' }} onClick={this.openSNPMarketReadinessNewForm}>
                            {strings.MarketReadinessLabelNewButton}
                        </ActionButton>
                    </div>}
                    {this.renderMarketReadiness()}
                </div>
                {
                    showFormDialog &&
                    this.renderFormDialog()
                }

                {showDeleteDialog && this.renderDeleteDialog()}

            </>

        );
    }

    private openSNPMarketReadinessNewForm = () => {
        this.setState({
            showFormDialog: true,
            itemId: undefined,
        });
    };

    private openSNPMarketReadinessEditForm = (itemId) => {
        this.setState({ showFormDialog: true, itemId });
    };

    private closeFormDialog = () => {
        this.setState({ showFormDialog: false, showDeleteDialog: false, documentToDelete: null, isDeleteDisabled: false });

    };

    private openDeleteDialog = (item: IMarketReadiness) => {
        this.setState({
            showDeleteDialog: true, documentToDelete: item,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'error',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        });

    };

    private deleteDataSheet = async () => {
        try {
            this.setState({ isFormProcessing: true });
            const { pnpService } = this.props;
            this.setState({ isDeleteDisabled: true });
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), Consts.LISTS.MA_READINESS_URL);
            await pnpService.getListItems(listUrl).getById(this.state.documentToDelete?.Id).recycle();
            const messageBanner: IMessageBanner = {
                message: strings.DeleteSuccessMessage,
                type: 'success',
                visible: true
            };
            this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
            this.refresh(true);
        } catch (error) {
            console.error(strings.DeleteErrorMessage, error);
            const messageBanner: IMessageBanner = {
                message: strings.DeleteErrorMessage,
                type: 'error',
                visible: true
            };
            this.setState({ isFormProcessing: false, messageBanner });
        }
    };

    private refresh = async (refreshData) => {
        this.closeFormDialog();
        if (refreshData) {
            await this.props.refresh();
        }
    };

    private renderFormDialog() {
        const {
            pnpService,
            marketItemId,
            userService
        } = this.props;
        const {
            showFormDialog,
            itemId,
        } = this.state;
        let formControl, formTitle = null;
        const formConfig = Form.config.find((f) => f.type.toLowerCase() === FormType.MarketReadiness);
        if (itemId) {
            formControl = (<MarketReadiness pnpService={pnpService} userService={userService} itemId={itemId} callback={() => this.refresh(true)}></MarketReadiness>);
            formTitle = `Edit ${formConfig.name}`;
        } else {
            formControl = (<MarketReadiness pnpService={pnpService} userService={userService}
                preselectedItemId={marketItemId} callback={() => this.refresh(true)}></MarketReadiness>);
            formTitle = `New ${formConfig.name}`;
        }
        return (<Modal
            isOpen={showFormDialog}
            onDismiss={this.closeFormDialog}
            containerClassName={styles.modalContainer}
            isBlocking={true} >
            <div>
                <div className={styles.modalHeader}>
                    {formTitle}
                    <ActionButton
                        iconProps={{ iconName: 'Cancel' }}
                        onClick={this.closeFormDialog}
                        className={styles.iconButton}
                    >Close</ActionButton>
                </div>
                <div className={styles.modalBody}>
                    {formControl}
                </div>
            </div>
        </Modal>);
    }

    private renderDeleteDialog() {
        const { documentToDelete } = this.state;
        return (
            <Dialog
                hidden={!documentToDelete}
                onDismiss={this.closeFormDialog}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: documentToDelete?.Title || 'Delete'
                }}
                modalProps={{ isBlocking: true }}
                styles={{
                    main: {
                        selectors: {
                            '@media (min-width: 480px)': {
                                maxWidth: '410px !important', // Override media query
                            },
                        },

                    }
                }}
            >
                {this.state.messageBanner.visible && (
                    <div className={styles.topMessageBanner}>
                        <MessageBanner
                            message={this.state.messageBanner.message}
                            type={this.state.messageBanner.type}
                            visible={this.state.messageBanner.visible}
                        />
                    </div>
                )}
                {this.state.messageBanner.type !== 'success' && (
                    <span>{strings.LibraryViewRecycleBinConfirmation}</span>
                )}
                <DialogFooter>
                    <PrimaryButton
                        disabled={this.state.isConfirmButtonDisabled}
                        onClick={this.deleteDataSheet} >
                        {this.state.isFormProcessing ? (
                            <Stack horizontal verticalAlign="center">
                                <Spinner size={SpinnerSize.xSmall} labelPosition="right" />
                                <span style={{ marginLeft: 8 }}>Processing</span>
                            </Stack>
                        ) : (
                            strings.DataSheetConfirmButton
                        )}
                    </PrimaryButton>
                    <DefaultButton
                        onClick={this.closeFormDialog}
                        text="Back"
                    />
                </DialogFooter>
            </Dialog>
        );
    }
}
