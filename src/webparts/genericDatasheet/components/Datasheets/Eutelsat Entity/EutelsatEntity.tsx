import * as React from 'react';
import { IEutelsatEntityDatasheetProps, IEutelsatEntityDatasheetState } from './EutelsatEntity.types';
import styles from '../../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import {
    ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, Icon, ITooltipHostStyles, Label,
    LayerHost, Modal, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, Stack, TooltipHost

} from '@fluentui/react';
import EutelsatEntity from '../../../../genericForm/components/Eutelsat Entity/EutelsatEntity';
import { Datasheet, DatasheetType, IDatasheetConfig } from '../../Datasheet.types';
import { UrlHelper } from '../../../../../common/helpers/UrlHelper';
import { Config } from '../../../../../common/config/Config';
import { Consts } from '../../../../../common/consts/Consts';
import { IEutelsatEntity, ICountry } from '../../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../../common/helpers/Util';
import { renderArrayPills, StatusPill } from '../../Datasheet.utility';
import { RichText } from '../../../../../common/components/RichText/RichText';
import LibraryView from '../../../../genericView/components/Library View/LibraryView';
import { UserRole } from '../../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../../common/components/Access Denied/AccessDenied';
import { IMessageBanner } from '../../../../genericForm/components/Form.types';
import { MessageBanner } from '../../../../../common/components/Message Banner/MessageBanner';
import DomHelper from '../../../../../common/helpers/DomHelper';
export default class EutelsatEntityDatasheet extends React.Component<
    IEutelsatEntityDatasheetProps,
    IEutelsatEntityDatasheetState
> {
    private config: IDatasheetConfig;
    //private allViewItems: ISNP[] = [];
    constructor(props: Readonly<IEutelsatEntityDatasheetProps>) {
        super(props);
        this.state = {
            isDatasheetReady: false,
            showEditFormDialog: false,
            showDeleteDialog: false,
            item: null,
            errorMessage: null,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'error',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        };
        this.config = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.EutelsatEntity);
        //console.log("Teleport partner view config:", this.config);
    }
    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };
    public async componentDidMount() {
        if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner,
            UserRole.Contributor, UserRole.Visitor, UserRole.Admin)) {
            await this.init();
            this.setState({ isDatasheetReady: true });
        }
    }
    public initItem = async () => {
        const { pnpService, itemId } = this.props;
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.EUT_ENTITY.STATUS,
            Consts.FIELDS.EUT_ENTITY.SUMMARY,
            Consts.FIELDS.EUT_ENTITY.COMMENTS,
            Consts.FIELDS.EUT_ENTITY.ADDRESS,
            Consts.FIELDS.COMMON.CREATION_TIME,
            Consts.FIELDS.COMMON.MODIFICATION_TIME,
            Consts.FIELDS.COMMON.DOCUMENTS_SPACE,
            `${Consts.FIELDS.EUT_ENTITY.COUNTRY}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.EUT_ENTITY.COUNTRY}/${Consts.FIELDS.COMMON.TITLE}`
        ];
        const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.EutelsatEntity).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        try {
            const item = await this.props.pnpService.getListItems(listUrl).
                getById(itemId).
                select(...selectedFields).
                expand(Consts.FIELDS.EUT_ENTITY.COUNTRY)();
            const eutelsatEntity: IEutelsatEntity = {
                Id: item[Consts.FIELDS.COMMON.ID],
                Title: item[Consts.FIELDS.COMMON.TITLE],
                Status: item[Consts.FIELDS.EUT_ENTITY.STATUS],
                Summary: item[Consts.FIELDS.EUT_ENTITY.SUMMARY],
                Comments: item[Consts.FIELDS.EUT_ENTITY.COMMENTS],
                Address: item[Consts.FIELDS.EUT_ENTITY.ADDRESS],
                Country: {
                    Id: item[Consts.FIELDS.EUT_ENTITY.COUNTRY][Consts.FIELDS.COMMON.ID],
                    Title: item[Consts.FIELDS.EUT_ENTITY.COUNTRY][Consts.FIELDS.COMMON.TITLE]
                },
                Created: item[Consts.FIELDS.COMMON.CREATION_TIME],
                Modified: item[Consts.FIELDS.COMMON.MODIFICATION_TIME],
                DocumentSpace: item[Consts.FIELDS.COMMON.DOCUMENTS_SPACE]?.Url
            };
            return eutelsatEntity;
        } catch (error) {
            if (error instanceof Error && error.message.includes(strings.DataSheetItemDoesNotExist)) {
                const errorMessage = strings.DataSheetItemDoesNotExist;
                this.setState({ errorMessage });
                return null;

            }
            throw error;
        }
    };
    private init = async () => {
        let item: IEutelsatEntity;
        const promises: Promise<void>[] = [];
        promises.push(this.initItem().then((opts) => { item = opts as IEutelsatEntity; }));
        //promises.push(this.initSanctionCategoriesChoices().then((opts) => { sanctions = opts as ISanctionOption[]; }));
        await Promise.all(promises);

        /*const sanctionCategory = sanctions.find((cat) => cat.key.toString() === item.SanctionCategory.Id.toString());
        item.SanctionCategory = {
          Id: item.SanctionCategory.Id,
          Title: sanctionCategory.text,
          Color: sanctionCategory.color,
          Summary: sanctionCategory.summary
        };*/
        this.setState({
            item,
        });
        //console.log('item', item);
    };

    private renderDatasheetInfo = (): JSX.Element | null => {
        const leftColumns = ['Summary', 'Address', 'Country'];
        const rightColumns = ['Status', 'Comments'];

        return (
            <div className={`${styles.datasheetInfoContainer}`}>
                <div className={styles.datasheetInfoContainerLeft}>
                    {leftColumns.map((key) =>
                        (this.renderDatasheetInfoRow(key))
                    )}
                </div>
                <div className={styles.datasheetInfoContainerRight}>
                    {rightColumns.map((key) =>
                        (this.renderDatasheetInfoRow(key))
                    )}
                </div>
            </div>
        );
    };

    private renderDatasheetInfoRow = (key: string) => {
        const {
            item
        } = this.state;
        const labelKey = `EutelsatEntityDataSheetInfoLabel${key}`;
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
                                {this.renderDatasheetInfoValue(item, key)}
                            </div>
                        </div>
                    </div>
                </div>
            </div>);
    };


    private renderDatasheetInfoValue = (item, fieldName) => {
        //const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        //const value = fieldName ? item[fieldName] : undefined;
        /*const swatchStyle = (color?: string) => ({
          width: 12,
          height: 12,
          borderRadius: 3,
          background: color || '#888',
          display: 'inline-block',
          verticalAlign: 'middle' as const,
          marginRight: 6,
          border: '1px solid rgba(0,0,0,.1)',
        });*/

        const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        const itemDatasheetUrl = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(this.config.type)}&itemId=${item.Id}`;
        //if (!column) return <span />;
        //const fieldName = column.fieldName as keyof IEutelsatEntity | undefined;
        const value = fieldName ? item[fieldName] : undefined;
        if (value === null || value === undefined || (Array.isArray(value) && value.length === 0)) {
            return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
        }
        if (fieldName === 'Title') {
            return (
                <a
                  data-interception="on"
                  rel="noreferrer"
                  onClick={(e) => {
                    e.preventDefault();
                    UrlHelper.navigate(itemDatasheetUrl, false);
                  }}
                >
                  {String(value)}
                </a>
            );
        }
        if (fieldName === 'Country') {
            const ct = item[fieldName] as ICountry;
            return (<div className={styles.termPills}>
                <span key={ct.Id} className={styles.termPill}>
                    <span className={styles.termPillText}>
                        <a data-interception="on" rel="noreferrer"
                            onClick={(e) => {
                                e.preventDefault();
                                const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Country)}&itemId=${ct.Id}`;
                                UrlHelper.navigate(url, false);
                            }}
                            title={String(ct?.['Title'])} >{String(ct?.['Title'])}</a>
                    </span>
                </span>

            </div>
            );
        }
        if (fieldName === 'Comments' || fieldName === 'Address' || fieldName === 'Summary') {
            //return <RichText isEditMode={false} value={item[fieldName]} />;
            const cleanValue = DomHelper.cleanRichHtml(item[fieldName]);
            if (!cleanValue) {
                return (<TooltipHost content='No data' styles={tooltipStyles}>
                    <span className={styles.detailsListValue}>
                        <Icon iconName="Remove" className={styles.mutedIcon} />
                        <span className={styles.srOnly}>No data</span>
                    </span>
                </TooltipHost>);
            }
            return <RichText isEditMode={false} value={cleanValue} />;
        }
        if (fieldName === 'Status') {
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

    public render(): React.ReactElement<IEutelsatEntityDatasheetProps> {
        const {
            showEditFormDialog,
            showDeleteDialog,
            item,
            errorMessage
        } = this.state;
        const { pnpService, userService } = this.props;
        const created = UtilHelper.formatDate(item?.Created, 'LL', 'en-us');
        const modified = UtilHelper.formatDate(item?.Modified, 'LL', 'en-us');
        if (!this.hasAnyRole(UserRole.Contributor, UserRole.Owner, UserRole.Visitor,
            UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Admin)) {
            return (
                <AccessDeniedMessage message={strings.FormAccessDeniedMessage}> </AccessDeniedMessage>
            );
        }
        if (errorMessage) {
            if (errorMessage.includes(strings.DataSheetItemDoesNotExist)) {
                return (
                    <div className={styles.noDataMessage}>
                        <Icon iconName="Remove" className={styles.mutedIcon} />
                        <span>{strings.DataSheetNoRecordFound || 'No record found'}</span>
                    </div>
                );
            }
        }
        return (
            <>
                {
                    this.state && this.state.isDatasheetReady ?

                        <div>
                            <div className={styles.datasheetHeader}>
                                <div className={styles.datasheetTypeContainer}>
                                    < h5 className={styles.datasheetType} >
                                        {strings.EutelsatEntityDataSheetTitle}
                                    </h5 >
                                </div>
                                {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.datasheetCommandsContainer}>
                                    {<div className={styles.datasheetCommand} title={strings.DataSheetEditButton}>
                                        <a href="#" onClick={this.openEditForm} >
                                            <Icon iconName="Edit" />
                                        </a>
                                    </div>
                                    }
                                    {showEditFormDialog &&
                                        this.renderEditFormDialog()
                                    }
                                    {<div className={styles.datasheetCommand} title={strings.DataSheetDeleteButton}>
                                        <a href="#" className="" onClick={this.openDeleteDialog}>
                                            <Icon iconName="Delete" />
                                        </a>
                                    </div>}
                                    {showDeleteDialog && this.renderDeleteDialog()}
                                </div>}
                            </div>
                            <div className={styles.datasheetTitleContainer}>
                                <Icon iconName='DateTime' className={styles.dateIcon} />
                                <span className={styles.dateText} >
                                    {`${strings.DataSheetLabelCreatedOn} ${created} / ${strings.DataSheetLabelUpdatedOn} ${modified}`}
                                </span>
                                <div className={styles.titleText}>
                                    <h1>{item?.Title}</h1>
                                </div>
                            </div>
                            <div className={styles.tabsContainer}>
                                <Pivot linkSize="large" >
                                    <PivotItem headerText="Info" itemIcon="Info" >
                                        <div className={styles.section}>
                                            <Icon iconName="TaskManager" className="sectionIcon" />
                                            <h3 >{strings.DataSheetSectionProperties}</h3>
                                        </div>
                                        <LayerHost
                                            id='sanctionCategoryLayer'
                                            style={{ position: 'relative', zIndex: 500000 }} // higher than page chrome
                                        />
                                        {this.renderDatasheetInfo()}
                                        <div className={styles.section}>
                                            <Icon iconName="Library" className="sectionIcon" />
                                            <h3 >{strings.DataSheetSectionDocs}</h3>
                                        </div>
                                        <LibraryView
                                            folderPath={item.DocumentSpace}
                                            pnpService={pnpService}
                                            userService={userService}
                                            pagination={{ isEnabled: true, showTopPagination: false, showBottomPagination: true }}>
                                        </LibraryView>
                                    </PivotItem>
                                </Pivot>
                            </div>
                        </div > :
                        <div className="ms-Grid-col ms-sm12 ms-md12">
                            <div className={styles.spinnerContainer}>
                                <Spinner label={'Loading...'} />
                            </div>
                        </div>
                }
            </>




        );
    }
    private openEditForm = () => {
        this.setState({ showEditFormDialog: true });
    };
    private closeEditForm = async (refreshData = false) => {
        this.setState({ showEditFormDialog: false });
        if (refreshData) {
            await this.init();
        }
    };
    private openDeleteDialog = () => {
        this.setState({ showDeleteDialog: true });
    };
    private closeDeleteDialog = () => {
        this.setState({ showDeleteDialog: false });
    };
    private deleteDataSheet = async () => {
        try {
            this.setState({ isFormProcessing: true });
            const { pnpService, itemId } = this.props;
            const list = Datasheet.config.find((f) => f.type.toLowerCase() === DatasheetType.EutelsatEntity).list;
            const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
            await pnpService.getListItems(listUrl).getById(itemId).recycle();
            const messageBanner: IMessageBanner = {
                message: strings.DeleteSuccessMessage,
                type: 'success',
                visible: true
            };
            this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
            //setTimeout(() => { window.location.href = Config.SITE.ROOT_WEB_URL; }, 2000);
            setTimeout(() => { window.location.href = Config.SITE.ROOT_WEB_URL; }, 2000);
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
    private renderEditFormDialog = () => {

        const {
            showEditFormDialog,
        } = this.state;
        const {
            pnpService,
            userService,
            itemId
        } = this.props;

        const editFormControl = (<EutelsatEntity pnpService={pnpService} userService={userService} itemId={itemId} callback={async () => { await this.closeEditForm(true); }}></EutelsatEntity>);
        const editFormTitle = strings.EutelsatEntityDataSheetEditFormTitle;

        return (<Modal
            isOpen={showEditFormDialog}
            onDismiss={() => this.closeEditForm(true)}
            containerClassName={styles.modalContainer}
            isBlocking={true}
        >
            <div>
                <div className={styles.modalHeader}>
                    {editFormTitle}
                    <ActionButton
                        iconProps={{ iconName: 'Cancel' }}
                        onClick={() => this.closeEditForm()}
                        className={styles.iconButton}
                    >{strings.DataSheetCancelButton}</ActionButton>
                </div>
                <div className={styles.modalBody}>
                    {editFormControl}
                </div>
            </div>
        </Modal>);
    };
    private renderDeleteDialog = () =>
    (<Dialog
        hidden={false}
        onDismiss={this.closeDeleteDialog}
        dialogContentProps={{
            type: DialogType.normal,
            title: strings.DataSheetDeleteModalTitle
        }}
        modalProps={{
            isBlocking: true
        }}
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
        <div>
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
            <span>{strings.DataSheetDeleteModalMessage}</span>
          )}
            <DialogFooter>
                <PrimaryButton
                    disabled={this.state.isConfirmButtonDisabled}
                    onClick={this.deleteDataSheet}
                //text={strings.DataSheetConfirmButton}
                >
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
                    onClick={this.closeDeleteDialog}
                    text={strings.DataSheetBackButton}
                />
            </DialogFooter>
        </div>
    </Dialog>);
}
