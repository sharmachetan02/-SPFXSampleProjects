import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { Fabric, Icon, Label, TextField } from 'office-ui-fabric-react';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { IMAFolderProps, IMAFolderState } from './MAFolder.types';
import { Form, FormOrigin, FormType, IFormConfig } from '../Form.types';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import strings from 'GenericFormWebPartStrings';
import { IItemAddResult } from '@pnp/sp/items';
import { UtilHelper } from '../../../../common/helpers/Util';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { checkFolderExists, createFolder } from '../Form.utility';
import { Consts } from '../../../../common/consts/Consts';
import { PNPService } from '../../../../common/services/PNPService';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
export default class MAFolder extends React.Component<IMAFolderProps, IMAFolderState> {
    private isEditMode: boolean;
    private config: IFormConfig;
    constructor(props: IMAFolderProps) {
        super(props);
        this.isEditMode = this.props.itemId !== undefined;
        this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.MAFolder);
        this.state = {
            isFormReady: false,
            name: '',
            errors: {}
        };
    }

    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };
    public async componentDidMount() {
        if (this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin)) {
            await this.init();
            this.setState({ isFormReady: true });
        }
    }

    private async init() {
        if (this.isEditMode)
            await this.setInitialFormValues();
    }

    public setInitialFormValues = async () => {
        const { pnpService, itemId } = this.props;
        const list = Consts.LISTS.DOCUMENT_SPACE_URL;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const folderItem = await pnpService.getListItems(listUrl).getById(itemId).select(Consts.FIELDS.COMMON.FILE_LEAFREF)();
        const folderName = folderItem[Consts.FIELDS.COMMON.FILE_LEAFREF];
        const initialValues = {
            exitingFolderName: folderName,
            name: folderName,
        };
        this.setState({
            ...initialValues
        });
    };

    public getFolderPropertiesByItemId = async (
        pnpService: PNPService,
        itemId: number
    ): Promise<{ id: number; name: string; path: string } | null> => {
        const list = Consts.LISTS.DOCUMENT_SPACE_URL;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.FILE_LEAFREF,
            Consts.FIELDS.COMMON.FILE_DIR_REF,
        ];
        const folderItem = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields)();
        if (!folderItem) return null;
        return {
            id: folderItem.Id,
            name: folderItem.FileLeafRef,
            path: folderItem.FileDirRef
        };
    };

    private processEdit = async () => {
        const {
            name
        } = this.state;

        const {
            pnpService,
            itemId,
        } = this.props;
        const itemData = {};
        const list = this.config.list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const folderItem = await pnpService.getListItems(listUrl).getById(itemId).select(Consts.FIELDS.COMMON.FILE_LEAFREF)();
        if (!folderItem.FileLeafRef || folderItem.FileLeafRef !== name) {
            itemData[Consts.FIELDS.COMMON.FILE_LEAFREF] = UtilHelper.sanitizeSharePointFolderName(name);
        }
        await pnpService.getListItems(listUrl).getById(itemId).update(itemData);
    };

    private processCreation = async (): Promise<IItemAddResult> => {
        const { name } = this.state;
        const {
            pnpService, currentFolderPath
        } = this.props;
        const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
        const folderUrl = await createFolder(pnpService, currentFolderPath, sanitizedFolderName);
        return { data: { FolderUrl: folderUrl } } as IItemAddResult;
    };

    private callback = async (redirectUrl) => {
        const {
            callback,
        } = this.props;
        callback(redirectUrl);
    };

    private getDuplicateItems = async (): Promise<IBaseItem[]> => {
        const duplicateItems = [];
        return duplicateItems;
    };

    private validateForm = async (): Promise<boolean> => {
        const errors: { [key: string]: string } = {};
        const { name, exitingFolderName } = this.state;
        const { pnpService, currentFolderPath } = this.props;

        if (!name || name.trim() === '') {
            errors.name = strings.FormMAFolderNameValidationError;
        }
        if (this.isEditMode && name === exitingFolderName) {
            // Name hasn't changed, no need to check for existence
            this.setState({ errors });
            return Object.keys(errors).length === 0;
        }
        // Only check for folder existence if there is no error so far
        if (Object.keys(errors).length === 0) {
            const sanitizedFolderName = UtilHelper.sanitizeSharePointFolderName(name);
            const parentFolderUrl = `${currentFolderPath ? `/${currentFolderPath}` : ''}`;
            if (this.isEditMode) {
                // Call getFolderPropertiesByItemId for edit mode
                if (this.props.itemId) {
                    const folderProps = await this.getFolderPropertiesByItemId(pnpService, this.props.itemId);
                    if (folderProps && folderProps.name !== sanitizedFolderName) {
                        const folderExists = await checkFolderExists(pnpService, folderProps.path, sanitizedFolderName);
                        if (folderExists) {
                            errors.name = strings.FormMAFolderExistsError;
                        }
                    }
                }
            } else {
                const folderExists = await checkFolderExists(pnpService, parentFolderUrl, sanitizedFolderName);
                if (folderExists) {
                    errors.name = strings.FormMAFolderExistsError;
                }
            }
        }

        this.setState({ errors });
        return Object.keys(errors).length === 0;
    };

    private getFormSummary = () => {
        const {
            name
        } = this.state;
        return (
            <div className={styles.summary}>
                <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                            <Label>{strings.FormMAFolderLabelName}</Label>
                        </div>
                    </div>
                    <div className={styles.fieldValue}>
                        {name}
                    </div>
                </div>
            </div>
        );
    };
    public render(): React.ReactElement<IMAFolderProps> {
        const { isFormReady, errors, name } = this.state;
        return (
            <div className="ms-Grid-row" >
                <div className="ms-Grid-col ms-sm12 ms-md12">
                    {!this.hasAnyRole(UserRole.Contributor, UserRole.Owner, UserRole.FinanceContributor, UserRole.Admin) && (
                        <AccessDeniedMessage message={strings.FormAccessDeniedMessage}> </AccessDeniedMessage>
                    )}
                    {/* Form  */}
                    {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <FormWizard
                        allowDuplicateItems={false}
                        getDuplicateItems={this.getDuplicateItems}
                        //getSummary={this.getFormSummary}
                        getSummary={() => this.getFormSummary()} // <-- wrap in arrow function
                        validateForm={this.validateForm}
                        //validateFormConcurrent={this.validateFormConcurrent}
                        isEdit={this.isEditMode}
                        isFormReady={isFormReady}
                        callback={this.callback}
                        processCreation={this.processCreation}
                        processEdit={this.processEdit}
                        formConfig={this.config}
                        formOrigin={FormOrigin.Datasheet}>
                        <>
                            {!isFormReady &&
                                <div className={styles.spinnerContainer}>
                                    <Spinner size={SpinnerSize.large} label={'Loading...'} />
                                </div>
                            }
                            {isFormReady &&
                                < Fabric >
                                    <form>
                                        {/* Name */}
                                        <div className={styles.field}>
                                            <div className={styles.fieldLabel}>
                                                <div className={styles.fieldLabelContainer}>
                                                    <Label required>{strings.FormMAFolderLabelName}</Label>
                                                </div>
                                                <Tooltip content={strings.FormMAFolderTooltipName} ></Tooltip>
                                            </div>
                                            <TextField
                                                value={name}
                                                onChange={this.onChangeName}
                                            />
                                            {errors.name &&
                                                <div className={styles.errorContainer}>
                                                    <Icon iconName="Error" className={styles.errorIcon} />
                                                    <span className={styles.errorMessage}>
                                                        {errors.name}
                                                    </span>
                                                </div>
                                            }
                                        </div>
                                    </form>
                                </Fabric>
                            }
                        </>
                    </FormWizard>}
                </div>
            </div >
        );

    }
    private onChangeName = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({ name: newValue || '' });
    };
}
