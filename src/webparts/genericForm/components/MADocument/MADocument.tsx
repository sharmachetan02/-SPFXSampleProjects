import * as React from 'react';

import { IMADocumentProps, IMADocumentState } from './MADocument.type';
import FormWizard from '../Form Wizard/FormWizard';
import styles from '../Form.module.scss';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { Form, FormOrigin, FormType, IFormConfig } from '../Form.types';
import * as strings from 'GenericFormWebPartStrings';
import { RichText } from '../../../../common/components/RichText/RichText';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import { Consts } from '../../../../common/consts/Consts';
import 'react-widgets/styles.css';
import { IItemAddResult } from '@pnp/sp/items';
import { IBaseItem } from '../../../../common/models/IBusiness';
import { Dropdown, Icon, IDropdownOption, Label, Spinner, SpinnerSize, TextField } from '@fluentui/react';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import FilePicker from '../../../../common/components/File Picker/FilePicker';
import { IAttachmentFileInfo } from '@pnp/sp/attachments';
import { UserRole } from '../../../../common/models/Enums';
import AccessDeniedMessage from '../../../../common/components/Access Denied/AccessDenied';
import DomHelper from '../../../../common/helpers/DomHelper';

export default class MADocument extends React.Component<IMADocumentProps, IMADocumentState> {
  private isEditMode: boolean;
  private config: IFormConfig;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private fileTypes: Map<string, any> = new Map();

  constructor(props: Readonly<IMADocumentProps>) {
    super(props);
    this.isEditMode = this.props.itemId !== undefined;
    this.state = {
      isFormReady: false,
      name: '',
      title: '',
      summary: '',
      comments: '',
      document: null,
      classification: '',
      classificationChoices: [],
      category: '',
      categoryChoices: [],
      fileType: null,
      fileTypeChoices: [],
      version: '',
      keywords: [],
      keywordsChoices: [],
      errors: {}
    };

    this.config = Form.config.find((f) => f.type.toLowerCase() === FormType.MADocument);
    this.fileTypes = new Map();
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
    const promises = [];
    promises.push(
      this.initFileTypes()
    );
    promises.push(
      this.initClassificationChoices()
    );
    promises.push(
      this.initKeywords()
    );
    await Promise.all(promises);
    if (this.isEditMode)
      await this.setInitialFormValues();
  }

  private initFileTypes = async () => {
    const { pnpService } = this.props;
    const selectedFields = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.MADOCUMENT.CATEGORY
    ];

    // Get the list based on FormType.FileType
    const list = Consts.LISTS.FILETYPE_URL;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const query = pnpService.getListItems(listUrl)
      .select(...selectedFields)
      .expand(Consts.FIELDS.COMMON.CONTENTTYPE);
    const items = await query.orderBy(Consts.FIELDS.COMMON.TITLE, true).top(5000)();

    // Flat array with category
    const fileTypes = items.map((item) => ({
      key: item[Consts.FIELDS.COMMON.ID] + "",
      text: item[Consts.FIELDS.COMMON.TITLE],
      category: item[Consts.FIELDS.MADOCUMENT.CATEGORY] || ""
    }));

    // Category-based map
    const groupedFileTypes = new Map();
    fileTypes.forEach((choice) => {
      const category = choice.category;
      if (!groupedFileTypes.has(category)) {
        groupedFileTypes.set(category, []);
      }
      groupedFileTypes.get(category).push({
        key: choice.key,
        text: choice.text
      });
    });

    // Unique category choices
    const categoryChoices: IDropdownOption[] = Array.from(groupedFileTypes.keys())
      .sort((a, b) => a?.localeCompare(b)) // Sort alphabetically
      .map((category) => ({
        key: category,
        text: category
      }));

    // Store all three structuress
    this.fileTypes.set("FileTypeFlat", fileTypes);
    this.fileTypes.set("FileTypeGrouped", groupedFileTypes);

    // Set initial selected category and bind its items
    const category = categoryChoices[0]?.key?.toString() || "";
    const fileTypeChoices: IDropdownOption[] = groupedFileTypes.get(category) || [];
    const fileType = fileTypeChoices[0]?.key?.toString() || "";
    this.setState({
      categoryChoices,
      category,
      fileType,
      fileTypeChoices
    });
  };

  private initClassificationChoices = async () => {
    const {
      pnpService
    } = this.props;
    let classificationChoices: IDropdownOption[] = [];
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
    const data = await pnpService.getListField(listUrl, Consts.FIELDS.MADOCUMENT.CLASSIFICATION)
      .select('Choices')();
    classificationChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
    if (!this.isEditMode) {
      const classification = classificationChoices.length > 0 ? classificationChoices[0].key as string : '';
      this.setState({
        classification, classificationChoices
      });
    } else {
      this.setState({
        classificationChoices
      });
    }
  };

  private initKeywords = async (): Promise<void> => {
    const {
      pnpService
    } = this.props;
    const keywordsChoices: IDropdownOption[] = [];
    const list = Consts.LISTS.KEYWORD_URL;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.ID,
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.MADOCUMENT.IS_ACTIVE
    ];
    const items = await pnpService.getListItems(listUrl).select(...selectedFields)
      .expand(Consts.FIELDS.COMMON.CONTENTTYPE)
      .filter(`${Consts.FIELDS.MADOCUMENT.IS_ACTIVE} eq 1`)
      .orderBy(Consts.FIELDS.COMMON.TITLE, true)();
    items.forEach((item) => {
      keywordsChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
    });
    this.setState({ keywordsChoices });
  };

  public setInitialFormValues = async () => {
    const { pnpService, itemId } = this.props;
    const list = Form.config.find((f) => f.type.toLowerCase() === FormType.MADocument).list;
    const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
    const selectedFields: string[] = [
      Consts.FIELDS.COMMON.TITLE,
      Consts.FIELDS.MADOCUMENT.SUMMARY,
      Consts.FIELDS.MADOCUMENT.CLASSIFICATION,
      Consts.FIELDS.MADOCUMENT.VERSION,
      Consts.FIELDS.COMMON.FILE_LEAFREF,
      Consts.FIELDS.COMMON.FILE_REF,
      `${Consts.FIELDS.MADOCUMENT.KEYWORDS}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.MADOCUMENT.KEYWORDS}/${Consts.FIELDS.COMMON.TITLE}`,
      `${Consts.FIELDS.MADOCUMENT.FILE_TYPE}/${Consts.FIELDS.COMMON.ID}`,
      `${Consts.FIELDS.MADOCUMENT.FILE_TYPE}/${Consts.FIELDS.COMMON.TITLE}`
    ];
    const expand = [Consts.FIELDS.MADOCUMENT.FILE_TYPE, Consts.FIELDS.MADOCUMENT.KEYWORDS];
    const item = await pnpService.getListItems(listUrl).getById(itemId).select(...selectedFields).expand(...expand)();
    const fileTypeId = item[Consts.FIELDS.MADOCUMENT.FILE_TYPE]?.[Consts.FIELDS.COMMON.ID] + "";
    const flatChoices = this.fileTypes.get("FileTypeFlat") as {
      key: string;
      text: string;
      category: string;
    }[];
    const matchedFileType = flatChoices.find((choice) => choice.key === fileTypeId);
    let fileTypeChoices = [];
    const fileType = item[Consts.FIELDS.MADOCUMENT.FILE_TYPE]?.[Consts.FIELDS.COMMON.ID] + "";
    let category = "";
    if (matchedFileType) {
      category = matchedFileType.category;
      fileTypeChoices = this.fileTypes.get("FileTypeGrouped").get(category) || [];
    }

    const fileName = item[Consts.FIELDS.COMMON.FILE_LEAFREF];
    const filepath = item[Consts.FIELDS.COMMON.FILE_REF].replace(item[Consts.FIELDS.COMMON.FILE_LEAFREF], "");
    const extension = fileName?.split('.').pop()?.toLowerCase() || "";
    const name = fileName?.split('.')?.[0] || "";
    // 🔽 Fetch the actual file content from SharePoint
    const fileBuffer = await pnpService.getFileByUrl(item[Consts.FIELDS.COMMON.FILE_REF]).getBuffer();
    const initialValues = {
      name,
      title: item[Consts.FIELDS.COMMON.TITLE],
      summary: item[Consts.FIELDS.MADOCUMENT.SUMMARY],
      classification: item[Consts.FIELDS.MADOCUMENT.CLASSIFICATION],
      fileType,
      category,
      fileTypeChoices,
      version: item[Consts.FIELDS.MADOCUMENT.VERSION],
      keywords: item[Consts.FIELDS.MADOCUMENT.KEYWORDS]?.map((keyItem) => ({
        key: keyItem[Consts.FIELDS.COMMON.ID],
        text: keyItem[Consts.FIELDS.COMMON.TITLE]
      })) || [],
      extension,
      filepath,
      document: [{
        name: fileName,
        content: fileBuffer
      }]
    };
    this.setState({
      ...initialValues
    });
  };

  private checkFileExists = async (fileUrl: string): Promise<boolean> => {
    const { pnpService } = this.props;
    try {
      const folder = await pnpService.getFileByUrl(fileUrl).select('Exists')();
      return folder?.Exists || false;
    } catch (error) {
      return false;
    }
  };

  private validateForm = async (): Promise<boolean> => {

    const {
      name,
      title,
      extension,
      document
    } = this.state;

    const {
      filepath,
    } = this.props;
    const errors: { [key: string]: string } = {};
    const parentFolderUrl = `${filepath}`;
    const displayName = `${name}.${extension}`;
    const fileUrlDocument = `${parentFolderUrl}/${displayName}`;
    // Validate name
    // if (!name || name.trim() === "") {
    //   errors.name = strings.FormMADocumentErrorNameRequired;
    // }

    // Validate Title
    if (!title || title.trim() === "") {
      errors.title = strings.FormMADocumentErrorTitleRequired;
    }

    // Validate document presence in create mode
    if (!document && this.isEditMode === false) {
      errors.file = strings.FormMADocumentErrorFileRequired;
    }

    // if (document && this.isEditMode && extension) {
    //   const actualExtension = document?.[0]?.name?.split('.').pop()?.toLowerCase();
    //   if (actualExtension && actualExtension !== extension.toLowerCase()) {
    //     errors.file = strings.FormMADocumentErrorFileExtenssionChange;
    //   }
    // }

    // Only check file existence if no file error yet and in create mode
    if (!errors.file && this.isEditMode === false) {
      const exists = await this.checkFileExists(fileUrlDocument);
      if (exists) {
        errors.file = strings.FormMADocumentErrorFileAlredyExist;
      }
    }
    this.setState({ errors });
    return Object.keys(errors).length === 0;
  };

  public render(): React.ReactElement<IMADocumentProps> {
    const {
      name,
      extension,
      title,
      summary,
      errors,
      classification,
      classificationChoices,
      category,
      categoryChoices,
      fileType,
      fileTypeChoices,
      version,
      keywords,
      keywordsChoices,
      isFormReady,
      document
    } = this.state;

    return (
      <div className="ms-Grid-row">
        <div className="ms-Grid-col ms-sm12 ms-md12">
          {!this.hasAnyRole(UserRole.Contributor, UserRole.Owner, UserRole.FinanceContributor, UserRole.Admin) && (
            <AccessDeniedMessage message={strings.FormAccessDeniedMessage}> </AccessDeniedMessage>
          )}
          {/* Form  */}
          {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <FormWizard
            allowDuplicateItems={true}
            getDuplicateItems={this.getDuplicateItems}
            getSummary={this.getFormSummary}
            validateForm={this.validateForm}
            isEdit={this.isEditMode}
            isFormReady={isFormReady}
            callback={this.callback}
            processCreation={this.processCreation}
            processEdit={this.processEdit}
            formConfig={this.config}
            formOrigin={FormOrigin.Datasheet}
          >
            <>
              {!isFormReady &&
                <div className={styles.spinnerContainer}>
                  <Spinner size={SpinnerSize.large} label={'Loading...'} />
                </div>
              }
              {isFormReady &&
                <form>
                  {/* File */}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label required>{strings.FormMADocumentFileLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormMADocumentTooltipFile} ></Tooltip>
                    </div>
                    <div className={styles.filePickerContainer}>
                      <FilePicker
                        multiple={false}
                        onChange={this.onDropDocument}
                        files={document}
                        disabled={false}
                      />
                    </div>
                    {errors.file &&
                      <div className={styles.errorContainer}>
                        <Icon iconName="Error" className={styles.errorIcon} />
                        <span className={styles.errorMessage}>
                          {errors.file}
                        </span>
                      </div>
                    }
                  </div>
                  {/* Name */}
                  {name &&
                    <div className={styles.field}>
                      <div className={styles.fieldLabel}>
                        <div className={styles.fieldLabelContainer}>
                          <Label>{strings.FormMADocumentNameLabel}</Label>
                        </div>
                        <Tooltip content={strings.FormMADocumentTooltipName} ></Tooltip>
                      </div>
                      <div className={styles.fileContainer}>
                        <div className={styles.fileNameContainer}>
                          <TextField
                            value={name}
                            disabled={true}
                            onChange={this.onChangeName}
                          />
                        </div>
                        <div className={styles.extensionContainer}>
                          <Label>{extension ? "." + extension : ".extension"}</Label>
                        </div>
                      </div>
                      {errors.name &&
                        <div className={styles.errorContainer}>
                          <Icon iconName="Error" className={styles.errorIcon} />
                          <span className={styles.errorMessage}>
                            {errors.name}
                          </span>
                        </div>
                      }
                    </div>
                  }

                  {/* Title */}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label required>{strings.FormMADocumentTitleLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormMADocumentTooltipTitle} ></Tooltip>
                    </div>
                    <TextField
                      value={title}
                      onChange={this.onChangeTitle}
                    />
                    {errors.title &&
                      <div className={styles.errorContainer}>
                        <Icon iconName="Error" className={styles.errorIcon} />
                        <span className={styles.errorMessage}>
                          {errors.title}
                        </span>
                      </div>
                    }
                  </div>
                  {/* Summary */}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label>{strings.FormMADocumentSummaryLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormMADocumentTooltipSummary} ></Tooltip>
                    </div>
                    <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
                  </div>
                  {/* Classification */}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label required>{strings.FormMADocumentClassificationLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormMADocumentTooltipClassification} ></Tooltip>
                    </div>
                    <Dropdown
                      selectedKey={classification}
                      onChange={this.onChangeClassification}
                      options={classificationChoices} />
                  </div>
                  {/* Category */}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label required>{strings.FormMADocumentDocumentTypeLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormMADocumentTooltipDocumentType} ></Tooltip>
                    </div>
                    <Dropdown
                      selectedKey={category}
                      onChange={this.onChangeCategory}
                      options={categoryChoices} />
                  </div>
                  {/* Category Type */}
                  <div className={styles.field} >
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label required>{strings.FormMADocumentTypeLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormMADocumentTooltipType} ></Tooltip>
                    </div>
                    <Dropdown
                      selectedKey={fileType}
                      onChange={this.onChangeDocumentType}
                      options={fileTypeChoices} />
                    {errors.fileType &&
                      <div className={styles.errorContainer}>
                        <Icon iconName="Error" className={styles.errorIcon} />
                        <span className={styles.errorMessage}>
                          {errors.fileType}
                        </span>
                      </div>
                    }
                  </div>
                  {/* Version */}
                  <div className={styles.field}>
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label>{strings.FormMADocumentVersionLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormMADocumentTooltipVersion} ></Tooltip>
                    </div>
                    <TextField
                      value={version}
                      onChange={this.onChangeVersion}
                    />
                  </div>
                  {/* keywords */}
                  <div className={styles.field} >
                    <div className={styles.fieldLabel}>
                      <div className={styles.fieldLabelContainer}>
                        <Label>{strings.FormMADocumentKeywordsLabel}</Label>
                      </div>
                      <Tooltip content={strings.FormMADocumentTooltipKeywords} ></Tooltip>
                    </div>
                    <MultiselectWrapper
                      data={keywordsChoices}
                      value={keywords}
                      textField="text"
                      filter="contains"
                      key="key"
                      allowCreate
                      onCreate={this.onCreate}
                      onChange={this.onChangeKeywords}
                    />
                  </div>
                </form>
              }
            </>
          </FormWizard>}
        </div>
      </div >
    );
  }

  private onCreate = async (name: string) => {
    const {
      pnpService
    } = this.props;
    if (name !== null && name !== undefined) {
      const itemData = {};
      itemData[Consts.FIELDS.COMMON.TITLE] = name;
      itemData[Consts.FIELDS.MADOCUMENT.IS_ACTIVE] = true;
      const list = Consts.LISTS.KEYWORD_URL;
      const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
      await pnpService.getListItems(listUrl).add(itemData);
      await this.initKeywords();
    }
  };

  private onDropDocument = (file: IAttachmentFileInfo[]) => {

    const document: IAttachmentFileInfo[] = file;
    if (document) {
      if (!this.isEditMode) {
        const name = document?.[0]?.name?.split(".")?.[0];
        const extension = document?.[0]?.name?.split('.')?.pop()?.toLowerCase();
        this.setState({ document, name, extension });
      } else {
        const {
          extension,
          errors
        } = this.state;
        const actualExtension = document?.[0]?.name?.split('.').pop()?.toLowerCase();
        if (actualExtension && actualExtension !== extension.toLowerCase()) {
          errors.file = strings.FormMADocumentErrorFileExtenssionChange;
          this.setState({ errors });
        } else {
          errors.file = null;
          this.setState({ errors, document });
        }
      }
    }
  };

  private onChangeName = (event, name: string) => {
    this.setState({ name });
  };

  private onChangeTitle = (event, title: string) => {
    this.setState({ title });
  };

  private onChangeVersion = (event, version: string) => {
    this.setState({ version });
  };

  private onChangeClassification = (event, option: IDropdownOption) => {
    this.setState({ classification: option.text.toString() });
  };

  private onChangeCategory = (event, option: IDropdownOption) => {
    this.setState({ category: option.text.toString() });
    const selectedCategory = option?.key as string;
    // Get the category map from fileTypes
    const categoryMap = this.fileTypes.get("FileTypeGrouped");
    if (categoryMap.has(selectedCategory)) {
      const fileTypeChoices = categoryMap.get(selectedCategory) || [];
      this.setState({
        fileTypeChoices,
        fileType: fileTypeChoices?.[0]?.key
      });
    } else {
      this.setState({
        fileTypeChoices: [],
        fileType: null
      });
    }
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private onChangeDocumentType = (event: any, fileType: IDropdownOption) => {
    this.setState({ fileType: fileType?.key?.toString() });
  };

  private onChangeSummary = (summary: string) => {
    this.setState({ summary });
    return summary;
  };

  private onChangeKeywords = (keywords: IDropdownOption[]): void => {
    this.setState({ keywords });
  };

  private getDuplicateItems = async (): Promise<IBaseItem[]> => {
    const duplicateItems = [];
    return duplicateItems;
  };

  private getFormSummary = () => {
    const {
      title,
      name,
      summary,
      classification,
      fileTypeChoices,
      fileType,
      version,
      keywords,
      category,
      extension
    } = this.state;
    const keywordsValues = keywords?.map((c) => c.text).join(", ");
    const displayName = `${name}.${extension}`;
    const summaryTextValue = DomHelper.cleanRichHtml(summary);
    return (
      <div className={styles.summary}>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMADocumentNameLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {displayName}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMADocumentTitleLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {title}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormPaymentLabelSummary}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {summaryTextValue.length > 0 ? <RichText isEditMode={false} value={summaryTextValue} /> : ''}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMADocumentClassificationLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {classification}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMADocumentDocumentTypeLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {category}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMADocumentTypeLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {fileTypeChoices.find((cat) => cat.key === fileType).text}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMADocumentVersionLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {version}
          </div>
        </div>
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMADocumentKeywordsLabel}</Label>
            </div>
          </div>
          <div className={styles.fieldValue}>
            {keywordsValues}
          </div>
        </div>
      </div>
    );
  };

  private callback = async () => {
    const {
      callback,
    } = this.props;
    callback(false);
  };

  private processCreation = async (): Promise<IItemAddResult> => {
    const {
      name,
      extension,
      title,
      summary,
      classification,
      fileType,
      version,
      document,
      keywords
    } = this.state;
    const {
      pnpService,
      filepath
    } = this.props;
    const itemData = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = title;
    itemData[Consts.FIELDS.MADOCUMENT.SUMMARY] = summary;
    itemData[Consts.FIELDS.MADOCUMENT.CLASSIFICATION] = classification;
    itemData[Consts.FIELDS.MADOCUMENT.FILE_TYPE_ID] = fileType;
    itemData[Consts.FIELDS.MADOCUMENT.VERSION] = version;
    itemData[Consts.FIELDS.MADOCUMENT.KEYWORDS_ID] = keywords.map((item) => parseInt(item?.key?.toString()));
    const parentFolderUrl = `${filepath}`;
    const fileUrlDocument = `${parentFolderUrl}/${name}.${extension}`;
    // Upload the file using path
    const fileAddResult = await pnpService.getFolder(parentFolderUrl)
      .files.addUsingPath(fileUrlDocument, document?.[0].content, { Overwrite: false });
    const item = await fileAddResult.file.getItem();
    const itemResult = await item.update(itemData);
    return itemResult;
  };

  private processEdit = async () => {
    const {
      name,
      title,
      summary,
      classification,
      fileType,
      version,
      document,
      keywords,
      extension,
      filepath
    } = this.state;

    const {
      pnpService,
    } = this.props;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const itemData: Record<string, any> = {};
    itemData[Consts.FIELDS.COMMON.TITLE] = title;
    itemData[Consts.FIELDS.MADOCUMENT.SUMMARY] = summary;
    itemData[Consts.FIELDS.MADOCUMENT.CLASSIFICATION] = classification;
    itemData[Consts.FIELDS.MADOCUMENT.FILE_TYPE_ID] = fileType;
    itemData[Consts.FIELDS.MADOCUMENT.VERSION] = version;
    itemData[Consts.FIELDS.MADOCUMENT.KEYWORDS_ID] = keywords.map((item) => parseInt(item?.key?.toString()));
    const parentFolderUrl = `${filepath}`;
    const fileUrlDocument = `${parentFolderUrl}/${name}.${extension}`;
    // Upload the file using path
    const fileAddResult = await pnpService.getFolder(parentFolderUrl)
      .files.addUsingPath(fileUrlDocument, document?.[0]?.content, { Overwrite: true });
    const item = await fileAddResult.file.getItem();
    await item.update(itemData);
  };

}
