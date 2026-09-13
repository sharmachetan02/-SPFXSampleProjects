import * as React from 'react';
import { ILibraryViewState, ILibraryViewProps } from './LibraryView.type';
import styles from '../View.module.scss';
import { IViewConfig, View, ViewType } from '../View.types';
import 'react-widgets/styles.css';
import {
    ActionButton,
    Breadcrumb,
    CommandButton,
    ConstrainMode, DefaultButton, DetailsList, DetailsListLayoutMode, Dialog,
    DialogFooter, DialogType, IBreadcrumbItem, IColumn, Icon, IconButton, IContextualMenuItem, IContextualMenuProps, IDropdownOption, ITooltipHostStyles, Label,
    LayerHost, Modal, Panel, PanelType, PrimaryButton, SearchBox, SelectionMode, Spinner, SpinnerSize, Stack, TooltipHost
} from '@fluentui/react';
import { IFilter, LogicalOperator } from '../../../../common/models/IFilter';
import * as strings from 'GenericViewPartStrings';
import { Consts } from '../../../../common/consts/Consts';
import { IFileType, IKeyword, ILibrary } from '../../../../common/models/IBusiness';
import { UtilHelper } from '../../../../common/helpers/Util';
import { SortDirection } from '@pnp/sp/search';
import { renderArrayPills } from '../View.utility';
import AppliedFilters from '../AppliedFilters/AppliedFilters';
import Pager from '../Pager/Pager';
import { FileHelper } from '../../../../common/helpers/FileHelper';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';
import MAFolder from '../../../genericForm/components/MA Folder/MAFolder';
import MADocument from '../../../genericForm/components/MADocument/MADocument';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import { CallOutButton } from '../../../../common/components/CallOutButton/CallOutButton';
import { RichText } from '../../../../common/components/RichText/RichText';
import DownloadFile from '../../../../common/components/Download File/DownlaodFile';
import { UserRole } from '../../../../common/models/Enums';
import { saveAs } from 'file-saver';
import { MessageBanner } from '../../../../common/components/Message Banner/MessageBanner';
import { IMessageBanner } from '../../../genericForm/components/Form.types';
import DomHelper from '../../../../common/helpers/DomHelper';

export default class LibraryView extends React.Component<
    ILibraryViewProps,
    ILibraryViewState
> {
    private config: IViewConfig;
    private allViewItems: ILibrary[] = [];
    private fileTypeItems: IFileType[] = [];
    private rootFolder: string;

    // File extension categories
    private readonly FILE_EXTENSIONS = {
        microsoftOffice: ['doc', 'docx', 'docm', 'xls', 'xlsx', 'xlsm', 'xlsb', 'ppt', 'pptx', 'pptm'],
        pdfsDocuments: ['pdf', 'txt'],
        images: ['jpg', 'jpeg', 'png', 'gif', 'dng', 'psd', 'svg', 'bmp'],
        audioVideo: ['mp4', 'mp3', 'mov', 'wmv', 'wav', 'mkv', 'webm', '3gp'],
        threeDModels: ['3mf', 'glb', 'obj', 'stl'],
        other: ['zip', 'eml', 'json', 'xml', 'html', 'msg', 'epub']
    };
    private classificationChoices: IDropdownOption[] = [];
    private fileTypeChoices: IDropdownOption[] = [];
    private categoryChoices: IDropdownOption[] = [];
    private keywordsChoices: IDropdownOption[] = [];
    constructor(props: Readonly<ILibraryViewProps>) {
        super(props);
        if (this.props.folderPath) {
            this.rootFolder = decodeURIComponent(new URL(this.props.folderPath).pathname?.toLowerCase() || "");
        } else {
            this.rootFolder = "";
        }
        this.state = {
            isViewReady: false,
            columns: [],
            filters: [],
            sort: null,
            searchKeyword: '',
            viewItems: [],
            pageIndex: 0,
            pageSize: 25,
            currentFolderPath: this.rootFolder,
            breadcrumbs: null,
            isDeleteDialogOpen: false,
            isDeleteDisabled: false,
            documentToDelete: undefined,
            isPreviewDocSource: false,
            previewDocSource: '',
            previewDocTitle: '',
            isFileVisible: false,
            isFolderVisible: false,
            editFromId: undefined,
            //fileTypeCatListItems: []
            isFilterPanelOpen: false,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'error',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        };
        this.config = View.config.find((f) => f.type.toLowerCase() === ViewType.Library);
    }
    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };
    public async componentDidMount() {
        if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.FinanceVisitor, UserRole.Owner,
            UserRole.Contributor, UserRole.Visitor, UserRole.Admin)) {
            if (this.props.folderPath)
                await this.init();
            this.setState({ isViewReady: true });
        }
    }

    public async componentDidUpdate(prev: ILibraryViewProps) {
        if (
            prev.columns !== this.props.columns ||
            prev.filters !== this.props.filters ||
            prev.sort !== this.props.sort
        ) {
            await this.init();
        }
    }

    private setSearchText = (searchKeyword?: string) => {
        const { filters } = this.state;
        const searchFilter = filters.find((f) => f.name === 'Search');
        searchFilter.selectedValues = searchKeyword ? [searchKeyword] : [];
        this.setState({ filters, searchKeyword });
    };

    private initColumns(): IColumn[] {
        const allColumns = [{
            key: 'FileType',
            name: strings.LibraryViewFileTypeColumnLabel,
            fieldName: 'FileType',
            minWidth: 15,
            maxWidth: 15,
            iconName: 'Page',
            isIconOnly: true,
            isResizable: true
        },
        {
            key: 'Name',
            name: strings.LibraryViewNameColumnLabel,
            fieldName: 'Name',
            minWidth: 100,
            maxWidth: 260,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Actions',
            name: strings.LibraryViewActionsColumnLabel,
            fieldName: 'Actions',
            minWidth: 50,
            maxWidth: 50,
            isResizable: true
        },
        {
            key: 'Title',
            name: strings.LibraryViewTitleColumnLabel,
            fieldName: 'Title',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Summary',
            name: strings.LibraryViewSummaryColumnLabel,
            fieldName: 'Summary',
            minWidth: 70,
            maxWidth: 70,
            isResizable: true,
        },
        {
            key: 'Classification',
            name: strings.LibraryViewClassificationColumnLabel,
            fieldName: 'Classification',
            minWidth: 100,
            maxWidth: 100,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Category',
            name: strings.LibraryViewCategoryColumnLabel,
            fieldName: 'Category',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Type',
            name: strings.LibraryViewTypeColumnLabel,
            fieldName: 'Type',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Version',
            name: strings.LibraryViewVersionColumnLabel,
            fieldName: 'Version',
            minWidth: 50,
            maxWidth: 70,
            isResizable: true,
            onColumnClick: this.onColumnClick
        },
        {
            key: 'Keywords',
            name: strings.LibraryViewKeywordsColumnLabel,
            fieldName: 'Keywords',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        },
        ];
        return this.props.columns ?
            allColumns.filter((c) => this.isColumnEnabled(c.key)) :
            allColumns;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private buildTreeFlat = (items: Record<string, any>[], parentPath: string): Record<string, any>[] => {
        // Lowercase copy for comparison; keep the original parentPath untouched
        const parentNorm = (parentPath ?? '').toLowerCase();

        // Group items by normalized FileDirRef; attach normalized helper properties
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const grouped: Record<string, Record<string, any>[]> = {};
        for (const itm of items) {
            const dirNorm = String(itm.FileDirRef ?? '').toLowerCase();
            const fileNorm = String(itm.FileRef ?? '').toLowerCase();

            (grouped[dirNorm] ??= []).push({
                ...itm,              // original fields preserved
                _dirNorm: dirNorm,   // helper: normalized dir
                _fileNorm: fileNorm, // helper: normalized file
            });
        }
        // Recursive builder that uses normalized keys for comparison,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        // but outputs nodes with original values + HasChildren
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const build = (normKey: string): Record<string, any>[] => {
            const nodes = grouped[normKey] || [];
            return nodes.reduce((acc, node) => {
                const children = build(node._fileNorm); // traverse by normalized file path

                const out = {
                    ...node,                  // keep originals
                    children,                 // original children list (may be empty)
                    HasChildren: children.length > 0, // flag based on normalized traversal
                };

                // Return a flat list: this node + all its children
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                return acc.concat([out, ...children]);
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
            }, [] as Record<string, any>[]);
        };
        // Group items by normalized FileDirRef; attach normalized helper properties
        return build(parentNorm);
    };

    public initViewItems = async (): Promise<ILibrary[]> => {
        const path = this.state.currentFolderPath;
        const selectedFields: string[] = [
            `${Consts.FIELDS.MADOCUMENT.UNIQUEID}`,
            `${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MADOCUMENT.CLASSIFICATION}`,
            `${Consts.FIELDS.MADOCUMENT.SUMMARY}`,
            `${Consts.FIELDS.MADOCUMENT.VERSION}`,
            `${Consts.FIELDS.COMMON.FILE_LEAFREF}`,
            `${Consts.FIELDS.COMMON.FILE_REF}`,
            `${Consts.FIELDS.COMMON.ISFOLDER}`,
            `${Consts.FIELDS.COMMON.FILE_TYPE}`,
            `${Consts.FIELDS.COMMON.FILE_DIR_REF}`,
            `${Consts.FIELDS.MADOCUMENT.KEYWORDS}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MADOCUMENT.KEYWORDS}/${Consts.FIELDS.COMMON.TITLE}`,
            `${Consts.FIELDS.MADOCUMENT.FILE_TYPE}/${Consts.FIELDS.COMMON.ID}`,
            `${Consts.FIELDS.MADOCUMENT.FILE_TYPE}/${Consts.FIELDS.COMMON.TITLE}`,
        ];
        const safePath = path.replace(/'/g, "''");
        const listUrl = UrlHelper.getListWebRelativeUrl(this.props.pnpService.getUrl(), this.config.list);
        const items = await this.props.pnpService.getListByUrl(listUrl)
            .items
            .select(...selectedFields).expand(Consts.FIELDS.MADOCUMENT.KEYWORDS, Consts.FIELDS.MADOCUMENT.FILE_TYPE)
            .filter(`
            startswith(${Consts.FIELDS.COMMON.FILE_REF}, '${safePath}') 
            or startswith(${Consts.FIELDS.COMMON.FILE_DIR_REF}, '${safePath}')
          `)
            .top(5000)();
        items.map((item) => {
            item[Consts.FIELDS.MADOCUMENT.CLASSIFICATION] ?
                this.classificationChoices.push({ key: item[Consts.FIELDS.MADOCUMENT.CLASSIFICATION], text: item[Consts.FIELDS.MADOCUMENT.CLASSIFICATION] }) : null;
            item[Consts.FIELDS.MADOCUMENT.FILE_TYPE] ?
                this.fileTypeChoices.push({
                    key: item[Consts.FIELDS.MADOCUMENT.FILE_TYPE][Consts.FIELDS.COMMON.ID],
                    text: item[Consts.FIELDS.MADOCUMENT.FILE_TYPE][Consts.FIELDS.COMMON.TITLE]
                }) : null;

            this.categoryChoices.push({
                key: this.fileTypeItems?.filter((ft) => ft.Id === item?.[Consts.FIELDS.MADOCUMENT.FILE_TYPE]?.[Consts.FIELDS.COMMON.ID])?.[0]?.Category || null,
                text: this.fileTypeItems?.filter((ft) => ft.Id === item?.[Consts.FIELDS.MADOCUMENT.FILE_TYPE]?.[Consts.FIELDS.COMMON.ID])?.[0]?.Category || null
            });
            this.keywordsChoices.push(...(item?.[Consts.FIELDS.MADOCUMENT.KEYWORDS]?.map((k) => ({
                key: k[Consts.FIELDS.COMMON.ID],
                text: k[Consts.FIELDS.COMMON.TITLE]
            })) || [])
            );
        });
        const treeItemsWithChildren = this.buildTreeFlat(items, safePath);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const viewItems: ILibrary[] = treeItemsWithChildren?.map((item: Record<string, any>) => ({

            Id: item?.[Consts.FIELDS.COMMON.ID],
            Name: item?.[Consts.FIELDS.COMMON.FILE_LEAFREF],
            Title: item?.[Consts.FIELDS.COMMON.TITLE],
            UniqueId: item[Consts.FIELDS.MADOCUMENT.UNIQUEID],
            Edit: "",
            Download: "",
            Preview: "",
            Delete: "",
            Open: "",
            Classification: item?.[Consts.FIELDS.MADOCUMENT.CLASSIFICATION],
            IsFolder: (item?.[Consts.FIELDS.COMMON.ISFOLDER] ?? 0) === 1,
            ServerRelativeUrl: item?.[Consts.FIELDS.COMMON.FILE_REF],
            FileType: item?.[Consts.FIELDS.COMMON.FILE_TYPE] ? item?.[Consts.FIELDS.COMMON.FILE_TYPE] : "",
            Children: item?.children,
            HasChildren: !!item?.children?.length,
            Summary: item?.[Consts.FIELDS.MADOCUMENT.SUMMARY],
            Version: item?.[Consts.FIELDS.MADOCUMENT.VERSION],
            Keywords: item?.[Consts.FIELDS.MADOCUMENT.KEYWORDS]?.map((k) => ({
                Id: k[Consts.FIELDS.COMMON.ID],
                Title: k[Consts.FIELDS.COMMON.TITLE]
            })) || [],
            Type: item?.[Consts.FIELDS.MADOCUMENT.FILE_TYPE] ? {
                Id: item?.[Consts.FIELDS.MADOCUMENT.FILE_TYPE][Consts.FIELDS.COMMON.ID],
                Title: item?.[Consts.FIELDS.MADOCUMENT.FILE_TYPE][Consts.FIELDS.COMMON.TITLE],
            } : undefined,
            Actions: "",
            Category: this.fileTypeItems?.filter((ft) => ft.Id === item?.[Consts.FIELDS.MADOCUMENT.FILE_TYPE]?.[Consts.FIELDS.COMMON.ID])?.[0]?.Category || undefined,

        } as ILibrary));
        return viewItems;
    };

    private initFilters = (): IFilter[] => {
        let allFilters: IFilter[] = [];
        const searchFilter: IFilter = {
            name: 'Search',
            displayName: strings.AuthorityViewSearchFilterLabel,
            placeHolder: strings.AuthortyViewSearchFilterPlaceHolder,
            fields: ['Name', 'Title', 'Summary', 'Classification', 'Category', 'Type', 'Version', 'Keywords'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.Or,
        };
        const classificationFilter: IFilter = {
            name: 'Classification',
            displayName: strings.LibraryViewClassificationFilterLabel,
            placeHolder: strings.LibraryViewClassificationFilterPlaceHolder,
            fields: ['Classification'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const keywordsFilter: IFilter = {
            name: 'Keywords',
            displayName: strings.LibraryViewKeywordsFilterLabel,
            placeHolder: strings.LibraryViewKeywordsFilterPlaceHolder,
            fields: ['Keywords'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const categoryFilter: IFilter = {
            name: 'Category',
            displayName: strings.LibraryViewCategoryFilterLabel,
            placeHolder: strings.LibraryViewCategoryFilterPlaceHolder,
            fields: ['Category'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        const fileTypeFilter: IFilter = {
            name: 'Type',
            displayName: strings.LibraryViewTypeFilterLabel,
            placeHolder: strings.LibraryViewTypeFilterPlaceHolder,
            fields: ['Type'],
            selectedValues: [],
            values: [],
            operatorBetweenValues: LogicalOperator.And,
        };
        allFilters = [searchFilter, classificationFilter, keywordsFilter, categoryFilter, fileTypeFilter];
        return this.props.filters ? allFilters.filter((f) => this.isFilterEnabled(f.name)) : allFilters;
    };

    private getFileExtension = (fileName: string): string => (fileName ? fileName.split('.').pop()?.toLowerCase() || '' : '');

    private openDocument = async (previewDocument: ILibrary): Promise<void> => {
        const { pnpService } = this.props;
        const fileExtension = this.getFileExtension(previewDocument.Name);
        const encodedId = encodeURIComponent(previewDocument.ServerRelativeUrl);
        const serverUrl = pnpService.getUrl();
        // Handle other files
        if (this.FILE_EXTENSIONS.other.includes(fileExtension) || this.FILE_EXTENSIONS.threeDModels.includes(fileExtension) || this.FILE_EXTENSIONS.images.includes(fileExtension)) {

            const parentPath = previewDocument.ServerRelativeUrl.substring(0, previewDocument.ServerRelativeUrl.lastIndexOf('/'));
            const encodedParent = encodeURIComponent(parentPath);
            const url = `${serverUrl}/documents_space/Forms/AllItems.aspx?id=${encodedId}&parent=${encodedParent}`;
            window.open(url, '_blank');
            return;
        }

        // Handle PDF and images
        if (this.FILE_EXTENSIONS.pdfsDocuments.includes(fileExtension)) {

            // Extract only the file path part (remove /sites/MADB-Dev prefix)
            const sitePathMatch = previewDocument.ServerRelativeUrl.match(/^\/sites\/[^/]+(.*)$/);
            const siteRelativePath = sitePathMatch ? sitePathMatch[1] : previewDocument.ServerRelativeUrl;
            const url = `${serverUrl}${siteRelativePath}`;
            window.open(url, '_blank');
            return;
        }

        // For MS Office documents
        if (this.FILE_EXTENSIONS.microsoftOffice.includes(fileExtension)) {
            const encodedSourceDoc = encodeURIComponent(`{${previewDocument.UniqueId}}`);
            const encodedFileName = encodeURIComponent(previewDocument.Name);
            const url: string =
                previewDocument.ServerRelativeUrl ? `${serverUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodedSourceDoc}&file=${encodedFileName}&action=default&mobileredirect=true` : '';
            window.open(url, '_blank');
            return;
        }

        // For audio/ video files
        if (this.FILE_EXTENSIONS.audioVideo.includes(fileExtension)) {
            const url: string =
                previewDocument.ServerRelativeUrl ? `${serverUrl}/_layouts/15/stream.aspx?id=${encodedId}` : '';
            window.open(url, '_blank');
            return;
        }

        const blob = await this.props.pnpService.getFileByUrl(previewDocument.ServerRelativeUrl).getBlob();
        saveAs(blob, `${previewDocument.Name}`);
    };

    private onClickedBreadCrumb = (currentFolderPath: string): void => {
        this.setState({ currentFolderPath }, this.getBreadcrumbItems);
    };

    private newOpen = (type: string): void => {
        if (type === "Folder") {
            this.setState({ isFolderVisible: true, editFromId: undefined });
        } else {
            this.setState({ isFileVisible: true, editFromId: undefined });
        }
    };

    private getBreadcrumbItems = () => {
        const root = this.rootFolder.replace(/\/+$/, "");
        const current = this.state?.currentFolderPath || root;
        let relativePath = current;
        if (current.toLowerCase().startsWith(root.toLowerCase())) {
            relativePath = current.substring(root.length + 1);
        }
        const segments = relativePath ? relativePath.split("/") : [];
        const items: IBreadcrumbItem[] = [
            {
                text: "Root",
                key: 'root',
                onClick:
                    (ev) => {
                        ev.preventDefault();
                        this.onClickedBreadCrumb(root);
                    }
            }
        ];
        segments.forEach((segment, index) => {
            const path = [root, ...segments.slice(0, index + 1)].join("/");
            items.push({
                text: segment,
                key: `crumb-${index}`,
                onClick: (ev) => {
                    ev.preventDefault();
                    this.onClickedBreadCrumb(path);
                }
            });
        });
        this.setState({ breadcrumbs: items });
    };

    private init = async () => {
        const columns = this.initColumns();
        let filters = this.initFilters();
        this.fileTypeItems = await this.initFileTypeItems();
        const items = await this.initViewItems();
        this.allViewItems = items;
        let classification: IDropdownOption[] = [];
        let types: IDropdownOption[] = [];
        let categories: IDropdownOption[] = [];
        let keywords: IDropdownOption[] = [];
        const promises: Promise<void>[] = [];
        if (this.isFilterEnabled('Classification')) {
            promises.push(this.initClassificationChoices().then((opts) => { classification = opts as IDropdownOption[]; }));
        }
        if (this.isFilterEnabled('Type')) {
            promises.push(this.initFileTypeChoices().then((opts) => { types = opts as IDropdownOption[]; }));
        }
        if (this.isFilterEnabled('Category')) {
            promises.push(this.initCategoryChoices().then((opts) => { categories = opts as IDropdownOption[]; }));
        }
        if (this.isFilterEnabled('Keywords')) {
            promises.push(this.initKeywordsChoices().then((opts) => { keywords = opts as IDropdownOption[]; }));
        }
        await Promise.all(promises);
        filters = filters.map((f) => {
            if (f.name === 'Classification') return { ...f, values: classification };
            if (f.name === 'Type') return { ...f, values: types };
            if (f.name === 'Category') return { ...f, values: categories };
            if (f.name === 'Keywords') return { ...f, values: keywords };
            return f;
        });
        const sort = this.props.sort ?? { Property: 'Name', Direction: SortDirection.Ascending };
        const sortedCols = columns.map((c) => ({
            ...c,
            isSorted: c.key === sort.Property || c.fieldName === sort.Property,
            isSortedDescending: sort.Direction === SortDirection.Descending
        }));
        this.setState({
            columns: sortedCols,
            filters,
            sort,
            pageIndex: 0,
            viewItems: items,
        }, this.getBreadcrumbItems);
    };

    private initClassificationChoices = async () => {
        /*const {
            pnpService
        } = this.props;
        let classificationChoices: IDropdownOption[] = [];
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), this.config.list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.MADOCUMENT.CLASSIFICATION)
            .select('Choices')();
        classificationChoices = data.Choices.map((choice) => ({ key: choice, text: choice }));
        return classificationChoices;*/
        const classificationChoices: IDropdownOption[] = [];
        this.classificationChoices.forEach((choice) => {
            if (!classificationChoices.find((c) => c.key === choice.key)) {
                classificationChoices.push(choice);
            }
        });
        classificationChoices.sort((a, b) => a.text?.localeCompare(b.text));
        return classificationChoices;
    };

    private initFileTypeItems = async () => {
        const {
            pnpService
        } = this.props;
        const fileTypeItems: IFileType[] = [];
        const list = View.config.find((f) => f.type.toLowerCase() === ViewType.FileType).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE,
            Consts.FIELDS.FILETYPE.CATEGORY
        ];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        items.forEach((item) => {
            fileTypeItems.push({ Id: item[Consts.FIELDS.COMMON.ID], Title: item[Consts.FIELDS.COMMON.TITLE], Category: item[Consts.FIELDS.FILETYPE.CATEGORY] });
        });
        return fileTypeItems;
    };

    private initFileTypeChoices = async () => {
        /*const {
            pnpService
        } = this.props;
        const fileTypeChoices: IDropdownOption[] = [];
        const list = View.config.find((f) => f.type.toLowerCase() === ViewType.FileType).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE
        ];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        items.forEach((item) => {
            fileTypeChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
        });
        return fileTypeChoices;*/
        const fileTypeChoices: IDropdownOption[] = [];
        this.fileTypeChoices.forEach((choice) => {
            if (!fileTypeChoices.find((c) => c.key === choice.key)) {
                fileTypeChoices.push(choice);
            }
        });
        fileTypeChoices.sort((a, b) => a.text?.localeCompare(b.text));
        return fileTypeChoices;
    };

    private initKeywordsChoices = async () => {
        /*const {
            pnpService
        } = this.props;
        const keywordsChoices: IDropdownOption[] = [];
        const list = View.config.find((f) => f.type.toLowerCase() === ViewType.Keyword).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const selectedFields: string[] = [
            Consts.FIELDS.COMMON.ID,
            Consts.FIELDS.COMMON.TITLE
        ];
        const items = await pnpService.getListItems(listUrl).select(...selectedFields).orderBy(Consts.FIELDS.COMMON.TITLE, true)();
        items.forEach((item) => {
            keywordsChoices.push({ key: item[Consts.FIELDS.COMMON.ID], text: item[Consts.FIELDS.COMMON.TITLE] });
        });
        return keywordsChoices;*/
        const keywordsChoices: IDropdownOption[] = [];
        this.keywordsChoices.forEach((choice) => {
            if (!keywordsChoices.find((c) => c.key === choice.key)) {
                keywordsChoices.push(choice);
            }
        });
        keywordsChoices.sort((a, b) => a.text?.localeCompare(b.text));
        return keywordsChoices;
    };

    private initCategoryChoices = async () => {
        /*const {
            pnpService
        } = this.props;
        const categoryChoices: IDropdownOption[] = [];
        const list = View.config.find((f) => f.type.toLowerCase() === ViewType.FileType).list;
        const listUrl = UrlHelper.getListWebRelativeUrl(pnpService.getUrl(), list);
        const data = await pnpService.getListField(listUrl, Consts.FIELDS.FILETYPE.CATEGORY).select('Choices')();
        categoryChoices.push(...data.Choices.map((choice) => ({ key: choice, text: choice })));
        return categoryChoices;*/
        const categoryChoices: IDropdownOption[] = [];
        this.categoryChoices.forEach((choice) => {
            choice.key !== null &&
                !categoryChoices.find((c) => c.key === choice.key) &&
                categoryChoices.push(choice);
        });
        categoryChoices.sort((a, b) => a.text?.localeCompare(b.text));
        return categoryChoices;
    };

    private isFilterEnabled = (name: string) =>
        !this.props.filters || this.props.filters.includes(name);

    private isColumnEnabled = (keyOrField: string) =>
        !this.props.columns || this.props.columns.includes(keyOrField);

    public applyFilters = (viewItems: ILibrary[]): ILibrary[] => {
        const { filters } = this.state;
        let filtered = viewItems;
        filters.forEach((filter) => {
            filtered = filtered.filter((item) =>
                filter.fields.some((field) => {
                    const selected = filter.selectedValues || [];
                    if (selected.length === 0) return true;

                    const assertions = selected.map((f) => {
                        if (filter.name === 'Search') {
                            const needle = UtilHelper.toNormalForm(f?.toString()?.toLowerCase());
                            if (!needle) return true;

                            if (field === 'Name' || field === 'Title' || field === 'Summary' || field === 'Classification' || field === 'Category' ||
                                field === 'Version') {
                                const hay = UtilHelper.toNormalForm((item[field])?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            } else if (field === 'Type') {
                                const hay = UtilHelper.toNormalForm((item[field] as IFileType)?.Title?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            } else if (field === 'Keywords') {
                                const keywords = (item[field] as IKeyword[]) || [];
                                const hay = UtilHelper.toNormalForm(keywords.map((k) => k.Title).join(' ').toLowerCase());
                                return hay?.includes(needle) === true;
                            } else {
                                const hay = UtilHelper.toNormalForm(item[field]?.toString()?.toLowerCase());
                                return hay?.includes(needle) === true;
                            }
                        }
                        if (field === 'Type') {
                            return String((item[field] as IFileType)?.Id) === String((f as IDropdownOption).key);
                        }
                        if (field === 'Keywords') {
                            const keywords = (item[field] as IKeyword[]) || [];
                            return keywords.some((k) => String(k.Id) === String((f as IDropdownOption).key));
                        }
                        if (filter.name === 'Classification' || filter.name === 'Category') {
                            return String(item[field]?.toString()) === String((f as IDropdownOption).key);
                        }
                        return true;
                    });
                    return filter.operatorBetweenValues === LogicalOperator.And ?
                        assertions.every(Boolean) :
                        assertions.some(Boolean);
                })
            );
        });
        return filtered;
    };

    private clearFilters = (): void => {
        let { filters } = this.state;
        filters = filters.map((filter) => {
            filter.selectedValues = [];
            return filter;
        });
        this.setState({ filters, searchKeyword: '', pageIndex: 0 });
    };

    private sort = (items) => {
        const { sort } = this.state;
        if (!sort) return items;

        const key = sort.Property as keyof ILibrary;
        const desc = sort.Direction === SortDirection.Descending;
        return items.slice().sort((a, b) => {
            let av = a[key];
            let bv = b[key];
            if (key === 'Type') {
                av = av?.Title ?? '';
                bv = bv?.Title ?? '';
            }
            const aStr = (av || '').toLowerCase();
            const bStr = (bv || '').toLowerCase();
            if (aStr === bStr) return 0;
            const cmp = aStr > bStr ? 1 : -1;
            return desc ? -cmp : cmp;
        });
    };

    private renderItemColumn = (item, index: number, column: IColumn) => {
        const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };
        if (!column) return <span />;

        const fieldName = column.fieldName as keyof ILibrary | undefined;
        const value = fieldName ? item[fieldName] : undefined;

        if (value === null || value === undefined || (Array.isArray(value) && value.length === 0)) {
            return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
        }
        if (fieldName === 'Name') {
            if (item['IsFolder']) {
                return (
                    <a
                        className={styles.folderColumn}
                        onClick={(e) => this.onChildFolderClick(e, item['ServerRelativeUrl'])} >
                        <span>{String(value)}</span>
                    </a>
                );
            } else {
                return (<a className={styles.fileColumn} data-interception="on" onClick={() => this.openDocument(item)} rel="noreferrer" >{String(value)}</a>);
            }
        }

        if (fieldName === 'Actions') {
            const getMenuItems = this.getItemLevelMenuProps(item);
            if (!getMenuItems || getMenuItems.items.length === 0) {
                return <span className={styles.detailsListValue}>-</span>;
            }
            return (
                <div>
                    <IconButton iconProps={{ iconName: 'More' }} menuProps={getMenuItems} className={styles.actionsButton} />
                    {/*<CommandButton iconProps={{ iconName: 'More' }} className={styles.newItemButton} text="Items" menuProps={this.getMenuProps()} />*/}
                </div>
            );
        }
        if (fieldName === 'Title') {
            return <span className={styles.detailsListValue}>{String(value)}</span>;
        }
        if (fieldName === 'Summary') {
            /*return (
                <CallOutButton icon={'View'} name={'Preview'} >
                    <RichText isEditMode={false} value={item[fieldName]} />
                </CallOutButton>
            );*/
            const cleanSummary = DomHelper.cleanRichHtml(item[fieldName]);
            if (!cleanSummary) return (<TooltipHost content='No data' styles={tooltipStyles}>
                <span className={styles.detailsListValue}>
                    <Icon iconName="Remove" className={styles.mutedIcon} />
                    <span className={styles.srOnly}>No data</span>
                </span>
            </TooltipHost>);
            return (
                <CallOutButton icon={'View'} name={'Preview'} >
                    <RichText isEditMode={false} value={cleanSummary} />
                </CallOutButton>
            );
        }
        if (fieldName === 'Classification') {
            return (
                <span className={styles.detailsListValue}>
                    {renderArrayPills([String(value)])}
                </span>
            );
        }
        if (fieldName === 'Category') {
            return (
                <span className={styles.detailsListValue}>
                    {renderArrayPills([String(value)])}
                </span>
            );
        }
        if (fieldName === 'Type') {
            const type = item['Type'] as { Id: number; Title: string };
            if (type) {
                return (
                    <span className={styles.detailsListValue}>
                        {renderArrayPills([type.Title])}
                    </span>
                );
            }
        }
        if (fieldName === 'Keywords') {
            const keywords = item['Keywords'] as { Id: number; Title: string }[];
            if (keywords && keywords.length > 0) {
                return renderArrayPills(keywords.map((k) => k.Title));
            }
        }
        if (fieldName === 'Download') {
            if (!item['IsFolder']) {
                return (
                    <DownloadFile document={item} pnpService={this.props.pnpService} />
                );
            }
        }
        {
            /*
        if (fieldName === 'Edit') {
            return (
                <IconButton
                    iconProps={{ iconName: 'Edit' }}
                    onClick={() => this.editOpen(item)}
                />
            );
        }
        if (fieldName === 'Download') {
            if (!item['IsFolder']) {
                return (
                    <DownloadFile document={item} pnpService={this.props.pnpService} />
                );
            }
        }
        if (fieldName === 'Preview') {
            if (!item['IsFolder']) {
                return (
                    <IconButton
                        iconProps={{ iconName: 'View' }}
                        onClick={() => this.showPreviewDoc(item)}
                    />
                );
            }
        }
        if (fieldName === 'Delete') {
            return (!item["HasChildren"] &&
                <IconButton
                    iconProps={{ iconName: 'Delete' }}
                    onClick={() => this.openDialogDelete(item)}
                />
            );
        }
            */
        }
        if (fieldName === 'FileType') {
            return (
                <>
                    {item['IsFolder'] ? (
                        <img
                            src={FileHelper.getUrlIconName('folder')}
                            className={styles.fileIconImg}
                            alt="Folder"
                        />
                    ) : (
                        <img
                            src={FileHelper.getUrlIconName(value)}
                            className={styles.fileIconImg}
                            alt={`${item.extension} file icon`}
                        />
                    )}
                </>
            );
        }
        {

            /*if (fieldName === 'Open') {
                if (!item['IsFolder']) {
                    if (item?.supportedExtensions?.includes(item?.extension)) {
                        return (
                            <IconButton
                                iconProps={{ iconName: 'OpenInNewWindow' }}
                                onClick={() => this.openDocument(item)}
                            />
                        );
                    } else {
                        return (<IconButton
                            iconProps={{ iconName: 'OpenInNewWindow' }}
                            onClick={() => this.openDocument(item)}
                        />);
                    }
                }
            }*/

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

    private onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        let { columns } = this.state;
        let isSortedDescending = column.isSortedDescending;
        if (column.isSorted) {
            isSortedDescending = !isSortedDescending;
        }
        columns = columns.map((col) => {
            col.isSorted = col.key === column.key;
            if (col.isSorted) {
                col.isSortedDescending = isSortedDescending;
            }
            return col;
        });
        const sort = {
            Property: column.key,
            Direction: isSortedDescending ? SortDirection.Descending : SortDirection.Ascending
        };

        this.setState({
            sort,
            columns,
            pageIndex: 0
        });
    };

    public refreshFilters = (filters: IFilter[]) => {
        const { searchKeyword } = this.state;
        const searchFilter = filters.find((f) => f.name === 'Search');
        this.setState({ filters, searchKeyword: searchFilter.selectedValues.length > 0 ? searchKeyword : '' });
    };

    public onSearch = (filter: IFilter, newValue): void => {
        let { filters } = this.state;
        filter.selectedValues = [].concat(newValue);
        filters = filters.map(((f) => ((f.name === filter.name) ? filter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    public onChange = (filter: IFilter, searchKeyword): void => {
        let { filters } = this.state;
        if (!searchKeyword) {
            filter.selectedValues = [];
            filters = filters.map(((f) => ((f.name === filter.name) ? filter : f)));
            this.setState({ filters, searchKeyword });
        } else {
            this.setState({ searchKeyword });
        }

    };

    // Go to a specific page (0-based)
    private setPage = (page: number) => {
        this.setState({ pageIndex: Math.max(0, page) });
    };

    // Change page size and reset to first page
    private setPageSize = (size: number) => {
        const pageSize = Math.max(1, size || 1);
        this.setState({ pageSize, pageIndex: 0 });
    };

    private onChildFolderClick = (e, ServerRelativeUrl: string) => {
        e.preventDefault();
        e.stopPropagation();
        this.setState({ currentFolderPath: ServerRelativeUrl }, this.getBreadcrumbItems);
    };

    private isSupportedExtension = (fileName: string): boolean => {
        const fileExtension = this.getFileExtension(fileName);
        const allSupportedExtensions = [
            ...this.FILE_EXTENSIONS.microsoftOffice,
            ...this.FILE_EXTENSIONS.pdfsDocuments,
            ...this.FILE_EXTENSIONS.images,
            ...this.FILE_EXTENSIONS.audioVideo,
            ...this.FILE_EXTENSIONS.threeDModels,
            ...this.FILE_EXTENSIONS.other
        ];
        return allSupportedExtensions.includes(fileExtension);
    };

    private getItemLevelMenuProps = (item: ILibrary): IContextualMenuProps => {
        const items: IContextualMenuItem[] = [];
        if (this.hasAnyRole(UserRole.FinanceContributor, UserRole.Owner,
            UserRole.Contributor, UserRole.Admin)) {
            items.push(
                {
                    key: 'Edit',
                    name: strings.LibraryViewEditColumnLabel,
                    iconProps: { iconName: 'Edit' },
                    onClick: () => this.editOpen(item),
                    className: styles.libraryMenuItem

                });
            if (!item.HasChildren) {
                items.push({
                    key: 'Delete',
                    name: strings.LibraryViewDeleteColumnLabel,
                    iconProps: { iconName: 'Delete' },
                    onClick: () => this.openDialogDelete(item),
                    className: styles.libraryMenuItem

                }
                );
            }
        }
        // Show Preview for non-folder files with supported extensions
        if (!item.IsFolder && this.isSupportedExtension(item.Name)) {
            items.push({
                key: 'Preview',
                name: strings.LibraryViewPreviewColumnLabel,
                iconProps: { iconName: 'View' },
                onClick: () => this.showPreviewDoc(item),
                className: styles.libraryMenuItem

            });
        }
        if (!item.IsFolder) {
            items.push({
                key: 'Download',
                name: strings.LibraryViewDownloadColumnLabel,
                iconProps: { iconName: 'Download' },
                onRender: () => (
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    <DownloadFile document={item as any} pnpService={this.props.pnpService} />
                ),
                className: styles.libraryMenuItem

            });
        }
        return { items };
    };

    private editOpen = (document: ILibrary): void => {
        if (document.IsFolder) {
            this.setState({ isFolderVisible: true, editFromId: document.Id });
        } else {
            this.setState({ isFileVisible: true, editFromId: document.Id });
        }
    };

    private openDialogDelete = (doc: ILibrary) => {
        this.setState({
            isDeleteDialogOpen: true, isDeleteDisabled: false, documentToDelete: doc,
            messageBanner: {
                message: strings.DeleteSuccessMessage,
                type: 'success',
                visible: false
            },
            isFormProcessing: false,
            isConfirmButtonDisabled: false
        });
    };

    // private getFileExtension = (fileName: string): string => (fileName ? fileName.split('.').pop()?.toLowerCase() || '' : '');

    private showPreviewDoc = (previewDocument: ILibrary): void => {
        const { pnpService } = this.props;
        //  const fileExtension = this.getFileExtension(previewDocument.Name);

        // Handle ZIP, PDF, and image files - display using SharePoint folder view
        //  if (fileExtension === 'zip' || fileExtension === 'png' || fileExtension === 'pdf') {
        const serverUrl = pnpService.getUrl();
        const encodedId = encodeURIComponent(previewDocument.ServerRelativeUrl);
        const parentPath = previewDocument.ServerRelativeUrl.substring(0, previewDocument.ServerRelativeUrl.lastIndexOf('/'));
        const encodedParent = encodeURIComponent(parentPath);
        const previewUrl = `${serverUrl}/documents_space/Forms/AllItems.aspx?id=${encodedId}&parent=${encodedParent}`;
        if (previewUrl === this.state.previewDocSource) {
            this.setState({ previewDocSource: '', previewDocTitle: '', isPreviewDocSource: false });
        } else {
            this.setState({
                isPreviewDocSource: true,
                previewDocSource: previewUrl,
                previewDocTitle: ''// previewDocument.Title
            });
        }
    };


    private getMenuProps = (): IContextualMenuProps => {
        const menuProps: IContextualMenuProps = {
            items: [
                {
                    key: 'File',
                    name: strings.LibraryViewFileLabel,
                    iconProps: { iconName: 'FileSymlink' },
                    onClick: () => (this.newOpen("File")),
                },
                {
                    key: 'Folder',
                    name: strings.LibraryViewFolderLabel,
                    iconProps: { iconName: 'Page' },
                    onClick: () => (this.newOpen("Folder")),
                },
            ]
        };
        return menuProps;
    };

    // normalize + filter by current folder (strict equality) in one place
    private filterByCurrentFolderNormalized = (): ILibrary[] => {
        const norm = (s?: string) => decodeURIComponent(s || '').toLowerCase();
        const current = norm(this.state.currentFolderPath);
        const filtered = (this.allViewItems || []).filter((i) => {
            const normalizedItemUrl = norm(i.ServerRelativeUrl);
            return normalizedItemUrl === current + `/${i.Name?.toLocaleLowerCase()}`;
        });
        return filtered;
    };

    // Open dialog with document name
    /*private openDialogDelete = (doc: ILibrary) => {
        this.setState({ isDeleteDialogOpen: true, isDeleteDisabled: false, documentToDelete: doc });
    };*/

    // Close dialog
    private closeDialogDelete = () => {
        this.setState({ isDeleteDialogOpen: false, isDeleteDisabled: false, documentToDelete: undefined });
    };

    // Confirm delete
    private deleteDocument = async () => {
        try {
            this.setState({ isFormProcessing: true });
            const { pnpService } = this.props;
            const { documentToDelete, currentFolderPath } = this.state;

            if (!documentToDelete?.ServerRelativeUrl) return;
            this.setState({ isDeleteDisabled: true });
            // Recycle the file
            await pnpService
                .getFileByUrl(documentToDelete.ServerRelativeUrl)
                .recycle();
            const messageBanner: IMessageBanner = {
                message: strings.DeleteSuccessMessage,
                type: 'success',
                visible: true
            };
            this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
            // Close dialog and refresh items without losing breadcrumbs
            this.setState({
                currentFolderPath: this.rootFolder, // temporarily set to props folderPath
            });
            setTimeout(async () => {
                await this.init(); // Refresh items
                // Restore previous folder path after refresh
                this.setState({ currentFolderPath, isDeleteDialogOpen: false, documentToDelete: undefined }, this.getBreadcrumbItems);
            }, 2000);

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

    private Callback = async () => {
        const { currentFolderPath } = this.state;
        // Close dialog and refresh items without losing breadcrumbs
        this.setState({
            currentFolderPath: this.rootFolder, // temporarily set to props folderPath
        });
        await this.init(); // Refresh items
        // Restore previous folder path after refresh
        this.setState({ currentFolderPath, isFileVisible: false, isFolderVisible: false, editFromId: undefined }, this.getBreadcrumbItems);
    };

    private onModalClose = async () => {
        this.setState({ isFolderVisible: false, isFileVisible: false, editFromId: undefined });
    };
    private onChangeClassification = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const classificationFilter = filters.find((f) => f.name === 'Classification');
        classificationFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === classificationFilter.name) ? classificationFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };
    private onChangeType = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const typeFilter = filters.find((f) => f.name === 'Type');
        typeFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === typeFilter.name) ? typeFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };
    private onChangeKeywords = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const keywordsFilter = filters.find((f) => f.name === 'Keywords');
        keywordsFilter.selectedValues = options;
        filters = filters.map(((f) => ((f.name === keywordsFilter.name) ? keywordsFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };
    private onChangeCategory = (options: IDropdownOption[]) => {
        let { filters } = this.state;
        const categoryFilter = filters.find((f) => f.name === 'Category');
        categoryFilter.selectedValues = options.length > 1 ? [options[options.length - 1]] : options;
        filters = filters.map(((f) => ((f.name === categoryFilter.name) ? categoryFilter : f)));
        this.setState({ filters, pageIndex: 0 });
    };

    private openFilterPanel = () => this.setState({ isFilterPanelOpen: true });
    private dismissFilterPanel = () => this.setState({ isFilterPanelOpen: false });

    public render(): React.ReactElement<ILibraryViewProps> {
        const menuProps = this.getMenuProps();
        const {
            filters, columns, searchKeyword, pageIndex, pageSize,
            breadcrumbs, isDeleteDialogOpen, documentToDelete,
            previewDocSource, previewDocTitle, isPreviewDocSource, isFolderVisible, isFileVisible,
            editFromId, currentFolderPath, isFilterPanelOpen

        } = this.state;
        const searchFilter = filters.find((f) => f.name === 'Search');
        const classificationFilter = filters.find((f) => f.name === 'Classification');
        const typeFilter = filters.find((f) => f.name === 'Type');
        const categoryFilter = filters.find((f) => f.name === 'Category');
        const keywordsFilter = filters.find((f) => f.name === 'Keywords');
        let viewItems = this.filterByCurrentFolderNormalized();
        viewItems = this.applyFilters(viewItems);
        viewItems = this.sort(viewItems);
        const pageCount = Math.max(1, Math.ceil(viewItems.length / pageSize));
        const safeIndex = Math.min(pageIndex, pageCount - 1);
        const start = safeIndex * pageSize;
        const end = start + pageSize;
        const pagedItems = viewItems.slice(start, end);
        const itemsToDisplay = this.props.pagination?.isEnabled ? pagedItems : viewItems;
        let formTitle = '';
        if (isFolderVisible && !editFromId)
            formTitle = strings.LibraryViewCreateFolderLabel;
        else if (isFolderVisible && editFromId)
            formTitle = strings.LibraryViewEditFolderLabel;
        else if (isFileVisible && !editFromId)
            formTitle = strings.LibraryViewCreateDocumentLabel;
        else if (isFileVisible && editFromId)
            formTitle = strings.LibraryViewEditDocumentLabel;
        return (
            <>
                {/* Search */}
                {!isPreviewDocSource &&
                    <> {this.state && this.state.isViewReady ?
                        <div className={styles.listView}>
                            {
                                !this.props.folderPath && (
                                    <div className={styles.noDataMessage}>
                                        <Icon iconName="Remove" className={styles.mutedIcon} />
                                        <span>No record found</span>
                                        <Icon iconName="Remove" className={styles.mutedIcon} />
                                    </div>
                                )
                            }
                            {this.props.folderPath && (
                                <>
                                    <>
                                        {this.allViewItems && this.allViewItems.length > 0 &&
                                            <div className={styles.applyFiltersBtn}>

                                                <PrimaryButton
                                                    text="Filters"
                                                    iconProps={{ iconName: 'Filter' }}
                                                    className={`${styles.clearButton} "form-button"`}
                                                    styles={{
                                                        icon: { order: 0, marginRight: 4 }, // moves icon after text
                                                    }}
                                                    onClick={() => this.openFilterPanel()}
                                                />

                                            </div>
                                        }
                                        <Panel                               // Always open
                                            isOpen={isFilterPanelOpen}
                                            // Let users interact with page behind it
                                            isBlocking={false}
                                            // Dismiss is disabled (always on)
                                            onDismiss={() => this.dismissFilterPanel()}
                                            type={PanelType.smallFixedFar} // pick: smallFixedFar / smallFixedNear / medium / large
                                            headerText={'Filters'}
                                            closeButtonAriaLabel="Close"
                                            // Optional: remove light dismiss since we never close
                                            isLightDismiss={false}

                                            styles={{
                                                scrollableContent: {
                                                    height: '100%'
                                                },
                                                main: {
                                                    selectors: {
                                                        '@media (min-width: 480px)': {
                                                            width: '300px !important', // Override media query
                                                        },
                                                    },

                                                }
                                            }}
                                        >

                                            <div className={styles.containerPaneldiv}>
                                                <div className={styles.clearFiltersBtn}>
                                                    <PrimaryButton
                                                        text="Clear Filters"
                                                        title="Clear"
                                                        disabled={false}
                                                        onClick={this.clearFilters}
                                                        className={`${styles.clearButton} "form-button"`}
                                                        iconProps={{ iconName: 'ClearFilter' }}
                                                        styles={{
                                                            icon: { order: 0, marginRight: 4 }, // moves icon after text
                                                        }}
                                                    /></div>
                                                {/* Search */}
                                                {
                                                    this.isFilterEnabled('Search') &&
                                                    <>
                                                        <div>
                                                            <Label className={styles.filterLabel}>{searchFilter.displayName}</Label>
                                                        </div>
                                                        <div className={styles.panelChildDiv}>
                                                            <SearchBox
                                                                placeholder={searchFilter.placeHolder}
                                                                autoComplete='off'
                                                                value={searchKeyword}
                                                                onSearch={(newValue) => this.onSearch(searchFilter, newValue)}
                                                                onChange={(_, newValue) => this.onChange(searchFilter, newValue)}
                                                                onClear={() => this.setSearchText('')} />
                                                        </div>
                                                    </>
                                                }
                                                {/* Classification */}
                                                {
                                                    this.isFilterEnabled('Classification') &&
                                                    <>
                                                        <div>
                                                            <Label className={styles.filterLabel}>{classificationFilter.displayName}</Label>
                                                        </div>

                                                        <div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper
                                                                placeholder={classificationFilter.placeHolder}
                                                                data={classificationFilter?.values as IDropdownOption[]}
                                                                dataKey={(item: IDropdownOption) => item.key}
                                                                textField={(item: IDropdownOption) => item.text}
                                                                value={classificationFilter?.selectedValues}
                                                                filter="contains"
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null}
                                                                onChange={this.onChangeClassification} />
                                                        </div>
                                                    </>
                                                }
                                                {/* Category */}
                                                {
                                                    this.isFilterEnabled('Category') &&
                                                    <>
                                                        <div>
                                                            <Label className={styles.filterLabel}>{categoryFilter.displayName}</Label>
                                                        </div>

                                                        <div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper
                                                                placeholder={categoryFilter.placeHolder}
                                                                data={categoryFilter?.values as IDropdownOption[]}
                                                                dataKey={(item: IDropdownOption) => item.key}
                                                                textField={(item: IDropdownOption) => item.text}
                                                                value={categoryFilter?.selectedValues}
                                                                filter="contains"
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null}
                                                                onChange={this.onChangeCategory} />
                                                        </div>
                                                    </>
                                                }
                                                {/* File Type */}
                                                {
                                                    this.isFilterEnabled('Type') &&
                                                    <>

                                                        <div>
                                                            <Label className={styles.filterLabel}>{typeFilter.displayName}</Label>
                                                        </div>

                                                        <div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper
                                                                placeholder={typeFilter.placeHolder}
                                                                data={typeFilter?.values as IDropdownOption[]}
                                                                dataKey={(item: IDropdownOption) => item.key}
                                                                textField={(item: IDropdownOption) => item.text}
                                                                value={typeFilter?.selectedValues}
                                                                filter="contains"
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null}
                                                                onChange={this.onChangeType} />
                                                        </div>
                                                    </>
                                                }
                                                {/* Keywords */}
                                                {
                                                    this.isFilterEnabled('FileType') &&
                                                    <>
                                                        <div >
                                                            <Label className={styles.filterLabel}>{keywordsFilter.displayName}</Label>
                                                        </div>

                                                        <div className={styles.panelChildDiv}>
                                                            <MultiselectWrapper
                                                                placeholder={keywordsFilter.placeHolder}
                                                                data={keywordsFilter?.values as IDropdownOption[]}
                                                                dataKey={(item: IDropdownOption) => item.key}
                                                                textField={(item: IDropdownOption) => item.text}
                                                                value={keywordsFilter?.selectedValues}
                                                                filter="contains"
                                                                showSelectedItemsInList
                                                                showPlaceholderWithValues
                                                                renderTagValue={() => null}
                                                                onChange={this.onChangeKeywords} />
                                                        </div>
                                                    </>
                                                }

                                            </div>
                                        </Panel>
                                        <div className={styles.tableContainer}>
                                            <AppliedFilters
                                                filters={filters}
                                                callback={this.refreshFilters}
                                            />
                                        </div>
                                    </>
                                    <div className="ms-Grid" dir="ltr">
                                        <div className="ms-Grid-row">
                                            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                                                <Breadcrumb items={breadcrumbs} onReduceData={() => undefined} />
                                            </div>
                                        </div>
                                    </div>
                                    {this.hasAnyRole(UserRole.FinanceContributor, UserRole.Owner,
                                        UserRole.Contributor, UserRole.Admin) && <div>
                                            <CommandButton iconProps={{ iconName: 'Add' }} className={styles.newItemButton} text="Items" menuProps={menuProps} />
                                        </div>}
                                    <div className={styles.listViewContainer}>
                                        {viewItems.length > 0 ? (<div className={styles.detailsList}>
                                            {this.props.pagination?.isEnabled && this.props.pagination.showTopPagination && (
                                                <Pager
                                                    total={viewItems.length}
                                                    pageIndex={pageIndex}
                                                    pageSize={pageSize}
                                                    onPageChange={this.setPage}
                                                    onPageSizeChange={this.setPageSize}
                                                />)}
                                            <LayerHost
                                                id='sanctionCategoryLayer'
                                                style={{ position: 'relative', zIndex: 500000 }} // higher than page chrome
                                            />
                                            <DetailsList
                                                items={itemsToDisplay}
                                                columns={columns}
                                                layoutMode={DetailsListLayoutMode.fixedColumns}
                                                constrainMode={ConstrainMode.horizontalConstrained}
                                                selectionMode={SelectionMode.none}
                                                onRenderItemColumn={this.renderItemColumn}
                                                onShouldVirtualize={() => false}
                                            />
                                            {this.props.pagination?.isEnabled && this.props.pagination.showBottomPagination && (
                                                <Pager
                                                    total={viewItems.length}
                                                    pageIndex={pageIndex}
                                                    pageSize={pageSize}
                                                    onPageChange={this.setPage}
                                                    onPageSizeChange={this.setPageSize}
                                                />
                                            )}
                                        </div>) : (
                                            <div className={styles.noDataMessage}>
                                                <Icon iconName="Remove" className={styles.mutedIcon} />
                                                <span>No record found</span>
                                                <Icon iconName="Remove" className={styles.mutedIcon} />
                                            </div>
                                        )
                                        }
                                    </div>
                                </>
                            )

                            }

                        </div > :
                        <div className="ms-Grid-col ms-sm12 ms-md12">
                            <div className={styles.spinnerContainer}>
                                <Spinner label={'Loading...'} />
                            </div>
                        </div>
                    }
                    </>
                }
                {isDeleteDialogOpen && <Dialog
                    hidden={documentToDelete === null}
                    onDismiss={this.closeDialogDelete}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: documentToDelete ? documentToDelete.Name : ''
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

                        {strings.LibraryViewRecycleBinConfirmation}
                        {this.state.messageBanner.visible && (
                            <div className={styles.topMessageBanner}>
                                <MessageBanner
                                    message={this.state.messageBanner.message}
                                    type={this.state.messageBanner.type}
                                    visible={this.state.messageBanner.visible}
                                />
                            </div>
                        )}
                        <DialogFooter>

                            <PrimaryButton
                                disabled={this.state.isConfirmButtonDisabled}
                                onClick={this.deleteDocument}
                            >
                                {this.state.isFormProcessing ? (
                                    <Stack horizontal verticalAlign="center">
                                        <Spinner size={SpinnerSize.xSmall} labelPosition="right" />
                                        <span style={{ marginLeft: 8 }}>Processing</span>
                                    </Stack>
                                ) : (
                                    'Confirm'
                                )}
                            </PrimaryButton>
                            <DefaultButton
                                onClick={this.closeDialogDelete}
                                text={'Back'}
                            />
                        </DialogFooter>
                    </div>
                </Dialog>
                }
                {isPreviewDocSource && (
                    <div className={styles.previewDoc}>
                        <div className={styles.previewContainer}>
                            <h3 className={styles.previewTitle}>{previewDocTitle}</h3>
                            <ActionButton
                                className={styles.closePreview}
                                iconProps={{ iconName: 'Cancel' }}
                                onClick={() => this.setState({ previewDocSource: '', isPreviewDocSource: false })}>
                                Close preview
                            </ActionButton>
                        </div>
                        <iframe
                            src={previewDocSource}
                            scrolling="auto"
                            frameBorder="0"
                            width="100%"
                            height="600"
                            onLoad={(e) => {
                                try {
                                    const iframeDoc = (e.target as HTMLIFrameElement).contentDocument || (e.target as HTMLIFrameElement).contentWindow?.document;
                                    if (iframeDoc) {
                                        // Hide OneUpCommandBar
                                        const commandBar = iframeDoc.getElementById('OneUpCommandBar');
                                        if (commandBar) {
                                            commandBar.style.display = 'none';
                                        }

                                        // Apply CSS to ms-ContextualMenu-itemText elements
                                        const styleElement = iframeDoc.createElement('style');
                                        styleElement.textContent = `
                                            .ms-ContextualMenu-itemText {
                                                font-family: "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif;
                                                -webkit-font-smoothing: antialiased;
                                                font-size: 14px;
                                                font-weight: 400;
                                                color: rgb(243, 242, 241);
                                                line-height: 36px;
                                                cursor: pointer;
                                                text-align: left;
                                            }
                                        `;
                                        iframeDoc.head.appendChild(styleElement);
                                    }
                                } catch (err) {
                                    console.error('Error styling iframe:', err);
                                }
                            }}
                        />
                    </div>
                )}
                {(isFolderVisible || isFileVisible) && (
                    <Modal
                        isOpen={isFolderVisible || isFileVisible}
                        layerProps={{ eventBubblingEnabled: true }}
                        onDismiss={this.onModalClose}
                        containerClassName="modalContainer"
                        isBlocking={true}>
                        <div className="document-form">
                            <div className="modalHeader">
                                {/* Dynamic Title */}
                                {formTitle}
                                <ActionButton
                                    iconProps={{ iconName: 'Cancel' }}
                                    onClick={this.onModalClose}
                                    className="iconButton">
                                    Close
                                </ActionButton>
                            </div>
                            <div className="modalBody">
                                {/* Render Folder or Document Form */}
                                {isFolderVisible ? (
                                    <MAFolder
                                        userService={this.props.userService}
                                        currentFolderPath={currentFolderPath}
                                        itemId={editFromId || undefined}
                                        pnpService={this.props.pnpService}
                                        callback={() => this.Callback()}
                                    />
                                ) : (
                                    <MADocument
                                        userService={this.props.userService}
                                        filepath={currentFolderPath}
                                        itemId={editFromId || undefined}
                                        pnpService={this.props.pnpService}
                                        callback={() => this.Callback()}
                                    />
                                )}
                            </div>
                        </div>
                    </Modal >
                )}
            </>
        );
    }
}

