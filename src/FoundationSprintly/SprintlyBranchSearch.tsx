import * as React from 'react';

import {
    getClient,
    IExtensionDataManager,
    IGlobalMessagesService,
    IProjectInfo,
    MessageBannerLevel,
} from 'azure-devops-extension-api';
import {
    GitBranchStats,
    GitRef,
    GitRefUpdate,
    GitRefUpdateResult,
    GitRefUpdateStatus,
    GitRepository,
    GitRestClient,
    GitVersionDescriptor,
    GitVersionOptions,
    GitVersionType,
} from 'azure-devops-extension-api/Git';

import {
    ObservableArray,
    ObservableValue,
} from 'azure-devops-ui/Core/Observable';

import { Button } from 'azure-devops-ui/Button';
import { ButtonGroup } from 'azure-devops-ui/ButtonGroup';
import { Card } from 'azure-devops-ui/Card';
import { Page } from 'azure-devops-ui/Page';
import { TextField, TextFieldWidth } from 'azure-devops-ui/TextField';
import {
    ColumnSelect,
    ColumnSorting,
    ITableColumn,
    SimpleTableCell,
    sortItems,
    SortOrder,
    Table,
} from 'azure-devops-ui/Table';

import { Icon } from 'azure-devops-ui/Icon';
import { Link } from 'azure-devops-ui/Link';
import { Tooltip } from 'azure-devops-ui/TooltipEx';
import { VssPersona } from 'azure-devops-ui/VssPersona';
import { Dialog } from 'azure-devops-ui/Dialog';
import { Observer } from 'azure-devops-ui/Observer';
import { ListSelection } from 'azure-devops-ui/List';
import { ISelectionRange } from 'azure-devops-ui/Utilities/Selection';
import { Spinner } from 'azure-devops-ui/Spinner';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import * as Common from './SprintlyCommon';

//#region "Observables"
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const searchObservable: ObservableValue<string> = new ObservableValue<string>(
    ''
);
const isLoadingObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const isDeleteSingleBranchDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const isDeleteBatchBranchDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const nameColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const repositoryColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const statsColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const branchCreatorColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const deleteBranchColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-40);

//#endregion "Observables"

const enum SearchType {
    AllBranches,
    JustMyBranches,
    AllMyBranches,
}

let repositoriesToProcess: string[] = [];

export interface ISprintlyBranchSearchPageState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    repositories: GitRepository[];
    searchResultBranchesObservable: ObservableArray<Common.ISearchResultBranch>;
}

export default class SprintlyBranchSearchPage extends React.Component<
    {
        dataManager: IExtensionDataManager;
        organizationName: string;
        userName: string;
        globalMessagesSvc: IGlobalMessagesService;
    },
    ISprintlyBranchSearchPageState
> {
    private dataManager: IExtensionDataManager;
    private organizationName: string;
    private userName: string;
    private globalMessagesSvc: IGlobalMessagesService;
    private branchToDelete?: Common.ISearchResultBranch;
    private searchResultsSelection: ListSelection = new ListSelection({
        selectOnFocus: false,
        multiSelect: true,
    });
    private columns: any = [];
    private sortingBehavior: ColumnSorting<Common.ISearchResultBranch> =
        new ColumnSorting<Common.ISearchResultBranch>(
            (
                columnIndex: number,
                proposedSortOrder: SortOrder,
                event:
                    | React.KeyboardEvent<HTMLElement>
                    | React.MouseEvent<HTMLElement>
            ) => {
                this.state.searchResultBranchesObservable.splice(
                    0,
                    this.state.searchResultBranchesObservable.length,
                    ...sortItems<Common.ISearchResultBranch>(
                        columnIndex,
                        proposedSortOrder,
                        this.sortFunctions,
                        this.columns,
                        this.state.searchResultBranchesObservable.value
                    )
                );
            }
        );
    private sortFunctions: any = [
        null,
        (
            a: Common.ISearchResultBranch,
            b: Common.ISearchResultBranch
        ): number => {
            return a.branchName.localeCompare(b.branchName);
        },
        (
            a: Common.ISearchResultBranch,
            b: Common.ISearchResultBranch
        ): number => {
            return a.repository.name.localeCompare(b.repository.name);
        },
        null,
        (
            a: Common.ISearchResultBranch,
            b: Common.ISearchResultBranch
        ): number => {
            return a.branchCreator.displayName.localeCompare(
                b.branchCreator.displayName
            );
        },
        null,
    ];

    constructor(props: {
        dataManager: IExtensionDataManager;
        globalMessagesSvc: IGlobalMessagesService;
        organizationName: string;
        userName: string;
    }) {
        super(props);

        this.onSize = this.onSize.bind(this);
        this.renderNameCell = this.renderNameCell.bind(this);
        this.renderRepositoryCell = this.renderRepositoryCell.bind(this);
        this.renderStatsCell = this.renderStatsCell.bind(this);
        this.renderDeleteBranchCell = this.renderDeleteBranchCell.bind(this);
        this.deleteSingleBranchAction =
            this.deleteSingleBranchAction.bind(this);
        this.deleteBatchBranchAction = this.deleteBatchBranchAction.bind(this);

        this.columns = [
            new ColumnSelect(),
            {
                id: 'name',
                name: 'Branch',
                onSize: this.onSize,
                renderCell: this.renderNameCell,
                sortProps: {
                    ariaLabelAscending: 'Sorted A to Z',
                    ariaLabelDescending: 'Sorted Z to A',
                },
                width: nameColumnWidthObservable,
            },
            {
                id: 'repository',
                name: 'Repository',
                onSize: this.onSize,
                renderCell: this.renderRepositoryCell,
                sortProps: {
                    ariaLabelAscending: 'Sorted A to Z',
                    ariaLabelDescending: 'Sorted Z to A',
                },
                width: repositoryColumnWidthObservable,
            },
            {
                id: 'stats',
                name: 'Behind Develop | Ahead Of Develop',
                onSize: this.onSize,
                renderCell: this.renderStatsCell,
                width: statsColumnWidthObservable,
            },
            {
                id: 'creator',
                name: 'Branch Creator',
                onSize: this.onSize,
                renderCell: this.renderBranchCreatorCell,
                sortProps: {
                    ariaLabelAscending: 'Sorted A to Z',
                    ariaLabelDescending: 'Sorted Z to A',
                },
                width: branchCreatorColumnWidthObservable,
            },
            {
                id: 'delete',
                name: 'Delete Branch',
                onSize: this.onSize,
                renderCell: this.renderDeleteBranchCell,
                width: deleteBranchColumnWidthObservable,
            },
        ];

        this.state = {
            repositories: [],
            searchResultBranchesObservable:
                new ObservableArray<Common.ISearchResultBranch>([]),
        };

        this.dataManager = props.dataManager;
        this.globalMessagesSvc = props.globalMessagesSvc;
        this.organizationName = props.organizationName;
        this.userName = props.userName;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        searchObservable.value = '';
        const userSettings: Common.IUserSettings | undefined =
            await Common.getUserSettings(
                this.dataManager,
                Common.USER_SETTINGS_DATA_MANAGER_KEY
            );
        const systemSettings: Common.ISystemSettings | undefined =
            await Common.getSystemSettings(
                this.dataManager,
                Common.SYSTEM_SETTINGS_DATA_MANAGER_KEY
            );

        this.setState({
            userSettings,
            systemSettings,
        });

        repositoriesToProcess = Common.getSavedRepositoriesToView(
            this.state.userSettings,
            this.state.systemSettings
        );

        totalRepositoriesToProcessObservable.value =
            repositoriesToProcess.length;
        if (repositoriesToProcess.length > 0) {
            const currentProject: IProjectInfo | undefined =
                await Common.getCurrentProject();
            await this.loadRepositoriesState(currentProject);
        }
    }

    private async loadRepositoriesState(
        currentProject: IProjectInfo | undefined
    ): Promise<void> {
        let repos: GitRepository[] = [];
        totalRepositoriesToProcessObservable.value = 0;
        if (currentProject !== undefined) {
            const filteredRepos: GitRepository[] =
                await Common.getFilteredProjectRepositories(
                    currentProject.id,
                    repositoriesToProcess
                );

            totalRepositoriesToProcessObservable.value += filteredRepos.length;
            repos = repos.concat(filteredRepos);
        }

        this.setState({
            repositories: repos,
        });
    }

    private renderNameCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.ISearchResultBranch>,
        tableItem: Common.ISearchResultBranch
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    <>
                        <Icon
                            iconName='OpenSource'
                            className='icon-margin'
                        ></Icon>
                        <u>
                            {Common.branchLinkJsxElement(
                                columnIndex.toString(),
                                tableItem.repository.webUrl,
                                Common.getBranchShortName(tableItem.branchName),
                                'bolt-table-link bolt-table-inline-link'
                            )}
                        </u>
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderRepositoryCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.ISearchResultBranch>,
        tableItem: Common.ISearchResultBranch
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    <>
                        <Icon iconName='Repo' className='icon-margin'></Icon>
                        <u>
                            {Common.repositoryLinkJsxElement(
                                tableItem.repository.webUrl,
                                '',
                                tableItem.repository.name
                            )}
                        </u>
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderStatsCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.ISearchResultBranch>,
        tableItem: Common.ISearchResultBranch
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    tableItem.branchStats ? (
                        <>
                            <Link
                                excludeTabStop
                                href={`https://dev.azure.com/${
                                    this.organizationName
                                }/${tableItem.projectId}/_git/${
                                    tableItem.repository.name
                                }/branchCompare?baseVersion=GB${Common.getBranchShortName(
                                    tableItem.branchName
                                )}&targetVersion=GBdevelop&_a=commits`}
                                subtle={true}
                                target='_blank'
                            >
                                <u>{tableItem.branchStats?.behindCount}</u>
                            </Link>
                            &nbsp;|&nbsp;
                            <Link
                                excludeTabStop
                                href={`https://dev.azure.com/${
                                    this.organizationName
                                }/${tableItem.projectId}/_git/${
                                    tableItem.repository.name
                                }/branchCompare?baseVersion=GB${
                                    Common.DEVELOP
                                }&targetVersion=GB${Common.getBranchShortName(
                                    tableItem.branchName
                                )}&_a=commits`}
                                subtle={true}
                                target='_blank'
                            >
                                <u>{tableItem.branchStats?.aheadCount}</u>
                            </Link>
                        </>
                    ) : (
                        <>No develop branch</>
                    )
                }
            ></SimpleTableCell>
        );
    }

    private renderBranchCreatorCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.ISearchResultBranch>,
        tableItem: Common.ISearchResultBranch
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    <>
                        <VssPersona
                            className='icon-margin'
                            imageUrl={
                                tableItem.branchCreator._links['avatar']['href']
                            }
                        />
                        <div className='flex-column text-ellipsis'>
                            <Tooltip overflowOnly={true}>
                                <div className='primary-text text-ellipsis'>
                                    {tableItem.branchCreator.displayName}
                                </div>
                            </Tooltip>
                            <Tooltip overflowOnly={true}>
                                <div className='primary-text text-ellipsis'>
                                    <Link
                                        excludeTabStop
                                        href={
                                            'mailto:' +
                                            tableItem.branchCreator.uniqueName
                                        }
                                        subtle={false}
                                        target='_blank'
                                    >
                                        {tableItem.branchCreator.uniqueName}
                                    </Link>
                                </div>
                            </Tooltip>
                        </div>
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderDeleteBranchCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.ISearchResultBranch>,
        tableItem: Common.ISearchResultBranch
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    <>
                        <Button
                            text='Delete branch'
                            iconProps={{ iconName: 'Delete' }}
                            onClick={() => {
                                isDeleteSingleBranchDialogOpenObservable.value =
                                    true;
                                this.branchToDelete = tableItem;
                            }}
                            danger={true}
                        />
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private async findBranchesInRepository(
        repositoryId: string,
        searchTerm: string,
        searchType: SearchType
    ): Promise<GitRef[]> {
        let repositoryBranches: GitRef[] = await getClient(
            GitRestClient
        ).getRefs(
            repositoryId,
            undefined,
            searchType === SearchType.AllBranches
                ? Common.repositoryHeadsFilter
                : undefined,
            undefined,
            undefined,
            searchType !== SearchType.AllBranches,
            undefined,
            undefined,
            searchType === SearchType.AllBranches ? searchTerm : undefined
        );
        if (searchType === SearchType.JustMyBranches) {
            repositoryBranches = repositoryBranches.filter(
                (branch: GitRef) =>
                    Common.getBranchShortName(branch.name).includes(
                        searchTerm
                    ) && branch.creator.uniqueName === this.userName
            );
        }
        if (searchType === SearchType.AllMyBranches) {
            repositoryBranches = repositoryBranches.filter(
                (branch: GitRef) => branch.creator.uniqueName === this.userName
            );
        }
        return repositoryBranches;
    }

    private async searchAction(searchType: SearchType): Promise<void> {
        isLoadingObservable.value = true;
        this.searchResultsSelection.clear();
        this.setState({
            searchResultBranchesObservable:
                new ObservableArray<Common.ISearchResultBranch>([]),
        });
        const searchTerm: string = searchObservable.value.trim();
        if (
            (searchTerm ||
                (!searchTerm && searchType === SearchType.AllMyBranches)) &&
            totalRepositoriesToProcessObservable.value > 0
        ) {
            const resultBranches: Common.ISearchResultBranch[] = [];
            for (const repositoryId of repositoriesToProcess) {
                const baseRepository: GitRepository | undefined =
                    this.state.repositories.find(
                        (repo: GitRepository) => repo.id === repositoryId
                    );
                if (baseRepository) {
                    const searchResultsBranches: GitRef[] =
                        await this.findBranchesInRepository(
                            repositoryId,
                            searchTerm,
                            searchType
                        );
                    if (searchResultsBranches.length > 0) {
                        let branchStatsBatch: GitBranchStats[] = [];
                        let repositoryDevelopBranch: GitBranchStats | undefined;
                        try {
                            repositoryDevelopBranch = await getClient(
                                GitRestClient
                            ).getBranch(repositoryId, Common.DEVELOP);
                        } catch (e) {
                            repositoryDevelopBranch = undefined;
                        }

                        if (repositoryDevelopBranch) {
                            const baseDevelopCommit: GitVersionDescriptor = {
                                version:
                                    repositoryDevelopBranch.commit.commitId,
                                versionOptions: GitVersionOptions.None,
                                versionType: GitVersionType.Commit,
                            };
                            const targetCommits: GitVersionDescriptor[] = [];
                            for (const branch of searchResultsBranches) {
                                targetCommits.push({
                                    version: branch.objectId,
                                    versionOptions: GitVersionOptions.None,
                                    versionType: GitVersionType.Commit,
                                });
                            }
                            branchStatsBatch = await getClient(
                                GitRestClient
                            ).getBranchStatsBatch(
                                {
                                    baseCommit: baseDevelopCommit,
                                    targetCommits,
                                },
                                repositoryId
                            );
                        }

                        for (const branch of searchResultsBranches) {
                            resultBranches.push({
                                branchName: branch.name,
                                repository: baseRepository,
                                branchCreator: branch.creator,
                                branchStats: repositoryDevelopBranch
                                    ? branchStatsBatch.find(
                                          (stat: GitBranchStats) =>
                                              stat.commit.commitId ===
                                              branch.objectId
                                      )
                                    : undefined,
                                projectId: baseRepository.project.id,
                            });
                        }
                    }
                }
            }
            this.setState({
                searchResultBranchesObservable:
                    new ObservableArray<Common.ISearchResultBranch>(
                        Common.sortSearchResultBranchesList(resultBranches)
                    ),
            });
        } else {
            this.globalMessagesSvc.addToast({
                duration: 5000,
                forceOverrideExisting: true,
                message:
                    'Please ensure you have typed in a search term or have repositories set up in the Settings tab.',
            });
        }
        isLoadingObservable.value = false;
    }

    private renderDeleteSingleBranchActionModal(): JSX.Element {
        return (
            <Observer
                isDeleteSingleBranchDialogOpen={
                    isDeleteSingleBranchDialogOpenObservable
                }
            >
                {(props: { isDeleteSingleBranchDialogOpen: boolean }) => {
                    return props.isDeleteSingleBranchDialogOpen ? (
                        <Dialog
                            titleProps={{
                                text: 'Delete branch',
                            }}
                            footerButtonProps={[
                                {
                                    text: 'Cancel',
                                    onClick:
                                        this
                                            .onDismissDeleteSingleBranchActionModal,
                                },
                                {
                                    text: 'Delete',
                                    onClick: this.deleteSingleBranchAction,
                                    danger: true,
                                },
                            ]}
                            onDismiss={
                                this.onDismissDeleteSingleBranchActionModal
                            }
                        >
                            <>
                                Branch{' '}
                                {Common.getBranchShortName(
                                    this.branchToDelete?.branchName ?? ''
                                )}{' '}
                                will be permanently deleted. Are you sure you
                                want to proceed?
                            </>
                        </Dialog>
                    ) : null;
                }}
            </Observer>
        );
    }

    private renderDeleteBatchBranchActionModal(): JSX.Element {
        return (
            <Observer
                isDeleteBatchBranchDialogOpen={
                    isDeleteBatchBranchDialogOpenObservable
                }
            >
                {(props: { isDeleteBatchBranchDialogOpen: boolean }) => {
                    return props.isDeleteBatchBranchDialogOpen ? (
                        <Dialog
                            titleProps={{
                                text: 'Delete branchs',
                            }}
                            footerButtonProps={[
                                {
                                    text: 'Cancel',
                                    onClick:
                                        this
                                            .onDismissDeleteBatchBranchActionModal,
                                },
                                {
                                    text: 'Delete',
                                    onClick: this.deleteBatchBranchAction,
                                    danger: true,
                                },
                            ]}
                            onDismiss={
                                this.onDismissDeleteBatchBranchActionModal
                            }
                        >
                            <>
                                These branches will be permanently deleted. Are
                                you sure you want to proceed?
                            </>
                        </Dialog>
                    ) : null;
                }}
            </Observer>
        );
    }

    private deleteBranchesAction(
        branchesToDelete: Common.ISearchResultBranch[]
    ): void {
        if (branchesToDelete.length > 0) {
            const createRefOptions: GitRefUpdate[] = [];

            const uniqueRepositories: string[] = branchesToDelete
                .map((branch: Common.ISearchResultBranch) => {
                    return branch.repository.id;
                })
                .filter(
                    (value: string, index: number, array: string[]) =>
                        array.indexOf(value) === index
                );

            for (const repositoryId of uniqueRepositories) {
                for (const branchToDelete of branchesToDelete) {
                    if (branchToDelete.branchStats) {
                        const branchShortName: string =
                            Common.getBranchShortName(
                                branchToDelete.branchName
                            );
                        if (
                            branchShortName === Common.DEVELOP ||
                            branchShortName === Common.MASTER ||
                            branchShortName === Common.MAIN
                        ) {
                            this.globalMessagesSvc.closeBanner();
                            this.globalMessagesSvc.addBanner({
                                dismissable: true,
                                level: MessageBannerLevel.error,
                                message:
                                    'Error: Will not delete "develop", "master", or "main" branches.',
                            });
                            return;
                        }
                        createRefOptions.push({
                            repositoryId: branchToDelete.repository.id,
                            name: branchToDelete.branchName,
                            isLocked: false,
                            oldObjectId:
                                branchToDelete.branchStats.commit.commitId,
                            newObjectId:
                                '0000000000000000000000000000000000000000',
                        });
                    }
                }

                if (createRefOptions.length > 0) {
                    getClient(GitRestClient)
                        .updateRefs(createRefOptions, repositoryId)
                        .then(async (result: GitRefUpdateResult[]) => {
                            for (const res of result) {
                                this.globalMessagesSvc.addToast({
                                    duration: 5000,
                                    forceOverrideExisting: true,
                                    message: res.success
                                        ? 'Branch(es) Deleted!'
                                        : 'Error Deleting Branch(es): ' +
                                          GitRefUpdateStatus[res.updateStatus],
                                });
                                if (res.success) {
                                    const searchResults: Common.ISearchResultBranch[] =
                                        this.state
                                            .searchResultBranchesObservable
                                            .value;
                                    const indexToRemove: number =
                                        searchResults.findIndex(
                                            (
                                                branch: Common.ISearchResultBranch
                                            ) =>
                                                branch.branchName ===
                                                    res.name &&
                                                branch.repository.id ===
                                                    res.repositoryId
                                        );
                                    searchResults.splice(indexToRemove, 1);
                                    this.setState({
                                        searchResultBranchesObservable:
                                            new ObservableArray<Common.ISearchResultBranch>(
                                                searchResults
                                            ),
                                    });
                                }
                            }
                        })
                        .catch((error: any) => {
                            if (error.response?.data?.message) {
                                this.globalMessagesSvc.closeBanner();
                                this.globalMessagesSvc.addBanner({
                                    dismissable: true,
                                    level: MessageBannerLevel.error,
                                    message: error.response.data.message,
                                });
                            } else {
                                this.globalMessagesSvc.addToast({
                                    duration: 5000,
                                    forceOverrideExisting: true,
                                    message:
                                        'Branch(es) deletion failed!' +
                                        error +
                                        ' ' +
                                        error.response?.data?.message,
                                });
                            }
                        });
                }
            }
        }
    }

    private deleteSingleBranchAction(): void {
        if (this.branchToDelete && this.branchToDelete.branchStats) {
            this.deleteBranchesAction([this.branchToDelete]);
        }

        this.searchResultsSelection.clear();
        this.onDismissDeleteSingleBranchActionModal();
    }

    private deleteBatchBranchAction(): void {
        const branchesToDelete: Common.ISearchResultBranch[] =
            this.getSelectedRange(
                this.searchResultsSelection.value,
                this.state.searchResultBranchesObservable.value
            );

        this.deleteBranchesAction(branchesToDelete);

        this.searchResultsSelection.clear();
        this.onDismissDeleteBatchBranchActionModal();
    }

    private onDismissDeleteSingleBranchActionModal(): void {
        isDeleteSingleBranchDialogOpenObservable.value = false;
    }

    private onDismissDeleteBatchBranchActionModal(): void {
        isDeleteBatchBranchDialogOpenObservable.value = false;
    }

    private getSelectedRange(
        selectionRange: ISelectionRange[],
        dataArray: Common.ISearchResultBranch[]
    ): Common.ISearchResultBranch[] {
        const selectedArray: Common.ISearchResultBranch[] = [];
        for (const rng of selectionRange) {
            const sliced: Common.ISearchResultBranch[] = dataArray.slice(
                rng.beginIndex,
                rng.endIndex + 1
            );
            for (const slice of sliced) {
                selectedArray.push(slice);
            }
        }
        return selectedArray;
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (this.columns[index].width as ObservableValue<number>).value = width;
    }

    public render(): JSX.Element {
        return (
            <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                <div>
                    Search for a branch name across all of the repositories you
                    have set up in the Settings tab.
                </div>
                <ButtonGroup>
                    <TextField
                        prefixIconProps={{ iconName: 'Search' }}
                        value={searchObservable}
                        onChange={(
                            event: React.ChangeEvent<
                                HTMLInputElement | HTMLTextAreaElement
                            >,
                            newValue: string
                        ) => (searchObservable.value = newValue)}
                        placeholder='Search Branch Name'
                        width={TextFieldWidth.standard}
                    />
                    <Button
                        text='Search All Branches'
                        primary={true}
                        onClick={async () =>
                            await this.searchAction(SearchType.AllBranches)
                        }
                    />
                    <Button
                        text='Just My Branches'
                        primary={true}
                        onClick={async () =>
                            await this.searchAction(SearchType.JustMyBranches)
                        }
                    />
                    <Button
                        text='All My Branches (Ignore Search Term)'
                        primary={true}
                        onClick={async () =>
                            await this.searchAction(SearchType.AllMyBranches)
                        }
                    />
                </ButtonGroup>
                <Observer isLoading={isLoadingObservable}>
                    {(observerProps: { isLoading: boolean }) => {
                        if (observerProps.isLoading) {
                            return <Spinner label='loading' />;
                        } else {
                            return <></>;
                        }
                    }}
                </Observer>
                <Observer
                    searchResults={this.state.searchResultBranchesObservable}
                    isLoading={isLoadingObservable}
                >
                    {(observerProps: {
                        searchResults: Common.ISearchResultBranch[];
                        isLoading: boolean;
                    }) => {
                        if (observerProps.searchResults.length > 0) {
                            return (
                                <>
                                    <ButtonGroup>
                                        <Button
                                            text='Delete Selected'
                                            iconProps={{ iconName: 'Delete' }}
                                            danger={true}
                                            onClick={() => {
                                                isDeleteBatchBranchDialogOpenObservable.value =
                                                    true;
                                            }}
                                        />
                                    </ButtonGroup>
                                    <Page>
                                        <Card className='bolt-table-card bolt-card-white'>
                                            <Table
                                                columns={this.columns}
                                                behaviors={[
                                                    this.sortingBehavior,
                                                ]}
                                                itemProvider={
                                                    this.state
                                                        .searchResultBranchesObservable
                                                }
                                                selection={
                                                    this.searchResultsSelection
                                                }
                                            />
                                        </Card>
                                        {this.renderDeleteSingleBranchActionModal()}
                                        {this.renderDeleteBatchBranchActionModal()}
                                    </Page>
                                </>
                            );
                        } else {
                            if (
                                searchObservable.value &&
                                !observerProps.isLoading
                            ) {
                                return (
                                    <div>
                                        <ZeroData
                                            primaryText='No results found.'
                                            secondaryText={
                                                <span>
                                                    Please update your search
                                                    term.
                                                </span>
                                            }
                                            imageAltText='No results found.'
                                            imagePath={'../static/notfound.png'}
                                        />
                                    </div>
                                );
                            }
                            return (
                                <>
                                    <div></div>
                                </>
                            );
                        }
                    }}
                </Observer>
            </div>
        );
    }
}
