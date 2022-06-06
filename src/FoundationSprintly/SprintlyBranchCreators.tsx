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

import { Card } from 'azure-devops-ui/Card';
import {
    ObservableArray,
    ObservableValue,
} from 'azure-devops-ui/Core/Observable';
import { Icon } from 'azure-devops-ui/Icon';
import { Link } from 'azure-devops-ui/Link';
import {
    IListItemDetails,
    List,
    ListItem,
    ListSelection,
} from 'azure-devops-ui/List';
import { bindSelectionToObservable } from 'azure-devops-ui/MasterDetailsContext';
import { Observer } from 'azure-devops-ui/Observer';
import { Page } from 'azure-devops-ui/Page';
import { Spinner } from 'azure-devops-ui/Spinner';
import {
    Splitter,
    SplitterDirection,
    SplitterElementPosition,
} from 'azure-devops-ui/Splitter';
import {
    ColumnSelect,
    ColumnSorting,
    ITableColumn,
    SimpleTableCell,
    sortItems,
    SortOrder,
    Table,
} from 'azure-devops-ui/Table';
import { Tooltip } from 'azure-devops-ui/TooltipEx';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { VssPersona } from 'azure-devops-ui/VssPersona';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import { Button } from 'azure-devops-ui/Button';
import { Dialog } from 'azure-devops-ui/Dialog';
import { ISelectionRange } from 'azure-devops-ui/Utilities/Selection';
import { ButtonGroup } from 'azure-devops-ui/ButtonGroup';

import * as Common from './SprintlyCommon';

export interface ISprintlyBranchCreatorsState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    repositories: ArrayItemProvider<GitRepository>;
    repositoryBranchesObservable: ObservableArray<Common.ISearchResultBranch>;
    repositoryListSelection: ListSelection;
    repositoryListSelectedItemObservable: ObservableValue<GitRepository>;
}

//#region "Observables"
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const isDeleteSingleBranchDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const isDeleteBatchBranchDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
//#endregion "Observables"

let repositoriesToProcess: string[] = [];

export default class SprintlyBranchCreators extends React.Component<
    {
        accessToken: string;
        organizationName: string;
        globalMessagesSvc: IGlobalMessagesService;
    },
    ISprintlyBranchCreatorsState
> {
    private dataManager!: IExtensionDataManager;
    private accessToken: string;
    private organizationName: string;
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
                this.state.repositoryBranchesObservable.splice(
                    0,
                    this.state.repositoryBranchesObservable.length,
                    ...sortItems<Common.ISearchResultBranch>(
                        columnIndex,
                        proposedSortOrder,
                        this.sortFunctions,
                        this.columns,
                        this.state.repositoryBranchesObservable.value
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
        null,
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
        accessToken: string;
        globalMessagesSvc: IGlobalMessagesService;
        organizationName: string;
    }) {
        super(props);

        this.onSize = this.onSize.bind(this);
        this.renderRepositoryMasterPageList =
            this.renderRepositoryMasterPageList.bind(this);
        this.renderDetailPageContent = this.renderDetailPageContent.bind(this);
        this.renderNameCell = this.renderNameCell.bind(this);
        this.renderLatestCommitCell = this.renderLatestCommitCell.bind(this);
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
                width: new ObservableValue<number>(-30),
            },
            {
                id: 'commit',
                name: 'Latest Commit',
                onSize: this.onSize,
                renderCell: this.renderLatestCommitCell,
                width: new ObservableValue<number>(-30),
            },
            {
                id: 'stats',
                name: 'Behind Develop | Ahead Of Develop',
                onSize: this.onSize,
                renderCell: this.renderStatsCell,
                width: new ObservableValue<number>(-30),
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
                width: new ObservableValue<number>(-30),
            },
            {
                id: 'delete',
                name: 'Delete Branch',
                onSize: this.onSize,
                renderCell: this.renderDeleteBranchCell,
                width: new ObservableValue<number>(-40),
            },
        ];

        this.state = {
            repositories: new ArrayItemProvider<GitRepository>([]),
            repositoryBranchesObservable:
                new ObservableArray<Common.ISearchResultBranch>([]),
            repositoryListSelection: new ListSelection({
                selectOnFocus: false,
            }),
            repositoryListSelectedItemObservable: new ObservableValue<any>({}),
        };

        this.accessToken = props.accessToken;
        this.globalMessagesSvc = props.globalMessagesSvc;
        this.organizationName = props.organizationName;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        this.dataManager = await Common.initializeDataManager(
            await Common.getOrRefreshToken(this.accessToken)
        );
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
            await this.loadRepositoriesDisplayState(currentProject);
        }
    }

    private async loadRepositoriesDisplayState(
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
            repositories: new ArrayItemProvider(
                Common.sortRepositoryList(repos)
            ),
        });

        bindSelectionToObservable(
            this.state.repositoryListSelection,
            this.state.repositories,
            this.state
                .repositoryListSelectedItemObservable as ObservableValue<GitRepository>
        );
    }

    private renderRepositoryMasterPageList(): JSX.Element {
        return !this.state.repositories ||
            this.state.repositories.length === 0 ? (
            <div className='page-content-top'>
                <Spinner label='loading' />
            </div>
        ) : (
            <List
                ariaLabel={'Repositories'}
                itemProvider={this.state.repositories}
                selection={this.state.repositoryListSelection}
                renderRow={this.renderRepositoryListItem}
                width='100%'
                singleClickActivation={true}
                onSelect={async () => {
                    await this.selectRepository();
                }}
            />
        );
    }

    private renderDetailPageContent(): JSX.Element {
        return (
            <Observer
                selectedItem={this.state.repositoryListSelectedItemObservable}
                repositoryBranches={this.state.repositoryBranchesObservable}
            >
                {(observerProps: {
                    selectedItem: GitRepository;
                    repositoryBranches: Common.ISearchResultBranch[];
                }) => (
                    <Page className='flex-grow single-layer-details'>
                        {this.state.repositoryListSelection.selectedCount ===
                            0 && (
                            <span className='single-layer-details-contents'>
                                Select a repository on the right to see its
                                branch creators.
                            </span>
                        )}
                        {observerProps.selectedItem &&
                            this.state.repositoryListSelection.selectedCount >
                                0 && (
                                <Page>
                                    <div className='page-content page-content-top'>
                                        <ButtonGroup>
                                            <Button
                                                text='Delete Selected'
                                                iconProps={{
                                                    iconName: 'Delete',
                                                }}
                                                danger={true}
                                                onClick={() => {
                                                    isDeleteBatchBranchDialogOpenObservable.value =
                                                        true;
                                                }}
                                            />
                                        </ButtonGroup>
                                        <br />
                                        <Card className='bolt-table-card bolt-card-white'>
                                            <Table
                                                columns={this.columns}
                                                behaviors={[
                                                    this.sortingBehavior,
                                                ]}
                                                itemProvider={
                                                    this.state
                                                        .repositoryBranchesObservable
                                                }
                                                selection={
                                                    this.searchResultsSelection
                                                }
                                            />
                                        </Card>
                                    </div>
                                </Page>
                            )}
                    </Page>
                )}
            </Observer>
        );
    }

    private renderRepositoryListItem(
        index: number,
        item: GitRepository,
        details: IListItemDetails<GitRepository>,
        key?: string
    ): JSX.Element {
        return (
            <ListItem
                className='master-row border-bottom'
                key={key || 'list-item' + index}
                index={index}
                details={details}
            >
                <div className='master-row-content flex-row flex-center h-scroll-hidden'>
                    <div className='flex-column text-ellipsis'>
                        <Tooltip overflowOnly={true}>
                            <div className='primary-text text-ellipsis'>
                                {Common.repositoryLinkJsxElement(
                                    item.webUrl,
                                    'font-size-1',
                                    item.name
                                )}
                            </div>
                        </Tooltip>
                    </div>
                </div>
            </ListItem>
        );
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
                                this.state.repositoryListSelectedItemObservable
                                    .value.webUrl,
                                tableItem.branchName.split('refs/heads/')[1],
                                'bolt-table-link bolt-table-inline-link'
                            )}
                        </u>
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderLatestCommitCell(
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
                    <Link
                        excludeTabStop
                        href={
                            this.state.repositoryListSelectedItemObservable
                                .value.webUrl +
                            '/commit/' +
                            tableItem.branchStats?.commit.commitId
                        }
                        subtle={true}
                        target='_blank'
                        className='bolt-table-link bolt-table-inline-link'
                    >
                        {tableItem.branchStats?.commit.commitId.substr(0, 8)}
                    </Link>
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

    private async selectRepository(): Promise<void> {
        const repositoryInfo: GitRepository =
            this.state.repositoryListSelectedItemObservable.value;
        const repositoryBranches: GitRef[] = await getClient(
            GitRestClient
        ).getRefs(
            repositoryInfo.id,
            undefined,
            Common.repositoryHeadsFilter,
            undefined,
            undefined,
            undefined,
            undefined,
            undefined,
            undefined
        );

        const resultBranches: Common.ISearchResultBranch[] = [];
        if (repositoryBranches.length > 0) {
            const repositoryDevelopBranch: GitRef | undefined =
                repositoryBranches.find(
                    (branch: GitRef) =>
                        Common.getBranchShortName(branch.name) ===
                        Common.DEVELOP
                );
            let branchStatsBatch: GitBranchStats[] = [];
            if (repositoryDevelopBranch) {
                const baseDevelopCommit: GitVersionDescriptor = {
                    version: repositoryDevelopBranch.objectId,
                    versionOptions: GitVersionOptions.None,
                    versionType: GitVersionType.Commit,
                };
                const targetCommits: GitVersionDescriptor[] = [];
                for (const branch of repositoryBranches) {
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
                    repositoryInfo.id
                );
            }

            for (const branch of repositoryBranches) {
                resultBranches.push({
                    branchName: branch.name,
                    repository: repositoryInfo,
                    branchCreator: branch.creator,
                    branchStats: repositoryDevelopBranch
                        ? branchStatsBatch.find(
                              (stat: GitBranchStats) =>
                                  stat.commit.commitId === branch.objectId
                          )
                        : undefined,
                    projectId: repositoryInfo.project.id,
                });
            }
        }

        this.setState({
            repositoryBranchesObservable:
                new ObservableArray<Common.ISearchResultBranch>(
                    Common.sortSearchResultBranchesList(resultBranches)
                ),
        });
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
                                text: 'Delete branches',
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
            let createRefOptions: GitRefUpdate[] = [];

            const uniqueRepositories: string[] = branchesToDelete
                .map((branch: Common.ISearchResultBranch) => {
                    return branch.repository.id;
                })
                .filter(
                    (value: string, index: number, array: string[]) =>
                        array.indexOf(value) === index
                );

            for (const repositoryId of uniqueRepositories) {
                createRefOptions = [];
                for (const branchToDelete of branchesToDelete) {
                    if (
                        branchToDelete.branchStats &&
                        branchToDelete.repository.id === repositoryId
                    ) {
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
                                        this.state.repositoryBranchesObservable
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
                                        repositoryBranchesObservable:
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
                this.state.repositoryBranchesObservable.value
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
            <Observer
                totalRepositoriesToProcess={
                    totalRepositoriesToProcessObservable
                }
            >
                {(props: { totalRepositoriesToProcess: number }) => {
                    if (props.totalRepositoriesToProcess > 0) {
                        return (
                            <div
                                className='flex-grow'
                                style={{
                                    display: 'flex',
                                    height: '0%',
                                }}
                            >
                                <Splitter
                                    fixedElement={SplitterElementPosition.Near}
                                    splitterDirection={
                                        SplitterDirection.Vertical
                                    }
                                    initialFixedSize={450}
                                    minFixedSize={100}
                                    nearElementClassName='v-scroll-auto custom-scrollbar'
                                    farElementClassName='v-scroll-auto custom-scrollbar'
                                    onRenderNearElement={
                                        this.renderRepositoryMasterPageList
                                    }
                                    onRenderFarElement={
                                        this.renderDetailPageContent
                                    }
                                />
                                {this.renderDeleteSingleBranchActionModal()}
                                {this.renderDeleteBatchBranchActionModal()}
                            </div>
                        );
                    }
                    return (
                        <ZeroData
                            primaryText='No repositories.'
                            secondaryText={
                                <span>
                                    Please select valid repositories from the
                                    Settings page.
                                </span>
                            }
                            imageAltText='No repositories.'
                            imagePath={'../static/notfound.png'}
                        />
                    );
                }}
            </Observer>
        );
    }
}
