import * as React from 'react';

import {
    getClient,
    IExtensionDataManager,
    IGlobalMessagesService,
    MessageBannerLevel,
} from 'azure-devops-extension-api';
import {
    GitBranchStats,
    GitCommitRef,
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
import { IdentityRef } from 'azure-devops-extension-api/WebApi';

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
    ColumnSorting,
    ITableColumn,
    SimpleTableCell,
    sortItems,
    SortOrder,
    Table,
} from 'azure-devops-ui/Table';

import * as Common from './SprintlyCommon';
import { TeamProjectReference } from 'azure-devops-extension-api/Core';
import { bindSelectionToObservable } from 'azure-devops-ui/MasterDetailsContext';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { Icon, IconSize } from 'azure-devops-ui/Icon';
import { Link } from 'azure-devops-ui/Link';
import { Tooltip } from 'azure-devops-ui/TooltipEx';
import { VssPersona } from 'azure-devops-ui/VssPersona';
import { Dialog } from 'azure-devops-ui/Dialog';
import { Observer } from 'azure-devops-ui/Observer';

//#region "Observables"
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const searchObservable = new ObservableValue<string>('');
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

const userSettingsDataManagerKey: string = 'user-settings';
const systemSettingsDataManagerKey: string = 'system-settings';

let repositoriesToProcess: string[] = [];

export interface ISearchResultBranch {
    branchName: string;
    branchStats?: GitBranchStats;
    branchCreator: IdentityRef;
    repository: GitRepository;
    projectId: string;
}

export interface ISprintlyBranchSearchPageState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    repositories: GitRepository[];
    searchResultBranchesObservable: ObservableArray<ISearchResultBranch>;
}

export default class SprintlyBranchSearchPage extends React.Component<
    {
        dataManager: IExtensionDataManager;
        organizationName: string;
        globalMessagesSvc: IGlobalMessagesService;
    },
    ISprintlyBranchSearchPageState
> {
    private dataManager: IExtensionDataManager;
    private organizationName: string;
    private globalMessagesSvc: IGlobalMessagesService;
    private branchToDelete?: ISearchResultBranch;
    private columns: any = [];
    private sortingBehavior: ColumnSorting<ISearchResultBranch> =
        new ColumnSorting<ISearchResultBranch>(
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
                    ...sortItems<ISearchResultBranch>(
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
        (a: ISearchResultBranch, b: ISearchResultBranch): number => {
            return a.branchName.localeCompare(b.branchName);
        },
        (a: ISearchResultBranch, b: ISearchResultBranch): number => {
            return a.repository.name.localeCompare(b.repository.name);
        },
        null,
        (a: ISearchResultBranch, b: ISearchResultBranch): number => {
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
    }) {
        super(props);

        this.onSize = this.onSize.bind(this);
        this.renderNameCell = this.renderNameCell.bind(this);
        this.renderRepositoryCell = this.renderRepositoryCell.bind(this);
        this.renderStatsCell = this.renderStatsCell.bind(this);
        this.renderBranchCreatorCell = this.renderBranchCreatorCell.bind(this);
        this.renderDeleteBranchCell = this.renderDeleteBranchCell.bind(this);
        this.deleteSingleBranchAction =
            this.deleteSingleBranchAction.bind(this);

        this.columns = [
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
                new ObservableArray<ISearchResultBranch>([]),
        };

        this.dataManager = props.dataManager;
        this.globalMessagesSvc = props.globalMessagesSvc;
        this.organizationName = props.organizationName;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        searchObservable.value = '';
        const userSettings: Common.IUserSettings | undefined =
            await Common.getUserSettings(
                this.dataManager,
                userSettingsDataManagerKey
            );
        const systemSettings: Common.ISystemSettings | undefined =
            await Common.getSystemSettings(
                this.dataManager,
                systemSettingsDataManagerKey
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
            const filteredProjects: TeamProjectReference[] =
                await Common.getFilteredProjects();
            await this.loadRepositoriesState(filteredProjects);
        }
    }

    private async loadRepositoriesState(
        projects: TeamProjectReference[]
    ): Promise<void> {
        let repos: GitRepository[] = [];
        totalRepositoriesToProcessObservable.value = 0;
        for (const project of projects) {
            const filteredRepos: GitRepository[] =
                await Common.getFilteredProjectRepositories(
                    project.id,
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
        tableColumn: ITableColumn<ISearchResultBranch>,
        tableItem: ISearchResultBranch
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
                                tableItem.branchName.split('refs/heads/')[1],
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
        tableColumn: ITableColumn<ISearchResultBranch>,
        tableItem: ISearchResultBranch
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
        tableColumn: ITableColumn<ISearchResultBranch>,
        tableItem: ISearchResultBranch
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    <>
                        <Link
                            excludeTabStop
                            href={`https://dev.azure.com/${
                                this.organizationName
                            }/${tableItem.projectId}/_git/${
                                tableItem.repository.name
                            }/branchCompare?baseVersion=GB${
                                tableItem.branchName.split('refs/heads/')[1]
                            }&targetVersion=GBdevelop&_a=commits`}
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
                            }&targetVersion=GB${
                                tableItem.branchName.split('refs/heads/')[1]
                            }&_a=commits`}
                            subtle={true}
                            target='_blank'
                        >
                            <u>{tableItem.branchStats?.aheadCount}</u>
                        </Link>
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderBranchCreatorCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<ISearchResultBranch>,
        tableItem: ISearchResultBranch
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
        tableColumn: ITableColumn<ISearchResultBranch>,
        tableItem: ISearchResultBranch
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
        searchTerm: string
    ): Promise<GitRef[]> {
        const repositoryBranches: GitRef[] = await getClient(
            GitRestClient
        ).getRefs(
            repositoryId,
            undefined,
            Common.repositoryHeadsFilter,
            undefined,
            undefined,
            undefined,
            undefined,
            undefined,
            searchTerm
        );
        return repositoryBranches;
    }

    private async searchAction(): Promise<void> {
        const searchTerm: string = searchObservable.value.trim();
        if (searchTerm && totalRepositoriesToProcessObservable.value > 0) {
            const resultBranches: ISearchResultBranch[] = [];
            for (const repositoryId of repositoriesToProcess) {
                const baseRepository = this.state.repositories.find(
                    (repo) => repo.id === repositoryId
                );
                if (baseRepository) {
                    const searchResultsBranches: GitRef[] =
                        await this.findBranchesInRepository(
                            repositoryId,
                            searchTerm
                        );
                    if (searchResultsBranches.length > 0) {
                        const repositoryDevelopBranch: GitBranchStats =
                            await getClient(GitRestClient).getBranch(
                                repositoryId,
                                Common.DEVELOP
                            );
                        const baseDevelopCommit: GitVersionDescriptor = {
                            version: repositoryDevelopBranch.commit.commitId,
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
                        const branchStatsBatch: GitBranchStats[] =
                            await getClient(GitRestClient).getBranchStatsBatch(
                                {
                                    baseCommit: baseDevelopCommit,
                                    targetCommits: targetCommits,
                                },
                                repositoryId
                            );
                        for (const branch of searchResultsBranches) {
                            resultBranches.push({
                                branchName: branch.name,
                                repository: baseRepository,
                                branchCreator: branch.creator,
                                branchStats: branchStatsBatch.find(
                                    (stat) =>
                                        stat.commit.commitId === branch.objectId
                                ),
                                projectId: baseRepository.project.id,
                            });
                        }
                    }
                }
            }
            this.setState({
                searchResultBranchesObservable:
                    new ObservableArray<ISearchResultBranch>(resultBranches),
            });
        }
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

    private deleteSingleBranchAction(): void {
        if (this.branchToDelete && this.branchToDelete.branchStats) {
            const createRefOptions: GitRefUpdate[] = [];

            createRefOptions.push({
                repositoryId: this.branchToDelete.repository.id,
                name: this.branchToDelete.branchName,
                isLocked: false,
                oldObjectId: this.branchToDelete.branchStats.commit.commitId,
                newObjectId: '0000000000000000000000000000000000000000',
            });

            getClient(GitRestClient)
                .updateRefs(createRefOptions, this.branchToDelete.repository.id)
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
                            const searchResults: ISearchResultBranch[] =
                                this.state.searchResultBranchesObservable.value;
                            const branchName =
                                this.branchToDelete?.branchName ?? '';
                            const repositoryId =
                                this.branchToDelete?.repository.id ?? '';
                            const indexToRemove: number =
                                searchResults.findIndex(
                                    (branch) =>
                                        branch.branchName === branchName &&
                                        branch.repository.id === repositoryId
                                );
                            searchResults.splice(indexToRemove, 1);
                            this.setState({
                                searchResultBranchesObservable:
                                    new ObservableArray<ISearchResultBranch>(
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

        this.onDismissDeleteSingleBranchActionModal();
    }

    private onDismissDeleteSingleBranchActionModal(): void {
        isDeleteSingleBranchDialogOpenObservable.value = false;
    }

    private onDismissDeleteBatchBranchActionModal(): void {
        isDeleteBatchBranchDialogOpenObservable.value = false;
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (this.columns[index].width as ObservableValue<number>).value = width;
    }

    public render(): JSX.Element {
        return (
            <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                <ButtonGroup>
                    <TextField
                        prefixIconProps={{ iconName: 'Search' }}
                        value={searchObservable}
                        onChange={(e, newValue) =>
                            (searchObservable.value = newValue)
                        }
                        placeholder='Search Branches'
                        width={TextFieldWidth.standard}
                    />
                    <Button
                        text='Search'
                        primary={true}
                        onClick={async () => await this.searchAction()}
                    />
                </ButtonGroup>
                <Observer
                    searchResults={this.state.searchResultBranchesObservable}
                >
                    {(observerProps: {
                        searchResults: ISearchResultBranch[];
                    }) => (
                        <Page>
                            <Card className='bolt-table-card bolt-card-white'>
                                <Table
                                    columns={this.columns}
                                    behaviors={[this.sortingBehavior]}
                                    itemProvider={
                                        this.state
                                            .searchResultBranchesObservable
                                    }
                                />
                            </Card>
                            {this.renderDeleteSingleBranchActionModal()}
                        </Page>
                    )}
                </Observer>
            </div>
        );
    }
}
