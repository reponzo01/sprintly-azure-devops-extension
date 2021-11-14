import * as React from 'react';

import { getClient, IExtensionDataManager } from 'azure-devops-extension-api';
import { TeamProjectReference } from 'azure-devops-extension-api/Core';
import {
    GitRef,
    GitRepository,
    GitRestClient,
} from 'azure-devops-extension-api/Git';

import { Card } from 'azure-devops-ui/Card';
import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Icon, IconSize } from 'azure-devops-ui/Icon';
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
    ColumnSorting,
    ITableColumn,
    SimpleTableCell,
    SortOrder,
    Table,
} from 'azure-devops-ui/Table';
import { Tooltip } from 'azure-devops-ui/TooltipEx';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { VssPersona } from 'azure-devops-ui/VssPersona';
import { ZeroData } from 'azure-devops-ui/ZeroData';

import * as Common from './SprintlyCommon';

export interface ISprintlyBranchCreatorsState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    repositories: ArrayItemProvider<GitRepository>;
    repositoryBranches: ArrayItemProvider<GitRef>;
    repositoryListSelection: ListSelection;
    repositoryListSelectedItemObservable: ObservableValue<GitRepository>;
}

//#region "Observables"
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const nameColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const latestCommitColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const branchCreatorColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-40);
//#endregion "Observables"

const userSettingsDataManagerKey: string = 'user-settings';
const systemSettingsDataManagerKey: string = 'system-settings';

let repositoriesToProcess: string[] = [];

export default class SprintlyBranchCreators extends React.Component<
    {
        dataManager: IExtensionDataManager;
    },
    ISprintlyBranchCreatorsState
> {
    private dataManager: IExtensionDataManager;
    private columns: any = [];
    // private sortingBehavior = new ColumnSorting<GitRef>(
    //     (
    //         columnIndex: number,
    //         proposedSortOrder: SortOrder,
    //         event:
    //             | React.KeyboardEvent<HTMLElement>
    //             | React.MouseEvent<HTMLElement>
    //     ) => {
    //         this.state.repositoryBranches.splice(
    //             0,
    //             tableItems.length,
    //             ...sortItems<GitRef>(
    //                 columnIndex,
    //                 proposedSortOrder,
    //                 sortFunctions,
    //                 columns,
    //                 rawTableItems
    //             )
    //         );
    //     }
    // );

    constructor(props: { dataManager: IExtensionDataManager }) {
        super(props);

        this.onSize = this.onSize.bind(this);
        this.renderRepositoryMasterPageList =
            this.renderRepositoryMasterPageList.bind(this);
        this.renderDetailPageContent = this.renderDetailPageContent.bind(this);
        this.renderNameCell = this.renderNameCell.bind(this);
        this.renderLatestCommitCell = this.renderLatestCommitCell.bind(this);

        this.columns = [
            {
                id: 'name',
                name: 'Branch',
                onSize: this.onSize,
                renderCell: this.renderNameCell,
                width: nameColumnWidthObservable,
            },
            {
                id: 'commit',
                name: 'Latest Commit',
                onSize: this.onSize,
                renderCell: this.renderLatestCommitCell,
                width: latestCommitColumnWidthObservable,
            },
            {
                id: 'creator',
                name: 'Branch Creator',
                onSize: this.onSize,
                renderCell: this.renderBranchCreatorCell,
                width: branchCreatorColumnWidthObservable,
            },
        ];

        this.state = {
            repositories: new ArrayItemProvider<GitRepository>([]),
            repositoryBranches: new ArrayItemProvider<GitRef>([]),
            repositoryListSelection: new ListSelection({
                selectOnFocus: false,
            }),
            repositoryListSelectedItemObservable: new ObservableValue<any>({}),
        };

        this.dataManager = props.dataManager;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
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
            await this.loadRepositoriesDisplayState(filteredProjects);
        }
    }

    private async loadRepositoriesDisplayState(
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
            >
                {(observerProps: { selectedItem: GitRepository }) => (
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
                                        <Card className='bolt-table-card bolt-card-white'>
                                            <Table
                                                columns={this.columns}
                                                // behaviors={[
                                                //     this.sortingBehavior,
                                                // ]}
                                                itemProvider={
                                                    this.state
                                                        .repositoryBranches
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
        tableColumn: ITableColumn<GitRef>,
        tableItem: GitRef
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
                                tableItem.name.split('refs/heads/')[1],
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
        tableColumn: ITableColumn<GitRef>,
        tableItem: GitRef
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
                            tableItem.objectId
                        }
                        subtle={true}
                        target='_blank'
                        className='bolt-table-link bolt-table-inline-link'
                    >
                        {tableItem.objectId.substr(0, 8)}
                    </Link>
                }
            ></SimpleTableCell>
        );
    }

    private renderBranchCreatorCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<GitRef>,
        tableItem: GitRef
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
                                tableItem.creator._links['avatar']['href']
                            }
                        />
                        <div className='flex-column text-ellipsis'>
                            <Tooltip overflowOnly={true}>
                                <div className='primary-text text-ellipsis'>
                                    {tableItem.creator.displayName}
                                </div>
                            </Tooltip>
                            <Tooltip overflowOnly={true}>
                                <div className='primary-text text-ellipsis'>
                                    <Link
                                        excludeTabStop
                                        href={
                                            'mailto:' +
                                            tableItem.creator.uniqueName
                                        }
                                        subtle={false}
                                        target='_blank'
                                    >
                                        {tableItem.creator.uniqueName}
                                    </Link>
                                </div>
                            </Tooltip>
                        </div>
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
        this.setState({
            repositoryBranches: new ArrayItemProvider<GitRef>(
                Common.sortBranchesList(repositoryBranches)
            ),
        });
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
