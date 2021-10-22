import * as React from 'react';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    getClient,
    IExtensionDataManager,
    IGlobalMessagesService,
} from 'azure-devops-extension-api';
import { GitRepository } from 'azure-devops-extension-api/Git';
import { TeamProjectReference } from 'azure-devops-extension-api/Core';
import { Release, ReleaseDefinition } from 'azure-devops-extension-api/Release';
import { BuildDefinition } from 'azure-devops-extension-api/Build';

import {
    ArrayItemProvider,
    IItemProvider,
} from 'azure-devops-ui/Utilities/Provider';
import {
    ITreeItemProvider,
    ITreeItemEx,
    ITreeItem,
    TreeItemProvider,
} from 'azure-devops-ui/Utilities/TreeItemProvider';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import {
    IReadonlyObservableValue,
    ObservableArray,
    ObservableValue,
} from 'azure-devops-ui/Core/Observable';
import { ISimpleTableCell, SimpleTableCell } from 'azure-devops-ui/Table';
import { Observer } from 'azure-devops-ui/Observer';
import { Card } from 'azure-devops-ui/Card';
import {
    Tree,
    renderExpandableTreeCell,
    renderTreeCell,
    ITreeColumn,
} from 'azure-devops-ui/TreeEx';
import { Spinner } from 'azure-devops-ui/Spinner';

import * as Common from './SprintlyCommon';
import { Icon } from 'azure-devops-ui/Icon';
import { Link } from 'azure-devops-ui/Link';
import { Dialog } from 'azure-devops-ui/Dialog';

export interface ISprintlyInReleaseState {
    releaseBranchDeployItemProvider: ITreeItemProvider<IReleaseBranchDeployTableItem>;
}

export interface IReleaseBranchDeployTableItem {
    name: string;
    webUrl?: string;
    releaseInfo?: Common.IReleaseInfo;
    isRepositoryItem: boolean;
}

//#region "Observables"
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const allBranchesReleaseInfoObservable: ObservableArray<Common.IReleaseInfo> =
    new ObservableArray<Common.IReleaseInfo>();
const isDeployDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const clickedDeployEnvironmentObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const clickedDeployBranchNameObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const clickedDeployReleaseIdObservable: ObservableValue<number> =
    new ObservableValue<number>(0);

const nameColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const deployColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-80);
//#endregion "Observables"

const repositoriesToProcessKey: string = 'repositories-to-process';
let repositoriesToProcess: string[] = [];

export default class SprintlyInRelease extends React.Component<
    {
        organizationName: string;
        dataManager: IExtensionDataManager;
    },
    ISprintlyInReleaseState
> {
    private dataManager: IExtensionDataManager;
    private accessToken: string = '';
    private organizationName: string;

    private releaseBranchDeployTreeColumns: any = [];
    private releaseDefinitions: ReleaseDefinition[] = [];
    private buildDefinitions: BuildDefinition[] = [];

    constructor(props: {
        organizationName: string;
        dataManager: IExtensionDataManager;
    }) {
        super(props);

        this.onSize = this.onSize.bind(this);

        this.releaseBranchDeployTreeColumns = [
            {
                id: 'name',
                name: 'Repository Release Branches',
                onSize: this.onSize,
                renderCell: this.renderBranchColumn,
                width: nameColumnWidthObservable,
            },
            {
                id: 'deploy',
                name: 'Deploy Status',
                onSize: this.onSize,
                renderCell: this.renderDeployStatusColumn,
                width: deployColumnWidthObservable,
            },
        ];

        this.state = {
            releaseBranchDeployItemProvider:
                new TreeItemProvider<IReleaseBranchDeployTableItem>([]),
        };

        this.organizationName = props.organizationName;
        this.dataManager = props.dataManager;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        this.accessToken = await SDK.getAccessToken();
        repositoriesToProcess = (
            await Common.getSavedRepositoriesToProcess(
                this.dataManager,
                repositoriesToProcessKey
            )
        ).map((item: Common.IAllowedEntity) => item.originId);
        totalRepositoriesToProcessObservable.value =
            repositoriesToProcess.length;
        if (repositoriesToProcess.length > 0) {
            const filteredProjects: TeamProjectReference[] =
                await Common.getFilteredProjects();
            this.releaseDefinitions = await Common.getReleaseDefinitions(
                filteredProjects,
                this.organizationName,
                this.accessToken
            );
            this.buildDefinitions = await Common.getBuildDefinitions(
                filteredProjects,
                this.organizationName,
                this.accessToken
            );
            await this.loadRepositoriesDisplayState(filteredProjects);
        }
    }

    private async loadRepositoriesDisplayState(
        projects: TeamProjectReference[]
    ): Promise<void> {
        const reposExtended: Common.IGitRepositoryExtended[] = [];
        for (const project of projects) {
            const filteredRepos: GitRepository[] =
                await Common.getFilteredProjectRepositories(
                    project.id,
                    repositoriesToProcess
                );
            const releaseBranchRootItems: Array<
                ITreeItem<IReleaseBranchDeployTableItem>
            > = [];
            for (const repo of filteredRepos) {
                const repositoryBranchInfo: Common.IRepositoryBranchInfo =
                    await Common.getRepositoryBranchesInfo(repo.id);

                const buildDefinitionForRepo: BuildDefinition | undefined =
                    this.buildDefinitions.find(
                        (buildDef: BuildDefinition) =>
                            buildDef.repository.id === repo.id
                    );

                const existingReleaseBranches: Common.IReleaseBranchInfo[] =
                    repositoryBranchInfo.releaseBranches.map<Common.IReleaseBranchInfo>(
                        (releaseBranch) => {
                            return {
                                targetBranch: releaseBranch,
                                repositoryId: repo.id,
                            };
                        }
                    );

                const releaseBranchDeployTableItem: ITreeItem<IReleaseBranchDeployTableItem> =
                    {
                        childItems: [],
                        data: {
                            name: repo.name,
                            isRepositoryItem: true,
                        },
                        expanded: true,
                    };

                for (const releaseBranch of existingReleaseBranches) {
                    if (buildDefinitionForRepo) {
                        await Common.fetchAndStoreBranchReleaseInfoIntoObservable(
                            allBranchesReleaseInfoObservable,
                            buildDefinitionForRepo,
                            this.releaseDefinitions,
                            releaseBranch,
                            repo.project.id,
                            repo.id,
                            this.organizationName,
                            this.accessToken
                        );
                    }
                    releaseBranchDeployTableItem.childItems!.push({
                        data: {
                            name: Common.getBranchShortName(
                                releaseBranch.targetBranch.name
                            ),
                            webUrl: repo.webUrl,
                            releaseInfo:
                                allBranchesReleaseInfoObservable.value.find(
                                    (ri) =>
                                        ri.repositoryId === repo.id &&
                                        ri.releaseBranch.targetBranch.name ===
                                            releaseBranch.targetBranch.name
                                ),
                            isRepositoryItem: false,
                        },
                    });
                }

                releaseBranchRootItems.push(releaseBranchDeployTableItem);
            }

            this.setState({
                releaseBranchDeployItemProvider: new TreeItemProvider(
                    releaseBranchRootItems
                ),
            });

            totalRepositoriesToProcessObservable.value = filteredRepos.length;
        }
    }

    private renderBranchColumn(
        rowIndex: number,
        columnIndex: number,
        treeColumn: ITreeColumn<IReleaseBranchDeployTableItem>,
        treeItem: ITreeItemEx<IReleaseBranchDeployTableItem>
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={columnIndex}
                columnIndex={columnIndex}
                tableColumn={treeColumn}
                children={
                    <>
                        {treeItem.depth === 0 ? (
                            <>
                                <Icon
                                    iconName={
                                        treeItem.underlyingItem.expanded
                                            ? 'ChevronDownMed'
                                            : 'ChevronRightMed'
                                    }
                                    className='bolt-tree-expand-button cursor-pointer'
                                ></Icon>
                                <Icon
                                    iconName='FabricFolderFill'
                                    className='icon-margin'
                                    style={{ color: '#DCB67A' }}
                                ></Icon>
                                <span className='icon-margin'>
                                    {treeItem.underlyingItem.data.name}
                                </span>
                            </>
                        ) : (
                            <>
                                <Icon
                                    iconName='ChevronRightMed'
                                    className='invisible'
                                    style={{
                                        marginLeft: `${treeItem.depth * 16}px`,
                                    }}
                                ></Icon>
                                <Icon
                                    iconName='OpenSource'
                                    className='icon-margin'
                                ></Icon>
                                <u>
                                    {Common.branchLinkJsxElement(
                                        columnIndex.toString(),
                                        treeItem.underlyingItem.data.webUrl ??
                                            '#',
                                        treeItem.underlyingItem.data.name,
                                        'bolt-table-link bolt-table-inline-link'
                                    )}
                                </u>
                            </>
                        )}
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderDeployStatusColumn(
        rowIndex: number,
        columnIndex: number,
        treeColumn: ITreeColumn<IReleaseBranchDeployTableItem>,
        treeItem: ITreeItemEx<IReleaseBranchDeployTableItem>
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={columnIndex}
                columnIndex={columnIndex}
                tableColumn={treeColumn}
                children={
                    <>
                        <Observer
                            releaseInfoForAllBranches={
                                allBranchesReleaseInfoObservable
                            }
                        >
                            {(observerProps: {
                                releaseInfoForAllBranches: Common.IReleaseInfo[];
                            }) => {
                                if (
                                    treeItem.underlyingItem.data
                                        .isRepositoryItem
                                ) {
                                    return <></>;
                                }
                                const mostRecentRelease: Release | undefined =
                                    Common.getMostRecentReleaseForBranch(
                                        treeItem.underlyingItem.data.releaseInfo
                                            ?.releaseBranch,
                                        observerProps.releaseInfoForAllBranches
                                    );
                                if (!mostRecentRelease) {
                                    return Common.noReleaseExistsPillJsxElement();
                                }

                                const environmentStatuses: JSX.Element[] = [];
                                for (const environment of mostRecentRelease.environments) {
                                    environmentStatuses.push(
                                        Common.getSingleEnvironmentStatusPillJsxElement(
                                            environment,
                                            () => {
                                                clickedDeployEnvironmentObservable.value =
                                                    environment.name;
                                                clickedDeployBranchNameObservable.value =
                                                    treeItem.underlyingItem.data.name;
                                                clickedDeployReleaseIdObservable.value =
                                                    mostRecentRelease.id;
                                                isDeployDialogOpenObservable.value =
                                                    true;
                                            }
                                        )
                                    );
                                }
                                return environmentStatuses;
                            }}
                        </Observer>
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderDeployModalAction(): JSX.Element {
        return (
            <Observer isDeployDialogOpen={isDeployDialogOpenObservable}>
                {(props: { isDeployDialogOpen: boolean }) => {
                    return props.isDeployDialogOpen ? (
                        <Dialog
                            titleProps={{
                                text: `Deploy to ${clickedDeployEnvironmentObservable.value}?`,
                            }}
                            footerButtonProps={[
                                {
                                    text: 'Cancel',
                                    onClick: this.onDismissDeployActionModal,
                                },
                                {
                                    text: 'Refresh Data',
                                    iconProps: {
                                        iconName: 'Refresh',
                                    },
                                    onClick: () => {
                                        window.location.reload();
                                    },
                                },
                                {
                                    text: 'Deploy',
                                    onClick: this.onDismissDeployActionModal,
                                    primary: true,
                                },
                            ]}
                            onDismiss={this.onDismissDeployActionModal}
                        >
                            You are about to deploy release #
                            {clickedDeployReleaseIdObservable.value} for branch{' '}
                            {clickedDeployBranchNameObservable.value} to{' '}
                            {clickedDeployEnvironmentObservable.value}
                            . Are you sure?
                            <Icon ariaLabel='Warning' iconName='Warning' />{' '}
                            Note: Please ensure you have refreshed the data on
                            this page to avoid deploying a potentially obsolete
                            release.
                        </Dialog>
                    ) : null;
                }}
            </Observer>
        );
    }

    private onDismissDeployActionModal(): void {
        isDeployDialogOpenObservable.value = false;
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (
            this.releaseBranchDeployTreeColumns[index]
                .width as ObservableValue<number>
        ).value = width;
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
                        if (
                            this.state.releaseBranchDeployItemProvider
                                .length === 0
                        ) {
                            return (
                                <div className='page-content-top'>
                                    <Spinner label='loading' />
                                </div>
                            );
                        }
                        return (
                            <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                                {this.state.releaseBranchDeployItemProvider && (
                                    <Card
                                        className='bolt-table-card bolt-card-white'
                                        titleProps={{ text: 'Deploy Actions' }}
                                        headerDescriptionProps={{
                                            text: 'This table displays ONLY the most recent release artifact for each release branch. Click an environment to deploy to that environment.',
                                        }}
                                        headerCommandBarItems={[
                                            {
                                                id: 'export',
                                                text: 'Export to CSV',
                                                onActivate: () => {
                                                    alert('Example text');
                                                },
                                                iconProps: {
                                                    iconName: 'Download',
                                                },
                                            },
                                        ]}
                                    >
                                        <Tree<IReleaseBranchDeployTableItem>
                                            ariaLabel='Basic tree'
                                            columns={
                                                this
                                                    .releaseBranchDeployTreeColumns
                                            }
                                            itemProvider={
                                                this.state
                                                    .releaseBranchDeployItemProvider
                                            }
                                            onToggle={(
                                                event,
                                                treeItem: ITreeItemEx<IReleaseBranchDeployTableItem>
                                            ) => {
                                                this.state.releaseBranchDeployItemProvider.toggle(
                                                    treeItem.underlyingItem
                                                );
                                            }}
                                            scrollable={true}
                                        />
                                    </Card>
                                )}
                                {this.renderDeployModalAction()}
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
