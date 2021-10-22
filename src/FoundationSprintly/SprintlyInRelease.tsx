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
import { ReleaseDefinition } from 'azure-devops-extension-api/Release';
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
import { ISimpleTableCell } from 'azure-devops-ui/Table';
import { Observer } from 'azure-devops-ui/Observer';
import { Card } from 'azure-devops-ui/Card';
import {
    Tree,
    renderExpandableTreeCell,
    renderTreeCell,
} from 'azure-devops-ui/TreeEx';
import { Spinner } from 'azure-devops-ui/Spinner';

import * as Common from './SprintlyCommon';

export interface ISprintlyInReleaseState {
    releaseBranchDeployItemProvider: ITreeItemProvider<IReleaseBranchDeployTableItem>;
}

// TODO: Do I still need this?
export interface IReleaseBranchDeployTableItem {
    name: string;
    releaseInfo?: Common.IReleaseInfo[];
}

//#region "Observables"
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const releaseInfoObservable: ObservableArray<Common.IReleaseInfo> =
    new ObservableArray<Common.IReleaseInfo>();

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
                name: 'Repository',
                onSize: this.onSize,
                renderCell: renderExpandableTreeCell,
                width: nameColumnWidthObservable,
            },
            {
                id: 'deploy',
                name: 'Deploy Status',
                onSize: this.onSize,
                renderCell: renderTreeCell,
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

                console.log('build defs: ', this.buildDefinitions);
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
                        },
                        expanded: true,
                    };

                console.log('about to go into release loop');
                for (const releaseBranch of existingReleaseBranches) {
                    if (buildDefinitionForRepo) {
                        console.log('build def exists for repo');
                        await Common.fetchAndStoreBranchReleaseInfoIntoObservable(
                            releaseInfoObservable,
                            buildDefinitionForRepo,
                            this.releaseDefinitions,
                            releaseBranch,
                            repo.project.id,
                            repo.id,
                            this.organizationName,
                            this.accessToken
                        );

                        console.log('about to push into child items');
                    }
                    releaseBranchDeployTableItem.childItems!.push({
                        data: {
                            name: Common.getBranchShortName(
                                releaseBranch.targetBranch.name
                            ),
                            releaseInfo: releaseInfoObservable.value,
                        },
                        expanded: false,
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

    private getReleaseBranchDeployItemProvider(): ITreeItemProvider<IReleaseBranchDeployTableItem> {
        const rootItems: Array<ITreeItem<IReleaseBranchDeployTableItem>> = [];

        return new TreeItemProvider<IReleaseBranchDeployTableItem>(rootItems);
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (
            this.releaseBranchDeployTreeColumns[index]
                .width as ObservableValue<number>
        ).value = width;
    }

    public render(): JSX.Element {
        console.log(this.state.releaseBranchDeployItemProvider);
        return (
            <Observer
                totalRepositoriesToProcess={
                    totalRepositoriesToProcessObservable
                }
            >
                {(props: { totalRepositoriesToProcess: number }) => {
                    if (props.totalRepositoriesToProcess > 0) {
                        return (
                            <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                                {this.state.releaseBranchDeployItemProvider && (
                                    // const mostRecentRelease: Release | undefined =
                                    // Common.getMostRecentReleaseForBranch(
                                    //     item,
                                    //     observerProps.releaseInfoForAllBranches
                                    // );
                                    <Card>
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
                            </div>
                        );
                    }
                    return (
                        <ZeroData
                            primaryText='Coming Soon!'
                            secondaryText={
                                <span>
                                    In-release (QA) functionality is coming
                                    soon!
                                </span>
                            }
                            imageAltText='Coming Soon'
                            imagePath={'../static/notfound.png'}
                        />
                    );
                }}
            </Observer>
        );
    }
}
