import * as React from 'react';

import * as SDK from 'azure-devops-extension-sdk';
import {
    IExtensionDataManager,
    IGlobalMessagesService,
    IProjectInfo,
    MessageBannerLevel,
} from 'azure-devops-extension-api';
import { GitRef, GitRepository } from 'azure-devops-extension-api/Git';
import {
    EnvironmentStatus,
    ProjectReference,
    Release,
    ReleaseDefinition,
    ReleaseEnvironment,
} from 'azure-devops-extension-api/Release';
import { BuildDefinition } from 'azure-devops-extension-api/Build';

import {
    ITreeItemProvider,
    ITreeItemEx,
    ITreeItem,
    TreeItemProvider,
} from 'azure-devops-ui/Utilities/TreeItemProvider';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import {
    ObservableArray,
    ObservableValue,
} from 'azure-devops-ui/Core/Observable';
import { SimpleTableCell } from 'azure-devops-ui/Table';
import { Observer } from 'azure-devops-ui/Observer';
import { Card } from 'azure-devops-ui/Card';
import { Tree, ITreeColumn } from 'azure-devops-ui/TreeEx';
import { Spinner } from 'azure-devops-ui/Spinner';

import * as Common from './SprintlyCommon';
import { Icon } from 'azure-devops-ui/Icon';
import { Dialog } from 'azure-devops-ui/Dialog';
import axios, { AxiosResponse } from 'axios';

export interface ISprintlyInReleaseState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    releaseBranchDeployItemProvider: ITreeItemProvider<IReleaseBranchDeployTableItem>;
    allBranchesReleaseInfo: Common.IReleaseInfo[];
    clickedDeployEnvironmentStatus?: EnvironmentStatus;
    ready: boolean;
}

export interface IReleaseBranchDeployTableItem {
    name: string;
    id: string;
    webUrl?: string;
    releaseInfo?: Common.IReleaseInfo;
    projectId?: string;
    isRepositoryItem: boolean;
}

//#region "Observables"
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const allBranchesReleaseInfoObservable: ObservableArray<Common.IReleaseInfo> =
    new ObservableArray<Common.IReleaseInfo>();
const isDeployDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const clickedDeployEnvironmentObservable: ObservableValue<ReleaseEnvironment> =
    new ObservableValue<any>({});
const clickedDeployBranchNameObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const clickedDeployReleaseIdObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const clickedDeployProjectReferenceObservable: ObservableValue<ProjectReference> =
    new ObservableValue<any>({});
//#endregion "Observables"

let repositoriesToProcess: string[] = [];

export default class SprintlyInRelease extends React.Component<
    {
        organizationName: string;
        globalMessagesSvc: IGlobalMessagesService;
        dataManager: IExtensionDataManager;
        releaseDefinitions: ReleaseDefinition[];
        buildDefinitions: BuildDefinition[];
    },
    ISprintlyInReleaseState
> {
    private dataManager: IExtensionDataManager;
    private globalMessagesSvc: IGlobalMessagesService;
    private accessToken: string = '';
    private organizationName: string;

    private releaseBranchDeployTreeColumns: any = [];
    private releaseDefinitions: ReleaseDefinition[];
    private buildDefinitions: BuildDefinition[];

    constructor(props: {
        organizationName: string;
        globalMessagesSvc: IGlobalMessagesService;
        dataManager: IExtensionDataManager;
        releaseDefinitions: ReleaseDefinition[];
        buildDefinitions: BuildDefinition[];
    }) {
        super(props);

        this.onSize = this.onSize.bind(this);
        this.deployAction = this.deployAction.bind(this);
        this.renderBranchColumn = this.renderBranchColumn.bind(this);
        this.renderDeployStatusColumn =
            this.renderDeployStatusColumn.bind(this);
        this.exportToCsvAction = this.exportToCsvAction.bind(this);

        this.releaseBranchDeployTreeColumns = [
            {
                id: 'name',
                name: 'Repository Release Branches',
                onSize: this.onSize,
                renderCell: this.renderBranchColumn,
                width: new ObservableValue<number>(-30),
            },
            {
                id: 'deploy',
                name: 'Deploy Status',
                onSize: this.onSize,
                renderCell: this.renderDeployStatusColumn,
                width: new ObservableValue<number>(-80),
            },
        ];

        this.state = {
            releaseBranchDeployItemProvider:
                new TreeItemProvider<IReleaseBranchDeployTableItem>([]),
            allBranchesReleaseInfo: [],
            ready: true,
        };

        this.organizationName = props.organizationName;
        this.globalMessagesSvc = props.globalMessagesSvc;
        this.dataManager = props.dataManager;
        this.releaseDefinitions = props.releaseDefinitions;
        this.buildDefinitions = props.buildDefinitions;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        this.accessToken = await SDK.getAccessToken();

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

    private async reloadComponent(silentRefresh: boolean): Promise<void> {
        this.setState({
            ready: silentRefresh,
        });
        await this.initializeComponent();
        this.setState({
            ready: true,
        });
        this.setState(this.state);
    }

    private async loadRepositoriesDisplayState(
        currentProject: IProjectInfo | undefined
    ): Promise<void> {
        const reposExtended: Common.IGitRepositoryExtended[] = [];
        if (currentProject !== undefined) {
            const filteredRepos: GitRepository[] =
                await Common.getFilteredProjectRepositories(
                    currentProject.id,
                    repositoriesToProcess
                );
            const releaseBranchRootItems: Array<
                ITreeItem<IReleaseBranchDeployTableItem>
            > = [];
            for (const repo of filteredRepos) {
                const repositoryBranchInfo: Common.IRepositoryBranchInfo =
                    await Common.getRepositoryBranchesInfo(
                        repo.id,
                        Common.repositoryHeadsFilter
                    );

                const existingReleaseBranches: Common.IReleaseBranchInfo[] =
                    repositoryBranchInfo.releaseBranches.map<Common.IReleaseBranchInfo>(
                        (releaseBranch: GitRef) => {
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
                            id: repo.id,
                            projectId: repo.project.id,
                            isRepositoryItem: true,
                        },
                        expanded: true,
                    };

                for (const releaseBranch of existingReleaseBranches) {
                    await Common.fetchAndStoreBranchReleaseInfoIntoObservable(
                        allBranchesReleaseInfoObservable,
                        this.buildDefinitions,
                        this.releaseDefinitions,
                        releaseBranch,
                        repo.project.id,
                        repo.id,
                        this.organizationName,
                        this.accessToken
                    );

                    releaseBranchDeployTableItem.childItems!.push({
                        data: {
                            name: Common.getBranchShortName(
                                releaseBranch.targetBranch.name
                            ),
                            id: releaseBranch.targetBranch.objectId,
                            webUrl: repo.webUrl,
                            releaseInfo:
                                allBranchesReleaseInfoObservable.value.find(
                                    (ri: Common.IReleaseInfo) =>
                                        ri.repositoryId === repo.id &&
                                        ri.releaseBranch.targetBranch.name ===
                                            releaseBranch.targetBranch.name
                                ),
                            projectId: repo.project.id,
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
                allBranchesReleaseInfo: allBranchesReleaseInfoObservable.value,
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
        let releaseUrl: string = `https://dev.azure.com/${this.organizationName}/${treeItem.underlyingItem.data.projectId}/_releaseProgress?_a=release-pipeline-progress&releaseId=`;

        if (!treeItem.underlyingItem.data.isRepositoryItem) {
            if (treeItem.underlyingItem.data.releaseInfo) {
                const mostRecentRelease: Release | undefined =
                    Common.getMostRecentReleaseForBranch(
                        {
                            repositoryId:
                                treeItem.underlyingItem.data.releaseInfo
                                    .releaseBranch.repositoryId,
                            targetBranch:
                                treeItem.underlyingItem.data.releaseInfo
                                    .releaseBranch.targetBranch,
                        },
                        allBranchesReleaseInfoObservable.value
                    );
                if (mostRecentRelease) {
                    releaseUrl += mostRecentRelease.id;
                }
            }
        }
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
                                    {!treeItem.underlyingItem.data
                                        .releaseInfo ? (
                                        <span className='bolt-table-link bolt-table-inline-link'>
                                            {treeItem.underlyingItem.data.name}
                                        </span>
                                    ) : (
                                        Common.branchLinkJsxElement(
                                            columnIndex.toString(),
                                            releaseUrl,
                                            treeItem.underlyingItem.data.name,
                                            'bolt-table-link bolt-table-inline-link',
                                            true
                                        )
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
                                    if (
                                        !environment.name
                                            .toLowerCase()
                                            .includes('feature')
                                    ) {
                                        environmentStatuses.push(
                                            Common.getSingleEnvironmentStatusPillJsxElement(
                                                environment,
                                                () => {
                                                    clickedDeployEnvironmentObservable.value =
                                                        environment;
                                                    clickedDeployBranchNameObservable.value =
                                                        treeItem.underlyingItem.data.name;
                                                    clickedDeployReleaseIdObservable.value =
                                                        mostRecentRelease.id;
                                                    clickedDeployProjectReferenceObservable.value =
                                                        {
                                                            id: mostRecentRelease
                                                                .projectReference
                                                                .id,
                                                            name: mostRecentRelease
                                                                .projectReference
                                                                .name,
                                                        };
                                                    this.setState({
                                                        clickedDeployEnvironmentStatus:
                                                            environment.status,
                                                    });
                                                    isDeployDialogOpenObservable.value =
                                                        true;
                                                }
                                            )
                                        );
                                    }
                                }
                                return environmentStatuses;
                            }}
                        </Observer>
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderDeployConfirmAction(): JSX.Element {
        return (
            <Observer isDeployDialogOpen={isDeployDialogOpenObservable}>
                {(props: { isDeployDialogOpen: boolean }) => {
                    return props.isDeployDialogOpen ? (
                        this.state.clickedDeployEnvironmentStatus ===
                            EnvironmentStatus.InProgress ||
                            this.state.clickedDeployEnvironmentStatus ===
                            EnvironmentStatus.Queued ? (
                            <Dialog
                                titleProps={{
                                    text: `Release in progress!`,
                                }}
                                footerButtonProps={[
                                    {
                                        text: 'Close',
                                        onClick:
                                            this.onDismissDeployActionModal,
                                    },
                                    {
                                        text: 'Refresh Data',
                                        iconProps: {
                                            iconName: 'Refresh',
                                        },
                                        onClick: async () => {
                                            isDeployDialogOpenObservable.value =
                                                false;
                                            await this.reloadComponent(false);
                                        },
                                    },
                                    {
                                        text: 'View Logs',
                                        href: `https://dev.azure.com/${this.organizationName}/${clickedDeployProjectReferenceObservable.value.id}/_releaseProgress?_a=release-environment-logs&releaseId=${clickedDeployReleaseIdObservable.value}&environmentId=${clickedDeployEnvironmentObservable.value.id}`,
                                        onClick: () =>
                                            (isDeployDialogOpenObservable.value =
                                                false),
                                        target: '_blank',
                                        primary: true,
                                    },
                                ]}
                                onDismiss={this.onDismissDeployActionModal}
                            >
                                <Icon
                                    ariaLabel='Warning'
                                    iconName='Warning'
                                    style={{ color: 'orange' }}
                                />{' '}
                                Note: A release is in progress but you can
                                reload the data to get its latest status.
                            </Dialog>
                        ) : (
                            <Dialog
                                titleProps={{
                                    text: `${
                                        this.state.clickedDeployEnvironmentStatus ===
                                        EnvironmentStatus.Succeeded
                                            ? 'Redeploy'
                                            : 'Deploy'
                                    } to ${
                                        clickedDeployEnvironmentObservable.value
                                            .name
                                    }?`,
                                }}
                                footerButtonProps={[
                                    {
                                        text: 'Cancel',
                                        onClick:
                                            this.onDismissDeployActionModal,
                                    },
                                    {
                                        text: 'Refresh Data',
                                        iconProps: {
                                            iconName: 'Refresh',
                                        },
                                        onClick: async () => {
                                            isDeployDialogOpenObservable.value =
                                                false;
                                            await this.reloadComponent(false);
                                        },
                                    },
                                    {
                                        text:
                                        this.state.clickedDeployEnvironmentStatus ===
                                            EnvironmentStatus.Succeeded
                                                ? 'Redeploy'
                                                : 'Deploy',
                                        onClick: this.deployAction,
                                        primary: true,
                                    },
                                ]}
                                onDismiss={this.onDismissDeployActionModal}
                            >
                                You are about to{' '}
                                {this.state.clickedDeployEnvironmentStatus ===
                                EnvironmentStatus.Succeeded
                                    ? 'redeploy'
                                    : 'deploy'}{' '}
                                release #
                                {clickedDeployReleaseIdObservable.value} for
                                branch {clickedDeployBranchNameObservable.value}{' '}
                                to{' '}
                                {clickedDeployEnvironmentObservable.value.name}
                                . Are you sure?
                                <Icon
                                    ariaLabel='Warning'
                                    iconName='Warning'
                                    style={{ color: 'orange' }}
                                />{' '}
                                Note: Please ensure you have refreshed the data
                                on this page to avoid deploying a potentially
                                obsolete release.
                            </Dialog>
                        )
                    ) : null;
                }}
            </Observer>
        );
    }

    private deployAction(): void {
        const url: string = `https://vsrm.dev.azure.com/${this.organizationName}/${clickedDeployProjectReferenceObservable.value.id}/_apis/Release/releases/${clickedDeployReleaseIdObservable.value}/environments/${clickedDeployEnvironmentObservable.value.id}?api-version=5.0-preview.6`;

        Common.getOrRefreshToken(this.accessToken).then(
            async (token: string) => {
                await axios
                    .patch(
                        url,
                        {
                            comment: '',
                            status: EnvironmentStatus.InProgress,
                        },
                        {
                            headers: {
                                Authorization: `Bearer ${token}`,
                            },
                        }
                    )
                    .then(async (result: void | AxiosResponse<any>) => {
                        this.globalMessagesSvc.addToast({
                            duration: 5000,
                            forceOverrideExisting: true,
                            message: 'Deploy started!',
                        });
                        await this.reloadComponent(true);
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
                                    'Deploy request failed!' +
                                    error +
                                    ' ' +
                                    error,
                            });
                        }
                    });
            }
        );

        isDeployDialogOpenObservable.value = false;
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

    private exportToCsvAction(): void {
        let data: string = '';
        const now: Date = new Date();
        data += `Release Candidate builds as of ${now.toString()}\r\n`;
        data += 'Repository,Version,Release Artifact\r\n';

        let projectId: string = '';
        let repoName: string = '';
        let version: string = '';
        let releaseArtifact: string = '';

        for (const repo of repositoriesToProcess) {
            for (const repoInfo of this.state.releaseBranchDeployItemProvider
                .value) {
                if (repoInfo.underlyingItem.data.id === repo) {
                    repoName = repoInfo.underlyingItem.data.name;
                    projectId = repoInfo.underlyingItem.data.projectId!;
                }
            }
            const repoReleaseBranchesInfo: Common.IReleaseInfo[] = [];
            for (const relInfo of this.state.allBranchesReleaseInfo) {
                if (relInfo.repositoryId === repo) {
                    repoReleaseBranchesInfo.push(relInfo);
                }
            }
            if (repoReleaseBranchesInfo && repoReleaseBranchesInfo.length > 0) {
                for (const releaseBranchInfo of repoReleaseBranchesInfo) {
                    version = Common.getBranchShortName(
                        releaseBranchInfo.releaseBranch.targetBranch.name
                    );
                    const mostRecentRelease: Release | undefined =
                        Common.getMostRecentReleaseForBranch(
                            releaseBranchInfo.releaseBranch,
                            this.state.allBranchesReleaseInfo
                        );
                    if (mostRecentRelease) {
                        releaseArtifact = `https://dev.azure.com/${this.organizationName}/${projectId}/_releaseProgress?_a=release-pipeline-progress&releaseId=${mostRecentRelease.id}`;
                    }

                    data += `${repoName},${version},${releaseArtifact}\r\n`;
                }
            }
        }

        const hiddenElement: HTMLAnchorElement = document.createElement('a');
        hiddenElement.setAttribute(
            'href',
            'data:text/csv;base64,' + window.btoa(data)
        );
        hiddenElement.target = '_blank';
        hiddenElement.download = `Release_Candidates_as_of_${now.getMonth()}-${now.getDate()}-${now.getFullYear()}.csv`;
        hiddenElement.click();
    }

    public render(): JSX.Element {
        return this.state.ready ? (
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
                                                    this.exportToCsvAction();
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
                                                event: React.SyntheticEvent<
                                                    HTMLElement,
                                                    Event
                                                >,
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
                                {this.renderDeployConfirmAction()}
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
        ) : (
            <div className='page-content-top'>
                <Spinner label='loading' />
            </div>
        );
    }
}
