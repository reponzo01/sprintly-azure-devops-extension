import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import {
    getClient,
    IColor,
    IExtensionDataManager,
} from 'azure-devops-extension-api';
import {
    CoreRestClient,
    TeamProjectReference,
} from 'azure-devops-extension-api/Core';
import {
    GitBaseVersionDescriptor,
    GitCommitDiffs,
    GitPullRequest,
    GitPullRequestSearchCriteria,
    GitRef,
    GitRepository,
    GitRestClient,
    GitTargetVersionDescriptor,
    PullRequestStatus,
} from 'azure-devops-extension-api/Git';

import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import { bindSelectionToObservable } from 'azure-devops-ui/MasterDetailsContext';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { ITableColumn, SimpleTableCell } from 'azure-devops-ui/Table';
import { Icon, IconSize } from 'azure-devops-ui/Icon';
import { Link } from 'azure-devops-ui/Link';
import { Button } from 'azure-devops-ui/Button';
import { Card } from 'azure-devops-ui/Card';
import {
    IListItemDetails,
    List,
    ListItem,
    ListSelection,
    SimpleList,
} from 'azure-devops-ui/List';
import {
    Splitter,
    SplitterDirection,
    SplitterElementPosition,
} from 'azure-devops-ui/Splitter';
import { Tooltip } from 'azure-devops-ui/TooltipEx';
import { Page } from 'azure-devops-ui/Page';

import {
    IAllowedEntity,
    IBranchAheadOf,
    IGitRepositoryExtended,
} from './FoundationSprintly';
import { Pill, PillSize, PillVariant } from 'azure-devops-ui/Pill';
import axios, { AxiosResponse } from 'axios';
import {
    CustomHeader,
    HeaderDescription,
    HeaderIcon,
    HeaderTitle,
    HeaderTitleArea,
    HeaderTitleRow,
    TitleSize,
} from 'azure-devops-ui/Header';
import { HeaderCommandBar } from 'azure-devops-ui/HeaderCommandBar';
import { Dialog } from 'azure-devops-ui/Dialog';

export interface ISprintlyPostReleaseState {
    repositories: ArrayItemProvider<IGitRepositoryExtended>;
    pullRequests: GitPullRequest[];
    selection: ListSelection;
    selectedItemObservable: ObservableValue<IGitRepositoryExtended>;
}

const isTagsDialogOpen: ObservableValue<boolean> = new ObservableValue<boolean>(
    false
);
const tagsRepoName: ObservableValue<string> = new ObservableValue<string>('');
const tags: ObservableValue<string[]> = new ObservableValue<string[]>([]);
const totalRepositoriesToProcess: ObservableValue<number> =
    new ObservableValue<number>(0);

const useFilteredRepos: boolean = true;
const repositoriesToProcessKey: string = 'repositories-to-process';
let repositoriesToProcess: string[] = [];

// TODO: Clean up arrow functions for the cases in which I thought I
// couldn't use regular functions because the this.* was undefined errors.
// The solution is to bind those functions to `this` in the constructor.
// See SprintlyPostRelease as an example.
export default class SprintlyPostRelease extends React.Component<
    { organizationName: string; dataManager: IExtensionDataManager },
    ISprintlyPostReleaseState
> {
    private dataManager: IExtensionDataManager;
    private accessToken: string = '';
    private organizationName: string;

    constructor(props: {
        organizationName: string;
        dataManager: IExtensionDataManager;
    }) {
        super(props);

        this.state = {
            repositories: new ArrayItemProvider<IGitRepositoryExtended>([]),
            pullRequests: [],
            selection: new ListSelection({ selectOnFocus: false }),
            selectedItemObservable: new ObservableValue<any>({}),
        };

        this.renderRepositoryList = this.renderRepositoryList.bind(this);
        this.renderRepositoryListItem = this.renderRepositoryListItem.bind(this);
        this.renderDetailPage = this.renderDetailPage.bind(this);

        this.organizationName = props.organizationName;
        this.dataManager = props.dataManager;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeSdk();
        await this.initializeComponent();
    }

    private async initializeSdk(): Promise<void> {
        await SDK.init();
        await SDK.ready();
    }

    private async initializeComponent(): Promise<void> {
        this.accessToken = await SDK.getAccessToken();

        await this.loadRepositoriesToProcess();
    }

    // TODO: This function is repeated in SprintlyPage. See about extracting.
    private async loadRepositoriesToProcess(): Promise<void> {
        this.dataManager!.getValue<IAllowedEntity[]>(repositoriesToProcessKey, {
            scopeType: 'User',
        }).then(async (repositories: IAllowedEntity[]) => {
            repositoriesToProcess = [];
            if (repositories) {
                for (const repository of repositories) {
                    repositoriesToProcess.push(repository.originId);
                }

                if (repositoriesToProcess.length > 0) {
                    const projects: TeamProjectReference[] = await getClient(
                        CoreRestClient
                    ).getProjects();

                    const filteredProjects: TeamProjectReference[] =
                        projects.filter((project: TeamProjectReference) => {
                            return (
                                project.name === 'Portfolio' ||
                                project.name === 'Sample Project'
                            );
                        });
                    await this.loadPullRequests(filteredProjects);
                    await this.loadRepositoriesDisplayState(filteredProjects);
                    await this.loadReleases(filteredProjects);

                    /**
                     * @param project - Project ID or project name
                     * @param taskGroupId -
                     * @param propertyFilters -
                     */
                    //getDefinitionEnvironments(project: string, taskGroupId?: string, propertyFilters?: string[]): Promise<Release.DefinitionEnvironmentReference[]>;
                    /**
                     * @param project - Project ID or project name
                     * @param definitionId -
                     * @param definitionEnvironmentId -
                     * @param createdBy -
                     * @param minModifiedTime -
                     * @param maxModifiedTime -
                     * @param deploymentStatus -
                     * @param operationStatus -
                     * @param latestAttemptsOnly -
                     * @param queryOrder -
                     * @param top -
                     * @param continuationToken -
                     * @param createdFor -
                     * @param minStartedTime -
                     * @param maxStartedTime -
                     * @param sourceBranch -
                     */
                    //getDeployments(project: string, definitionId?: number, definitionEnvironmentId?: number, createdBy?: string, minModifiedTime?: Date, maxModifiedTime?: Date, deploymentStatus?: Release.DeploymentStatus, operationStatus?: Release.DeploymentOperationStatus, latestAttemptsOnly?: boolean, queryOrder?: Release.ReleaseQueryOrder, top?: number, continuationToken?: number, createdFor?: string, minStartedTime?: Date, maxStartedTime?: Date, sourceBranch?: string): Promise<Release.Deployment[]>;
                    /**
                     * @param queryParameters -
                     * @param project - Project ID or project name
                     */
                    //getDeploymentsForMultipleEnvironments(queryParameters: Release.DeploymentQueryParameters, project: string): Promise<Release.Deployment[]>;
                    /**
                     * Get a release environment.
                     *
                     * @param project - Project ID or project name
                     * @param releaseId - Id of the release.
                     * @param environmentId - Id of the release environment.
                     */
                    //getReleaseEnvironment(project: string, releaseId: number, environmentId: number): Promise<Release.ReleaseEnvironment>;
                    /**
                     * @param project - Project ID or project name
                     * @param releaseId -
                     */
                    //getReleaseHistory(project: string, releaseId: number): Promise<Release.ReleaseRevision[]>;
                    /**
                     * Get a list of releases
                     *
                     * @param project - Project ID or project name
                     * @param definitionId - Releases from this release definition Id.
                     * @param definitionEnvironmentId -
                     * @param searchText - Releases with names containing searchText.
                     * @param createdBy - Releases created by this user.
                     * @param statusFilter - Releases that have this status.
                     * @param environmentStatusFilter -
                     * @param minCreatedTime - Releases that were created after this time.
                     * @param maxCreatedTime - Releases that were created before this time.
                     * @param queryOrder - Gets the results in the defined order of created date for releases. Default is descending.
                     * @param top - Number of releases to get. Default is 50.
                     * @param continuationToken - Gets the releases after the continuation token provided.
                     * @param expand - The property that should be expanded in the list of releases.
                     * @param artifactTypeId - Releases with given artifactTypeId will be returned. Values can be Build, Jenkins, GitHub, Nuget, Team Build (external), ExternalTFSBuild, Git, TFVC, ExternalTfsXamlBuild.
                     * @param sourceId - Unique identifier of the artifact used. e.g. For build it would be \{projectGuid\}:\{BuildDefinitionId\}, for Jenkins it would be \{JenkinsConnectionId\}:\{JenkinsDefinitionId\}, for TfsOnPrem it would be \{TfsOnPremConnectionId\}:\{ProjectName\}:\{TfsOnPremDefinitionId\}. For third-party artifacts e.g. TeamCity, BitBucket you may refer 'uniqueSourceIdentifier' inside vss-extension.json https://github.com/Microsoft/vsts-rm-extensions/blob/master/Extensions.
                     * @param artifactVersionId - Releases with given artifactVersionId will be returned. E.g. in case of Build artifactType, it is buildId.
                     * @param sourceBranchFilter - Releases with given sourceBranchFilter will be returned.
                     * @param isDeleted - Gets the soft deleted releases, if true.
                     * @param tagFilter - A comma-delimited list of tags. Only releases with these tags will be returned.
                     * @param propertyFilters - A comma-delimited list of extended properties to be retrieved. If set, the returned Releases will contain values for the specified property Ids (if they exist). If not set, properties will not be included. Note that this will not filter out any Release from results irrespective of whether it has property set or not.
                     * @param releaseIdFilter - A comma-delimited list of releases Ids. Only releases with these Ids will be returned.
                     * @param path - Releases under this folder path will be returned
                     */
                    //getReleases(project?: string, definitionId?: number, definitionEnvironmentId?: number, searchText?: string, createdBy?: string, statusFilter?: Release.ReleaseStatus, environmentStatusFilter?: number, minCreatedTime?: Date, maxCreatedTime?: Date, queryOrder?: Release.ReleaseQueryOrder, top?: number, continuationToken?: number, expand?: Release.ReleaseExpands, artifactTypeId?: string, sourceId?: string, artifactVersionId?: string, sourceBranchFilter?: string, isDeleted?: boolean, tagFilter?: string[], propertyFilters?: string[], releaseIdFilter?: number[], path?: string): Promise<Release.Release[]>;
                }
            }
        });
    }

    private async loadReleases(
        projects: TeamProjectReference[]
    ): Promise<void> {
        for (const project of projects) {
            // axios
            //     .get(
            //         `https://vsrm.dev.azure.com/${this.organizationName}/${project.id}/_apis/release/releases?api-version=6.0`,
            //         {
            //             headers: {
            //                 Authorization: `Bearer ${this.accessToken}`,
            //             },
            //         }
            //     )
            //     .then((res: AxiosResponse<never>) => {
            //         console.log('releases: ', res.data);
            //     })
            //     .catch((error: any) => {
            //         console.error(error);
            //         throw error;
            //     });
            // axios
            //     .get(
            //         `https://vsrm.dev.azure.com/${this.organizationName}/${project.id}/_apis/release/deployments?api-version=6.0`,
            //         {
            //             headers: {
            //                 Authorization: `Bearer ${this.accessToken}`,
            //             },
            //         }
            //     )
            //     .then((res: AxiosResponse<never>) => {
            //         console.log('deployments: ', res.data);
            //     })
            //     .catch((error: any) => {
            //         console.error(error);
            //         throw error;
            //     });
            // axios
            //     .get(
            //         `https://vsrm.dev.azure.com/${this.organizationName}/${project.id}/_apis/release/definitions?$expand=environments,artifacts&api-version=6.0`,
            //         {
            //             headers: {
            //                 Authorization: `Bearer ${this.accessToken}`,
            //             },
            //         }
            //     )
            //     .then((res: AxiosResponse<never>) => {
            //         console.log('definitions: ', res.data);
            //     })
            //     .catch((error: any) => {
            //         console.error(error);
            //         throw error;
            //     });
            // axios
            //     .get(
            //         `https://dev.azure.com/${this.organizationName}/${project.id}/_apis/pipelines?api-version=6.0-preview.1`,
            //         {
            //             headers: {
            //                 Authorization: `Bearer ${this.accessToken}`,
            //             },
            //         }
            //     )
            //     .then((res: AxiosResponse<never>) => {
            //         console.log('pipelines: ', res.data);
            //     })
            //     .catch((error: any) => {
            //         console.error(error);
            //         throw error;
            //     });
            //     axios
            //     .get(
            //         `https://dev.azure.com/${this.organizationName}/${project.id}/_apis/build/definitions?includeAllProperties=true&api-version=6.0`,
            //         {
            //             headers: {
            //                 Authorization: `Bearer ${this.accessToken}`,
            //             },
            //         }
            //     )
            //     .then((res: AxiosResponse<never>) => {
            //         console.log('builds: ', res.data);
            //     })
            //     .catch((error: any) => {
            //         console.error(error);
            //         throw error;
            //     });
        }
    }

    private async loadPullRequests(
        projects: TeamProjectReference[]
    ): Promise<void> {
        // Statuses:
        // 1 = Queued, 2 = Conflicts, 3 = Premerge Succeeded, 4 = RejectedByPolicy, 5 = Failure
        const pullRequestCriteria: GitPullRequestSearchCriteria = {
            includeLinks: false,
            creatorId: '',
            repositoryId: '',
            reviewerId: '',
            sourceRefName: '',
            sourceRepositoryId: '',
            status: PullRequestStatus.Active,
            targetRefName: '',
        };
        for (const project of projects) {
            const pullRequests: GitPullRequest[] = await getClient(
                GitRestClient
            ).getPullRequestsByProject(project.id, pullRequestCriteria);
            this.setState({
                pullRequests: this.state.pullRequests.concat(pullRequests),
            });
        }
    }

    // TODO: This function is repeated in SprintlyPage. See about extracting.
    private async loadRepositoriesDisplayState(
        projects: TeamProjectReference[]
    ): Promise<void> {
        let reposExtended: IGitRepositoryExtended[] = [];
        projects.forEach(async (project: TeamProjectReference) => {
            const repos: GitRepository[] = await getClient(
                GitRestClient
            ).getRepositories(project.id);
            console.log('repos: ', repos);
            let filteredRepos: GitRepository[] = repos;
            if (useFilteredRepos) {
                filteredRepos = repos.filter((repo: GitRepository) =>
                    repositoriesToProcess.includes(repo.id)
                );
            }

            totalRepositoriesToProcess.value = filteredRepos.length;

            for (const repo of filteredRepos) {
                const branchesAndTags: GitRef[] = await this.getRepositoryInfo(
                    repo.id
                );

                let hasDevelopBranch: boolean = false;
                let hasMasterBranch: boolean = false;
                let hasMainBranch: boolean = false;

                for (const ref of branchesAndTags) {
                    if (ref.name.includes('heads/develop')) {
                        hasDevelopBranch = true;
                    } else if (ref.name.includes('heads/master')) {
                        hasMasterBranch = true;
                    } else if (ref.name.includes('heads/main')) {
                        hasMainBranch = true;
                    }
                }

                const processRepo: boolean =
                    hasDevelopBranch && (hasMasterBranch || hasMainBranch);
                if (processRepo === true) {
                    //TODO: base = master/main, target = each release branch.
                    // base = develop, target = each release branch.
                    // if code changes, flag ahead of develop/main/master

                    const existingReleaseBranches: IBranchAheadOf[] = [];
                    let hasExistingRelease: boolean = false;
                    for (const branch of branchesAndTags) {
                        if (branch.name.includes('heads/release')) {
                            hasExistingRelease = true;

                            const branchName = branch.name.split('heads/')[1];

                            // TODO: maybe extract this for readibility
                            const masterMainBranchDescriptor: GitBaseVersionDescriptor =
                                {
                                    baseVersion: hasMasterBranch
                                        ? 'master'
                                        : 'main',
                                    baseVersionOptions: 0,
                                    baseVersionType: 0,
                                    version: hasMasterBranch
                                        ? 'master'
                                        : 'main',
                                    versionOptions: 0,
                                    versionType: 0,
                                };
                            const developBranchDescriptor: GitBaseVersionDescriptor =
                                {
                                    baseVersion: 'develop',
                                    baseVersionOptions: 0,
                                    baseVersionType: 0,
                                    version: 'develop',
                                    versionOptions: 0,
                                    versionType: 0,
                                };
                            const releaseBranchDescriptor: GitTargetVersionDescriptor =
                                {
                                    targetVersion: branchName,
                                    targetVersionOptions: 0,
                                    targetVersionType: 0,
                                    version: branchName,
                                    versionOptions: 0,
                                    versionType: 0,
                                };

                            const masterMainCommitsDiff: GitCommitDiffs =
                                await this.getCommitDiffs(
                                    repo.id,
                                    masterMainBranchDescriptor,
                                    releaseBranchDescriptor
                                );

                            const developCommitsDiff: GitCommitDiffs =
                                await this.getCommitDiffs(
                                    repo.id,
                                    developBranchDescriptor,
                                    releaseBranchDescriptor
                                );

                            const aheadOfMasterMain =
                                this.codeChangesInCommitDiffs(
                                    masterMainCommitsDiff
                                );
                            const aheadOfDevelop =
                                this.codeChangesInCommitDiffs(
                                    developCommitsDiff
                                );

                            const branchInfo: IBranchAheadOf = {
                                targetBranch: branch,
                                aheadOfDevelop,
                                aheadOfMasterMain,
                            };

                            for (const pullRequest of this.state.pullRequests) {
                                if (
                                    pullRequest.repository.id === repo.id &&
                                    pullRequest.sourceRefName === branch.name
                                ) {
                                    if (
                                        pullRequest.targetRefName.includes(
                                            'heads/develop'
                                        )
                                    ) {
                                        branchInfo.developPR = pullRequest;
                                    }

                                    if (
                                        pullRequest.targetRefName.includes(
                                            'heads/master'
                                        ) ||
                                        pullRequest.targetRefName.includes(
                                            'heads/main'
                                        )
                                    ) {
                                        branchInfo.masterMainPR = pullRequest;
                                    }
                                }
                            }

                            existingReleaseBranches.push(branchInfo);
                        }
                    }

                    reposExtended.push({
                        _links: repo._links,
                        defaultBranch: repo.defaultBranch,
                        id: repo.id,
                        isFork: repo.isFork,
                        name: repo.name,
                        parentRepository: repo.parentRepository,
                        project: repo.project,
                        remoteUrl: repo.remoteUrl,
                        size: repo.size,
                        sshUrl: repo.sshUrl,
                        url: repo.url,
                        validRemoteUrls: repo.validRemoteUrls,
                        webUrl: repo.webUrl,
                        createRelease: false,
                        hasExistingRelease,
                        hasMainBranch,
                        existingReleaseBranches,
                        branchesAndTags,
                    });
                }
            }

            if (reposExtended.length > 0) {
                reposExtended = reposExtended.sort(
                    (a: IGitRepositoryExtended, b: IGitRepositoryExtended) => {
                        return a.name.localeCompare(b.name);
                    }
                );
            }
            this.setState({
                repositories: new ArrayItemProvider(reposExtended),
            });

            bindSelectionToObservable(
                this.state.selection,
                this.state.repositories,
                this.state
                    .selectedItemObservable as ObservableValue<IGitRepositoryExtended>
            );
        });
    }

    // TODO: This function is repeated in SprintlyPage. See about extracting.
    private async getRepositoryInfo(repoId: string): Promise<GitRef[]> {
        return await getClient(GitRestClient).getRefs(
            repoId,
            undefined,
            undefined,
            false,
            true,
            undefined,
            true,
            true,
            undefined
        );
    }

    // TODO: This function is repeated in SprintlyPage. See about extracting.
    private async getCommitDiffs(
        repoId: string,
        baseVersion: GitBaseVersionDescriptor,
        targetVersion: GitTargetVersionDescriptor
    ): Promise<GitCommitDiffs> {
        return await getClient(GitRestClient).getCommitDiffs(
            repoId,
            undefined,
            undefined,
            1000,
            0,
            baseVersion,
            targetVersion
        );
    }

    // TODO: This function is repeated in SprintlyPage. See about extracting.
    private codeChangesInCommitDiffs(commitsDiff: GitCommitDiffs): boolean {
        return (
            Object.keys(commitsDiff.changeCounts).length > 0 ||
            commitsDiff.changes.length > 0
        );
    }

    private renderRepositoryList(): JSX.Element {
        return (
            <List
                ariaLabel={'Repositories'}
                itemProvider={this.state.repositories}
                selection={this.state.selection}
                renderRow={this.renderRepositoryListItem}
                width="100%"
                singleClickActivation={true}
            />
        );
    }

    private renderRepositoryListItem(
        index: number,
        item: IGitRepositoryExtended,
        details: IListItemDetails<IGitRepositoryExtended>,
        key?: string
    ): JSX.Element {
        const primaryColor: IColor = {
            red: 0,
            green: 120,
            blue: 114,
        };
        const primaryColorShade30: IColor = {
            red: 0,
            green: 69,
            blue: 120,
        };
        const releaseLinks: JSX.Element[] = [];
        let counter: number = 0;
        for (const releaseBranch of item.existingReleaseBranches) {
            const releaseBranchName =
                releaseBranch.targetBranch.name.split('heads/')[1];
            releaseLinks.push(
                <div className="flex-row padding-vertical-10" key={counter}>
                    <Link
                        excludeTabStop
                        href={
                            item.webUrl +
                            '?version=GB' +
                            encodeURI(releaseBranchName)
                        }
                        subtle={false}
                        target="_blank"
                        className="padding-horizontal-6"
                    >
                        {releaseBranchName}
                    </Link>
                    {releaseBranch.aheadOfDevelop && (
                        <Pill
                            color={primaryColor}
                            size={PillSize.regular}
                            className="bolt-list-overlay margin-horizontal-3"
                        >
                            <div style={{ color: 'white' }}>
                                Ahead of develop{' '}
                                {releaseBranch.developPR && (
                                    <i>
                                        <Icon
                                            ariaLabel="Pull Request"
                                            iconName="BranchPullRequest"
                                            size={IconSize.small}
                                        />{' '}
                                        #{releaseBranch.developPR.pullRequestId}
                                    </i>
                                )}
                            </div>
                        </Pill>
                    )}
                    {releaseBranch.aheadOfMasterMain && (
                        <Pill
                            color={primaryColorShade30}
                            size={PillSize.regular}
                            className="bolt-list-overlay margin-horizontal-3"
                            variant={PillVariant.outlined}
                        >
                            <div style={{ color: 'white' }}>
                                Ahead of{' '}
                                {item.hasMainBranch ? 'main' : 'master'}{' '}
                                {releaseBranch.masterMainPR && (
                                    <i>
                                        <Icon
                                            ariaLabel="Pull Request"
                                            iconName="BranchPullRequest"
                                            size={IconSize.small}
                                        />{' '}
                                        #
                                        {
                                            releaseBranch.masterMainPR
                                                .pullRequestId
                                        }
                                    </i>
                                )}
                            </div>
                        </Pill>
                    )}
                </div>
            );
            counter++;
        }

        return (
            <ListItem
                className="master-row border-bottom"
                key={key || 'list-item' + index}
                index={index}
                details={details}
            >
                <div className="master-row-content flex-row flex-center h-scroll-hidden">
                    <div className="flex-column text-ellipsis">
                        <Tooltip overflowOnly={true}>
                            <div className="primary-text text-ellipsis">
                                <Link
                                    excludeTabStop
                                    href={item.webUrl + '/branches'}
                                    subtle={true}
                                    target="_blank"
                                >
                                    <u>{item.name}</u>
                                </Link>
                            </div>
                        </Tooltip>
                        <Tooltip overflowOnly={true}>
                            <div className="flex-column primary-text text-ellipsis">
                                {<>{releaseLinks}</>}
                            </div>
                        </Tooltip>
                    </div>
                </div>
            </ListItem>
        );
    }

    private renderDetailPage(): JSX.Element {
        return (
            <Observer selectedItem={this.state.selectedItemObservable}>
                {(observerProps: { selectedItem: IGitRepositoryExtended }) => (
                    <Page className="flex-grow single-layer-details">
                        {this.state.selection.selectedCount == 0 && (
                            <span className="single-layer-details-contents">
                                Select a repository on the right to get started.
                            </span>
                        )}
                        {observerProps.selectedItem &&
                            this.state.selection.selectedCount > 0 && (
                                <Page>
                                    <CustomHeader className="bolt-header-with-commandbar">
                                        <HeaderIcon
                                            className="bolt-table-status-icon-large"
                                            iconProps={{
                                                iconName: 'Repo',
                                                size: IconSize.large,
                                            }}
                                            titleSize={TitleSize.Large}
                                        />
                                        <HeaderTitleArea>
                                            <HeaderTitleRow>
                                                <HeaderTitle
                                                    ariaLevel={3}
                                                    className="text-ellipsis"
                                                    titleSize={TitleSize.Large}
                                                >
                                                    <Link
                                                        excludeTabStop
                                                        href={
                                                            observerProps
                                                                .selectedItem
                                                                .webUrl +
                                                            '/branches'
                                                        }
                                                        subtle={false}
                                                        target="_blank"
                                                    >
                                                        {
                                                            observerProps
                                                                .selectedItem
                                                                .name
                                                        }
                                                    </Link>
                                                </HeaderTitle>
                                            </HeaderTitleRow>
                                            <HeaderDescription>
                                                Select a release branch below to
                                                perform actions on it.
                                            </HeaderDescription>
                                        </HeaderTitleArea>
                                        <HeaderCommandBar
                                            items={[
                                                {
                                                    iconProps: {
                                                        iconName: 'Tag',
                                                    },
                                                    id: 'testSave',
                                                    important: true,
                                                    onActivate: () => {
                                                        isTagsDialogOpen.value =
                                                            true;
                                                        tagsRepoName.value =
                                                            observerProps
                                                                .selectedItem
                                                                .name + ' Tags';
                                                        tags.value = [];
                                                        observerProps.selectedItem.branchesAndTags.forEach(
                                                            (branch) => {
                                                                if (
                                                                    branch.name.includes(
                                                                        'refs/tags'
                                                                    )
                                                                ) {
                                                                    tags.value.push(
                                                                        branch.name
                                                                    );
                                                                }
                                                            }
                                                        );
                                                    },
                                                    text: 'View Tags',
                                                },
                                            ]}
                                        />
                                    </CustomHeader>

                                    <div className="page-content page-content-top">
                                        <Card>Page content</Card>
                                    </div>
                                </Page>
                            )}
                    </Page>
                )}
            </Observer>
        );
    }

    public render(): JSX.Element {
        const onDismiss: () => void = () => {
            isTagsDialogOpen.value = false;
        };
        return (
            /* tslint:disable */
            <Observer totalRepositoriesToProcess={totalRepositoriesToProcess}>
                {(props: { totalRepositoriesToProcess: number }) => {
                    if (totalRepositoriesToProcess.value > 0) {
                        return (
                            <div
                                style={{
                                    height: '85%',
                                    width: '100%',
                                    display: 'flex',
                                }}
                            >
                                <Splitter
                                    fixedElement={SplitterElementPosition.Near}
                                    splitterDirection={
                                        SplitterDirection.Vertical
                                    }
                                    initialFixedSize={450}
                                    minFixedSize={100}
                                    nearElementClassName="v-scroll-auto custom-scrollbar light-grey"
                                    farElementClassName="v-scroll-auto custom-scrollbar"
                                    onRenderNearElement={
                                        this.renderRepositoryList
                                    }
                                    onRenderFarElement={this.renderDetailPage}
                                />
                                <Observer
                                    isTagsDialogOpen={isTagsDialogOpen}
                                    tagsRepoName={tagsRepoName}
                                >
                                    {(props: {
                                        isTagsDialogOpen: boolean;
                                        tagsRepoName: string;
                                    }) => {
                                        return props.isTagsDialogOpen ? (
                                            <Dialog
                                                titleProps={{
                                                    text: props.tagsRepoName,
                                                }}
                                                footerButtonProps={[
                                                    {
                                                        text: 'Close',
                                                        onClick: onDismiss,
                                                    },
                                                ]}
                                                onDismiss={onDismiss}
                                            >
                                                <SimpleList
                                                    itemProvider={
                                                        new ArrayItemProvider<string>(
                                                            tags.value
                                                        )
                                                    }
                                                />
                                            </Dialog>
                                        ) : null;
                                    }}
                                </Observer>
                            </div>
                        );
                    }
                    return (
                        <ZeroData
                            primaryText="No repositories."
                            secondaryText={
                                <span>
                                    Please select valid repositories from the
                                    Settings page.
                                </span>
                            }
                            imageAltText="No repositories."
                            imagePath={'../static/notfound.png'}
                        />
                    );
                }}
            </Observer>

            /* tslint:disable */
        );
    }
}

// TODO: May be able to remove this function here
// TODO: This function is repeated in SprintlyPage. See about extracting.
function renderName(
    rowIndex: number,
    columnIndex: number,
    tableColumn: ITableColumn<IGitRepositoryExtended>,
    tableItem: IGitRepositoryExtended
): JSX.Element {
    return (
        <SimpleTableCell
            key={'col-' + columnIndex}
            columnIndex={columnIndex}
            tableColumn={tableColumn}
            children={
                <>
                    <Icon ariaLabel="Repository" iconName="Repo" />
                    &nbsp;
                    <Link
                        excludeTabStop
                        href={tableItem.webUrl + '/branches'}
                        subtle={true}
                        target="_blank"
                    >
                        <u>{tableItem.name}</u>
                    </Link>
                </>
            }
        ></SimpleTableCell>
    );
}

// TODO: This function is repeated in SprintlyPage. See about extracting.
function renderTags(
    rowIndex: number,
    columnIndex: number,
    tableColumn: ITableColumn<IGitRepositoryExtended>,
    tableItem: IGitRepositoryExtended
): JSX.Element {
    return (
        <SimpleTableCell
            key={'col-' + columnIndex}
            columnIndex={columnIndex}
            tableColumn={tableColumn}
            children={
                <>
                    <Button
                        text="View Tags"
                        subtle={true}
                        iconProps={{ iconName: 'Tag' }}
                        onClick={() => {
                            isTagsDialogOpen.value = true;
                            tagsRepoName.value = tableItem.name + ' Tags';
                            tags.value = [];
                            tableItem.branchesAndTags.forEach((branch) => {
                                if (branch.name.includes('refs/tags')) {
                                    tags.value.push(branch.name);
                                }
                            });
                        }}
                    />
                </>
            }
        ></SimpleTableCell>
    );
}

// The following code would go on the onclick of a merge button
/*
const createRefOptions: GitRefUpdate[] = [];
const developBranch = await getClient(
    GitRestClient
).getBranch(tableItem.id, 'develop');

// new test code
const mainBranch = await getClient(
    GitRestClient
).getBranch(tableItem.id, 'main');

console.log;

//TODO: Try this page: https://docs.microsoft.com/en-us/rest/api/azure/devops/git/merges/create?view=azure-devops-rest-6.0 And try using regular axios instead of the api.

const newMainObjectId = mainBranch.commit.commitId;
const newDevObjectId =
    developBranch.commit.commitId;
console.log(mainBranch);
const gitMergeParams: GitMergeParameters = {
    comment: 'Merging dev to main hopefully',
    parents: [newMainObjectId, newDevObjectId],
};
//POST https://dev.azure.com/{organization}/{project}/_apis/git/repositories/{repositoryNameOrId}/merges?api-version=6.0-preview.1

const mergeRequest: GitMerge = await getClient(
    GitRestClient
).createMergeRequest(
    gitMergeParams,
    tableItem.project.id,
    tableItem.id
);
console.log(mergeRequest);

let mergeCommitId = '';
const mergeCheckInterval = setInterval(async () => {
    const mergeRequestStatus: GitMerge =
        await getClient(
            GitRestClient
        ).getMergeRequest(
            tableItem.project.id,
            tableItem.id,
            mergeRequest.mergeOperationId
        );
    console.log(mergeRequestStatus);
    // TODO: check for other errors (detailedStatus has failure message)
    if (
        mergeRequestStatus.status ===
        GitAsyncOperationStatus.Completed
    ) {
        clearInterval(mergeCheckInterval);
        mergeCommitId =
            mergeRequestStatus.detailedStatus
                .mergeCommitId;

        // TODO: This is ugly. this is inside a set interval
        createRefOptions.push({
            repositoryId: tableItem.id,
            name: 'refs/heads/main',
            isLocked: false,
            newObjectId: mergeCommitId,
            oldObjectId: newMainObjectId,
        });
        const createRef = await getClient(
            GitRestClient
        ).updateRefs(
            createRefOptions,
            tableItem.id
        );
    }
}, 500);
// This is async. Need a callback above.
console.log(
    'outside the interval, merge commit id: ',
    mergeCommitId
);
*/
