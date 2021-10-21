import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IColor, IExtensionDataManager } from 'azure-devops-extension-api';
import { TeamProjectReference } from 'azure-devops-extension-api/Core';
import {
    GitBaseVersionDescriptor,
    GitCommitDiffs,
    GitPullRequest,
    GitRef,
    GitRepository,
    GitTargetVersionDescriptor,
} from 'azure-devops-extension-api/Git';
import {
    EnvironmentStatus,
    Release,
    ReleaseDefinition,
} from 'azure-devops-extension-api/Release';
import { BuildDefinition } from 'azure-devops-extension-api/Build';

import {
    ObservableArray,
    ObservableValue,
} from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import { bindSelectionToObservable } from 'azure-devops-ui/MasterDetailsContext';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { Icon, IconSize } from 'azure-devops-ui/Icon';
import { Link } from 'azure-devops-ui/Link';
import { Card } from 'azure-devops-ui/Card';
import { Status, Statuses, StatusSize } from 'azure-devops-ui/Status';
import {
    IListItemDetails,
    List,
    ListItem,
    ListSelection,
} from 'azure-devops-ui/List';
import {
    Splitter,
    SplitterDirection,
    SplitterElementPosition,
} from 'azure-devops-ui/Splitter';
import { Tooltip } from 'azure-devops-ui/TooltipEx';
import { Page } from 'azure-devops-ui/Page';
import { Pill, PillSize, PillVariant } from 'azure-devops-ui/Pill';
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

import * as Common from './SprintlyCommon';
import { TagsModal, ITagsModalContent, getTagsModalContent } from './TagsModal';
import { Spinner } from 'azure-devops-ui/Spinner';

// TODO: Instead of a state, consider just global observables
export interface ISprintlyPostReleaseState {
    repositories: ArrayItemProvider<Common.IGitRepositoryExtended>;
    pullRequests: GitPullRequest[];
    repositoryListSelection: ListSelection;
    releaseBranchListSelection: ListSelection;
    repositoryListSelectedItemObservable: ObservableValue<Common.IGitRepositoryExtended>;
    releaseBranchListSelectedItemObservable: ObservableValue<Common.IReleaseBranchInfo>;
}

const tagsModalKeyObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const isTagsDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const tagsRepoNameObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const tagsObservable: ObservableValue<string[]> = new ObservableValue<string[]>(
    []
);
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const releaseInfoObservable: ObservableArray<Common.IReleaseInfo> =
    new ObservableArray<Common.IReleaseInfo>();

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

    private releaseDefinitions: ReleaseDefinition[] = [];
    private buildDefinitions: BuildDefinition[] = [];

    constructor(props: {
        organizationName: string;
        dataManager: IExtensionDataManager;
    }) {
        super(props);

        this.state = {
            repositories: new ArrayItemProvider<Common.IGitRepositoryExtended>(
                []
            ),
            pullRequests: [],
            repositoryListSelection: new ListSelection({
                selectOnFocus: false,
            }),
            repositoryListSelectedItemObservable: new ObservableValue<any>({}),
            releaseBranchListSelection: new ListSelection({
                selectOnFocus: false,
            }),
            releaseBranchListSelectedItemObservable: new ObservableValue<any>(
                {}
            ),
        };

        this.renderRepositoryMasterPageList =
            this.renderRepositoryMasterPageList.bind(this);
        this.renderRepositoryListItem =
            this.renderRepositoryListItem.bind(this);
        this.renderReleaseBranchDetailList =
            this.renderReleaseBranchDetailList.bind(this);
        this.renderReleaseBranchDetailListItem =
            this.renderReleaseBranchDetailListItem.bind(this);
        this.renderDetailPageContent = this.renderDetailPageContent.bind(this);

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
        ).map((item) => item.originId);
        totalRepositoriesToProcessObservable.value =
            repositoriesToProcess.length;
        if (repositoriesToProcess.length > 0) {
            const filteredProjects = await Common.getFilteredProjects();
            await this.loadRepositoriesDisplayState(filteredProjects);
            this.setState({
                pullRequests: this.state.pullRequests.concat(
                    await Common.getPullRequests(filteredProjects)
                ),
            });
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
        }
    }

    // TODO: This function is repeated in SprintlyPage. See about extracting.
    private async loadRepositoriesDisplayState(
        projects: TeamProjectReference[]
    ): Promise<void> {
        let reposExtended: Common.IGitRepositoryExtended[] = [];
        for (const project of projects) {
            const filteredRepos: GitRepository[] =
                await Common.getFilteredProjectRepositories(
                    project.id,
                    repositoriesToProcess
                );

            totalRepositoriesToProcessObservable.value = filteredRepos.length;

            for (const repo of filteredRepos) {
                const repositoryBranchInfo =
                    await Common.getRepositoryBranchesInfo(repo.id);

                const processRepo: boolean =
                    repositoryBranchInfo.hasDevelopBranch &&
                    (repositoryBranchInfo.hasMasterBranch ||
                        repositoryBranchInfo.hasMainBranch);

                if (processRepo === true) {
                    const existingReleaseBranches: Common.IReleaseBranchInfo[] =
                        [];
                    for (const releaseBranch of repositoryBranchInfo.releaseBranches) {
                        const releaseBranchName =
                            releaseBranch.name.split('heads/')[1];

                        const branchInfo: Common.IReleaseBranchInfo = {
                            targetBranch: releaseBranch,
                            repositoryId: repo.id,
                            aheadOfDevelop: await this.isBranchAheadOfDevelop(
                                releaseBranchName,
                                repo.id
                            ),
                            aheadOfMasterMain:
                                await this.isBranchAheadOMasterMain(
                                    repositoryBranchInfo,
                                    releaseBranchName,
                                    repo.id
                                ),
                        };

                        for (const pullRequest of this.state.pullRequests) {
                            if (
                                pullRequest.repository.id === repo.id &&
                                pullRequest.sourceRefName === releaseBranch.name
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
                        hasExistingRelease:
                            repositoryBranchInfo.releaseBranches.length > 0,
                        hasMainBranch: repositoryBranchInfo.hasMainBranch,
                        existingReleaseBranches,
                        branchesAndTags:
                            repositoryBranchInfo.allBranchesAndTags,
                    });
                }
            }

            this.setState({
                repositories: new ArrayItemProvider(
                    Common.sortRepositoryList(reposExtended)
                ),
            });

            bindSelectionToObservable(
                this.state.repositoryListSelection,
                this.state.repositories,
                this.state
                    .repositoryListSelectedItemObservable as ObservableValue<Common.IGitRepositoryExtended>
            );
        }
    }

    private async isBranchAheadOfDevelop(
        branchName: string,
        repositoryId: string
    ): Promise<boolean> {
        const developBranchDescriptor: GitBaseVersionDescriptor = {
            baseVersion: 'develop',
            baseVersionOptions: 0,
            baseVersionType: 0,
            version: 'develop',
            versionOptions: 0,
            versionType: 0,
        };
        const releaseBranchDescriptor: GitTargetVersionDescriptor = {
            targetVersion: branchName,
            targetVersionOptions: 0,
            targetVersionType: 0,
            version: branchName,
            versionOptions: 0,
            versionType: 0,
        };

        const developCommitsDiff: GitCommitDiffs = await Common.getCommitDiffs(
            repositoryId,
            developBranchDescriptor,
            releaseBranchDescriptor
        );

        return Common.codeChangesInCommitDiffs(developCommitsDiff);
    }

    private async isBranchAheadOMasterMain(
        repositoryBranchInfo: Common.IRepositoryBranchInfo,
        branchName: string,
        repositoryId: string
    ): Promise<boolean> {
        const masterMainBranchDescriptor: GitBaseVersionDescriptor = {
            baseVersion: repositoryBranchInfo.hasMasterBranch
                ? 'master'
                : 'main',
            baseVersionOptions: 0,
            baseVersionType: 0,
            version: repositoryBranchInfo.hasMasterBranch ? 'master' : 'main',
            versionOptions: 0,
            versionType: 0,
        };
        const releaseBranchDescriptor: GitTargetVersionDescriptor = {
            targetVersion: branchName,
            targetVersionOptions: 0,
            targetVersionType: 0,
            version: branchName,
            versionOptions: 0,
            versionType: 0,
        };

        const masterMainCommitsDiff: GitCommitDiffs =
            await Common.getCommitDiffs(
                repositoryId,
                masterMainBranchDescriptor,
                releaseBranchDescriptor
            );

        return Common.codeChangesInCommitDiffs(masterMainCommitsDiff);
    }

    private async selectRepository(): Promise<void> {
        this.state.releaseBranchListSelection.clear();
        if (
            this.state.repositoryListSelectedItemObservable.value
                .existingReleaseBranches.length == 1
        ) {
            this.state.releaseBranchListSelection.select(0);
        }

        const buildDefinitionForRepo: BuildDefinition | undefined =
            this.buildDefinitions.find(
                (buildDef) =>
                    buildDef.repository.id ===
                    this.state.repositoryListSelectedItemObservable.value.id
            );

        for (const releaseBranch of this.state
            .repositoryListSelectedItemObservable.value
            .existingReleaseBranches) {
            if (buildDefinitionForRepo) {
                await Common.storeBranchReleaseInfoIntoObservable(
                    releaseInfoObservable,
                    buildDefinitionForRepo,
                    this.releaseDefinitions,
                    releaseBranch,
                    this.state.repositoryListSelectedItemObservable.value
                        .project.id,
                    this.state.repositoryListSelectedItemObservable.value.id,
                    this.organizationName,
                    this.accessToken
                );
            }
        }
    }

    private renderRepositoryMasterPageList(): JSX.Element {
        console.log('inside render list: ', this.state.repositories.length);
        return !this.state.repositories ||
            this.state.repositories.length == 0 ? (
            <div className="page-content-top">
                <Spinner label="loading" />
            </div>
        ) : (
            <List
                ariaLabel={'Repositories'}
                itemProvider={this.state.repositories}
                selection={this.state.repositoryListSelection}
                renderRow={this.renderRepositoryListItem}
                width="100%"
                singleClickActivation={true}
                onSelect={async () => {
                    await this.selectRepository();
                }}
            />
        );
    }

    private renderRepositoryListItem(
        index: number,
        item: Common.IGitRepositoryExtended,
        details: IListItemDetails<Common.IGitRepositoryExtended>,
        key?: string
    ): JSX.Element {
        console.log('in render repository list item');
        const releaseBranchLinks: JSX.Element[] = [];
        let counter: number = 0;
        for (const releaseBranch of item.existingReleaseBranches) {
            const releaseBranchName =
                releaseBranch.targetBranch.name.split('heads/')[1];
            releaseBranchLinks.push(
                <div className="flex-row padding-vertical-10" key={counter}>
                    {Common.branchLinkJsxElement(
                        counter + 'link',
                        item.webUrl,
                        releaseBranchName,
                        'padding-horizontal-6'
                    )}
                    {releaseBranch.aheadOfDevelop &&
                        this.renderAheadOfPillJsxElement(
                            Common.primaryColor,
                            PillVariant.standard,
                            'develop',
                            releaseBranch.developPR
                        )}
                    {releaseBranch.aheadOfMasterMain &&
                        this.renderAheadOfPillJsxElement(
                            Common.primaryColorShade30,
                            PillVariant.outlined,
                            item.hasMainBranch ? 'main' : 'master',
                            releaseBranch.masterMainPR
                        )}
                </div>
            );
            counter++;
        }
        console.log('release branches ready: ', releaseBranchLinks.length);
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
                                {Common.repositoryLinkJsxElement(
                                    item.webUrl,
                                    'font-size-1',
                                    item.name
                                )}
                            </div>
                        </Tooltip>
                        <Tooltip overflowOnly={true}>
                            <div className="flex-column primary-text text-ellipsis">
                                {<>{releaseBranchLinks}</>}
                            </div>
                        </Tooltip>
                    </div>
                </div>
            </ListItem>
        );
    }

    private renderAheadOfPillJsxElement(
        color: IColor,
        varient: PillVariant,
        aheadOfText: string,
        pullRequest?: GitPullRequest
    ): JSX.Element {
        return (
            <Pill
                color={color}
                size={PillSize.regular}
                className="bolt-list-overlay margin-horizontal-3"
                variant={varient}
            >
                <div className="sprintly-text-white">
                    Ahead of {aheadOfText}{' '}
                    {pullRequest && (
                        <i>
                            <Icon
                                ariaLabel="Pull Request"
                                iconName="BranchPullRequest"
                                size={IconSize.small}
                            />{' '}
                            #{pullRequest.pullRequestId}
                        </i>
                    )}
                </div>
            </Pill>
        );
    }

    private renderDetailPageHeaderTitle(
        repositoryName: string,
        repositoryWebUrl: string
    ): JSX.Element {
        return (
            <HeaderTitleArea>
                <HeaderTitleRow>
                    <HeaderTitle
                        ariaLevel={3}
                        className="text-ellipsis"
                        titleSize={TitleSize.Large}
                    >
                        <Link
                            excludeTabStop
                            href={repositoryWebUrl + '/branches'}
                            subtle={false}
                            target="_blank"
                        >
                            {repositoryName}
                        </Link>
                    </HeaderTitle>
                </HeaderTitleRow>
                <HeaderDescription>
                    Select a release branch below to perform actions on it.
                </HeaderDescription>
            </HeaderTitleArea>
        );
    }

    private renderDetailPageHeaderCommandBar(
        repositoryName: string,
        repositoryBranchesAndTags: GitRef[]
    ): JSX.Element {
        return (
            <HeaderCommandBar
                items={[
                    {
                        iconProps: {
                            iconName: 'Tag',
                        },
                        id: 'testSave',
                        important: true,
                        text: 'View Tags',
                        onActivate: () => {
                            tagsModalKeyObservable.value = new Date()
                                .getTime()
                                .toString();
                            isTagsDialogOpenObservable.value = true;
                            const modalContent: ITagsModalContent =
                                getTagsModalContent(
                                    repositoryName,
                                    repositoryBranchesAndTags
                                );
                            tagsRepoNameObservable.value =
                                modalContent.modalName;
                            tagsObservable.value = modalContent.modalValues;
                        },
                    },
                ]}
            />
        );
    }

    private renderDetailPageHeader(
        repositoryName: string,
        repositoryWebUrl: string,
        repositoryBranchesAndTags: GitRef[]
    ): JSX.Element {
        return (
            <CustomHeader className="bolt-header-with-commandbar">
                <HeaderIcon
                    className="bolt-table-status-icon-large"
                    iconProps={{
                        iconName: 'Repo',
                        size: IconSize.large,
                    }}
                    titleSize={TitleSize.Large}
                />
                {this.renderDetailPageHeaderTitle(
                    repositoryName,
                    repositoryWebUrl
                )}
                {this.renderDetailPageHeaderCommandBar(
                    repositoryName,
                    repositoryBranchesAndTags
                )}
            </CustomHeader>
        );
    }

    private renderDetailPageContent(): JSX.Element {
        return (
            <Observer
                selectedItem={this.state.repositoryListSelectedItemObservable}
            >
                {(observerProps: {
                    selectedItem: Common.IGitRepositoryExtended;
                }) => (
                    <Page className="flex-grow single-layer-details">
                        {this.state.repositoryListSelection.selectedCount ==
                            0 && (
                            <span className="single-layer-details-contents">
                                Select a repository on the right to get started.
                            </span>
                        )}
                        {observerProps.selectedItem &&
                            this.state.repositoryListSelection.selectedCount >
                                0 && (
                                <Page>
                                    {this.renderDetailPageHeader(
                                        observerProps.selectedItem.name,
                                        observerProps.selectedItem.webUrl,
                                        observerProps.selectedItem
                                            .branchesAndTags
                                    )}
                                    <div className="page-content page-content-top">
                                        <Card>
                                            {this.renderReleaseBranchDetailList(
                                                new ArrayItemProvider(
                                                    observerProps.selectedItem.existingReleaseBranches
                                                )
                                            )}
                                        </Card>
                                    </div>
                                </Page>
                            )}
                    </Page>
                )}
            </Observer>
        );
    }

    private renderReleaseBranchDetailList(
        items: ArrayItemProvider<Common.IReleaseBranchInfo>
    ): JSX.Element {
        return (
            <List
                ariaLabel={'Release Branches'}
                itemProvider={items}
                selection={this.state.releaseBranchListSelection}
                renderRow={this.renderReleaseBranchDetailListItem}
                width="100%"
                singleClickActivation={true}
            />
        );
    }

    private renderReleaseBranchDetailListItem(
        index: number,
        item: Common.IReleaseBranchInfo,
        details: IListItemDetails<Common.IReleaseBranchInfo>,
        key?: string
    ): JSX.Element {
        return (
            <ListItem
                className="master-row border-bottom"
                key={key || 'list-item' + index}
                index={index}
                details={details}
            >
                <div className="master-row-content flex-row flex-center h-scroll-hidden">
                    <Observer releaseInfoForAllBranches={releaseInfoObservable}>
                        {(observerProps: {
                            releaseInfoForAllBranches: Common.IReleaseInfo[];
                        }) => {
                            let sortedReleases: Release[] = [];
                            if (
                                observerProps.releaseInfoForAllBranches.length >
                                0
                            ) {
                                sortedReleases =
                                    Common.getSortedReleasesForBranch(
                                        item,
                                        observerProps.releaseInfoForAllBranches
                                    );
                            }
                            if (sortedReleases.length == 0) {
                                return (
                                    <div className="flex-row">
                                        <div className="margin-horizontal-10">
                                            {Common.getBranchShortName(
                                                item.targetBranch.name
                                            )}
                                        </div>
                                        {Common.noReleaseExistsPillJsxElement()}
                                    </div>
                                );
                            }
                            console.log('sortedReleases: ', sortedReleases);
                            const mostRecentRelease: Release =
                                sortedReleases[0];
                            const environmentStatuses: JSX.Element[] =
                                Common.getAllEnvironmentStatusPillJsxElements(
                                    mostRecentRelease.environments
                                );
                            return (
                                <div className="flex-row">
                                    <div className="margin-horizontal-10">
                                        {Common.getBranchShortName(
                                            item.targetBranch.name
                                        )}
                                    </div>
                                    {environmentStatuses}
                                </div>
                            );
                        }}
                    </Observer>
                </div>
            </ListItem>
        );
    }

    private renderTagsModalActionButton(): JSX.Element {
        return (
            <Observer
                isTagsDialogOpen={isTagsDialogOpenObservable}
                tagsRepoName={tagsRepoNameObservable}
                tagsModalKey={tagsModalKeyObservable}
            >
                {(props: {
                    isTagsDialogOpen: boolean;
                    tagsRepoName: string;
                    tagsModalKey: string;
                }) => {
                    return (
                        <TagsModal
                            key={props.tagsModalKey}
                            isTagsDialogOpen={props.isTagsDialogOpen}
                            tagsRepoName={props.tagsRepoName}
                            tags={tagsObservable.value}
                            closeMe={() => {
                                isTagsDialogOpenObservable.value = false;
                            }}
                        ></TagsModal>
                    );
                }}
            </Observer>
        );
    }

    public render(): JSX.Element {
        return (
            /* tslint:disable */
            <Observer
                totalRepositoriesToProcess={
                    totalRepositoriesToProcessObservable
                }
            >
                {(props: { totalRepositoriesToProcess: number }) => {
                    if (props.totalRepositoriesToProcess > 0) {
                        return (
                            <div
                                className="flex-grow"
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
                                    nearElementClassName="v-scroll-auto custom-scrollbar light-grey"
                                    farElementClassName="v-scroll-auto custom-scrollbar"
                                    onRenderNearElement={
                                        this.renderRepositoryMasterPageList
                                    }
                                    onRenderFarElement={
                                        this.renderDetailPageContent
                                    }
                                />
                                {this.renderTagsModalActionButton()}
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
