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
import { Release, ReleaseDefinition } from 'azure-devops-extension-api/Release';
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
    Header,
    HeaderDescription,
    HeaderIcon,
    HeaderTitle,
    HeaderTitleArea,
    HeaderTitleRow,
    TitleSize,
} from 'azure-devops-ui/Header';
import { HeaderCommandBar } from 'azure-devops-ui/HeaderCommandBar';

import { TagsModal, ITagsModalContent, getTagsModalContent } from './TagsModal';
import { Spinner } from 'azure-devops-ui/Spinner';
import {
    IStatusProps,
    Status,
    Statuses,
    StatusSize,
} from 'azure-devops-ui/Status';
import { Button } from 'azure-devops-ui/Button';
import { ButtonGroup } from 'azure-devops-ui/ButtonGroup';

import * as Common from './SprintlyCommon';
import { Panel } from 'azure-devops-ui/Panel';
import { Dialog } from 'azure-devops-ui/Dialog';

// TODO: Instead of a state, consider just global observables
export interface ISprintlyPostReleaseState {
    repositories: ArrayItemProvider<Common.IGitRepositoryExtended>;
    pullRequests: GitPullRequest[];
    repositoryListSelection: ListSelection;
    releaseBranchListSelection: ListSelection;
    repositoryListSelectedItemObservable: ObservableValue<Common.IGitRepositoryExtended>;
    releaseBranchListSelectedItemObservable: ObservableValue<Common.IReleaseBranchInfo>;
    selectedRepositoryWebUrl?: string;
    releaseBranchSafeToDelete?: boolean;
}

//#region "Observables"
const tagsModalKeyObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const isTagsDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const isPRCreatePanelOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const isDeleteBranchDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const tagsRepoNameObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const tagsObservable: ObservableValue<string[]> = new ObservableValue<string[]>(
    []
);
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const allBranchesReleaseInfoObservable: ObservableArray<Common.IReleaseInfo> =
    new ObservableArray<Common.IReleaseInfo>();
//#endregion "Observables"

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
        ).map((item: Common.IAllowedEntity) => item.originId);
        totalRepositoriesToProcessObservable.value =
            repositoriesToProcess.length;
        if (repositoriesToProcess.length > 0) {
            const filteredProjects: TeamProjectReference[] =
                await Common.getFilteredProjects();
            const pullRequests: GitPullRequest[] = await Common.getPullRequests(
                filteredProjects
            );
            this.setState({
                pullRequests: pullRequests,
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

            totalRepositoriesToProcessObservable.value = filteredRepos.length;

            for (const repo of filteredRepos) {
                const repositoryBranchInfo: Common.IRepositoryBranchInfo =
                    await Common.getRepositoryBranchesInfo(repo.id);

                const processRepo: boolean =
                    repositoryBranchInfo.hasDevelopBranch &&
                    (repositoryBranchInfo.hasMasterBranch ||
                        repositoryBranchInfo.hasMainBranch);

                if (processRepo === true) {
                    const existingReleaseBranches: Common.IReleaseBranchInfo[] =
                        [];
                    for (const releaseBranch of repositoryBranchInfo.releaseBranches) {
                        const releaseBranchName: string =
                            Common.getBranchShortName(releaseBranch.name);

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
        const buildDefinitionForRepo: BuildDefinition | undefined =
            this.buildDefinitions.find(
                (buildDef: BuildDefinition) =>
                    buildDef.repository.id ===
                    this.state.repositoryListSelectedItemObservable.value.id
            );

        for (const releaseBranch of this.state
            .repositoryListSelectedItemObservable.value
            .existingReleaseBranches) {
            if (buildDefinitionForRepo) {
                await Common.fetchAndStoreBranchReleaseInfoIntoObservable(
                    allBranchesReleaseInfoObservable,
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

        bindSelectionToObservable(
            this.state.releaseBranchListSelection,
            new ArrayItemProvider(
                this.state.repositoryListSelectedItemObservable.value.existingReleaseBranches
            ),
            this.state
                .releaseBranchListSelectedItemObservable as ObservableValue<Common.IReleaseBranchInfo>
        );

        this.setState({
            selectedRepositoryWebUrl:
                this.state.repositoryListSelectedItemObservable.value.webUrl,
        });
        this.autoSelectIfSingleBranch();
    }

    private autoSelectIfSingleBranch(): void {
        if (
            this.state.repositoryListSelectedItemObservable.value
                .existingReleaseBranches.length === 1
        ) {
            this.state.releaseBranchListSelection.select(0);
        }
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

    private renderRepositoryListItem(
        index: number,
        item: Common.IGitRepositoryExtended,
        details: IListItemDetails<Common.IGitRepositoryExtended>,
        key?: string
    ): JSX.Element {
        const releaseBranchLinks: JSX.Element[] = [];
        let counter: number = 0;
        for (const releaseBranch of item.existingReleaseBranches) {
            const releaseBranchName: string =
                releaseBranch.targetBranch.name.split('heads/')[1];
            releaseBranchLinks.push(
                <div className='flex-row padding-vertical-10' key={counter}>
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
                        <Tooltip overflowOnly={true}>
                            <div className='flex-column primary-text text-ellipsis'>
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
                className='bolt-list-overlay margin-horizontal-3'
                variant={varient}
            >
                <div className='sprintly-text-white'>
                    Ahead of {aheadOfText}{' '}
                    {pullRequest && (
                        <i>
                            <Icon
                                ariaLabel='Pull Request'
                                iconName='BranchPullRequest'
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
                        className='text-ellipsis'
                        titleSize={TitleSize.Large}
                    >
                        <Link
                            excludeTabStop
                            href={repositoryWebUrl + '/branches'}
                            subtle={false}
                            target='_blank'
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
            <CustomHeader className='bolt-header-with-commandbar'>
                <HeaderIcon
                    className='bolt-table-status-icon-large'
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
                    <Page className='flex-grow single-layer-details'>
                        {this.state.repositoryListSelection.selectedCount ===
                            0 && (
                            <span className='single-layer-details-contents'>
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
                                    <div className='page-content page-content-top'>
                                        <Card className='bolt-card-white'>
                                            {this.renderReleaseBranchDetailList(
                                                new ArrayItemProvider(
                                                    observerProps.selectedItem.existingReleaseBranches
                                                )
                                            )}
                                        </Card>
                                    </div>
                                    {this.renderDetailPageActionsContent()}
                                </Page>
                            )}
                    </Page>
                )}
            </Observer>
        );
    }

    private renderDetailPageActionsContent(): JSX.Element {
        return (
            <Observer
                selectedItem={
                    this.state.releaseBranchListSelectedItemObservable
                }
            >
                {(observerProps: { selectedItem: Common.IReleaseBranchInfo }) =>
                    this.state.releaseBranchListSelection.selectedCount ===
                    0 ? (
                        <></>
                    ) : (
                        <Page>
                            <div className='page-content'>
                                <div>
                                    <Card
                                        className='bolt-table-card bolt-card-white'
                                        titleProps={{
                                            text: 'Develop Branch Actions',
                                        }}
                                    >
                                        <div className='page-content page-content-top'>
                                            {this.renderDetailPageTopActionContent(
                                                observerProps.selectedItem
                                            )}
                                        </div>
                                    </Card>
                                </div>
                                <div style={{ marginTop: '16px' }}>
                                    <Card
                                        className='bolt-table-card bolt-card-white'
                                        titleProps={{
                                            text: 'Master/Main Branch Actions',
                                        }}
                                    >
                                        <div className='page-content page-content-top'>
                                            {this.renderDetailPageBottomActionContent(
                                                observerProps.selectedItem
                                            )}
                                        </div>
                                    </Card>
                                </div>
                                <div className='page-content-top'>
                                    <Button
                                        text='Delete branch'
                                        iconProps={{ iconName: 'Delete' }}
                                        onClick={() => {
                                            isDeleteBranchDialogOpenObservable.value =
                                                true;
                                            this.setState({
                                                releaseBranchSafeToDelete:
                                                    !observerProps.selectedItem
                                                        .aheadOfDevelop &&
                                                    !observerProps.selectedItem
                                                        .aheadOfMasterMain,
                                            });
                                        }}
                                        danger={true}
                                    />
                                </div>
                            </div>
                        </Page>
                    )
                }
            </Observer>
        );
    }

    private renderPullRequestActionSection(
        baseBranch: string,
        aheadOfStatus?: boolean,
        pullRequest?: GitPullRequest
    ): JSX.Element {
        let statusText = '';
        let status: IStatusProps = { ...Statuses.Success };
        let prLink: JSX.Element = <></>;
        let actionButtons: JSX.Element = <></>;
        if (aheadOfStatus && aheadOfStatus.valueOf()) {
            if (pullRequest) {
                statusText = `This branch is ahead of ${baseBranch} and an open PR exists.`;
                status = { ...Statuses.Information };
                prLink = (
                    <Link
                        excludeTabStop
                        href={`${this.state.selectedRepositoryWebUrl}/pullrequest/${pullRequest.pullRequestId}`}
                        subtle={false}
                        target='_blank'
                    >
                        <Icon
                            ariaLabel='Pull Request'
                            iconName='BranchPullRequest'
                            size={IconSize.small}
                        />{' '}
                        #{pullRequest.pullRequestId}
                    </Link>
                );
                actionButtons = (
                    <ButtonGroup>
                        <Button
                            iconProps={{ iconName: 'Accept' }}
                            text='Complete PR'
                            onClick={() => alert('Default button clicked!')}
                        />
                    </ButtonGroup>
                );
            } else {
                statusText = `This branch is ahead of ${baseBranch} and no open PR exists.`;
                status = { ...Statuses.Warning };
                actionButtons = (
                    <ButtonGroup>
                        <Button
                            iconProps={{ iconName: 'Add' }}
                            text='Create New PR'
                            primary={true}
                            onClick={() =>
                                (isPRCreatePanelOpenObservable.value = true)
                            }
                        />
                    </ButtonGroup>
                );
            }
        } else {
            statusText = `This branch is not ahead of ${baseBranch}.`;
            status = { ...Statuses.Success };
        }
        return (
            <>
                <div className='flex-row'>
                    <Status {...status} size={StatusSize.l} animated={false} />
                </div>
                <div className='flex-row page-content-top'>
                    {statusText}&nbsp;{prLink}
                </div>
                <div className='flex-row page-content-top'>{actionButtons}</div>
            </>
        );
    }

    private renderDetailPageTopActionContent(
        selectedBranch: Common.IReleaseBranchInfo
    ): JSX.Element {
        return this.renderPullRequestActionSection(
            'develop',
            selectedBranch.aheadOfDevelop,
            selectedBranch.developPR
        );
    }

    private renderDetailPageBottomActionContent(
        selectedBranch: Common.IReleaseBranchInfo
    ): JSX.Element {
        return this.renderPullRequestActionSection(
            'master/main',
            selectedBranch.aheadOfMasterMain,
            selectedBranch.masterMainPR
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
                width='100%'
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
                className='master-row border-bottom'
                key={key || 'list-item' + index}
                index={index}
                details={details}
            >
                <div className='master-row-content flex-row flex-center h-scroll-hidden'>
                    <Observer
                        releaseInfoForAllBranches={
                            allBranchesReleaseInfoObservable
                        }
                    >
                        {(observerProps: {
                            releaseInfoForAllBranches: Common.IReleaseInfo[];
                        }) => {
                            const mostRecentRelease: Release | undefined =
                                Common.getMostRecentReleaseForBranch(
                                    item,
                                    observerProps.releaseInfoForAllBranches
                                );
                            if (!mostRecentRelease) {
                                return (
                                    <div className='flex-row'>
                                        <div className='margin-horizontal-10'>
                                            {Common.getBranchShortName(
                                                item.targetBranch.name
                                            )}
                                        </div>
                                        {Common.noReleaseExistsPillJsxElement()}
                                    </div>
                                );
                            }
                            const environmentStatuses: JSX.Element[] =
                                Common.getAllEnvironmentStatusPillJsxElements(
                                    mostRecentRelease.environments
                                );
                            return (
                                <div className='flex-row'>
                                    <div className='margin-horizontal-10'>
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

    private renderPRCreatePanelActionButton(): JSX.Element {
        return (
            <Observer isPRCreatePanelOpen={isPRCreatePanelOpenObservable}>
                {(observerProps: { isPRCreatePanelOpen: boolean }) => {
                    return observerProps.isPRCreatePanelOpen ? (
                        <Panel
                            onDismiss={() =>
                                (isPRCreatePanelOpenObservable.value = false)
                            }
                            titleProps={{ text: 'Sample Panel Title' }}
                            description={
                                'A description of the header. It can expand to multiple lines. Consumers should try to limit this to a maximum of three lines.'
                            }
                            footerButtonProps={[
                                {
                                    text: 'Cancel',
                                    onClick: () =>
                                        (isPRCreatePanelOpenObservable.value =
                                            false),
                                },
                                { text: 'Create', primary: true },
                            ]}
                        >
                            <div>Panel Content</div>
                        </Panel>
                    ) : null;
                }}
            </Observer>
        );
    }

    private renderDeleteBranchConfirmAction(): JSX.Element {
        return (
            <Observer
                isDeleteBranchDialogOpen={isDeleteBranchDialogOpenObservable}
            >
                {(props: { isDeleteBranchDialogOpen: boolean }) => {
                    return props.isDeleteBranchDialogOpen ? (
                        <Dialog
                            titleProps={{
                                text: 'Are you sure?',
                            }}
                            footerButtonProps={[
                                {
                                    text: 'Cancel',
                                    onClick:
                                        this.onDismissDeleteBranchActionModal,
                                },
                                {
                                    text: 'Delete',
                                    onClick:
                                        this.onDismissDeleteBranchActionModal,
                                    danger: true,
                                },
                            ]}
                            onDismiss={this.onDismissDeleteBranchActionModal}
                        >
                            {this.state.releaseBranchSafeToDelete ? (
                                <>
                                    You are about to delete{' '}
                                    {Common.getBranchShortName(
                                        this.state
                                            .releaseBranchListSelectedItemObservable
                                            .value.targetBranch.name
                                    )}
                                </>
                            ) : (
                                <>
                                    <Icon
                                        style={{ color: 'orange' }}
                                        ariaLabel='Warning'
                                        iconName='Warning'
                                        size={IconSize.large}
                                    />{' '}
                                    Note:{' '}
                                    {Common.getBranchShortName(
                                        this.state
                                            .releaseBranchListSelectedItemObservable
                                            .value.targetBranch.name
                                    )}{' '}
                                    may need to still be back-merged into
                                    develop or main/master. Are you sure you
                                    want to delete it?
                                </>
                            )}
                        </Dialog>
                    ) : null;
                }}
            </Observer>
        );
    }

    private onDismissDeleteBranchActionModal(): void {
        isDeleteBranchDialogOpenObservable.value = false;
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
                                    nearElementClassName='v-scroll-auto custom-scrollbar light-grey'
                                    farElementClassName='v-scroll-auto custom-scrollbar'
                                    onRenderNearElement={
                                        this.renderRepositoryMasterPageList
                                    }
                                    onRenderFarElement={
                                        this.renderDetailPageContent
                                    }
                                />
                                {this.renderTagsModalActionButton()}
                                {this.renderPRCreatePanelActionButton()}
                                {this.renderDeleteBranchConfirmAction()}
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
