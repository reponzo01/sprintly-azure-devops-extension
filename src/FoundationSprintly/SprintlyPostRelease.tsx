import * as React from 'react';
import axios, { AxiosResponse } from 'axios';
import * as SDK from 'azure-devops-extension-sdk';
import {
    getClient,
    IColor,
    IExtensionDataManager,
    IGlobalMessagesService,
    MessageBannerLevel,
} from 'azure-devops-extension-api';
import { TeamProjectReference } from 'azure-devops-extension-api/Core';
import {
    GitAnnotatedTag,
    GitBaseVersionDescriptor,
    GitCommitDiffs,
    GitObjectType,
    GitPullRequest,
    GitRef,
    GitRefUpdate,
    GitRefUpdateResult,
    GitRefUpdateStatus,
    GitRepository,
    GitRestClient,
    GitTargetVersionDescriptor,
    PullRequestAsyncStatus,
    PullRequestStatus,
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
import { Dialog } from 'azure-devops-ui/Dialog';
import { ISelectionRange } from 'azure-devops-ui/Utilities/Selection';
import {
    TextField,
    TextFieldStyle,
    TextFieldWidth,
} from 'azure-devops-ui/TextField';
import { FormItem } from 'azure-devops-ui/FormItem';

// TODO: Instead of a state, consider just global observables
export interface ISprintlyPostReleaseState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    repositories: ArrayItemProvider<Common.IGitRepositoryExtended>;
    pullRequests: GitPullRequest[];
    repositoryListSelection: ListSelection;
    releaseBranchListSelection: ListSelection;
    repositoryListSelectedItemObservable: ObservableValue<Common.IGitRepositoryExtended>;
    releaseBranchListSelectedItemObservable: ObservableValue<Common.IReleaseBranchInfo>;
    baseDevelopBranch?: GitRef;
    baseMasterMainBranch?: GitRef;
    selectedRepositoryWebUrl?: string;
    releaseBranchSafeToDelete?: boolean;
    pullRequestSourceBranchName?: string;
    pullRequestTargetBranchName?: string;
    pullRequestToComplete?: GitPullRequest;
    createTagTitle?: string;
    createTagDescription?: string;
    createPullRequestTitle?: string;
    createPullRequestDescription?: string;
}

//#region "Observables"
const tagsModalKeyObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const isTagsDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const isCreatePullRequestDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const isCompletePullRequestDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const isDeleteBranchDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const isCreateTagDialogOpenObservable: ObservableValue<boolean> =
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
const createPullRequestTitleObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const createPullRequestDescriptionObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const createTagTitleObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const createTagDescriptionObservable: ObservableValue<string> =
    new ObservableValue<string>('');
//#endregion "Observables"

const userSettingsDataManagerKey: string = 'user-settings';
const systemSettingsDataManagerKey: string = 'system-settings';

let repositoriesToProcess: string[] = [];

// TODO: Clean up arrow functions for the cases in which I thought I
// couldn't use regular functions because the this.* was undefined errors.
// The solution is to bind those functions to `this` in the constructor.
// See SprintlyPostRelease as an example.
export default class SprintlyPostRelease extends React.Component<
    {
        organizationName: string;
        globalMessagesSvc: IGlobalMessagesService;
        dataManager: IExtensionDataManager;
    },
    ISprintlyPostReleaseState
> {
    _isMounted: boolean = false;
    private dataManager: IExtensionDataManager;
    private globalMessagesSvc: IGlobalMessagesService;
    private accessToken: string = '';
    private organizationName: string;

    private releaseDefinitions: ReleaseDefinition[] = [];
    private buildDefinitions: BuildDefinition[] = [];

    constructor(props: {
        organizationName: string;
        globalMessagesSvc: IGlobalMessagesService;
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
        this.deleteBranchAction = this.deleteBranchAction.bind(this);
        this.completePullRequestAction =
            this.completePullRequestAction.bind(this);

        this.organizationName = props.organizationName;
        this.globalMessagesSvc = props.globalMessagesSvc;
        this.dataManager = props.dataManager;
    }

    public async componentDidMount(): Promise<void> {
        this._isMounted = true;
        await this.initializeComponent();
    }

    public async componentWillUnmount() {
        this._isMounted = false;
    }

    private async initializeComponent(): Promise<void> {
        if (this._isMounted) {
            this.accessToken = await SDK.getAccessToken();

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
                userSettings: userSettings,
                systemSettings: systemSettings,
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
                const pullRequests: GitPullRequest[] =
                    await Common.getPullRequests(filteredProjects);
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
    }

    private async reloadComponent(): Promise<void> {
        this.state.releaseBranchListSelection.clear();
        const repoSelectionIndex: ISelectionRange[] =
            this.state.repositoryListSelection.value;
        this.state.repositoryListSelection.clear();
        await this.initializeComponent();
        this.setState(this.state);
        this.state.repositoryListSelection.select(
            repoSelectionIndex[0].beginIndex
        );
        await this.selectRepository();
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
                    const baseDevelopBranch: GitRef | undefined =
                        repositoryBranchInfo.hasDevelopBranch
                            ? repositoryBranchInfo.allBranchesAndTags.find(
                                  (branch) =>
                                      branch.name === 'refs/heads/develop'
                              )
                            : undefined;
                    const baseMasterMainBranch: GitRef | undefined =
                        repositoryBranchInfo.hasMasterBranch
                            ? repositoryBranchInfo.allBranchesAndTags.find(
                                  (branch) =>
                                      branch.name === 'refs/heads/master'
                              )
                            : repositoryBranchInfo.allBranchesAndTags.find(
                                  (branch) => branch.name === 'refs/heads/main'
                              );
                    const existingReleaseBranches: Common.IReleaseBranchInfo[] =
                        [];
                    for (const releaseBranch of repositoryBranchInfo.releaseBranches) {
                        const releaseBranchName: string =
                            Common.getBranchShortName(releaseBranch.name);

                        const branchInfo: Common.IReleaseBranchInfo = {
                            targetBranch: releaseBranch,
                            repositoryId: repo.id,
                            aheadOfDevelop: await Common.isBranchAheadOfDevelop(
                                releaseBranchName,
                                repo.id
                            ),
                            aheadOfMasterMain:
                                await Common.isBranchAheadOMasterMain(
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
                        baseDevelopBranch,
                        baseMasterMainBranch,
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

    private async selectRepository(): Promise<void> {
        this.state.releaseBranchListSelection.clear();
        this.state.releaseBranchListSelectedItemObservable.value = {} as any;
        const repositoryInfo: Common.IGitRepositoryExtended =
            this.state.repositoryListSelectedItemObservable.value;
        const buildDefinitionForRepo: BuildDefinition | undefined =
            this.buildDefinitions.find(
                (buildDef: BuildDefinition) =>
                    buildDef.repository.id === repositoryInfo.id
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
                    repositoryInfo.project.id,
                    repositoryInfo.id,
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
            baseDevelopBranch: repositoryInfo.baseDevelopBranch,
            baseMasterMainBranch: repositoryInfo.baseMasterMainBranch,
        });

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
                            {pullRequest.mergeStatus ===
                                PullRequestAsyncStatus.Conflicts && (
                                <Icon
                                    ariaLabel='Merge Conflicts'
                                    iconName='Warning'
                                    size={IconSize.small}
                                    style={{ color: 'orange' }}
                                />
                            )}{' '}
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
                        id: 'viewtags',
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
                    {
                        iconProps: {
                            iconName: 'Add',
                        },
                        id: 'createtag',
                        important: true,
                        text: 'Create Tag',
                        isPrimary: true,
                        onActivate: () => {
                            isCreateTagDialogOpenObservable.value = true;
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
                    this.state.releaseBranchListSelection.selectedCount === 0 ||
                    !observerProps.selectedItem ? (
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
                                                    (!observerProps.selectedItem
                                                        .aheadOfDevelop ||
                                                        !observerProps.selectedItem.aheadOfDevelop.valueOf()) &&
                                                    (!observerProps.selectedItem
                                                        .aheadOfMasterMain ||
                                                        !observerProps.selectedItem.aheadOfMasterMain.valueOf()),
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
        // TODO: Extract the baseBranch ('develop', 'master/main') into enum
        baseBranch: string,
        sourceBranchName: string,
        aheadOfStatus?: boolean,
        pullRequest?: GitPullRequest
    ): JSX.Element {
        let statusText = '';
        let status: IStatusProps = { ...Statuses.Success };
        let prLink: JSX.Element = <></>;
        let actionButtons: JSX.Element = <></>;
        if (aheadOfStatus && aheadOfStatus.valueOf()) {
            if (pullRequest) {
                const mergeConflict: boolean =
                    pullRequest.mergeStatus ===
                    PullRequestAsyncStatus.Conflicts;
                const mergeFailed: boolean =
                    pullRequest.mergeStatus === PullRequestAsyncStatus.Failure;
                const mergeQueued: boolean =
                    pullRequest.mergeStatus === PullRequestAsyncStatus.Queued;

                statusText = `This branch is ahead of ${baseBranch} and an open PR exists.`;
                statusText = mergeConflict
                    ? `${statusText} There are merge conflicts.`
                    : statusText;
                statusText = mergeFailed
                    ? `${statusText} There are merge failures.`
                    : statusText;
                statusText = mergeQueued
                    ? `${statusText} The initial creation pre-merge is queued.`
                    : statusText;
                status =
                    mergeConflict || mergeFailed
                        ? { ...Statuses.Failed }
                        : mergeQueued
                        ? { ...Statuses.Queued }
                        : { ...Statuses.Information };
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
                        {mergeConflict || mergeFailed || mergeQueued ? (
                            <Button
                                iconProps={{ iconName: 'BranchPullRequest' }}
                                text='Review PR'
                                onClick={() =>
                                    window.open(
                                        `${this.state.selectedRepositoryWebUrl}/pullrequest/${pullRequest.pullRequestId}`,
                                        '_blank'
                                    )
                                }
                            />
                        ) : (
                            <Button
                                iconProps={{ iconName: 'Accept' }}
                                text='Complete PR'
                                onClick={() => {
                                    this.setState({
                                        pullRequestToComplete: pullRequest,
                                    });
                                    isCompletePullRequestDialogOpenObservable.value =
                                        true;
                                }}
                            />
                        )}
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
                            onClick={() => {
                                this.setState({
                                    pullRequestTargetBranchName:
                                        baseBranch === 'develop'
                                            ? this.state.baseDevelopBranch?.name
                                            : this.state.baseMasterMainBranch
                                                  ?.name,
                                    pullRequestSourceBranchName:
                                        sourceBranchName,
                                });

                                isCreatePullRequestDialogOpenObservable.value =
                                    true;
                            }}
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
                    <Status {...status} size={StatusSize.l} />
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
            selectedBranch.targetBranch.name,
            selectedBranch.aheadOfDevelop,
            selectedBranch.developPR
        );
    }

    private renderDetailPageBottomActionContent(
        selectedBranch: Common.IReleaseBranchInfo
    ): JSX.Element {
        return this.renderPullRequestActionSection(
            'master/main',
            selectedBranch.targetBranch.name,
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

    private renderViewTagsModal(): JSX.Element {
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

    private renderDeleteBranchActionModal(): JSX.Element {
        return (
            <Observer
                isDeleteBranchDialogOpen={isDeleteBranchDialogOpenObservable}
            >
                {(props: { isDeleteBranchDialogOpen: boolean }) => {
                    return props.isDeleteBranchDialogOpen ? (
                        <Dialog
                            titleProps={{
                                text: 'Delete branch',
                            }}
                            footerButtonProps={[
                                {
                                    text: 'Cancel',
                                    onClick:
                                        this.onDismissDeleteBranchActionModal,
                                },
                                {
                                    text: 'Delete',
                                    onClick: this.deleteBranchAction,
                                    danger: true,
                                },
                            ]}
                            onDismiss={this.onDismissDeleteBranchActionModal}
                        >
                            {this.state.releaseBranchSafeToDelete ? (
                                <>
                                    Branch{' '}
                                    {Common.getBranchShortName(
                                        this.state
                                            .releaseBranchListSelectedItemObservable
                                            .value.targetBranch.name
                                    )}{' '}
                                    will be permanently deleted. Are you sure
                                    you want to proceed?
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

    private deleteBranchAction(): void {
        const createRefOptions: GitRefUpdate[] = [];

        createRefOptions.push({
            repositoryId:
                this.state.repositoryListSelectedItemObservable.value.id,
            name: this.state.releaseBranchListSelectedItemObservable.value
                .targetBranch.name,
            isLocked: false,
            oldObjectId:
                this.state.releaseBranchListSelectedItemObservable.value
                    .targetBranch.objectId,
            newObjectId: '0000000000000000000000000000000000000000',
        });

        // TODO: Error handling delete permissions
        getClient(GitRestClient)
            .updateRefs(
                createRefOptions,
                this.state.repositoryListSelectedItemObservable.value.id
            )
            .then(async (result) => {
                for (const res of result) {
                    this.globalMessagesSvc.addToast({
                        duration: 5000,
                        forceOverrideExisting: true,
                        message: res.success
                            ? 'Branch Deleted!'
                            : 'Error Deleting Branch: ' +
                              GitRefUpdateStatus[res.updateStatus],
                    });
                }
                await this.reloadComponent();
            })
            .catch((error: any) => {
                if (error.response?.data?.message) {
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
                            'Branch deletion failed!' +
                            error +
                            ' ' +
                            error.response?.data?.message,
                    });
                }
            });
        this.onDismissDeleteBranchActionModal();
    }

    private onDismissDeleteBranchActionModal(): void {
        isDeleteBranchDialogOpenObservable.value = false;
    }

    private renderCreatePullRequestActionModal(): JSX.Element {
        return (
            <Observer
                isCreatePullRequestDialogOpen={
                    isCreatePullRequestDialogOpenObservable
                }
                pullRequestTitle={createPullRequestTitleObservable}
                pullRequestDescription={createPullRequestDescriptionObservable}
            >
                {(observerProps: {
                    isCreatePullRequestDialogOpen: boolean;
                    pullRequestTitle: string;
                    pullRequestDescription: string;
                }) => {
                    return observerProps.isCreatePullRequestDialogOpen ? (
                        <Dialog
                            onDismiss={() =>
                                this.onDismissCreatePullRequestActionModal()
                            }
                            titleProps={{
                                text: `Create a PR from ${Common.getBranchShortName(
                                    this.state.pullRequestSourceBranchName ?? ''
                                )} to ${Common.getBranchShortName(
                                    this.state.pullRequestTargetBranchName ?? ''
                                )}`,
                                size: TitleSize.Medium,
                                iconProps: { iconName: 'BranchPullRequest' },
                            }}
                            footerButtonProps={[
                                {
                                    text: 'Cancel',
                                    onClick: () =>
                                        this.onDismissCreatePullRequestActionModal(),
                                },
                                {
                                    text: 'Create',
                                    primary: true,
                                    disabled:
                                        createPullRequestTitleObservable.value.trim() ===
                                            '' ||
                                        createPullRequestDescriptionObservable.value.trim() ===
                                            '',
                                    onClick: () => {
                                        this.onDismissCreatePullRequestActionModal();
                                        this.createPullRequestAction();
                                    },
                                },
                            ]}
                        >
                            <Page className='flex-column rhythm-vertical-16 flex-grow scroll-auto'>
                                <FormItem label='Title *'>
                                    <TextField
                                        required={true}
                                        value={createPullRequestTitleObservable}
                                        onChange={(e, newValue) => {
                                            createPullRequestTitleObservable.value =
                                                newValue;
                                            this.setState({
                                                createPullRequestTitle:
                                                    createPullRequestTitleObservable.value,
                                            });
                                        }}
                                        style={TextFieldStyle.normal}
                                    />
                                </FormItem>
                                <FormItem label='Description *'>
                                    <TextField
                                        required={true}
                                        value={
                                            createPullRequestDescriptionObservable
                                        }
                                        onChange={(e, newValue) => {
                                            createPullRequestDescriptionObservable.value =
                                                newValue;
                                            this.setState({
                                                createPullRequestDescription:
                                                    createPullRequestDescriptionObservable.value,
                                            });
                                        }}
                                        multiline={true}
                                        style={TextFieldStyle.normal}
                                    />
                                </FormItem>
                            </Page>
                        </Dialog>
                    ) : null;
                }}
            </Observer>
        );
    }

    private createPullRequestAction(): void {
        const pullRequest: any = {
            title: this.state.createPullRequestTitle,
            description: this.state.createPullRequestDescription,
            isDraft: false,
            labels: [],
            reviewers: [],
            sourceRefName: this.state.pullRequestSourceBranchName,
            targetRefName: this.state.pullRequestTargetBranchName,
        };
        getClient(GitRestClient)
            .createPullRequest(
                pullRequest as GitPullRequest,
                this.state.repositoryListSelectedItemObservable.value.id
            )
            .then(async (result) => {
                this.globalMessagesSvc.addToast({
                    duration: 5000,
                    forceOverrideExisting: true,
                    message: 'PR creation started!',
                });
                await this.reloadComponent();
            })
            .catch((error: any) => {
                this.globalMessagesSvc.addToast({
                    duration: 5000,
                    forceOverrideExisting: true,
                    message: 'PR creation failed!' + error,
                });
            });
        isCreatePullRequestDialogOpenObservable.value = false;
    }

    private onDismissCreatePullRequestActionModal(): void {
        createPullRequestTitleObservable.value = '';
        createPullRequestDescriptionObservable.value = '';
        isCreatePullRequestDialogOpenObservable.value = false;
    }

    private renderCompletePullRequestActionModal(): JSX.Element {
        return (
            <Observer
                isCompletePullRequestDialogOpen={
                    isCompletePullRequestDialogOpenObservable
                }
            >
                {(props: { isCompletePullRequestDialogOpen: boolean }) => {
                    return props.isCompletePullRequestDialogOpen ? (
                        <Dialog
                            titleProps={{
                                text: 'Complete Pull Request',
                            }}
                            footerButtonProps={[
                                {
                                    text: 'Cancel',
                                    onClick:
                                        this
                                            .onDismissCompletePullRequestActionModal,
                                },
                                {
                                    text: 'Complete',
                                    onClick: this.completePullRequestAction,
                                    danger: true,
                                },
                            ]}
                            onDismiss={
                                this.onDismissCompletePullRequestActionModal
                            }
                        >
                            <>
                                <Icon
                                    style={{ color: 'orange' }}
                                    ariaLabel='Warning'
                                    iconName='Warning'
                                    size={IconSize.large}
                                />{' '}
                                Note: If you have permissions to override branch
                                policies, this action will complete this pull
                                request overriding any branch policies.
                            </>
                        </Dialog>
                    ) : null;
                }}
            </Observer>
        );
    }

    private completePullRequestAction(): void {
        if (!this.state.pullRequestToComplete) {
            isCompletePullRequestDialogOpenObservable.value = false;
            return;
        }
        const url: string = `https://dev.azure.com/${this.organizationName}/${this.state.pullRequestToComplete.repository.project.id}/_apis/git/repositories/${this.state.pullRequestToComplete.repository.id}/pullrequests/${this.state.pullRequestToComplete.pullRequestId}?api-version=5.0`;
        const requestBody: any = {
            completionOptions: {
                autoCompleteIgnoreConfigIds: [],
                bypassPolicy: true,
                bypassReason: 'Post release cleanup',
                deleteSourceBranch: false,
                mergeCommitMessage: `Merge ${Common.getBranchShortName(
                    this.state.pullRequestToComplete.sourceRefName
                )} into ${Common.getBranchShortName(
                    this.state.pullRequestToComplete.targetRefName
                )}`,
                mergeStrategy: 1,
                transitionWorkItems: false,
            },
            lastMergeSourceCommit: {
                commitId: `${this.state.pullRequestToComplete.lastMergeSourceCommit.commitId}`,
                url: `${this.state.pullRequestToComplete.lastMergeSourceCommit.url}`,
            },
            status: PullRequestStatus.Completed,
        };

        Common.getOrRefreshToken(this.accessToken).then(async (token) => {
            await axios
                .patch(url, requestBody, {
                    headers: {
                        Authorization: `Bearer ${token}`,
                    },
                })
                .then(async (result: void | AxiosResponse<any>) => {
                    this.globalMessagesSvc.addToast({
                        duration: 5000,
                        forceOverrideExisting: true,
                        message: 'PR completion queued!',
                    });
                    await this.reloadComponent();
                })
                .catch((error: any) => {
                    if (error.response.status == 403) {
                        this.globalMessagesSvc.addBanner({
                            dismissable: true,
                            level: MessageBannerLevel.error,
                            message:
                                'Permission denied! The target branch may have policies that you do not have permissions to override.',
                            buttons: [
                                {
                                    text: 'Check Permissions',
                                    href: `https://dev.azure.com/${this.organizationName}/${this.state.pullRequestToComplete?.repository.project.id}/_settings/repositories?_a=permissions`,
                                    target: '_blank',
                                },
                            ],
                        });
                    } else {
                        this.globalMessagesSvc.addToast({
                            duration: 5000,
                            forceOverrideExisting: true,
                            message:
                                'PR completion failed!' +
                                error +
                                ' ' +
                                error.response.data?.message,
                        });
                    }
                });
        });

        isCompletePullRequestDialogOpenObservable.value = false;
    }

    private onDismissCompletePullRequestActionModal(): void {
        isCompletePullRequestDialogOpenObservable.value = false;
    }

    private renderCreateTagActionModal(): JSX.Element {
        return (
            <Observer
                isCreateTagDialogOpen={isCreateTagDialogOpenObservable}
                tagTitle={createTagTitleObservable}
                tagDescription={createTagDescriptionObservable}
            >
                {(observerProps: {
                    isCreateTagDialogOpen: boolean;
                    tagTitle: string;
                    tagDescription: string;
                }) => {
                    return observerProps.isCreateTagDialogOpen ? (
                        <Dialog
                            onDismiss={() =>
                                this.onDismissCreateTagActionModal()
                            }
                            titleProps={{
                                text: `Create a Tag from ${Common.getBranchShortName(
                                    this.state.baseMasterMainBranch?.name ?? ''
                                )}`,
                                size: TitleSize.Medium,
                                iconProps: { iconName: 'Tag' },
                            }}
                            footerButtonProps={[
                                {
                                    text: 'Cancel',
                                    onClick: () =>
                                        this.onDismissCreateTagActionModal(),
                                },
                                {
                                    text: 'Create',
                                    primary: true,
                                    disabled:
                                        createTagTitleObservable.value.trim() ===
                                            '' ||
                                        createTagDescriptionObservable.value.trim() ===
                                            '',
                                    onClick: () => {
                                        this.onDismissCreateTagActionModal();
                                        this.createTagAction();
                                    },
                                },
                            ]}
                        >
                            <Page className='flex-column rhythm-vertical-16 flex-grow scroll-auto'>
                                <FormItem label='Title *'>
                                    <TextField
                                        required={true}
                                        value={createTagTitleObservable}
                                        onChange={(e, newValue) => {
                                            createTagTitleObservable.value =
                                                newValue;
                                            this.setState({
                                                createTagTitle:
                                                    createTagTitleObservable.value,
                                            });
                                        }}
                                        style={TextFieldStyle.normal}
                                    />
                                </FormItem>
                                <FormItem label='Description *'>
                                    <TextField
                                        required={true}
                                        value={createTagDescriptionObservable}
                                        onChange={(e, newValue) => {
                                            createTagDescriptionObservable.value =
                                                newValue;
                                            this.setState({
                                                createTagDescription:
                                                    createTagDescriptionObservable.value,
                                            });
                                        }}
                                        multiline={true}
                                        style={TextFieldStyle.normal}
                                    />
                                </FormItem>
                            </Page>
                        </Dialog>
                    ) : null;
                }}
            </Observer>
        );
    }

    private createTagAction(): void {
        if (this.state.baseMasterMainBranch) {
            const tag: any = {
                taggedObject: {
                    objectId: this.state.baseMasterMainBranch.objectId,
                    objectType: GitObjectType.Commit,
                },
                objectId: this.state.baseMasterMainBranch.objectId,
                name: this.state.createTagTitle,
                message: this.state.createTagDescription,
            };
            getClient(GitRestClient)
                .createAnnotatedTag(
                    tag as GitAnnotatedTag,
                    this.state.repositoryListSelectedItemObservable.value
                        .project.id,
                    this.state.repositoryListSelectedItemObservable.value.id
                )
                .then(async (result) => {
                    this.globalMessagesSvc.addToast({
                        duration: 5000,
                        forceOverrideExisting: true,
                        message: 'Tag creation started!',
                    });
                    await this.reloadComponent();
                })
                .catch((error: any) => {
                    this.globalMessagesSvc.addToast({
                        duration: 5000,
                        forceOverrideExisting: true,
                        message: 'Tag creation failed!' + error,
                    });
                });
        }

        isCreatePullRequestDialogOpenObservable.value = false;
    }

    private onDismissCreateTagActionModal(): void {
        createTagTitleObservable.value = '';
        createTagDescriptionObservable.value = '';
        this.setState({
            createTagTitle: createTagTitleObservable.value,
            createTagDescription: createTagDescriptionObservable.value,
        });
        isCreateTagDialogOpenObservable.value = false;
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
                                {this.renderViewTagsModal()}
                                {this.renderCreatePullRequestActionModal()}
                                {this.renderDeleteBranchActionModal()}
                                {this.renderCompletePullRequestActionModal()}
                                {this.renderCreateTagActionModal()}
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
