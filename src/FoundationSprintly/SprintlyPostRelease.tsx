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
    GitRef,
    GitRepository,
    GitRestClient,
    GitTargetVersionDescriptor,
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

import { AllowedEntity, GitRepositoryExtended } from './FoundationSprintly';
import { Pill, PillSize, PillVariant } from 'azure-devops-ui/Pill';

export interface ISprintlyPostReleaseState {
    repositories: ArrayItemProvider<GitRepositoryExtended>;
    selection: ListSelection;
    selectedItemObservable: ObservableValue<GitRepositoryExtended>;
}

const isTagsDialogOpen: ObservableValue<boolean> = new ObservableValue<boolean>(
    false
);
const tagsRepoName: ObservableValue<string> = new ObservableValue<string>('');
const tags: ObservableValue<string[]> = new ObservableValue<string[]>([]);
const totalRepositoriesToProcess: ObservableValue<number> =
    new ObservableValue<number>(0);

const columns: any = [
    {
        id: 'name',
        name: 'Repository',
        onSize,
        renderCell: renderName,
        width: new ObservableValue(-30),
    },
    {
        id: 'tags',
        name: 'Tags',
        onSize,
        renderCell: renderTags,
        width: new ObservableValue(-30),
    },
];

const useFilteredRepos: boolean = true;
const repositoriesToProcessKey: string = 'repositories-to-process';
let repositoriesToProcess: string[] = [];

// TODO: Clean up arrow functions for the cases in which I thought I
// couldn't use regular functions because the this.* was undefined errors.
// The solution is to bind those functions to `this` in the constructor.
// See SprintlyPostRelease as an example.
export default class SprintlyPostRelease extends React.Component<
    { dataManager: IExtensionDataManager },
    ISprintlyPostReleaseState
> {
    private dataManager: IExtensionDataManager;
    private accessToken: string = '';

    constructor(props: { dataManager: IExtensionDataManager }) {
        super(props);

        this.state = {
            repositories: new ArrayItemProvider<GitRepositoryExtended>([]),
            selection: new ListSelection({ selectOnFocus: false }),
            selectedItemObservable: new ObservableValue<any>({}),
        };

        this.renderRepositoryList = this.renderRepositoryList.bind(this);
        this.renderListItem = this.renderListItem.bind(this);
        this.renderDetailPage = this.renderDetailPage.bind(this);

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
        this.dataManager!.getValue<AllowedEntity[]>(repositoriesToProcessKey, {
            scopeType: 'User',
        }).then(async (repositories: AllowedEntity[]) => {
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
                    await this.loadRepositoriesDisplayState(filteredProjects);
                }
            }
        });
    }

    // TODO: This function is repeated in SprintlyPage. See about extracting.
    private async loadRepositoriesDisplayState(
        projects: TeamProjectReference[]
    ): Promise<void> {
        let reposExtended: GitRepositoryExtended[] = [];
        projects.forEach(async (project: TeamProjectReference) => {
            const repos: GitRepository[] = await getClient(
                GitRestClient
            ).getRepositories(project.id);
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

                let hasDevelop: boolean = false;
                let hasMaster: boolean = false;
                let hasMain: boolean = false;

                for (const ref of branchesAndTags) {
                    if (ref.name.includes('heads/develop')) {
                        hasDevelop = true;
                    } else if (ref.name.includes('heads/master')) {
                        hasMaster = true;
                    } else if (ref.name.includes('heads/main')) {
                        hasMain = true;
                    }
                }

                const processRepo: boolean =
                    hasDevelop && (hasMaster || hasMain);
                if (processRepo === true) {
                    const baseVersion: GitBaseVersionDescriptor = {
                        baseVersion: hasMaster ? 'master' : 'main',
                        baseVersionOptions: 0,
                        baseVersionType: 0,
                        version: hasMaster ? 'master' : 'main',
                        versionOptions: 0,
                        versionType: 0,
                    };
                    const targetVersion: GitTargetVersionDescriptor = {
                        targetVersion: 'develop',
                        targetVersionOptions: 0,
                        targetVersionType: 0,
                        version: 'develop',
                        versionOptions: 0,
                        versionType: 0,
                    };

                    const commitsDiff: GitCommitDiffs =
                        await this.getCommitDiffs(
                            repo.id,
                            baseVersion,
                            targetVersion
                        );

                    let createRelease: boolean = true;
                    if (!this.codeChangesInCommitDiffs(commitsDiff)) {
                        createRelease = false;
                    }

                    const existingReleaseNames: string[] = [];
                    let hasExistingRelease: boolean = false;
                    branchesAndTags.forEach((ref: GitRef) => {
                        if (ref.name.includes('heads/release')) {
                            hasExistingRelease = true;
                            const refNameSplit: string[] =
                                ref.name.split('heads/');
                            existingReleaseNames.push(refNameSplit[1]);
                        }
                    });

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
                        createRelease,
                        hasExistingRelease,
                        existingReleaseNames,
                        branchesAndTags,
                    });
                }
            }

            if (reposExtended.length > 0) {
                reposExtended = reposExtended.sort(
                    (a: GitRepositoryExtended, b: GitRepositoryExtended) => {
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
                    .selectedItemObservable as ObservableValue<GitRepositoryExtended>
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
                renderRow={this.renderListItem}
                width="100%"
                singleClickActivation={true}
            />
        );
    }

    private renderListItem(
        index: number,
        item: GitRepositoryExtended,
        details: IListItemDetails<GitRepositoryExtended>,
        key?: string
    ): JSX.Element {
        const primaryColor: IColor = {
            red: 0,
            green: 90,
            blue: 158,
        };
        const primaryColorShade30: IColor = {
            red: 0,
            green: 69,
            blue: 120,
        };
        const releaseLinks: JSX.Element[] =
            item.existingReleaseNames.map<JSX.Element>((release, index) => (
                <div className="flex-row padding-vertical-10" key={index}>
                    <Link
                        excludeTabStop
                        href={item.webUrl + '?version=GB' + encodeURI(release)}
                        subtle={false}
                        target="_blank"
                        className="padding-horizontal-6"
                    >
                        {release}
                    </Link>
                    <Pill
                        color={primaryColor}
                        size={PillSize.regular}
                        className="bolt-list-overlay margin-horizontal-3"
                    >
                        Ahead of develop{' '}
                        <i>
                            <Icon
                                ariaLabel="Pull Request"
                                iconName="BranchPullRequest"
                                size={IconSize.small}
                            />{' '}
                            #8542
                        </i>
                    </Pill>
                    <Pill
                        color={primaryColorShade30}
                        size={PillSize.regular}
                        className="bolt-list-overlay margin-horizontal-3"
                        variant={PillVariant.outlined}
                    >
                        Ahead of master{' '}
                        <i>
                            <Icon
                                ariaLabel="Pull Request"
                                iconName="BranchPullRequest"
                                size={IconSize.small}
                            />{' '}
                            #8542
                        </i>
                    </Pill>
                </div>
            ));
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
                {(observerProps: { selectedItem: GitRepositoryExtended }) => (
                    <Page className="flex-grow single-layer-details">
                        {this.state.selection.selectedCount == 0 && (
                            <span className="single-layer-details-contents">
                                Select a repository on the right to get started.
                            </span>
                        )}
                        {observerProps.selectedItem &&
                            this.state.selection.selectedCount > 0 && (
                                <Tooltip
                                    text={observerProps.selectedItem.name}
                                    overflowOnly={true}
                                >
                                    <span className="single-layer-details-contents">
                                        {observerProps.selectedItem.name} This
                                        is the Detail Page
                                    </span>
                                </Tooltip>
                            )}
                    </Page>
                )}
            </Observer>
        );
    }

    public render(): JSX.Element {
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

// TODO: This function is repeated in SprintlyPage. See about extracting.
function renderName(
    rowIndex: number,
    columnIndex: number,
    tableColumn: ITableColumn<GitRepositoryExtended>,
    tableItem: GitRepositoryExtended
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
    tableColumn: ITableColumn<GitRepositoryExtended>,
    tableItem: GitRepositoryExtended
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

function onSize(event: MouseEvent, index: number, width: number) {
    (columns[index].width as ObservableValue<number>).value = width;
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
