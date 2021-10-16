import './FoundationSprintly.scss';

import * as React from 'react';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    getClient,
    IExtensionDataManager,
    IExtensionDataService,
    IGlobalMessagesService,
} from 'azure-devops-extension-api';

import {
    CoreRestClient,
    TeamProjectReference,
} from 'azure-devops-extension-api/Core';
import {
    GitRestClient,
    GitBaseVersionDescriptor,
    GitTargetVersionDescriptor,
    GitRepository,
    GitRefUpdate,
    GitRef,
    GitCommitDiffs,
    GitRefUpdateStatus,
} from 'azure-devops-extension-api/Git';
import {
    Table,
    ITableColumn,
    SimpleTableCell,
    TwoLineTableCell,
} from 'azure-devops-ui/Table';
import { Link } from 'azure-devops-ui/Link';
import { Pill, PillVariant, PillSize } from 'azure-devops-ui/Pill';
import { Button } from 'azure-devops-ui/Button';
import { TextField } from 'azure-devops-ui/TextField';
import { IColor } from 'azure-devops-ui/Utilities/Color';
import { Spinner } from 'azure-devops-ui/Spinner';
import { Icon, IconSize } from 'azure-devops-ui/Icon';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { Dialog } from 'azure-devops-ui/Dialog';
import { SimpleList } from 'azure-devops-ui/List';
import { AllowedEntity, GitRepositoryExtended } from './FoundationSprintly';
import { ZeroData } from 'azure-devops-ui/ZeroData';

export interface ISprintlyPageState {
    repositories?: ArrayItemProvider<GitRepositoryExtended>;
}

const newReleaseBranchNamesObservable: Array<ObservableValue<string>> = [];
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
        id: 'releaseNeeded',
        name: 'Release Needed?',
        onSize,
        renderCell: renderReleaseNeeded,
        width: new ObservableValue(-30),
    },
    {
        id: 'tags',
        name: 'Tags',
        onSize,
        renderCell: renderTags,
        width: new ObservableValue(-30),
    },
    {
        id: 'createReleaseBranch',
        name: 'Create Release Branch',
        renderCell: renderCreateReleaseBranch,
        width: new ObservableValue(-40),
    },
];

const useFilteredRepos: boolean = true;
const repositoriesToProcessKey: string = 'repositories-to-process';
let repositoriesToProcess: string[] = [];

// TODO: Clean up arrow functions for the cases in which I thought I
// couldn't use regular functions because the this.* was undefined errors.
// The solution is to bind those functions to `this` in the constructor.
// See SprintlyPostRelease as an example.
export default class SprintlyPage extends React.Component<
    {
        dataManager: IExtensionDataManager;
    },
    ISprintlyPageState
> {
    private dataManager: IExtensionDataManager;

    constructor(props: { dataManager: IExtensionDataManager }) {
        super(props);

        this.state = {};
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
        await this.loadRepositoriesToProcess();
    }

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
        });
    }

    private async getRepositoryInfo(repoId: string): Promise<GitRef[]> {
        return await getClient(GitRestClient).getRefs(
            repoId,
            undefined,
            undefined,
            false,
            false,
            undefined,
            true,
            false,
            undefined
        );
    }

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

    private codeChangesInCommitDiffs(commitsDiff: GitCommitDiffs): boolean {
        return (
            Object.keys(commitsDiff.changeCounts).length > 0 ||
            commitsDiff.changes.length > 0
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
                            <div className="page-content page-content-top flex-column rhythm-vertical-16">
                                {!this.state.repositories && (
                                    <div className="flex-row">
                                        <Spinner label="loading" />
                                    </div>
                                )}
                                {this.state.repositories && (
                                    <Table
                                        columns={columns}
                                        itemProvider={this.state.repositories}
                                    />
                                )}
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
                            imageAltText="No repositories"
                            imagePath={'../static/notfound.png'}
                        />
                    );
                }}
            </Observer>

            /* tslint:disable */
        );
    }
}

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
                    <Icon
                        ariaLabel="Repository"
                        iconName="Repo"
                        size={IconSize.large}
                    />
                    {' '}
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

function renderReleaseNeeded(
    rowIndex: number,
    columnIndex: number,
    tableColumn: ITableColumn<GitRepositoryExtended>,
    tableItem: GitRepositoryExtended
): JSX.Element {
    const redColor: IColor = {
        red: 191,
        green: 65,
        blue: 65,
    };
    const greenColor: IColor = {
        red: 109,
        green: 210,
        blue: 109,
    };
    const orangeColor: IColor = {
        red: 225,
        green: 172,
        blue: 74,
    };
    let color: IColor = redColor;
    let text: string = 'No';
    if (tableItem.createRelease === true) {
        color = greenColor;
        text = 'Yes';
    }
    if (tableItem.hasExistingRelease === true) {
        color = orangeColor;
        text = 'Release Exists';
    }
    if (tableItem.hasExistingRelease) {
        const releaseLinks: JSX.Element[] =
            tableItem.existingReleaseNames.map<JSX.Element>(
                (release, index) => (
                    <Link
                        key={index}
                        excludeTabStop
                        href={
                            tableItem.webUrl +
                            '?version=GB' +
                            encodeURI(release)
                        }
                        target="_blank"
                    >
                        {release}
                    </Link>
                )
            );

        return (
            <TwoLineTableCell
                className={'flex-direction-col'}
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                line1={
                    <>
                        <Pill
                            color={color}
                            size={PillSize.large}
                            variant={PillVariant.colored}
                            iconProps={{ iconName: 'Warning' }}
                            className="bolt-list-overlay"
                        >
                            <b>{text}</b>
                        </Pill>
                    </>
                }
                line2={<>{releaseLinks}</>}
            ></TwoLineTableCell>
        );
    }
    return (
        <SimpleTableCell
            key={'col-' + columnIndex}
            columnIndex={columnIndex}
            tableColumn={tableColumn}
            children={
                <>
                    <Pill
                        color={color}
                        size={PillSize.large}
                        variant={PillVariant.colored}
                        className="bolt-list-overlay"
                    >
                        <b>{text}</b>
                    </Pill>
                </>
            }
        ></SimpleTableCell>
    );
}

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

function renderCreateReleaseBranch(
    rowIndex: number,
    columnIndex: number,
    tableColumn: ITableColumn<GitRepositoryExtended>,
    tableItem: GitRepositoryExtended
): JSX.Element {
    newReleaseBranchNamesObservable[rowIndex] = new ObservableValue<string>('');
    return (
        <SimpleTableCell
            key={'col-' + columnIndex}
            columnIndex={columnIndex}
            tableColumn={tableColumn}
            children={
                <>
                    release /&nbsp;
                    <TextField
                        value={newReleaseBranchNamesObservable[rowIndex]}
                        onChange={(e, newValue) =>
                            (newReleaseBranchNamesObservable[rowIndex].value =
                                newValue.trim())
                        }
                    />
                    &nbsp;
                    <Button
                        text="Create Branch"
                        iconProps={{ iconName: 'OpenSource' }}
                        primary={true}
                        onClick={async () => {
                            const createRefOptions: GitRefUpdate[] = [];
                            const developBranch = await getClient(
                                GitRestClient
                            ).getBranch(tableItem.id, 'develop');

                            const newDevObjectId =
                                developBranch.commit.commitId;

                            createRefOptions.push({
                                repositoryId: tableItem.id,
                                name:
                                    'refs/heads/release/' +
                                    newReleaseBranchNamesObservable[rowIndex]
                                        .value,
                                isLocked: false,
                                newObjectId: newDevObjectId,
                                oldObjectId:
                                    '0000000000000000000000000000000000000000',
                            });
                            const createRef = await getClient(
                                GitRestClient
                            ).updateRefs(createRefOptions, tableItem.id);

                            newReleaseBranchNamesObservable[rowIndex].value =
                                '';
                            createRef.forEach(async (ref) => {
                                const globalMessagesSvc =
                                    await SDK.getService<IGlobalMessagesService>(
                                        CommonServiceIds.GlobalMessagesService
                                    );
                                globalMessagesSvc.addToast({
                                    duration: 5000,
                                    forceOverrideExisting: true,
                                    message: ref.success
                                        ? 'Branch Created!'
                                        : 'Error Creating Branch: ' +
                                          GitRefUpdateStatus[ref.updateStatus],
                                });
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
