import './Pivot.scss';

import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { showRootComponent } from '../../Common';

import {
    CommonServiceIds,
    getClient,
    IGlobalMessagesService,
    IHostNavigationService,
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
import { Icon } from 'azure-devops-ui/Icon';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { Dialog } from 'azure-devops-ui/Dialog';
import { SimpleList } from 'azure-devops-ui/List';

export interface IPivotContentState {
    projects?: ArrayItemProvider<TeamProjectReference>;
    repositories?: ArrayItemProvider<GitRepositoryExtended>;
}

export interface GitRepositoryExtended extends GitRepository {
    hasExistingRelease: boolean;
    existingReleaseName: string;
    createRelease: boolean;
    refs: GitRef[];
}

const newReleaseBranchNamesObservable: ObservableValue<string>[] = [];
const isTagsDialogOpen = new ObservableValue<boolean>(false);
const tagsRepoName = new ObservableValue<string>('');
const tags = new ObservableValue<string[]>([]);
const columns = [
    {
        id: 'name',
        name: 'Repository',
        onSize: onSize,
        renderCell: renderName,
        width: new ObservableValue(-30),
    },
    {
        id: 'releaseNeeded',
        name: 'Release Needed?',
        onSize: onSize,
        renderCell: renderReleaseNeeded,
        width: new ObservableValue(-30),
    },
    {
        id: 'tags',
        name: 'Tags',
        onSize: onSize,
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

export class PivotContent extends React.Component<{}, IPivotContentState> {
    constructor(props: {}) {
        super(props);

        this.state = {};
    }

    public componentDidMount() {
        SDK.init();
        this.initializeComponent();
    }

    private async initializeComponent() {
        const _this: this = this;
        const projects = await getClient(CoreRestClient).getProjects();
        console.log('projects: ');
        console.log(projects);
        const reposExtended: GitRepositoryExtended[] = [];
        projects.forEach(async (project) => {
            const repos = await getClient(GitRestClient).getRepositories(
                project.id
            );
            console.log('repos: ');
            console.log(repos);
            repos.forEach(async (repo) => {
                const refs = await getClient(GitRestClient).getRefs(
                    repo.id,
                    undefined,
                    undefined,
                    true,
                    true,
                    undefined,
                    undefined,
                    false,
                    undefined
                );
                console.log('refs: ');
                console.log(refs);
                const branches = await getClient(GitRestClient).getBranches(
                    repo.id
                );
                console.log('branches: ');
                console.log(branches);
                const processRepo: boolean = branches.some(
                    (branch) =>
                        branch.name.includes('develop') ||
                        branch.name.includes('master')
                );
                console.log('process repo: ', processRepo);
                if (processRepo === true) {
                    const baseVersion: GitBaseVersionDescriptor = {
                        baseVersion: 'master',
                        baseVersionOptions: 0,
                        baseVersionType: 0,
                        version: 'master',
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

                    const commitsDiff = await getClient(
                        GitRestClient
                    ).getCommitDiffs(
                        repo.id,
                        undefined,
                        undefined,
                        1000,
                        0,
                        baseVersion,
                        targetVersion
                    );
                    console.log('getCommitDiffs: ');
                    console.log(commitsDiff);

                    let createRelease: boolean = true;
                    if (
                        Object.keys(commitsDiff.changeCounts).length === 0 &&
                        commitsDiff.changes.length === 0
                    ) {
                        createRelease = false;
                    }

                    let existingReleaseName: string = '';
                    let hasExistingRelease: boolean = false;
                    branches.forEach((branch) => {
                        if (branch.name.includes('release')) {
                            hasExistingRelease = true;
                            existingReleaseName = branch.name;
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
                        createRelease: createRelease,
                        hasExistingRelease: hasExistingRelease,
                        existingReleaseName: existingReleaseName,
                        refs: refs,
                    });
                }

                _this.setState({
                    repositories: new ArrayItemProvider(
                        reposExtended.sort(
                            (
                                a: GitRepositoryExtended,
                                b: GitRepositoryExtended
                            ) => {
                                return a.name.localeCompare(b.name);
                            }
                        )
                    ),
                });
            });
        });
        this.setState({
            projects: new ArrayItemProvider(projects),
        });
    }

    public render(): JSX.Element {
        const onDismiss = () => {
            isTagsDialogOpen.value = false;
        };
        return (
            <div className="sample-pivot">
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
                                titleProps={{ text: props.tagsRepoName }}
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
                    <Icon ariaLabel="Repository" iconName="Repo" />
                    &nbsp;
                    <Link
                        excludeTabStop
                        href={tableItem.webUrl + '/branches'}
                        target="_blank"
                    >
                        {tableItem.name}
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
        red: 151,
        green: 30,
        blue: 79,
    };
    const greenColor: IColor = {
        red: 0,
        green: 255,
        blue: 0,
    };
    const orangeColor: IColor = {
        red: 255,
        green: 165,
        blue: 0,
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
        const releaseUrl = encodeURI(tableItem.existingReleaseName);
        return (
            <TwoLineTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                line1={
                    <>
                        <Pill
                            color={color}
                            size={PillSize.large}
                            variant={PillVariant.colored}
                        >
                            {text}
                        </Pill>
                    </>
                }
                line2={
                    <>
                        <Link
                            excludeTabStop
                            href={tableItem.webUrl + '?version=GB' + releaseUrl}
                            target="_blank"
                        >
                            {tableItem.existingReleaseName}
                        </Link>
                    </>
                }
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
                    >
                        {text}
                    </Pill>
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
                        primary={true}
                        onClick={async () => {
                            console.log(
                                'release/' +
                                    newReleaseBranchNamesObservable[rowIndex]
                                        .value
                            );
                            const createRefOptions: GitRefUpdate[] = [];
                            const developBranch = await getClient(
                                GitRestClient
                            ).getBranch(tableItem.id, 'develop');
                            const newObjectId = developBranch.commit.commitId;
                            createRefOptions.push({
                                repositoryId: tableItem.id,
                                name:
                                    'refs/heads/release/' +
                                    newReleaseBranchNamesObservable[rowIndex]
                                        .value,
                                isLocked: false,
                                newObjectId: newObjectId,
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
                                    duration: 3000,
                                    forceOverrideExisting: true,
                                    message: ref.success
                                        ? 'Branch Created!'
                                        : 'Error Creating Branch: ' +
                                          ref.customMessage,
                                });
                            });
                        }}
                    />
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
                        iconProps={{ iconName: 'Tag' }}
                        onClick={() => {
                            isTagsDialogOpen.value = true;
                            tagsRepoName.value = tableItem.name + ' Tags';
                            tags.value = [];
                            tableItem.refs.forEach((ref) => {
                                if (ref.name.includes('refs/tags')) {
                                    tags.value.push(ref.name);
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

showRootComponent(<PivotContent />);
