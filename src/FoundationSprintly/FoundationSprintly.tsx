import './FoundationSprintly.scss';

import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { showRootComponent } from '../Common';

import {
    CommonServiceIds,
    getClient,
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
    GitBranchStats,
    GitCommitDiffs,
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

export interface IFoundationSprintlyContentState {
    repositories?: ArrayItemProvider<GitRepositoryExtended>;
}

export interface GitRepositoryExtended extends GitRepository {
    hasExistingRelease: boolean;
    existingReleaseName: string;
    createRelease: boolean;
    refs: GitRef[];
}

const newReleaseBranchNamesObservable: Array<ObservableValue<string>> = [];
const isTagsDialogOpen: ObservableValue<boolean> = new ObservableValue<boolean>(
    false
);
const tagsRepoName: ObservableValue<string> = new ObservableValue<string>('');
const tags: ObservableValue<string[]> = new ObservableValue<string[]>([]);
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
const reposToProcess: string[] = [
    'repository-1.git',
    'repository-2.git',
    'repository-3.git',
    'fsi.myprojecthq.jobcosting.api',
    'fsi.myprojecthq.reports.api',
    'fsi.myprojecthq.reports.database',
    'fsi.myprojecthq.reports.web',
    'fsi.myprojecthq.reportsint.api',
    'fsi.myprojecthq.stimulsoftviewer.api',
    'fsi.myprojecthq.web',
    'fsi.pm.api',
    'fsi.pm.database',
    'fsi.pm.messageprocessor.functionapp',
    'fsi.pm.webhookemail.functionapp',
    'fsi.pmint.api',
    'fsl.auth.b2c.apim',
    'fsl.auth.b2c.appregistration',
    'fsl.auth.b2c.functionapp',
    'fsl.auth.b2c.userflows',
    'fsl.auth.b2c.migration',
    'fsl.authz.cosmosdb',
    'fsl.myprojecthq.apim',
    'fsl.myprojecthq.dailylogs.api',
    'fsl.myprojecthq.purchaseorders.api',
    'fsl.myprojecthq.reportdataint.api',
    'fsl.myprojecthq.reports.stimulsoftviewer.api',
    'fsl.myprojecthq.weatherautomation.functionapp',
    'fsl.systemnotifications.api',
    'fsl.systemnotifications.apim',
    'fsl.systemnotifications.database',
];

export class FoundationSprintlyContent extends React.Component<
    {},
    IFoundationSprintlyContentState
> {
    constructor(props: {}) {
        super(props);

        this.state = {};
    }

    public componentDidMount(): void {
        SDK.init();
        this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        const _this: this = this;
        const projects: TeamProjectReference[] = await getClient(
            CoreRestClient
        ).getProjects();
        const reposExtended: GitRepositoryExtended[] = [];
        projects.forEach(async (project: TeamProjectReference) => {
            const repos: GitRepository[] = await getClient(
                GitRestClient
            ).getRepositories(project.id);
            let filteredRepos: GitRepository[] = repos;
            if (useFilteredRepos) {
                filteredRepos = repos.filter((repo: GitRepository) =>
                    reposToProcess.includes(repo.name.toLowerCase())
                );
            }

            filteredRepos.forEach(async (repo: GitRepository) => {
                const refs: GitRef[] = await getClient(GitRestClient).getRefs(
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
                let hasDevelop: boolean = false;
                let hasMaster: boolean = false;
                let hasMain: boolean = false;

                for (const ref of refs) {
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

                    const commitsDiff: GitCommitDiffs = await getClient(
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

                    let createRelease: boolean = true;
                    if (
                        Object.keys(commitsDiff.changeCounts).length === 0 &&
                        commitsDiff.changes.length === 0
                    ) {
                        createRelease = false;
                    }

                    let existingReleaseName: string = '';
                    let hasExistingRelease: boolean = false;
                    refs.forEach((ref: GitRef) => {
                        if (ref.name.includes('heads/release')) {
                            hasExistingRelease = true;
                            const refNameSplit: string[] =
                                ref.name.split('heads/');
                            existingReleaseName = refNameSplit[1];
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
                        existingReleaseName,
                        refs,
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
    }

    public render(): JSX.Element {
        const onDismiss: () => void = () => {
            isTagsDialogOpen.value = false;
        };
        return (
            /* tslint:disable */
            <div className="foundation-sprintly">
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
                            iconProps={{ iconName: 'Warning' }}
                            className="bolt-list-overlay"
                        >
                            <b>{text}</b>
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
                        className="bolt-list-overlay"
                    >
                        <b>{text}</b>
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
                        iconProps={{ iconName: 'OpenSource' }}
                        primary={true}
                        onClick={async () => {
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
                        subtle={true}
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

showRootComponent(<FoundationSprintlyContent />);
