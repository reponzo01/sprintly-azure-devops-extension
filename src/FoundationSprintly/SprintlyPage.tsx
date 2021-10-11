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
    GitBranchStats,
    GitCommitDiffs,
    GitMergeParameters,
    GitMerge,
    GitAsyncOperationStatus,
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
import { AllowedEntity } from './SprintlySettings';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import axios from 'axios';

export interface ISprintlyPageState {
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
// TODO: Extract this magic string
const selectedTabId: ObservableValue<string> = new ObservableValue<string>(
    'sprintly-page'
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
let repositoriesToProcess: string[] = [
    /*'repository-1.git',
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
    'fsl.systemnotifications.database',*/
];
let accessToken: string = '';

export class SprintlyPage extends React.Component<
    { loggedInUserDescriptor: string },
    ISprintlyPageState
> {
    private _dataManager?: IExtensionDataManager;
    private loggedInUserDescriptor: string;

    constructor(props: { loggedInUserDescriptor: string }) {
        super(props);

        this.state = {};
        this.loggedInUserDescriptor = props.loggedInUserDescriptor;
    }

    public async componentDidMount(): Promise<void> {
        await SDK.init();
        this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        // TODO: Extract this into methods to make more readable
        await SDK.ready();
        // TODO: Get this access token at the parent page and pass in as prop
        accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>(
            CommonServiceIds.ExtensionDataService
        );
        this._dataManager = await extDataService.getExtensionDataManager(
            SDK.getExtensionContext().id,
            accessToken
        );

        this._dataManager
            .getValue<AllowedEntity[]>(
                this.loggedInUserDescriptor.replace('.', '-') +
                    '-repositories-to-process'
            )
            .then(async (repositories) => {
                repositoriesToProcess = [];
                if (repositories) {
                    for (const repository of repositories) {
                        repositoriesToProcess.push(repository.originId);
                    }

                    if (repositoriesToProcess.length > 0) {
                        const projects: TeamProjectReference[] =
                            await getClient(CoreRestClient).getProjects();
                        // TODO: Limit projects to 'Portfolio' or 'Sample Project'
                        const reposExtended: GitRepositoryExtended[] = [];
                        projects.forEach(
                            async (project: TeamProjectReference) => {
                                const repos: GitRepository[] = await getClient(
                                    GitRestClient
                                ).getRepositories(project.id);
                                let filteredRepos: GitRepository[] = repos;
                                if (useFilteredRepos) {
                                    filteredRepos = repos.filter(
                                        (repo: GitRepository) =>
                                            repositoriesToProcess.includes(
                                                repo.id
                                            )
                                    );
                                }

                                totalRepositoriesToProcess.value =
                                    filteredRepos.length;

                                filteredRepos.forEach(
                                    async (repo: GitRepository) => {
                                        const refs: GitRef[] = await getClient(
                                            GitRestClient
                                        ).getRefs(
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
                                            if (
                                                ref.name.includes(
                                                    'heads/develop'
                                                )
                                            ) {
                                                hasDevelop = true;
                                            } else if (
                                                ref.name.includes(
                                                    'heads/master'
                                                )
                                            ) {
                                                hasMaster = true;
                                            } else if (
                                                ref.name.includes('heads/main')
                                            ) {
                                                hasMain = true;
                                            }
                                        }

                                        const processRepo: boolean =
                                            hasDevelop &&
                                            (hasMaster || hasMain);
                                        if (processRepo === true) {
                                            const baseVersion: GitBaseVersionDescriptor =
                                                {
                                                    baseVersion: hasMaster
                                                        ? 'master'
                                                        : 'main',
                                                    baseVersionOptions: 0,
                                                    baseVersionType: 0,
                                                    version: hasMaster
                                                        ? 'master'
                                                        : 'main',
                                                    versionOptions: 0,
                                                    versionType: 0,
                                                };
                                            const targetVersion: GitTargetVersionDescriptor =
                                                {
                                                    targetVersion: 'develop',
                                                    targetVersionOptions: 0,
                                                    targetVersionType: 0,
                                                    version: 'develop',
                                                    versionOptions: 0,
                                                    versionType: 0,
                                                };

                                            const commitsDiff: GitCommitDiffs =
                                                await getClient(
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
                                                Object.keys(
                                                    commitsDiff.changeCounts
                                                ).length === 0 &&
                                                commitsDiff.changes.length === 0
                                            ) {
                                                createRelease = false;
                                            }

                                            let existingReleaseName: string =
                                                '';
                                            let hasExistingRelease: boolean =
                                                false;
                                            refs.forEach((ref: GitRef) => {
                                                if (
                                                    ref.name.includes(
                                                        'heads/release'
                                                    )
                                                ) {
                                                    hasExistingRelease = true;
                                                    const refNameSplit: string[] =
                                                        ref.name.split(
                                                            'heads/'
                                                        );
                                                    existingReleaseName =
                                                        refNameSplit[1];
                                                }
                                            });

                                            reposExtended.push({
                                                _links: repo._links,
                                                defaultBranch:
                                                    repo.defaultBranch,
                                                id: repo.id,
                                                isFork: repo.isFork,
                                                name: repo.name,
                                                parentRepository:
                                                    repo.parentRepository,
                                                project: repo.project,
                                                remoteUrl: repo.remoteUrl,
                                                size: repo.size,
                                                sshUrl: repo.sshUrl,
                                                url: repo.url,
                                                validRemoteUrls:
                                                    repo.validRemoteUrls,
                                                webUrl: repo.webUrl,
                                                createRelease,
                                                hasExistingRelease,
                                                existingReleaseName,
                                                refs,
                                            });
                                        }

                                        this.setState({
                                            repositories: new ArrayItemProvider(
                                                reposExtended.sort(
                                                    (
                                                        a: GitRepositoryExtended,
                                                        b: GitRepositoryExtended
                                                    ) => {
                                                        return a.name.localeCompare(
                                                            b.name
                                                        );
                                                    }
                                                )
                                            ),
                                        });
                                    }
                                );
                            }
                        );
                    }
                }
            });
    }

    public render(): JSX.Element {
        const onDismiss: () => void = () => {
            isTagsDialogOpen.value = false;
        };
        return (
            // TODO: If repos to process = 0, add ZeroData page.
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
                                    Please select valid repositories from the{' '}
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

                            // axios
                            //     .post(
                            //         `https://dev.azure.com/reponzo01/${tableItem.project.id}/_apis/git/repositories/${tableItem.id}/merges?api-version=6.0-preview.1`,
                            //         {
                            //             comment:
                            //                 'Merging dev(p2) to main(p1) hopefully',
                            //             parents: [
                            //                 newMainObjectId,
                            //                 newDevObjectId,
                            //             ],
                            //         },
                            //         {
                            //             headers: {
                            //                 Authorization: `Bearer ${accessToken}`,
                            //             },
                            //         }
                            //     )
                            //     .then((res: any) => {
                            //         console.log(res);
                            //     })
                            //     .catch((error) => {
                            //         console.error(error);
                            //     });

                            // createRefOptions.push({
                            //     repositoryId: tableItem.id,
                            //     name:
                            //         'refs/heads/main/',
                            //     isLocked: false,
                            //     newObjectId: "this will have to be the result of true async call above",
                            //     oldObjectId:
                            //     newMainObjectId,
                            // });
                            // const createRef = await getClient(
                            //     GitRestClient
                            // ).updateRefs(createRefOptions, tableItem.id);

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
                            // regular code
                            /*createRefOptions.push({
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
                            });*/
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

function onSelectedTabChanged(newTabId: string) {
    selectedTabId.value = newTabId;
}
