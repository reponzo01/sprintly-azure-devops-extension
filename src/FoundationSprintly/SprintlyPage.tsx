import './FoundationSprintly.scss';

import * as React from 'react';

import {
    getClient,
    IExtensionDataManager,
    IGlobalMessagesService,
    IProjectInfo,
} from 'azure-devops-extension-api';

import {
    GitRestClient,
    GitBaseVersionDescriptor,
    GitTargetVersionDescriptor,
    GitRepository,
    GitRefUpdate,
    GitCommitDiffs,
    GitRefUpdateStatus,
    GitBranchStats,
    GitRefUpdateResult,
} from 'azure-devops-extension-api/Git';
import {
    Table,
    ITableColumn,
    SimpleTableCell,
    TwoLineTableCell,
} from 'azure-devops-ui/Table';
import { Pill, PillVariant, PillSize } from 'azure-devops-ui/Pill';
import { Button } from 'azure-devops-ui/Button';
import { TextField } from 'azure-devops-ui/TextField';
import { IColor } from 'azure-devops-ui/Utilities/Color';
import { Spinner } from 'azure-devops-ui/Spinner';
import { Icon, IconSize } from 'azure-devops-ui/Icon';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { ZeroData } from 'azure-devops-ui/ZeroData';

import * as Common from './SprintlyCommon';
import { TagsModal, ITagsModalContent, getTagsModalContent } from './TagsModal';
import { Card } from 'azure-devops-ui/Card';
import { Link } from 'azure-devops-ui/Link';

export interface ISprintlyPageState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    repositories?: ArrayItemProvider<Common.IGitRepositoryExtended>;
}

//#region "Observables"
const newReleaseBranchNamesObservable: Array<ObservableValue<string>> = [];
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
//#endregion "Observables"

let repositoriesToProcess: string[] = [];

// TODO: Clean up arrow functions for the cases in which I thought I
// couldn't use regular functions because the this.* was undefined errors.
// The solution is to bind those functions to `this` in the constructor.
// See SprintlyPostRelease as an example.
export default class SprintlyPage extends React.Component<
    {
        accessToken: string;
        globalMessagesSvc: IGlobalMessagesService;
    },
    ISprintlyPageState
> {
    private dataManager!: IExtensionDataManager;
    private accessToken: string;
    private globalMessagesSvc: IGlobalMessagesService;

    private columns: any = [];

    constructor(props: {
        accessToken: string;
        globalMessagesSvc: IGlobalMessagesService;
    }) {
        super(props);

        this.onSize = this.onSize.bind(this);
        this.renderCreateReleaseBranchCell =
            this.renderCreateReleaseBranchCell.bind(this);

        this.columns = [
            {
                id: 'name',
                name: 'Repository',
                onSize: this.onSize,
                renderCell: this.renderNameCell,
                width: new ObservableValue<number>(-30),
            },
            {
                id: 'releaseNeeded',
                name: 'Release Needed?',
                onSize: this.onSize,
                renderCell: this.renderReleaseNeededCell,
                width: new ObservableValue<number>(-30),
            },
            {
                id: 'tags',
                name: 'Tags',
                onSize: this.onSize,
                renderCell: this.renderTagsCell,
                width: new ObservableValue<number>(-30),
            },
            {
                id: 'createReleaseBranch',
                name: 'Create Release Branch',
                renderCell: this.renderCreateReleaseBranchCell,
                width: new ObservableValue<number>(-40),
            },
        ];

        this.state = {};
        this.accessToken = props.accessToken;
        this.globalMessagesSvc = props.globalMessagesSvc;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        this.dataManager = await Common.initializeDataManager(
            await Common.getOrRefreshToken(this.accessToken)
        );
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
            await this.loadRepositoriesDisplayState(
                await Common.getCurrentProject()
            );
        }
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

            totalRepositoriesToProcessObservable.value = filteredRepos.length;

            for (const repo of filteredRepos) {
                const repositoryBranchInfo: Common.IRepositoryBranchInfo =
                    await Common.getRepositoryBranchesInfo(
                        repo.id,
                        Common.repositoryHeadsFilter
                    );

                const processRepo: boolean =
                    repositoryBranchInfo.hasDevelopBranch &&
                    (repositoryBranchInfo.hasMasterBranch ||
                        repositoryBranchInfo.hasMainBranch);

                if (processRepo === true) {
                    const createRelease: boolean =
                        await this.isDevelopAheadOfMasterMain(
                            repositoryBranchInfo,
                            repo.id
                        );

                    const existingReleaseBranches: Common.IReleaseBranchInfo[] =
                        [];
                    for (const releaseBranch of repositoryBranchInfo.releaseBranches) {
                        existingReleaseBranches.push({
                            targetBranch: releaseBranch,
                            repositoryId: repo.id,
                        });
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
                        createRelease,
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
                    Common.sortRepositoryExtendedList(reposExtended)
                ),
            });
        }
    }

    private async isDevelopAheadOfMasterMain(
        repositoryBranchInfo: Common.IRepositoryBranchInfo,
        repositoryId: string
    ): Promise<boolean> {
        const baseVersion: GitBaseVersionDescriptor = {
            baseVersion: repositoryBranchInfo.hasMasterBranch
                ? 'master'
                : 'main',
            baseVersionOptions: 0,
            baseVersionType: 0,
            version: repositoryBranchInfo.hasMasterBranch ? 'master' : 'main',
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

        const commitsDiff: GitCommitDiffs = await Common.getCommitDiffs(
            repositoryId,
            baseVersion,
            targetVersion
        );

        return Common.codeChangesInCommitDiffs(commitsDiff);
    }

    private tagsModal(): JSX.Element {
        return (
            <Observer
                isTagsDialogOpen={isTagsDialogOpenObservable}
                tagsRepoName={tagsRepoNameObservable}
                tags={tagsObservable}
                tagsModalKey={tagsModalKeyObservable}
            >
                {(props: {
                    isTagsDialogOpen: boolean;
                    tagsRepoName: string;
                    tags: string[];
                    tagsModalKey: string;
                }) => {
                    return (
                        <TagsModal
                            key={props.tagsModalKey}
                            isTagsDialogOpen={props.isTagsDialogOpen}
                            tagsRepoName={props.tagsRepoName}
                            tags={props.tags}
                            closeMe={() => {
                                isTagsDialogOpenObservable.value = false;
                            }}
                        ></TagsModal>
                    );
                }}
            </Observer>
        );
    }

    private renderNameCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.IGitRepositoryExtended>,
        tableItem: Common.IGitRepositoryExtended
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    <>
                        <Icon
                            ariaLabel='Repository'
                            iconName='Repo'
                            size={IconSize.large}
                        />{' '}
                        {Common.repositoryLinkJsxElement(
                            tableItem.webUrl,
                            '',
                            tableItem.name
                        )}
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderReleaseNeededCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.IGitRepositoryExtended>,
        tableItem: Common.IGitRepositoryExtended
    ): JSX.Element {
        let color: IColor = Common.redColor;
        let text: string = 'No Release Needed';
        let viewChangesUrl: string = `${tableItem.webUrl}/branchCompare?baseVersion=GB${tableItem.hasMainBranch ? 'main' : 'master'}&targetVersion=GBdevelop&_a=files`;

        const compareChangesLink = <Link
            key='compareChangesLink'
            excludeTabStop
            href={viewChangesUrl}
            target='_blank'
        >
            Compare Changes
        </Link>

        if (tableItem.createRelease === true) {
            color = Common.greenColor;
            text = 'Yes, Release Needed';
        }
        if (tableItem.hasExistingRelease === true) {
            color = tableItem.createRelease === true
                ? Common.drakOrangeColor
                : Common.orangeColor;
            text = tableItem.createRelease === true
                ? 'Release Exists, but there are changes'
                : 'Release Exists';
        }
        if (tableItem.hasExistingRelease) {
            const lineTwoLinks: JSX.Element[] = [];
            let counter: number = 0;
            for (const releaseBranch of tableItem.existingReleaseBranches) {
                const releaseBranchName: string =
                    releaseBranch.targetBranch.name.split('heads/')[1];
                lineTwoLinks.push(
                    Common.branchLinkJsxElement(
                        counter + 'link',
                        tableItem.webUrl,
                        `${releaseBranchName}?${releaseBranch.targetBranch.creator.displayName}`,
                        ''
                    )
                );
                counter++;
            }
            if (tableItem.createRelease === true) {
                lineTwoLinks.push(compareChangesLink);
            }

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
                                className='bolt-list-overlay'
                            >
                                <b>{text}</b>
                            </Pill>
                        </>
                    }
                    line2={<>{lineTwoLinks}</>}
                ></TwoLineTableCell>
            );
        }

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
                            className='bolt-list-overlay'
                        >
                            <b>{text}</b>
                        </Pill>
                    </>
                }
                line2={tableItem.createRelease === true ? compareChangesLink : <></>}
            ></TwoLineTableCell>
        );
    }

    private renderTagsCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.IGitRepositoryExtended>,
        tableItem: Common.IGitRepositoryExtended
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    <>
                        <Button
                            text='View Tags'
                            subtle={true}
                            iconProps={{ iconName: 'Tag' }}
                            onClick={() => {
                                tagsModalKeyObservable.value = new Date()
                                    .getTime()
                                    .toString();
                                isTagsDialogOpenObservable.value = true;
                                tagsRepoNameObservable.value =
                                    'Loading tags...';
                                tagsObservable.value = [];
                                getTagsModalContent(
                                    tableItem.name,
                                    tableItem.id
                                ).then((modalContent: ITagsModalContent) => {
                                    tagsRepoNameObservable.value =
                                        modalContent.modalName;
                                    tagsObservable.value =
                                        modalContent.modalValues;
                                });
                            }}
                        />
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderCreateReleaseBranchCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.IGitRepositoryExtended>,
        tableItem: Common.IGitRepositoryExtended
    ): JSX.Element {
        newReleaseBranchNamesObservable[rowIndex] = new ObservableValue<string>(
            ''
        );
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
                            onChange={(
                                e: React.ChangeEvent<
                                    HTMLInputElement | HTMLTextAreaElement
                                >,
                                newValue: string
                            ) =>
                                (newReleaseBranchNamesObservable[
                                    rowIndex
                                ].value = newValue.trim())
                            }
                        />
                        &nbsp;
                        <Observer
                            enteredValue={
                                newReleaseBranchNamesObservable[rowIndex]
                            }
                        >
                            {() => {
                                return (
                                    <Button
                                        disabled={
                                            newReleaseBranchNamesObservable[
                                                rowIndex
                                            ].value.trim() === ''
                                        }
                                        text='Create Branch'
                                        iconProps={{ iconName: 'OpenSource' }}
                                        primary={true}
                                        onClick={async () => {
                                            const createRefOptions: GitRefUpdate[] =
                                                [];
                                            const developBranch: GitBranchStats =
                                                await getClient(
                                                    GitRestClient
                                                ).getBranch(
                                                    tableItem.id,
                                                    'develop'
                                                );

                                            const newDevObjectId: string =
                                                developBranch.commit.commitId;

                                            createRefOptions.push({
                                                repositoryId: tableItem.id,
                                                name:
                                                    'refs/heads/release/' +
                                                    newReleaseBranchNamesObservable[
                                                        rowIndex
                                                    ].value,
                                                isLocked: false,
                                                newObjectId: newDevObjectId,
                                                oldObjectId:
                                                    '0000000000000000000000000000000000000000',
                                            });
                                            const createRef: GitRefUpdateResult[] =
                                                await getClient(
                                                    GitRestClient
                                                ).updateRefs(
                                                    createRefOptions,
                                                    tableItem.id
                                                );

                                            newReleaseBranchNamesObservable[
                                                rowIndex
                                            ].value = '';
                                            createRef.forEach(
                                                async (
                                                    ref: GitRefUpdateResult
                                                ) => {
                                                    this.globalMessagesSvc.addToast(
                                                        {
                                                            duration: 5000,
                                                            forceOverrideExisting:
                                                                true,
                                                            message: ref.success
                                                                ? 'Branch Created!'
                                                                : 'Error Creating Branch: ' +
                                                                  GitRefUpdateStatus[
                                                                      ref
                                                                          .updateStatus
                                                                  ],
                                                        }
                                                    );
                                                    await this.initializeComponent();
                                                }
                                            );
                                        }}
                                    />
                                );
                            }}
                        </Observer>
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (this.columns[index].width as ObservableValue<number>).value = width;
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
                        return !this.state.repositories ? (
                            <Spinner label='loading' />
                        ) : (
                            <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                                {this.state.repositories && (
                                    <Card className='bolt-table-card bolt-card-white'>
                                        <Table
                                            columns={this.columns}
                                            itemProvider={
                                                this.state.repositories
                                            }
                                        />
                                    </Card>
                                )}
                                {this.tagsModal()}
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
                            imageAltText='No repositories'
                            imagePath={'../static/notfound.png'}
                        />
                    );
                }}
            </Observer>
        );
    }
}
