import './FoundationSprintly.scss';

import * as React from 'react';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    getClient,
    IExtensionDataManager,
    IGlobalMessagesService,
} from 'azure-devops-extension-api';

import { TeamProjectReference } from 'azure-devops-extension-api/Core';
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

export interface ISprintlyPageState {
    repositories?: ArrayItemProvider<Common.IGitRepositoryExtended>;
}

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

const nameColumnWidthObservable: ObservableValue<number> = new ObservableValue<number>(-30);
const releaseNeededColumnWidthObservable: ObservableValue<number> = new ObservableValue<number>(-30);
const tagsColumnWidthObservable: ObservableValue<number> = new ObservableValue<number>(-30);
const createReleaseBranchColumnWidthObservable: ObservableValue<number> = new ObservableValue<number>(-40);

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

    public columns: any = [];

    constructor(props: { dataManager: IExtensionDataManager }) {
        super(props);

        this.onSize = this.onSize.bind(this);
        this.renderCreateReleaseBranchCell = this.renderCreateReleaseBranchCell.bind(this);

        this.columns = [
            {
                id: 'name',
                name: 'Repository',
                onSize: this.onSize,
                renderCell: this.renderNameCell,
                width: nameColumnWidthObservable,
            },
            {
                id: 'releaseNeeded',
                name: 'Release Needed?',
                onSize: this.onSize,
                renderCell: this.renderReleaseNeededCell,
                width: releaseNeededColumnWidthObservable,
            },
            {
                id: 'tags',
                name: 'Tags',
                onSize: this.onSize,
                renderCell: this.renderTagsCell,
                width: tagsColumnWidthObservable,
            },
            {
                id: 'createReleaseBranch',
                name: 'Create Release Branch',
                renderCell: this.renderCreateReleaseBranchCell,
                width: createReleaseBranchColumnWidthObservable,
            },
        ];

        this.state = {};
        this.dataManager = props.dataManager;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        repositoriesToProcess = (
            await Common.getSavedRepositoriesToProcess(
                this.dataManager,
                repositoriesToProcessKey
            )
        ).map((item: Common.IAllowedEntity) => item.originId);
        totalRepositoriesToProcessObservable.value =
            repositoriesToProcess.length;
        if (repositoriesToProcess.length > 0) {
            await this.loadRepositoriesDisplayState(
                await Common.getFilteredProjects()
            );
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
                    Common.sortRepositoryList(reposExtended)
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
                                    <Table
                                        columns={this.columns}
                                        itemProvider={this.state.repositories}
                                    />
                                )}
                                <Observer
                                    isTagsDialogOpen={
                                        isTagsDialogOpenObservable
                                    }
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
                                                isTagsDialogOpen={
                                                    props.isTagsDialogOpen
                                                }
                                                tagsRepoName={
                                                    props.tagsRepoName
                                                }
                                                tags={tagsObservable.value}
                                                closeMe={() => {
                                                    isTagsDialogOpenObservable.value =
                                                        false;
                                                }}
                                            ></TagsModal>
                                        );
                                    }}
                                </Observer>
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
        let text: string = 'No';
        if (tableItem.createRelease === true) {
            color = Common.greenColor;
            text = 'Yes';
        }
        if (tableItem.hasExistingRelease === true) {
            color = Common.orangeColor;
            text = 'Release Exists';
        }
        if (tableItem.hasExistingRelease) {
            const releaseBranchLinks: JSX.Element[] = [];
            let counter: number = 0;
            for (const releaseBranch of tableItem.existingReleaseBranches) {
                const releaseBranchName: string =
                    releaseBranch.targetBranch.name.split('heads/')[1];
                releaseBranchLinks.push(
                    Common.branchLinkJsxElement(
                        counter + 'link',
                        tableItem.webUrl,
                        releaseBranchName,
                        ''
                    )
                );
                counter++;
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
                    line2={<>{releaseBranchLinks}</>}
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
                            className='bolt-list-overlay'
                        >
                            <b>{text}</b>
                        </Pill>
                    </>
                }
            ></SimpleTableCell>
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
                                const modalContent: ITagsModalContent =
                                    getTagsModalContent(
                                        tableItem.name,
                                        tableItem.branchesAndTags
                                    );
                                tagsRepoNameObservable.value =
                                    modalContent.modalName;
                                tagsObservable.value = modalContent.modalValues;
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
                            onChange={(
                                e: React.ChangeEvent<
                                    HTMLInputElement | HTMLTextAreaElement
                                >,
                                newValue: string
                            ) =>
                                (newReleaseBranchNamesObservable[rowIndex].value =
                                    newValue.trim())
                            }
                        />
                        &nbsp;
                        <Button
                            text='Create Branch'
                            iconProps={{ iconName: 'OpenSource' }}
                            primary={true}
                            onClick={async () => {
                                const createRefOptions: GitRefUpdate[] = [];
                                const developBranch: GitBranchStats =
                                    await getClient(GitRestClient).getBranch(
                                        tableItem.id,
                                        'develop'
                                    );

                                const newDevObjectId: string =
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
                                const createRef: GitRefUpdateResult[] =
                                    await getClient(GitRestClient).updateRefs(
                                        createRefOptions,
                                        tableItem.id
                                    );

                                newReleaseBranchNamesObservable[rowIndex].value =
                                    '';
                                createRef.forEach(
                                    async (ref: GitRefUpdateResult) => {
                                        const globalMessagesSvc: IGlobalMessagesService =
                                            await SDK.getService<IGlobalMessagesService>(
                                                CommonServiceIds.GlobalMessagesService
                                            );
                                        globalMessagesSvc.addToast({
                                            duration: 5000,
                                            forceOverrideExisting: true,
                                            message: ref.success
                                                ? 'Branch Created!'
                                                : 'Error Creating Branch: ' +
                                                  GitRefUpdateStatus[
                                                      ref.updateStatus
                                                  ],
                                        });
                                        await this.initializeComponent();
                                    }
                                );
                            }}
                        />
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (this.columns[index].width as ObservableValue<number>).value = width;
    }
}
