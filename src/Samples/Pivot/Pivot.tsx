import './Pivot.scss';

import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';

import { showRootComponent } from '../../Common';

import { getClient } from 'azure-devops-extension-api';
import {
    CoreRestClient,
    TeamProjectReference,
} from 'azure-devops-extension-api/Core';
import {
    GitRestClient,
    GitBaseVersionDescriptor,
    GitTargetVersionDescriptor,
    GitRepository,
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
import { Spinner, SpinnerSize } from 'azure-devops-ui/Spinner';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';

export interface IPivotContentState {
    projects?: ArrayItemProvider<TeamProjectReference>;
    repositories?: ArrayItemProvider<GitRepositoryExtended>;
    columns: ITableColumn<any>[];
}

export interface GitRepositoryExtended extends GitRepository {
    hasExistingRelease: boolean;
    existingReleaseName: string;
    newReleaseBranchName: string;
    createRelease: boolean;
    _this: any;
}

export class PivotContent extends React.Component<{}, IPivotContentState> {
    constructor(props: {}) {
        super(props);

        this.state = {
            columns: [
                {
                    id: 'name',
                    name: 'Repository',
                    renderCell: this.renderName,
                    width: 200,
                },
                {
                    id: 'createRelease',
                    name: 'Release Needed?',
                    renderCell: this.renderCreateRelease,
                    width: 150,
                },
                {
                    id: 'createReleaseBranch',
                    name: 'Create Release Branch',
                    renderCell: this.renderCreateReleaseBranch,
                    width: 500,
                },
            ],
        };
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
                        newReleaseBranchName: '',
                        _this: this,
                    });
                }

                _this.setState({
                    repositories: new ArrayItemProvider(reposExtended),
                });
            });
        });
        this.setState({
            projects: new ArrayItemProvider(projects),
        });
    }

    public render(): JSX.Element {
        return (
            <div className="sample-pivot">
                {!this.state.repositories && (
                    <div className="flex-row">
                        <Spinner label="loading" />
                    </div>
                )}
                {this.state.repositories && (
                    <Table
                        columns={this.state.columns}
                        itemProvider={this.state.repositories}
                    />
                )}
            </div>
        );
    }

    private renderName(
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

    private renderCreateRelease(
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
                                href={
                                    tableItem.webUrl +
                                    '?version=GB' +
                                    releaseUrl
                                }
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

    private renderCreateReleaseBranch(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<GitRepositoryExtended>,
        tableItem: GitRepositoryExtended
    ): JSX.Element {
        console.log(tableItem._this.state);
        // TODO: Just use local storage as the persistent storage
        tableItem._this.setState({
            repositories: null,
        });
        let newReleaseBranchName = tableItem.newReleaseBranchName;
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    <>
                        release /&nbsp;
                        <TextField
                            value={newReleaseBranchName}
                            onChange={(e) => {
                                console.log('changing text');
                                console.log(newReleaseBranchName);
                                newReleaseBranchName = e.target.value;
                            }}
                            disabled={false}
                        />
                        &nbsp;
                        <Button
                            text="Create Branch"
                            primary={true}
                            onClick={() => {
                                console.log(
                                    'release/' + newReleaseBranchName
                                );
                            }}
                        />
                    </>
                }
            ></SimpleTableCell>
        );
    }
}

showRootComponent(<PivotContent />);

