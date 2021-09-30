import "./Pivot.scss";

import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";

import { showRootComponent } from "../../Common";

import { getClient } from "azure-devops-extension-api";
import { CoreRestClient, ProjectVisibility, TeamProjectReference } from "azure-devops-extension-api/Core";
import { GitRestClient, GitQueryCommitsCriteria, GitHistoryMode, GitBaseVersionDescriptor, GitTargetVersionDescriptor } from "azure-devops-extension-api/Git";

import { Table, ITableColumn, renderSimpleCell, renderSimpleCellValue } from "azure-devops-ui/Table";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";

interface IPivotContentState {
    projects?: ArrayItemProvider<TeamProjectReference>;
    columns: ITableColumn<any>[];
}

class PivotContent extends React.Component<{}, IPivotContentState> {

    constructor(props: {}) {
        super(props);

        this.state = {
            columns: [{
                id: "name",
                name: "Project",
                renderCell: renderSimpleCell,
                width: 200
            },
            {
                id: "description",
                name: "Description",
                renderCell: renderSimpleCell,
                width: 300
            },
            {
                id: "visibility",
                name: "Visibility",
                renderCell: (rowIndex: number, columnIndex: number, tableColumn: ITableColumn<TeamProjectReference>, tableItem: TeamProjectReference): JSX.Element => {
                    return renderSimpleCellValue<any>(columnIndex, tableColumn, tableItem.visibility === ProjectVisibility.Public ? "Public" : "Private");
                },
                width: 100
            }]
        };
    }

    public componentDidMount() {
        SDK.init();
        this.initializeComponent();
    }

    private async initializeComponent() {
        const projects = await getClient(CoreRestClient).getProjects();
        console.log('projects: ');
        console.log(projects);
        projects.forEach(async (project) => {
            const repos = await getClient(GitRestClient).getRepositories(project.id);
            console.log('repos: ');
            console.log(repos);
            const criteria: GitQueryCommitsCriteria = {
                $skip: 0,
                $top: 1000,
                author: "",
                compareVersion: {
                    version: "develop",
                    versionOptions: 0,
                    versionType: 0
                },
                excludeDeletes: false,
                fromCommitId: "",
                fromDate: "",
                historyMode: GitHistoryMode.FullHistorySimplifyMerges,
                ids: [],
                includeLinks: true,
                includePushData: false,
                includeUserImageUrl: false,
                includeWorkItems: false,
                itemPath: "",
                itemVersion: {
                    version: "master",
                    versionOptions: 0,
                    versionType: 0
                },
                toCommitId: "",
                toDate: "",
                user: ""
            }
            repos.forEach(async (repo) => {
                const commitsBatch = await getClient(GitRestClient).getCommitsBatch(criteria, repo.id, "", 0, 1000, true);
                console.log("getCommitsBatch: ");
                console.log(commitsBatch);

                const baseVersion: GitBaseVersionDescriptor = {
                    baseVersion: "master",
                    baseVersionOptions: 0,
                    baseVersionType: 0,
                    version: "master",
                    versionOptions: 0,
                    versionType: 0
                };
                const targetVersion: GitTargetVersionDescriptor = {
                    targetVersion: "develop",
                    targetVersionOptions: 0,
                    targetVersionType: 0,
                    version: "develop",
                    versionOptions: 0,
                    versionType: 0
                };

                const commitsDiff = await getClient(GitRestClient).getCommitDiffs(repo.id, undefined, undefined, 1000, 0, baseVersion, targetVersion);
                console.log("getCommitDiffs: ");
                console.log(commitsDiff);
            });

        });
        this.setState({
            projects: new ArrayItemProvider(projects)
        });
    }

    public render(): JSX.Element {
        return (
            <div className="sample-pivot">
                {
                    !this.state.projects &&
                    <p>Loading...</p>
                }
                {
                    this.state.projects &&
                    <Table
                        columns={this.state.columns}
                        itemProvider={this.state.projects}
                    />
                }
            </div>
        );
    }
}

showRootComponent(<PivotContent />);