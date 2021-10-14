import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';

import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import { getClient, IExtensionDataManager } from 'azure-devops-extension-api';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { AllowedEntity, GitRepositoryExtended } from './FoundationSprintly';
import { ITableColumn, SimpleTableCell } from 'azure-devops-ui/Table';
import { Icon } from 'azure-devops-ui/Icon';
import { Link } from 'azure-devops-ui/Link';
import { Button } from 'azure-devops-ui/Button';
import { CoreRestClient, TeamProjectReference } from 'azure-devops-extension-api/Core';

export interface ISprintlyPostReleaseState {
    repositories?: ArrayItemProvider<GitRepositoryExtended>;
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

const repositoriesToProcessKey: string = 'repositories-to-process';
let repositoriesToProcess: string[] = [];

export default class SprintlyPostRelease extends React.Component<
    { dataManager: IExtensionDataManager },
    ISprintlyPostReleaseState
> {
    private dataManager: IExtensionDataManager;
    private accessToken: string = '';

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
        this.accessToken = await SDK.getAccessToken();

        this.loadRepositoriesToProcess();
    }

    // TODO: This function is repeated in SprintlyPage. See about extracting.
    private loadRepositoriesToProcess(): void {
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
                    //this.loadRepositoriesDisplayState(filteredProjects);
                }
            }
        });
    }

    public render(): JSX.Element {
        return (
            /* tslint:disable */
            <Observer totalRepositoriesToProcess={totalRepositoriesToProcess}>
                {(props: { totalRepositoriesToProcess: number }) => {
                    if (totalRepositoriesToProcess.value > 0) {
                        return <></>;
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
