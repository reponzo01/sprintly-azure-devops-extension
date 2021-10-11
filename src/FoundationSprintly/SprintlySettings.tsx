import * as React from 'react';

import axios from 'axios';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    getClient,
    IExtensionDataManager,
    IExtensionDataService,
    IGlobalMessagesService,
} from 'azure-devops-extension-api';
import {
    GraphRestClient,
    GraphSubject,
    GraphSubjectQuery,
} from 'azure-devops-extension-api/Graph';

import { Button } from 'azure-devops-ui/Button';
import { TextField } from 'azure-devops-ui/TextField';
import { Dropdown } from 'azure-devops-ui/Dropdown';
import { Observer } from 'azure-devops-ui/Observer';
import { DropdownMultiSelection } from 'azure-devops-ui/Utilities/DropdownSelection';
import { ISelectionRange } from 'azure-devops-ui/Utilities/Selection';
import {
    CoreRestClient,
    TeamProjectReference,
} from 'azure-devops-extension-api/Core';
import { resolveTypeReferenceDirective } from 'typescript';

import { GitRepository, GitRestClient } from 'azure-devops-extension-api/Git';

// TODO: Remove all this extension data from here and pass in as prop from Parent page
export interface AllowedEntity {
    displayName: string;
    originId: string;
    descriptor?: string;
}

export interface ISprintlySettingsState {
    dataAllowedUserGroups?: AllowedEntity[];
    dataAllowedUsers?: AllowedEntity[];
    dataRepositoriesToProcess?: AllowedEntity[];

    persistedAllowedUserGroups?: AllowedEntity[];
    persistedAllowedUsers?: AllowedEntity[];
    persistedRepositoriesToProcess?: AllowedEntity[];

    ready?: boolean;
}

export default class SprintlySettings extends React.Component<
    { sampleProp: string; loggedInUserDescriptor: string },
    ISprintlySettingsState
> {
    private userGroupsSelection = new DropdownMultiSelection();
    private usersSelection = new DropdownMultiSelection();
    private repositoriesToProcessSelection = new DropdownMultiSelection();

    private allUserGroups: AllowedEntity[] = [];
    private allUsers: AllowedEntity[] = [];
    private allRepositories: AllowedEntity[] = [];

    private _dataManager?: IExtensionDataManager;
    private sampleProp: string;
    private loggedInUserDescriptor: string;

    constructor(props: { sampleProp: string; loggedInUserDescriptor: string }) {
        super(props);

        this.state = {};
        this.sampleProp = props.sampleProp;
        this.loggedInUserDescriptor = props.loggedInUserDescriptor;
    }

    public async componentDidMount() {
        await SDK.init();
        this.initializeState();
        console.log(SDK.getUser());
    }

    private async initializeState(): Promise<void> {
        await SDK.ready();
        // TODO: Get this access token at the parent page and pass in as prop
        const accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>(
            CommonServiceIds.ExtensionDataService
        );
        this._dataManager = await extDataService.getExtensionDataManager(
            SDK.getExtensionContext().id,
            accessToken
        );

        await this.getGroups(accessToken);
        await this.getUsers(accessToken);
        await this.getRepositories(accessToken);

        this.setState({ ready: true });

        // TODO: Extract these into their own methods
        this._dataManager.getValue<AllowedEntity[]>('allowed-user-groups').then(
            (userGroups) => {
                console.log('data is this ', userGroups);
                this.userGroupsSelection.clear();
                for (const selectedUserGroup of userGroups) {
                    console.log(
                        'searching the user gruops for ',
                        selectedUserGroup
                    );
                    const idx = this.allUserGroups.findIndex(
                        (item) => item.originId === selectedUserGroup.originId
                    );
                    if (idx >= 0) {
                        this.userGroupsSelection.select(idx);
                    }
                    console.log('would have selected gruop ', idx);
                }
                this.setState({
                    dataAllowedUserGroups: userGroups,
                    persistedAllowedUserGroups: userGroups,
                    ready: true,
                });
            },
            () => {
                this.setState({
                    dataAllowedUserGroups: [],
                    ready: true,
                });
            }
        );

        this._dataManager.getValue<AllowedEntity[]>('allowed-users').then(
            (users) => {
                this.usersSelection.clear();
                for (const selectedUser of users) {
                    const idx = this.allUsers.findIndex(
                        (user) => user.originId === selectedUser.originId
                    );
                    if (idx >= 0) {
                        this.usersSelection.select(idx);
                    }
                    console.log('would have selected user ', idx);
                }
                this.setState({
                    dataAllowedUsers: users,
                    persistedAllowedUsers: users,
                    ready: true,
                });
            },
            () => {
                this.setState({
                    dataAllowedUsers: [],
                    ready: true,
                });
            }
        );

        // TODO: Extract magic strings
        this._dataManager
            .getValue<AllowedEntity[]>(
                this.loggedInUserDescriptor.replace('.', '-') +
                    '-repositories-to-process'
            )
            .then(
                (repositories) => {
                    this.repositoriesToProcessSelection.clear();
                    for (const selectedRepository of repositories) {
                        const idx = this.allRepositories.findIndex(
                            (repository) =>
                                repository.originId ===
                                selectedRepository.originId
                        );
                        if (idx >= 0) {
                            this.repositoriesToProcessSelection.select(idx);
                        }
                    }
                    this.setState({
                        dataRepositoriesToProcess: repositories,
                        persistedRepositoriesToProcess: repositories,
                        ready: true,
                    });
                },
                () => {
                    this.setState({
                        dataRepositoriesToProcess: [],
                        ready: true,
                    });
                }
            );
    }

    private async getGraphResource(
        resouce: string,
        accessToken: string,
        callback: (data: any) => void
    ): Promise<void> {
        // TODO: extract the organization name globally
        axios
            .get(
                `https://vssps.dev.azure.com/reponzo01/_apis/graph/${resouce}`,
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                    },
                }
            )
            .then((res) => {
                console.log(res.data);
                callback(res.data);
            })
            .catch((error) => {
                console.error(error);
                throw error;
            });
    }

    private async getGroups(accessToken: string): Promise<void> {
        return new Promise((resolve) => {
            this.getGraphResource('groups', accessToken, (data: any) => {
                this.allUserGroups = [];
                for (const group in data.value) {
                    this.allUserGroups.push({
                        displayName: data.value[group].displayName,
                        originId: data.value[group].originId,
                        descriptor: data.value[group].descriptor,
                    });
                }
                console.log('resolving getGroups with ', this.allUserGroups);
                resolve();
            });
        });
    }

    private async getUsers(accessToken: string): Promise<void> {
        return new Promise((resolve) => {
            this.getGraphResource('users', accessToken, (data: any) => {
                this.allUsers = [];
                for (const user in data.value) {
                    this.allUsers.push({
                        displayName: data.value[user].displayName,
                        originId: data.value[user].originId,
                        descriptor: data.value[user].descriptor,
                    });
                }
                console.log('resolving getUsers with ', this.allUsers);
                resolve();
            });
        });
    }

    private async getRepositories(accessToken: string): Promise<void> {
        return new Promise(async (resolve) => {
            this.allRepositories = [];
            const projects: TeamProjectReference[] = await getClient(
                CoreRestClient
            ).getProjects();
            // TODO: Limit project to 'Portfolio' or 'Sample Project'
            for (const project of projects) {
                const repos: GitRepository[] = await getClient(
                    GitRestClient
                ).getRepositories(project.id);
                repos.forEach((repo: GitRepository) => {
                    this.allRepositories.push({
                        originId: repo.id,
                        displayName: repo.name,
                    });
                });
            }
            resolve();
        });
    }

    public render() {
        const {
            dataAllowedUserGroups,
            dataAllowedUsers,
            ready,
            persistedAllowedUserGroups,
            persistedAllowedUsers,
        } = this.state;

        console.log('returning a render');
        return (
            <div className="page-content page-content-top flex-column rhythm-vertical-16">
                <div>
                    By default the Azure groups{' '}
                    <u>
                        <code>Dev Team Leads</code>
                    </u>{' '}
                    and{' '}
                    <u>
                        <code>DevOps</code>
                    </u>{' '}
                    have access to this extension. Use the dropdowns to add more{' '}
                    groups or individual users. These two settings are global{' '}
                    settings.
                </div>
                <div className="flex-column">
                    <Observer selection={this.userGroupsSelection}>
                        {() => {
                            return (
                                <Dropdown
                                    ariaLabel="Multiselect"
                                    actions={[
                                        {
                                            className:
                                                'bolt-dropdown-action-right-button',
                                            disabled:
                                                this.userGroupsSelection
                                                    .selectedCount === 0,
                                            iconProps: { iconName: 'Clear' },
                                            text: 'Clear',
                                            onClick: () => {
                                                this.userGroupsSelection.clear();
                                            },
                                        },
                                    ]}
                                    className="example-dropdown flex-column"
                                    items={this.allUserGroups.map(
                                        (item) => item.displayName
                                    )}
                                    selection={this.userGroupsSelection}
                                    placeholder="Select User Groups"
                                    showFilterBox={true}
                                />
                            );
                        }}
                    </Observer>
                </div>
                <div className="flex-column">
                    <Observer selection={this.usersSelection}>
                        {() => {
                            return (
                                <Dropdown
                                    ariaLabel="Multiselect"
                                    actions={[
                                        {
                                            className:
                                                'bolt-dropdown-action-right-button',
                                            disabled:
                                                this.usersSelection
                                                    .selectedCount === 0,
                                            iconProps: { iconName: 'Clear' },
                                            text: 'Clear',
                                            onClick: () => {
                                                this.usersSelection.clear();
                                            },
                                        },
                                    ]}
                                    className="example-dropdown flex-column"
                                    items={this.allUsers.map(
                                        (item) => item.displayName
                                    )}
                                    selection={this.usersSelection}
                                    placeholder="Select Individual Users"
                                    showFilterBox={true}
                                />
                            );
                        }}
                    </Observer>
                </div>
                <div>
                    Select the repositories you want to process. This is a{' '}
                    user-based setting. Everyone with access to this extension{' '}
                    can select a different list.
                </div>
                <div className="flex-column">
                    <Observer selection={this.repositoriesToProcessSelection}>
                        {() => {
                            return (
                                <Dropdown
                                    ariaLabel="Multiselect"
                                    actions={[
                                        {
                                            className:
                                                'bolt-dropdown-action-right-button',
                                            disabled:
                                                this
                                                    .repositoriesToProcessSelection
                                                    .selectedCount === 0,
                                            iconProps: { iconName: 'Clear' },
                                            text: 'Clear',
                                            onClick: () => {
                                                this.repositoriesToProcessSelection.clear();
                                            },
                                        },
                                    ]}
                                    className="example-dropdown flex-column"
                                    items={this.allRepositories.map(
                                        (item) => item.displayName
                                    )}
                                    selection={
                                        this.repositoriesToProcessSelection
                                    }
                                    placeholder="Select Individual Repositories"
                                    showFilterBox={true}
                                />
                            );
                        }}
                    </Observer>
                </div>
                <div className="bolt-button-group flex-row rhythm-horizontal-8">
                    <Button
                        text="Save"
                        primary={true}
                        onClick={this.onSaveData}
                        disabled={!ready}
                    />
                </div>
            </div>
        );
    }

    private onSaveData = (): void => {
        /*const {
            dataAllowedUserGroups,
            dataAllowedUsers,
            ready,
            persistedAllowedUserGroups,
            persistedAllowedUsers,
        } = this.state;*/

        this.setState({ ready: false });

        const userGroupsSelectedArray: AllowedEntity[] = this.setSelectionRange(
            this.userGroupsSelection.value,
            this.allUserGroups
        );
        const usersSelectedArray: AllowedEntity[] = this.setSelectionRange(
            this.usersSelection.value,
            this.allUsers
        );
        const repositoriesSelectedArray: AllowedEntity[] =
            this.setSelectionRange(
                this.repositoriesToProcessSelection.value,
                this.allRepositories
            );

        // TODO: Extract magic strings
        this._dataManager!.setValue<AllowedEntity[]>(
            'allowed-user-groups',
            userGroupsSelectedArray || []
        ).then(() => {
            this._dataManager!.setValue<AllowedEntity[]>(
                'allowed-users',
                usersSelectedArray || []
            ).then(() => {
                this._dataManager!.setValue<AllowedEntity[]>(
                    this.loggedInUserDescriptor.replace('.', '-') +
                        '-repositories-to-process',
                    repositoriesSelectedArray || []
                ).then(async () => {
                    this.setState({
                        ready: true,
                        persistedRepositoriesToProcess:
                            repositoriesSelectedArray,
                        persistedAllowedUsers: usersSelectedArray,
                        persistedAllowedUserGroups: userGroupsSelectedArray,
                    });
                    const globalMessagesSvc =
                        await SDK.getService<IGlobalMessagesService>(
                            CommonServiceIds.GlobalMessagesService
                        );
                    globalMessagesSvc.addToast({
                        duration: 3000,
                        forceOverrideExisting: true,
                        message: 'Settings saved successfully!',
                    });
                });
            });
        });
    };

    private setSelectionRange(
        selectionRange: ISelectionRange[],
        dataArray: AllowedEntity[]
    ): AllowedEntity[] {
        const selectedArray: AllowedEntity[] = [];
        for (const rng of selectionRange) {
            var sliced = dataArray.slice(rng.beginIndex, rng.endIndex + 1);
            for (const slic of sliced) {
                selectedArray.push(slic);
            }
        }
        return selectedArray;
    }
}
