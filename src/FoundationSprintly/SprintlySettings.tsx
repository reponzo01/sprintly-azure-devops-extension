import * as React from 'react';
import axios, { AxiosResponse } from 'axios';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    getClient,
    IExtensionDataManager,
    IExtensionDataService,
    IGlobalMessagesService,
} from 'azure-devops-extension-api';

import { Button } from 'azure-devops-ui/Button';
import { Dropdown } from 'azure-devops-ui/Dropdown';
import { Observer } from 'azure-devops-ui/Observer';
import { DropdownMultiSelection } from 'azure-devops-ui/Utilities/DropdownSelection';
import { ISelectionRange } from 'azure-devops-ui/Utilities/Selection';
import {
    CoreRestClient,
    TeamProjectReference,
} from 'azure-devops-extension-api/Core';

import { GitRepository, GitRestClient } from 'azure-devops-extension-api/Git';

import { IAllowedEntity } from './FoundationSprintly';

const allowedUserGroupsKey: string = 'allowed-user-groups';
const allowedUsersKey: string = 'allowed-users';
const repositoriesToProcessKey: string = 'repositories-to-process';

export interface ISprintlySettingsState {
    dataAllowedUserGroups?: IAllowedEntity[];
    dataAllowedUsers?: IAllowedEntity[];
    dataRepositoriesToProcess?: IAllowedEntity[];

    persistedAllowedUserGroups?: IAllowedEntity[];
    persistedAllowedUsers?: IAllowedEntity[];
    persistedRepositoriesToProcess?: IAllowedEntity[];

    ready?: boolean;
}

// TODO: Clean up arrow functions for the cases in which I thought I
// couldn't use regular functions because the this.* was undefined errors.
// The solution is to bind those functions to `this` in the constructor.
// See SprintlyPostRelease as an example.
export default class SprintlySettings extends React.Component<
    {
        organizationName: string;
        dataManager?: IExtensionDataManager;
    },
    ISprintlySettingsState
> {
    private userGroupsSelection: DropdownMultiSelection =
        new DropdownMultiSelection();
    private usersSelection: DropdownMultiSelection =
        new DropdownMultiSelection();
    private repositoriesToProcessSelection: DropdownMultiSelection =
        new DropdownMultiSelection();

    private allUserGroups: IAllowedEntity[] = [];
    private allUsers: IAllowedEntity[] = [];
    private allRepositories: IAllowedEntity[] = [];

    private dataManager: IExtensionDataManager;
    private accessToken: string = '';
    private organizationName: string;

    constructor(props: { organizationName: string; dataManager: IExtensionDataManager }) {
        super(props);

        this.state = {};
        this.organizationName = props.organizationName;
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

        await this.getGroups();
        await this.getUsers();
        await this.getRepositories();

        this.setState({ ready: true });

        this.loadAllowedUserGroupsUsers();
        this.loadAllowedUsers();
        this.loadRepositoriesToProcess();
    }

    private async getGraphResource(
        resouce: string,
        callback: (data: any) => void
    ): Promise<void> {
        axios
            .get(
                `https://vssps.dev.azure.com/${this.organizationName}/_apis/graph/${resouce}`,
                {
                    headers: {
                        Authorization: `Bearer ${this.accessToken}`,
                    },
                }
            )
            .then((res: AxiosResponse<never>) => {
                callback(res.data);
            })
            .catch((error: any) => {
                console.error(error);
                throw error;
            });
    }

    private async getGroups(): Promise<void> {
        return new Promise(
            (resolve: (value: void | PromiseLike<void>) => void) => {
                this.getGraphResource('groups', (data: any) => {
                    this.allUserGroups = [];
                    for (const group of data.value) {
                        this.allUserGroups.push({
                            displayName: group.displayName,
                            originId: group.originId,
                            descriptor: group.descriptor,
                        });
                    }
                    resolve();
                });
            }
        );
    }

    private async getUsers(): Promise<void> {
        return new Promise(
            (resolve: (value: void | PromiseLike<void>) => void) => {
                this.getGraphResource('users', (data: any) => {
                    this.allUsers = [];
                    for (const user of data.value) {
                        this.allUsers.push({
                            displayName: user.displayName,
                            originId: user.originId,
                            descriptor: user.descriptor,
                        });
                    }
                    resolve();
                });
            }
        );
    }

    private async getRepositories(): Promise<void> {
        return new Promise(
            async (resolve: (value: void | PromiseLike<void>) => void) => {
                this.allRepositories = [];
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
                for (const project of filteredProjects) {
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
            }
        );
    }

    private loadAllowedUserGroupsUsers(): void {
        this.dataManager!.getValue<IAllowedEntity[]>(allowedUserGroupsKey).then(
            (userGroups: IAllowedEntity[]) => {
                this.userGroupsSelection.clear();
                if (userGroups) {
                    for (const selectedUserGroup of userGroups) {
                        const idx: number = this.allUserGroups.findIndex(
                            (item: IAllowedEntity) =>
                                item.originId === selectedUserGroup.originId
                        );
                        if (idx >= 0) {
                            this.userGroupsSelection.select(idx);
                        }
                    }
                    this.setState({
                        dataAllowedUserGroups: userGroups,
                        persistedAllowedUserGroups: userGroups,
                        ready: true,
                    });
                }
            },
            () => {
                this.setState({
                    dataAllowedUserGroups: [],
                    ready: true,
                });
            }
        );
    }

    private loadAllowedUsers(): void {
        this.dataManager!.getValue<IAllowedEntity[]>(allowedUsersKey).then(
            (users: IAllowedEntity[]) => {
                this.usersSelection.clear();
                if (users) {
                    for (const selectedUser of users) {
                        const idx: number = this.allUsers.findIndex(
                            (user: IAllowedEntity) =>
                                user.originId === selectedUser.originId
                        );
                        if (idx >= 0) {
                            this.usersSelection.select(idx);
                        }
                    }
                    this.setState({
                        dataAllowedUsers: users,
                        persistedAllowedUsers: users,
                        ready: true,
                    });
                }
            },
            () => {
                this.setState({
                    dataAllowedUsers: [],
                    ready: true,
                });
            }
        );
    }

    private loadRepositoriesToProcess(): void {
        this.dataManager!.getValue<IAllowedEntity[]>(repositoriesToProcessKey, {
            scopeType: 'User',
        }).then(
            (repositories: IAllowedEntity[]) => {
                this.repositoriesToProcessSelection.clear();
                if (repositories) {
                    for (const selectedRepository of repositories) {
                        const idx: number = this.allRepositories.findIndex(
                            (repository: IAllowedEntity) =>
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
                }
            },
            () => {
                this.setState({
                    dataRepositoriesToProcess: [],
                    ready: true,
                });
            }
        );
    }

    private onSaveData = (): void => {
        this.setState({ ready: false });

        const userGroupsSelectedArray: IAllowedEntity[] = this.setSelectionRange(
            this.userGroupsSelection.value,
            this.allUserGroups
        );
        const usersSelectedArray: IAllowedEntity[] = this.setSelectionRange(
            this.usersSelection.value,
            this.allUsers
        );
        const repositoriesSelectedArray: IAllowedEntity[] =
            this.setSelectionRange(
                this.repositoriesToProcessSelection.value,
                this.allRepositories
            );

        this.dataManager!.setValue<IAllowedEntity[]>(
            allowedUserGroupsKey,
            userGroupsSelectedArray || []
        ).then(() => {
            this.dataManager!.setValue<IAllowedEntity[]>(
                allowedUsersKey,
                usersSelectedArray || []
            ).then(() => {
                this.dataManager!.setValue<IAllowedEntity[]>(
                    repositoriesToProcessKey,
                    repositoriesSelectedArray || [],
                    { scopeType: 'User' }
                ).then(async () => {
                    this.setState({
                        ready: true,
                        persistedRepositoriesToProcess:
                            repositoriesSelectedArray,
                        persistedAllowedUsers: usersSelectedArray,
                        persistedAllowedUserGroups: userGroupsSelectedArray,
                    });
                    const globalMessagesSvc: IGlobalMessagesService =
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
        dataArray: IAllowedEntity[]
    ): IAllowedEntity[] {
        const selectedArray: IAllowedEntity[] = [];
        for (const rng of selectionRange) {
            const sliced: IAllowedEntity[] = dataArray.slice(
                rng.beginIndex,
                rng.endIndex + 1
            );
            for (const slic of sliced) {
                selectedArray.push(slic);
            }
        }
        return selectedArray;
    }

    private renderUserGroupsDropdown(): JSX.Element {
        return (
            /* tslint:disable */
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
            /* tslint:disable */
        );
    }

    private renderUsersDropdown(): JSX.Element {
        return (
            /* tslint:disable */
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
            /* tslint:disable */
        );
    }

    private renderRepositoriesDropdown(): JSX.Element {
        return (
            /* tslint:disable */
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
                                        iconProps: { iconName: 'Accept' },
                                        text: 'Select All',
                                        onClick: () => {
                                            this.repositoriesToProcessSelection.select(
                                                0,
                                                this.allRepositories.length
                                            );
                                        },
                                    },
                                    {
                                        className:
                                            'bolt-dropdown-action-right-button',
                                        disabled:
                                            this.repositoriesToProcessSelection
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
                                selection={this.repositoriesToProcessSelection}
                                placeholder="Select Individual Repositories"
                                showFilterBox={true}
                            />
                        );
                    }}
                </Observer>
            </div>
            /* tslint:disable */
        );
    }

    public render() {
        const { ready } = this.state;

        return (
            /* tslint:disable */
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
                {this.renderUserGroupsDropdown()}
                {this.renderUsersDropdown()}
                <div>
                    Select the repositories you want to process. This is a{' '}
                    user-based setting. Everyone with access to this extension{' '}
                    can select a different list.
                </div>
                {this.renderRepositoriesDropdown()}

                <div className="bolt-button-group flex-row rhythm-horizontal-8">
                    <Button
                        text="Save Settings"
                        primary={true}
                        onClick={this.onSaveData}
                        disabled={!ready}
                    />
                </div>
            </div>
            /* tslint:disable */
        );
    }
}
