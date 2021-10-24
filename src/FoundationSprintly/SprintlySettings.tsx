import * as React from 'react';
import axios, { AxiosResponse } from 'axios';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    getClient,
    IExtensionDataManager,
    IGlobalMessagesService,
} from 'azure-devops-extension-api';

import { TeamProjectReference } from 'azure-devops-extension-api/Core';
import { GitRepository, GitRestClient } from 'azure-devops-extension-api/Git';

import { Button } from 'azure-devops-ui/Button';
import { Dropdown } from 'azure-devops-ui/Dropdown';
import { Observer } from 'azure-devops-ui/Observer';
import { DropdownMultiSelection } from 'azure-devops-ui/Utilities/DropdownSelection';
import { ISelectionRange } from 'azure-devops-ui/Utilities/Selection';

import * as Common from './SprintlyCommon';
import { Card } from 'azure-devops-ui/Card';
import { Page } from 'azure-devops-ui/Page';
import { Header, TitleSize, CustomHeader } from 'azure-devops-ui/Header';
import { HeaderCommandBar } from 'azure-devops-ui/HeaderCommandBar';

const allowedUserGroupsKey: string = 'allowed-user-groups';
const allowedUsersKey: string = 'allowed-users';
const repositoriesToProcessKey: string = 'repositories-to-process';

export interface ISprintlySettingsState {
    dataAllowedUserGroups?: Common.IAllowedEntity[];
    dataAllowedUsers?: Common.IAllowedEntity[];
    dataRepositoriesToProcess?: Common.IAllowedEntity[];

    persistedAllowedUserGroups?: Common.IAllowedEntity[];
    persistedAllowedUsers?: Common.IAllowedEntity[];
    persistedRepositoriesToProcess?: Common.IAllowedEntity[];

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

    private allUserGroups: Common.IAllowedEntity[] = [];
    private allUsers: Common.IAllowedEntity[] = [];
    private allRepositories: Common.IAllowedEntity[] = [];

    private dataManager: IExtensionDataManager;
    private accessToken: string = '';
    private organizationName: string;

    constructor(props: {
        organizationName: string;
        dataManager: IExtensionDataManager;
    }) {
        super(props);

        this.state = {};
        this.organizationName = props.organizationName;
        this.dataManager = props.dataManager;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        this.accessToken = await SDK.getAccessToken();

        await this.loadGroups();
        await this.loadUsers();
        await this.loadRepositories();

        this.setState({ ready: true });

        this.loadAllowedUserGroupsUsers();
        this.loadAllowedUsers();
        this.loadRepositoriesToProcess();
    }

    private async getGraphResource(resouce: string): Promise<any> {
        this.accessToken = await Common.getOrRefreshToken(this.accessToken);
        const response: AxiosResponse<never> = await axios
            .get(
                `https://vssps.dev.azure.com/${this.organizationName}/_apis/graph/${resouce}`,
                {
                    headers: {
                        Authorization: `Bearer ${this.accessToken}`,
                    },
                }
            )
            .catch((error: any) => {
                console.error(error);
                throw error;
            });
        return response.data;
    }

    private async loadGroups(): Promise<void> {
        this.allUserGroups = [];
        const data: any = await this.getGraphResource('groups');
        for (const group of data.value) {
            this.allUserGroups.push({
                displayName: group.displayName,
                originId: group.originId,
                descriptor: group.descriptor,
            });
        }
    }

    private async loadUsers(): Promise<void> {
        this.allUsers = [];
        const data: any = await this.getGraphResource('users');
        for (const user of data.value) {
            this.allUsers.push({
                displayName: user.displayName,
                originId: user.originId,
                descriptor: user.descriptor,
            });
        }
    }

    private async loadRepositories(): Promise<void> {
        this.allRepositories = [];
        const filteredProjects: TeamProjectReference[] =
            await Common.getFilteredProjects();
        for (const project of filteredProjects) {
            const repos: GitRepository[] = await getClient(
                GitRestClient
            ).getRepositories(project.id);
            for (const repo of repos) {
                this.allRepositories.push({
                    originId: repo.id,
                    displayName: repo.name,
                });
            }
        }
    }

    private loadAllowedUserGroupsUsers(): void {
        this.dataManager!.getValue<Common.IAllowedEntity[]>(
            allowedUserGroupsKey
        ).then(
            (userGroups: Common.IAllowedEntity[]) => {
                this.userGroupsSelection.clear();
                if (userGroups) {
                    for (const selectedUserGroup of userGroups) {
                        const idx: number = this.allUserGroups.findIndex(
                            (item: Common.IAllowedEntity) =>
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
        this.dataManager!.getValue<Common.IAllowedEntity[]>(
            allowedUsersKey
        ).then(
            (users: Common.IAllowedEntity[]) => {
                this.usersSelection.clear();
                if (users) {
                    for (const selectedUser of users) {
                        const idx: number = this.allUsers.findIndex(
                            (user: Common.IAllowedEntity) =>
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

    private async loadRepositoriesToProcess(): Promise<void> {
        const repositories: Common.IAllowedEntity[] =
            await Common.getSavedRepositoriesToProcess(
                this.dataManager,
                repositoriesToProcessKey
            );
        this.repositoriesToProcessSelection.clear();
        if (repositories) {
            for (const selectedRepository of repositories) {
                const idx: number = this.allRepositories.findIndex(
                    (repository: Common.IAllowedEntity) =>
                        repository.originId === selectedRepository.originId
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
        } else {
            this.setState({
                dataRepositoriesToProcess: [],
                ready: true,
            });
        }
    }

    private onSaveData = (): void => {
        this.setState({ ready: false });

        const userGroupsSelectedArray: Common.IAllowedEntity[] =
            this.setSelectionRange(
                this.userGroupsSelection.value,
                this.allUserGroups
            );
        const usersSelectedArray: Common.IAllowedEntity[] =
            this.setSelectionRange(this.usersSelection.value, this.allUsers);
        const repositoriesSelectedArray: Common.IAllowedEntity[] =
            this.setSelectionRange(
                this.repositoriesToProcessSelection.value,
                this.allRepositories
            );

        this.dataManager!.setValue<Common.IAllowedEntity[]>(
            allowedUserGroupsKey,
            userGroupsSelectedArray || []
        ).then(() => {
            this.dataManager!.setValue<Common.IAllowedEntity[]>(
                allowedUsersKey,
                usersSelectedArray || []
            ).then(() => {
                this.dataManager!.setValue<Common.IAllowedEntity[]>(
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
        dataArray: Common.IAllowedEntity[]
    ): Common.IAllowedEntity[] {
        const selectedArray: Common.IAllowedEntity[] = [];
        for (const rng of selectionRange) {
            const sliced: Common.IAllowedEntity[] = dataArray.slice(
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
            <div className='page-content'>
                <Observer selection={this.userGroupsSelection}>
                    {() => {
                        return (
                            <Dropdown
                                ariaLabel='Multiselect'
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
                                className='example-dropdown flex-column'
                                items={this.allUserGroups.map(
                                    (item: Common.IAllowedEntity) =>
                                        item.displayName
                                )}
                                selection={this.userGroupsSelection}
                                placeholder='Select User Groups'
                                showFilterBox={true}
                            />
                        );
                    }}
                </Observer>
            </div>
        );
    }

    private renderUsersDropdown(): JSX.Element {
        return (
            <div className='page-content'>
                <Observer selection={this.usersSelection}>
                    {() => {
                        return (
                            <Dropdown
                                ariaLabel='Multiselect'
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
                                className='example-dropdown flex-column'
                                items={this.allUsers.map(
                                    (item: Common.IAllowedEntity) =>
                                        item.displayName
                                )}
                                selection={this.usersSelection}
                                placeholder='Select Individual Users'
                                showFilterBox={true}
                            />
                        );
                    }}
                </Observer>
            </div>
        );
    }

    private renderRepositoriesDropdown(): JSX.Element {
        return (
            <div className='page-content'>
                <Observer selection={this.repositoriesToProcessSelection}>
                    {() => {
                        return (
                            <Dropdown
                                ariaLabel='Multiselect'
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
                                className='example-dropdown flex-column'
                                items={this.allRepositories.map(
                                    (item: Common.IAllowedEntity) =>
                                        item.displayName
                                )}
                                selection={this.repositoriesToProcessSelection}
                                placeholder='Select Individual Repositories'
                                showFilterBox={true}
                            />
                        );
                    }}
                </Observer>
            </div>
        );
    }

    public render(): JSX.Element {
        return (
            <Page>
                <Header commandBarItems={[
                            {
                                iconProps: {
                                    iconName: 'Save',
                                },
                                id: 'savesettings',
                                important: true,
                                text: 'Save Settings',
                                isPrimary: true,
                                onActivate: this.onSaveData,
                                disabled: !this.state.ready
                            },
                        ]}>
                    </Header>
                <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                    <Card className='bolt-card-white'>
                        <Page className='sprintly-width-100'>
                            <Header
                                title='Permissions'
                                titleSize={TitleSize.Medium}
                                titleIconProps={{ iconName: 'People' }}
                            />
                            <div className='page-content page-content-top'>
                                By default the Azure groups{' '}
                                <u>
                                    <code>Dev Team Leads</code>
                                </u>{' '}
                                and{' '}
                                <u>
                                    <code>DevOps</code>
                                </u>{' '}
                                have access to this extension. Use the dropdowns
                                to add more groups or individual users. These
                                two settings are global settings.
                            </div>
                            {this.renderUserGroupsDropdown()}
                            {this.renderUsersDropdown()}
                        </Page>
                    </Card>
                    <Card className='bolt-card-white'>
                        <Page className='sprintly-width-100'>
                            <Header
                                title='My Repositories'
                                titleSize={TitleSize.Medium}
                                titleIconProps={{ iconName: 'Contact' }}
                            />
                            <div className='page-content page-content-top'>
                                Select the repositories you want to process.
                                This is a user-based setting. Everyone with
                                access to this extension can select a different
                                list.
                            </div>
                            {this.renderRepositoriesDropdown()}
                        </Page>
                    </Card>

                    {/* <div className='bolt-button-group flex-row rhythm-horizontal-8'>
                        <Button
                            text='Save Settings'
                            primary={true}
                            onClick={this.onSaveData}
                            disabled={!this.state.ready}
                        />
                    </div> */}
                </div>
            </Page>
        );
    }
}
