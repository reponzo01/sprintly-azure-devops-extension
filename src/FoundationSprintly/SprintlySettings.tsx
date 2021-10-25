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
import {
    Splitter,
    SplitterDirection,
    SplitterElementPosition,
} from 'azure-devops-ui/Splitter';

const userSettingsDataManagerKey: string = 'user-settings';
const systemSettingsDataManagerKey: string = 'system-settings';

export interface ISprintlySettingsState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;

    ready?: boolean;
}

// TODO: Clean up arrow functions for the cases in which I thought I
// couldn't use regular functions because the this.* was undefined errors.
// The solution is to bind those functions to `this` in the constructor.
// See SprintlyPostRelease as an example.
export default class SprintlySettings extends React.Component<
    {
        organizationName: string;
        globalMessagesSvc: IGlobalMessagesService;
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
    private globalMessagesSvc: IGlobalMessagesService;
    private accessToken: string = '';
    private organizationName: string;

    constructor(props: {
        organizationName: string;
        globalMessagesSvc: IGlobalMessagesService;
        dataManager: IExtensionDataManager;
    }) {
        super(props);

        this.state = {};

        this.renderUserSettings = this.renderUserSettings.bind(this);
        this.renderSystemSettings = this.renderSystemSettings.bind(this);

        this.organizationName = props.organizationName;
        this.globalMessagesSvc = props.globalMessagesSvc;
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

        const userSettings: Common.IUserSettings | undefined =
            await Common.getUserSettings(
                this.dataManager,
                userSettingsDataManagerKey
            );
        const systemSettings: Common.ISystemSettings | undefined =
            await Common.getSystemSettings(
                this.dataManager,
                systemSettingsDataManagerKey
            );

        this.setState({
            userSettings: userSettings,
            systemSettings: systemSettings,
            ready: true,
        });

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
        const userGroups: Common.IAllowedEntity[] | undefined =
            this.state.systemSettings?.allowedUserGroups;
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
        }
        this.setState({
            ready: true,
        });
    }

    private loadAllowedUsers(): void {
        const users: Common.IAllowedEntity[] | undefined =
            this.state.systemSettings?.allowedUsers;
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
        }
        this.setState({
            ready: true,
        });
    }

    private loadRepositoriesToProcess(): void {
        this.repositoriesToProcessSelection.clear();
        if (this.state.userSettings?.myRepositories) {
            for (const selectedRepository of this.state.userSettings
                .myRepositories) {
                const idx: number = this.allRepositories.findIndex(
                    (repository: Common.IAllowedEntity) =>
                        repository.originId === selectedRepository.originId
                );
                if (idx >= 0) {
                    this.repositoriesToProcessSelection.select(idx);
                }
            }
        }
        this.setState({
            ready: true,
        });
    }

    private onSaveUserSettingsData = (): void => {
        this.setState({ ready: false });

        const repositoriesSelectedArray: Common.IAllowedEntity[] =
            this.getSelectedRange(
                this.repositoriesToProcessSelection.value,
                this.allRepositories
            );

        let userSettings: Common.IUserSettings;

        if (this.state.userSettings) {
            userSettings = this.state.userSettings;
            userSettings.myRepositories = repositoriesSelectedArray;
        } else {
            userSettings = {
                myRepositories: repositoriesSelectedArray,
                projectRepositoriesKey: '',
            };
        }

        this.dataManager!.setValue<Common.IUserSettings>(
            userSettingsDataManagerKey,
            userSettings,
            { scopeType: 'User' }
        ).then(() => {
            this.setState({
                userSettings: userSettings,
                ready: true,
            });
            this.globalMessagesSvc.addToast({
                duration: 3000,
                forceOverrideExisting: true,
                message: 'User Settings saved successfully!',
            });
        });
    };

    private onSaveSystemSettingsData = (): void => {
        this.setState({ ready: false });

        const userGroupsSelectedArray: Common.IAllowedEntity[] =
            this.getSelectedRange(
                this.userGroupsSelection.value,
                this.allUserGroups
            );
        const usersSelectedArray: Common.IAllowedEntity[] =
            this.getSelectedRange(this.usersSelection.value, this.allUsers);

        let systemSettings: Common.ISystemSettings;

        if (this.state.systemSettings) {
            systemSettings = this.state.systemSettings;
            systemSettings.allowedUserGroups = userGroupsSelectedArray;
            systemSettings.allowedUsers = usersSelectedArray;
            systemSettings.projectRepositories = []; // TODO: blank for now but will save this once I create this setting
        } else {
            systemSettings = {
                allowedUserGroups: userGroupsSelectedArray,
                allowedUsers: usersSelectedArray,
                projectRepositories: [], // TODO: blank for now but will save this once I create this setting
            };
        }

        this.dataManager!.setValue<Common.ISystemSettings>(
            systemSettingsDataManagerKey,
            systemSettings
        ).then(() => {
            this.setState({
                systemSettings: systemSettings,
                ready: true,
            });
            this.globalMessagesSvc.addToast({
                duration: 3000,
                forceOverrideExisting: true,
                message: 'System Settings saved successfully!',
            });
        });
    };

    private getSelectedRange(
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

    private renderUserSettings(): JSX.Element {
        return (
            <Page>
                <Header
                    title='User Settings'
                    titleSize={TitleSize.Medium}
                    titleIconProps={{
                        iconName: 'Contact',
                        tooltipProps: {
                            text: 'These settings affect just you ',
                        },
                    }}
                    commandBarItems={[
                        {
                            iconProps: {
                                iconName: 'Save',
                            },
                            id: 'savesuserettings',
                            important: true,
                            text: 'Save User Settings',
                            isPrimary: true,
                            onActivate: this.onSaveUserSettingsData,
                            disabled: !this.state.ready,
                        },
                    ]}
                />
                <div className='page-content page-content-top'>
                    <Card className='bolt-card-white'>
                        <Page className='sprintly-width-100'>
                            <Header
                                title='My Repositories'
                                titleSize={TitleSize.Medium}
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
                </div>
            </Page>
        );
    }

    private renderSystemSettings(): JSX.Element {
        return (
            <Page>
                <Header
                    title='System Settings'
                    titleSize={TitleSize.Medium}
                    titleIconProps={{
                        iconName: 'People',
                        tooltipProps: {
                            text: 'These settings affect all users',
                        },
                    }}
                    commandBarItems={[
                        {
                            iconProps: {
                                iconName: 'Save',
                            },
                            id: 'savessystemettings',
                            important: true,
                            text: 'Save System Settings',
                            isPrimary: true,
                            onActivate: this.onSaveSystemSettingsData,
                            disabled: !this.state.ready,
                        },
                    ]}
                />
                <div className='page-content page-content-top'>
                    <Card className='bolt-card-white'>
                        <Page className='sprintly-width-100'>
                            <Header
                                title='Permissions'
                                titleSize={TitleSize.Medium}
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
                                to add more groups or individual users.
                            </div>
                            {this.renderUserGroupsDropdown()}
                            {this.renderUsersDropdown()}
                        </Page>
                    </Card>
                </div>
            </Page>
        );
    }

    public render(): JSX.Element {
        return (
            <Page>
                <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                    <Splitter
                        fixedElement={SplitterElementPosition.Near}
                        splitterDirection={SplitterDirection.Vertical}
                        nearElementClassName='v-scroll-auto custom-scrollbar'
                        farElementClassName='v-scroll-auto custom-scrollbar'
                        onRenderNearElement={this.renderUserSettings}
                        onRenderFarElement={this.renderSystemSettings}
                    />
                </div>
            </Page>
        );
    }
}
