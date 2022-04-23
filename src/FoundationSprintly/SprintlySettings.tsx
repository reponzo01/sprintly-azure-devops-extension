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
import { Header, TitleSize } from 'azure-devops-ui/Header';
import {
    Splitter,
    SplitterDirection,
    SplitterElementPosition,
} from 'azure-devops-ui/Splitter';
import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { FormItem } from 'azure-devops-ui/FormItem';
import { TextField, TextFieldStyle } from 'azure-devops-ui/TextField';
import { ITableColumn, SimpleTableCell, Table } from 'azure-devops-ui/Table';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { Dialog } from 'azure-devops-ui/Dialog';

export interface ISprintlySettingsState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    addProjectRepositoriesLabel?: string;
    projectLabelIdToDelete?: string;

    ready?: boolean;
}

const addProjectRepositoriesLabelObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const isDeleteProjectLabelDialogOpenObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);

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

    private projectRepositoriesTableColumns: any = [];

    constructor(props: {
        organizationName: string;
        globalMessagesSvc: IGlobalMessagesService;
        dataManager: IExtensionDataManager;
    }) {
        super(props);

        this.state = {};

        this.renderUserSettings = this.renderUserSettings.bind(this);
        this.renderSystemSettings = this.renderSystemSettings.bind(this);
        this.renderProjectRepositoriesTableRepositoriesCell =
            this.renderProjectRepositoriesTableRepositoriesCell.bind(this);
        this.renderProjectRepositoriesTableDeleteCell =
            this.renderProjectRepositoriesTableDeleteCell.bind(this);
        this.addProjectRepositoriesLabelAction =
            this.addProjectRepositoriesLabelAction.bind(this);
        this.deleteProjectLabelAction =
            this.deleteProjectLabelAction.bind(this);

        this.organizationName = props.organizationName;
        this.globalMessagesSvc = props.globalMessagesSvc;
        this.dataManager = props.dataManager;

        this.projectRepositoriesTableColumns = [
            {
                id: 'delete',
                name: 'Delete',
                renderCell: this.renderProjectRepositoriesTableDeleteCell,
                width: new ObservableValue(-10),
            },
            {
                id: 'projectLabel',
                name: 'Project Label',
                renderCell: this.renderProjectRepositoriesTableLabelCell,
                width: new ObservableValue(-20),
            },
            {
                id: 'repositories',
                name: 'Repositories',
                renderCell: this.renderProjectRepositoriesTableRepositoriesCell,
                width: new ObservableValue(-50),
            },
        ];
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
                Common.USER_SETTINGS_DATA_MANAGER_KEY
            );
        const systemSettings: Common.ISystemSettings | undefined =
            await Common.getSystemSettings(
                this.dataManager,
                Common.SYSTEM_SETTINGS_DATA_MANAGER_KEY
            );
        if (systemSettings && systemSettings.projectRepositories) {
            for (const item of systemSettings.projectRepositories) {
                item.selections = new DropdownMultiSelection();
            }
        }

        this.setState({
            userSettings,
            systemSettings,
            ready: true,
        });

        this.loadSystemSettingsValues();
        this.loadUserSettingsValues();

        this.setState({
            ready: true,
        });
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
        this.allUserGroups = Common.sortAllowedEntityList(this.allUserGroups);
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
        this.allUsers = Common.sortAllowedEntityList(this.allUsers);
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
            this.allRepositories = Common.sortAllowedEntityList(
                this.allRepositories
            );
        }
    }

    private loadSystemSettingsValues(): void {
        this.loadAllowedUserGroupsUsers();
        this.loadAllowedUsers();
        this.loadProjectRepositories();
    }

    private loadUserSettingsValues(): void {
        this.loadRepositoriesToProcess();
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
                if (idx > -1) {
                    this.userGroupsSelection.select(idx);
                }
            }
        }
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
                if (idx > -1) {
                    this.usersSelection.select(idx);
                }
            }
        }
    }

    private loadProjectRepositories(): void {
        if (this.state.systemSettings?.projectRepositories) {
            const systemSettings: Common.ISystemSettings =
                this.state.systemSettings;
            for (const projectRepository of systemSettings.projectRepositories) {
                for (const selectedRepository of projectRepository.repositories) {
                    const idx: number = this.allRepositories.findIndex(
                        (repository: Common.IAllowedEntity) =>
                            repository.originId === selectedRepository.originId
                    );
                    if (idx > -1) {
                        projectRepository.selections.select(idx);
                    }
                }
            }
            this.setState({
                systemSettings,
            });
        }
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
                if (idx > -1) {
                    this.repositoriesToProcessSelection.select(idx);
                }
            }
        }
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
                projectRepositoriesId: '',
            };
        }

        this.dataManager!.setValue<Common.IUserSettings>(
            Common.USER_SETTINGS_DATA_MANAGER_KEY,
            userSettings,
            { scopeType: 'User' }
        ).then(() => {
            this.setState({
                userSettings,
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

            const projectRepos: Common.IProjectRepositories[] =
                this.state.systemSettings.projectRepositories;
            for (const projectRepo of projectRepos) {
                const projectRepoSelectedArray: Common.IAllowedEntity[] =
                    this.getSelectedRange(
                        projectRepo.selections.value,
                        this.allRepositories
                    );
                projectRepo.repositories = projectRepoSelectedArray;
            }
            systemSettings.projectRepositories = projectRepos;
        } else {
            systemSettings = {
                allowedUserGroups: userGroupsSelectedArray,
                allowedUsers: usersSelectedArray,
                projectRepositories: [],
            };
        }

        this.dataManager!.setValue<Common.ISystemSettings>(
            Common.SYSTEM_SETTINGS_DATA_MANAGER_KEY,
            systemSettings
        ).then(() => {
            this.setState({
                systemSettings,
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
            for (const slice of sliced) {
                selectedArray.push(slice);
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
                                className='flex-column'
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
                                className='flex-column'
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

    private renderMyRepositoriesDropdown(): JSX.Element {
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
                                className='flex-column'
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

    private renderAddProjectRepositoriesLabel(): JSX.Element {
        return (
            <div className='page-content'>
                <Observer
                    addProjectRepositoriesKey={
                        addProjectRepositoriesLabelObservable
                    }
                >
                    {(observerProps: { addProjectRepositoriesKey: string }) => {
                        return (
                            <>
                                <FormItem label='Project Label *'>
                                    <div className='flex-row rhythm-horizontal-16'>
                                        <TextField
                                            required={true}
                                            value={
                                                addProjectRepositoriesLabelObservable
                                            }
                                            onChange={(
                                                event: React.ChangeEvent<
                                                    | HTMLInputElement
                                                    | HTMLTextAreaElement
                                                >,
                                                newValue: string
                                            ) => {
                                                addProjectRepositoriesLabelObservable.value =
                                                    newValue;
                                                this.setState({
                                                    addProjectRepositoriesLabel:
                                                        addProjectRepositoriesLabelObservable.value,
                                                });
                                            }}
                                            style={TextFieldStyle.normal}
                                        />
                                        <Button
                                            text='Add Project Label'
                                            iconProps={{ iconName: 'Add' }}
                                            primary={true}
                                            disabled={
                                                addProjectRepositoriesLabelObservable.value.trim() ===
                                                ''
                                            }
                                            onClick={
                                                this
                                                    .addProjectRepositoriesLabelAction
                                            }
                                        />
                                    </div>
                                </FormItem>
                                {this.state.systemSettings &&
                                    this.state.systemSettings
                                        .projectRepositories.length > 0 && (
                                        <Table
                                            ariaLabel='Project Repositories'
                                            columns={
                                                this
                                                    .projectRepositoriesTableColumns
                                            }
                                            itemProvider={
                                                new ArrayItemProvider<Common.IProjectRepositories>(
                                                    this.state.systemSettings.projectRepositories
                                                )
                                            }
                                            role='table'
                                            containerClassName='h-scroll-auto'
                                        />
                                    )}
                            </>
                        );
                    }}
                </Observer>
            </div>
        );
    }

    private addProjectRepositoriesLabelAction(): void {
        let systemSettings: Common.ISystemSettings | undefined =
            this.state.systemSettings;
        if (!systemSettings) {
            systemSettings = {
                allowedUserGroups: [],
                allowedUsers: [],
                projectRepositories: [],
            };
        }
        systemSettings.projectRepositories.push({
            id: new Date().getTime().toString(),
            label: this.state.addProjectRepositoriesLabel ?? '',
            selections: new DropdownMultiSelection(),
            repositories: [],
        });

        this.setState({
            systemSettings,
        });

        addProjectRepositoriesLabelObservable.value = '';
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
                                titleIconProps={{ iconName: 'Repo' }}
                            />
                            <div className='page-content page-content-top'>
                                Select the repositories you want to view and
                                process.
                            </div>
                            {this.renderMyRepositoriesDropdown()}
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
                                titleIconProps={{ iconName: 'Permissions' }}
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
                <div className='page-content'>
                    <Card className='bolt-card-white'>
                        <Page className='sprintly-width-100'>
                            <Header
                                title='Project Repositories'
                                titleSize={TitleSize.Medium}
                                titleIconProps={{ iconName: 'Repo' }}
                            />
                            <div className='page-content page-content-top'>
                                Select predefined lists of repositories for
                                projects/teams to help quickly identify which
                                repositories are needed for a release.
                            </div>
                            {this.renderAddProjectRepositoriesLabel()}
                        </Page>
                    </Card>
                </div>
            </Page>
        );
    }

    private renderProjectRepositoriesTableDeleteCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.IProjectRepositories>,
        tableItem: Common.IProjectRepositories
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    <>
                        <Button
                            iconProps={{
                                iconName: 'Delete',
                            }}
                            tooltipProps={{ text: 'Delete Project Label' }}
                            subtle={true}
                            onClick={() => {
                                this.setState({
                                    projectLabelIdToDelete: tableItem.id,
                                });
                                isDeleteProjectLabelDialogOpenObservable.value =
                                    true;
                            }}
                        />
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderDeleteProjectRepositoriesAction(): JSX.Element {
        return (
            <Observer
                isDeleteProjectLabelDialogOpen={
                    isDeleteProjectLabelDialogOpenObservable
                }
            >
                {(observerProps: {
                    isDeleteProjectLabelDialogOpen: boolean;
                }) => {
                    return observerProps.isDeleteProjectLabelDialogOpen ? (
                        <Dialog
                            titleProps={{ text: 'Delete Project Label' }}
                            footerButtonProps={[
                                {
                                    text: 'Cancel',
                                    onClick:
                                        this
                                            .onDismissDeleteProjectLabelActionModal,
                                },
                                {
                                    text: 'Delete',
                                    onClick: this.deleteProjectLabelAction,
                                    danger: true,
                                },
                            ]}
                            onDismiss={
                                this.onDismissDeleteProjectLabelActionModal
                            }
                        >
                            This is a safe operation. Only the label and its
                            predefined list of repositories for viewing will be
                            removed. Deletion will not persist until Save System
                            Settings is clicked.
                        </Dialog>
                    ) : null;
                }}
            </Observer>
        );
    }

    private onDismissDeleteProjectLabelActionModal(): void {
        isDeleteProjectLabelDialogOpenObservable.value = false;
    }

    private deleteProjectLabelAction(): void {
        if (this.state.systemSettings?.projectRepositories) {
            const systemSettings: Common.ISystemSettings =
                this.state.systemSettings;
            const projectRepositories: Common.IProjectRepositories[] =
                systemSettings.projectRepositories;
            const projectLabelIdx: number = projectRepositories.findIndex(
                (item: Common.IProjectRepositories) =>
                    item.id === this.state.projectLabelIdToDelete!
            );
            if (projectLabelIdx > -1) {
                projectRepositories.splice(projectLabelIdx, 1);
            }
            systemSettings.projectRepositories = projectRepositories;
            this.setState({
                systemSettings,
            });
        }
        this.onDismissDeleteProjectLabelActionModal();
    }

    private renderProjectRepositoriesTableLabelCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.IProjectRepositories>,
        tableItem: Common.IProjectRepositories
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={<>{tableItem.label}</>}
            ></SimpleTableCell>
        );
    }

    private renderProjectRepositoriesTableRepositoriesCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<Common.IProjectRepositories>,
        tableItem: Common.IProjectRepositories
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={
                    <>
                        <Observer selection={tableItem.selections}>
                            {() => {
                                return (
                                    <Dropdown
                                        ariaLabel='Multiselect'
                                        actions={[
                                            {
                                                className:
                                                    'bolt-dropdown-action-right-button',
                                                iconProps: {
                                                    iconName: 'Accept',
                                                },
                                                text: 'Select All',
                                                onClick: () => {
                                                    tableItem.selections.select(
                                                        0,
                                                        this.allRepositories
                                                            .length
                                                    );
                                                },
                                            },
                                            {
                                                className:
                                                    'bolt-dropdown-action-right-button',
                                                disabled:
                                                    tableItem.selections
                                                        .selectedCount === 0,
                                                iconProps: {
                                                    iconName: 'Clear',
                                                },
                                                text: 'Clear',
                                                onClick: () => {
                                                    tableItem.selections.clear();
                                                },
                                            },
                                        ]}
                                        className='sprintly-dropdown-width-100'
                                        items={this.allRepositories.map(
                                            (item: Common.IAllowedEntity) =>
                                                item.displayName
                                        )}
                                        selection={tableItem.selections}
                                        placeholder='Select Individual Repositories'
                                        showFilterBox={true}
                                        onSelect={() => {
                                            const systemSettings: Common.ISystemSettings =
                                                this.state.systemSettings!;
                                            const projectRepoIdx: number =
                                                systemSettings.projectRepositories.findIndex(
                                                    (
                                                        item: Common.IProjectRepositories
                                                    ) =>
                                                        item.id === tableItem.id
                                                );
                                            if (projectRepoIdx > -1) {
                                                const projectRepo: Common.IProjectRepositories =
                                                    systemSettings
                                                        .projectRepositories[
                                                        projectRepoIdx
                                                    ];
                                                projectRepo.selections =
                                                    tableItem.selections;
                                                systemSettings.projectRepositories.splice(
                                                    projectRepoIdx,
                                                    1,
                                                    projectRepo
                                                );
                                            }
                                            this.setState({
                                                systemSettings,
                                            });
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
                    {this.renderDeleteProjectRepositoriesAction()}
                </div>
            </Page>
        );
    }
}
