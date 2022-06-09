import * as React from 'react';
import axios from 'axios';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    IExtensionDataManager,
    IGlobalMessagesService,
    IProjectInfo,
} from 'azure-devops-extension-api';
import { ReleaseDefinition } from 'azure-devops-extension-api/Release';
import { BuildDefinition } from 'azure-devops-extension-api/Build';

import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { Tab, TabBar, TabSize } from 'azure-devops-ui/Tabs';
import { Page } from 'azure-devops-ui/Page';
import { Header, TitleSize } from 'azure-devops-ui/Header';
import { IHeaderCommandBarItem } from 'azure-devops-ui/HeaderCommandBar';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import { IMenuItem } from 'azure-devops-ui/Menu';
import { Link } from 'azure-devops-ui/Link';
import { Icon, IconSize } from 'azure-devops-ui/Icon';

import SprintlyPage from './SprintlyPage';
import SprintlyInRelease from './SprintlyInRelease';
import SprintlyPostRelease from './SprintlyPostRelease';
import SprintlySettings from './SprintlySettings';
import SprintlyBranchCreators from './SprintlyBranchCreators';
import SprintlyBranchNameSearch from './SprintlyBranchNameSearch';
import SprintlyEnvironmentVariableViewer from './SprintlyEnvironmentVariableViewer';
import * as Common from './SprintlyCommon';

import { showRootComponent } from '../Common';

const selectedTabKey: string = 'selected-tab';
const sprintlyPageTabKey: string = 'sprintly-page';
const sprintlyPageTabName: string = 'Sprintly';
const sprintlyInReleaseTabKey: string = 'sprintly-in-release';
const sprintlyInReleaseTabName: string = 'In-Release (QA)';
const sprintlyPostReleaseTabKey: string = 'sprintly-post-release';
const sprintlyPostReleaseTabName: string = 'Post Release';
const sprintlySettingsTabKey: string = 'sprintly-settings';
const sprintlySettingsTabName: string = 'Settings';
const sprintlyBranchCreatorsTabKey: string = 'sprintly-branch-creators';
const sprintlyBranchCreatorsTabName: string = 'Branch Creators';
const sprintlyBranchNameSearchTabKey: string = 'sprintly-branch-name-search';
const sprintlyBranchNameSearchTabName: string = 'Branch Name Search';
const sprintlyEnvironmentVariableViewerTabKey: string =
    'sprintly-environment-variable-viewer';
const sprintlyEnvironmentVariableViewerTabName: string =
    'Environment Variable Viewer';

const selectedTabIdObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const userIsAllowedObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const loggedInUserDescriptorObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const loggedInUserNameObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const organizationNameObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const isReadyObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);

export interface IFoundationSprintlyState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    allAllowedUsersDescriptors: string[];
}

// TODO: Clean up arrow functions for the cases in which I thought I
// couldn't use regular functions because the this.* was undefined errors.
// The solution is to bind those functions to `this` in the constructor.
// See SprintlyPostRelease as an example.
export default class FoundationSprintly extends React.Component<
    {},
    IFoundationSprintlyState
> {
    private dataManager!: IExtensionDataManager;
    private globalMessagesSvc!: IGlobalMessagesService;
    private accessToken: string = '';
    private releaseDefinitions: ReleaseDefinition[] = [];
    private buildDefinitions: BuildDefinition[] = [];

    constructor(props: {}) {
        super(props);
        this.state = {
            allAllowedUsersDescriptors: [],
        };

        this.selectViewRepositoriesCommandBarItem =
            this.selectViewRepositoriesCommandBarItem.bind(this);
        this.selectRepositoriesAction =
            this.selectRepositoriesAction.bind(this);
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
        const user: SDK.IUserContext = SDK.getUser();
        loggedInUserDescriptorObservable.value = user.descriptor;
        loggedInUserNameObservable.value = user.name;
        organizationNameObservable.value = SDK.getHost().name;

        this.accessToken = await SDK.getAccessToken();
        this.globalMessagesSvc = await SDK.getService<IGlobalMessagesService>(
            CommonServiceIds.GlobalMessagesService
        );
        this.dataManager = await Common.initializeDataManager(this.accessToken);

        selectedTabIdObservable.value = getUserSelectedTab();

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

        const currentProject: IProjectInfo | undefined =
            await Common.getCurrentProject();
        this.releaseDefinitions = await Common.getReleaseDefinitions(
            currentProject,
            organizationNameObservable.value,
            this.accessToken
        );
        this.buildDefinitions = await Common.getBuildDefinitions(
            currentProject,
            organizationNameObservable.value,
            this.accessToken
        );

        this.setState({
            userSettings,
            systemSettings,
        });

        await this.loadAllowedUserGroupsUsers();
        this.loadAllowedUsers();
        isReadyObservable.value = true;
    }

    private async loadAllowedUserGroupsUsers(): Promise<void> {
        let userGroups: Common.IAllowedEntity[] | undefined =
            this.state.systemSettings?.allowedUserGroups;
        if (!userGroups) {
            userGroups = Common.ALWAYS_ALLOWED_GROUPS;
        } else {
            userGroups = userGroups.concat(Common.ALWAYS_ALLOWED_GROUPS);
        }
        if (userGroups) {
            for (const group of userGroups) {
                this.accessToken = await Common.getOrRefreshToken(
                    this.accessToken
                );
                axios
                    .get(
                        `https://vsaex.dev.azure.com/${organizationNameObservable.value}/_apis/GroupEntitlements/${group.originId}/members`,
                        {
                            headers: {
                                Authorization: `Bearer ${this.accessToken}`,
                            },
                        }
                    )
                    .then((res: any) => {
                        const allAllowedUsersDescriptors: string[] = res.data[
                            'members'
                        ].map((item: any) => item['user']['descriptor']);
                        this.setState({
                            allAllowedUsersDescriptors:
                                allAllowedUsersDescriptors.concat(
                                    this.state.allAllowedUsersDescriptors
                                ),
                        });
                        userIsAllowedObservable.value =
                            this.state.allAllowedUsersDescriptors.includes(
                                loggedInUserDescriptorObservable.value
                            );
                    })
                    .catch((error: any) => {
                        console.error(error);
                    });
            }
        }
    }

    private loadAllowedUsers(): void {
        const users: Common.IAllowedEntity[] | undefined =
            this.state.systemSettings?.allowedUsers;
        if (users) {
            const allAllowedUsersDescriptors: string[] = users.map(
                (user: Common.IAllowedEntity) => user.descriptor || ''
            );
            this.setState({
                allAllowedUsersDescriptors: allAllowedUsersDescriptors.concat(
                    this.state.allAllowedUsersDescriptors
                ),
            });
            userIsAllowedObservable.value =
                this.state.allAllowedUsersDescriptors.includes(
                    loggedInUserDescriptorObservable.value
                );
        }
    }

    private getCommandBarItems(): IHeaderCommandBarItem[] {
        const items: IHeaderCommandBarItem[] = [];
        if (
            this.state.systemSettings?.projectRepositories &&
            this.state.systemSettings.projectRepositories.length > 0
        ) {
            items.push(this.selectViewRepositoriesCommandBarItem());
        }
        items.push(this.refreshButtonCommandBarItem());
        return items;
    }

    private refreshButtonCommandBarItem(): IHeaderCommandBarItem {
        return {
            id: 'refresh',
            text: 'Refresh Data',
            onActivate: () => {
                window.location.reload();
            },
            iconProps: {
                iconName: 'Refresh',
            },
            tooltipProps: {
                text: 'Refresh the data on the page',
            },
        };
    }

    // TODO: Extract this into two buttons and show one or the other
    private selectViewRepositoriesCommandBarItem(): IHeaderCommandBarItem {
        const subMenuItems: IMenuItem[] = [];
        subMenuItems.push({
            id: '0',
            text: ' My Repositories',
            onActivate: () => {
                this.selectRepositoriesAction('');
            },
        });
        if (this.state.systemSettings?.projectRepositories) {
            for (const projectRepository of this.state.systemSettings
                .projectRepositories) {
                subMenuItems.push({
                    id: projectRepository.id,
                    text: projectRepository.label,
                    onActivate: (item: IMenuItem) => {
                        this.selectRepositoriesAction(item.id);
                    },
                });
            }
        }
        return {
            id: 'selectViewProjectRepositories',
            text: 'View Project Repositories',
            iconProps: {
                iconName: 'Repo',
            },
            subMenuProps: {
                id: 'submenu',
                items: subMenuItems,
            },
        };
    }

    private selectRepositoriesAction(projectRepositoriesId: string): void {
        Common.getUserSettings(
            this.dataManager,
            Common.USER_SETTINGS_DATA_MANAGER_KEY
        ).then((userSettings: Common.IUserSettings | undefined) => {
            if (!userSettings) {
                userSettings = {
                    myRepositories: [],
                    projectRepositoriesId,
                };
            } else {
                userSettings.projectRepositoriesId = projectRepositoriesId;
            }

            this.dataManager!.setValue<Common.IUserSettings>(
                Common.USER_SETTINGS_DATA_MANAGER_KEY,
                userSettings,
                { scopeType: 'User' }
            ).then(() => {
                window.location.reload();
            });
        });
    }

    private renderSelectedTabPage(): JSX.Element {
        switch (selectedTabIdObservable.value) {
            case sprintlyPageTabKey:
            case '':
                return (
                    <SprintlyPage
                        accessToken={this.accessToken}
                        globalMessagesSvc={this.globalMessagesSvc}
                    />
                );
            case sprintlySettingsTabKey:
                return (
                    <SprintlySettings
                        organizationName={organizationNameObservable.value}
                        globalMessagesSvc={this.globalMessagesSvc}
                    />
                );
            case sprintlyInReleaseTabKey:
                return (
                    <SprintlyInRelease
                        organizationName={organizationNameObservable.value}
                        globalMessagesSvc={this.globalMessagesSvc}
                        releaseDefinitions={this.releaseDefinitions}
                        buildDefinitions={this.buildDefinitions}
                    />
                );
            case sprintlyPostReleaseTabKey:
                return (
                    <SprintlyPostRelease
                        organizationName={organizationNameObservable.value}
                        globalMessagesSvc={this.globalMessagesSvc}
                        releaseDefinitions={this.releaseDefinitions}
                        buildDefinitions={this.buildDefinitions}
                    />
                );
            case sprintlyBranchCreatorsTabKey:
                return (
                    <SprintlyBranchCreators
                        accessToken={this.accessToken}
                        globalMessagesSvc={this.globalMessagesSvc}
                        organizationName={organizationNameObservable.value}
                    />
                );
            case sprintlyBranchNameSearchTabKey:
                return (
                    <SprintlyBranchNameSearch
                        accessToken={this.accessToken}
                        globalMessagesSvc={this.globalMessagesSvc}
                        organizationName={organizationNameObservable.value}
                        userName={loggedInUserNameObservable.value}
                    />
                );
            case sprintlyEnvironmentVariableViewerTabKey:
                return (
                    <SprintlyEnvironmentVariableViewer
                        organizationName={organizationNameObservable.value}
                        userDescriptor={loggedInUserDescriptorObservable.value}
                        releaseDefinitions={this.releaseDefinitions}
                        buildDefinitions={this.buildDefinitions}
                    />
                );
            default:
                return <div></div>;
        }
    }

    public render(): JSX.Element {
        let title: string = 'Foundation Sprintly';
        if (this.state.userSettings) {
            if (this.state.userSettings.projectRepositoriesId.trim() === '') {
                title += ' (My Repositories)';
            } else {
                if (!this.state.systemSettings?.projectRepositories) {
                    title += ' (My Repositories)';
                } else {
                    const projectRepo: Common.IProjectRepositories | undefined =
                        this.state.systemSettings?.projectRepositories.find(
                            (item: Common.IProjectRepositories) =>
                                item.id ===
                                this.state.userSettings!.projectRepositoriesId
                        );
                    if (!projectRepo) {
                        title += ' (My Repositories)';
                    } else {
                        title += ` (${projectRepo.label})`;
                    }
                }
            }
        }
        return (
            <Page className='flex-grow foundation-sprintly'>
                <Header
                    title={title}
                    commandBarItems={this.getCommandBarItems()}
                    titleSize={TitleSize.Large}
                    description={
                        <>
                            <Link
                                href='#'
                                tooltipProps={{
                                    text: 'Repos without a develop or master/main branch will not be shown.',
                                }}
                            >
                                Not seeing your repositories?{' '}
                                <Icon
                                    iconName='Info'
                                    size={IconSize.medium}
                                    className='sprintly-vertical-align-bottom'
                                />
                            </Link>
                        </>
                    }
                />
                <Observer userIsAllowedObservable={userIsAllowedObservable}>
                    {(props: {
                        userIsAllowedObservable: boolean;
                        refreshDataObservable: boolean;
                    }) => {
                        if (userIsAllowedObservable.value) {
                            return renderTabBar();
                        }
                        return <div></div>;
                    }}
                </Observer>

                <Observer
                    selectedTabIdObservable={selectedTabIdObservable}
                    userIsAllowedObservable={userIsAllowedObservable}
                    isReadyObservable={isReadyObservable}
                >
                    {(props: {
                        selectedTabIdObservable: string;
                        userIsAllowedObservable: boolean;
                        isReadyObservable: boolean;
                        refreshDataObservable: boolean;
                    }) => {
                        if (userIsAllowedObservable.value) {
                            return this.renderSelectedTabPage();
                        }
                        if (isReadyObservable.value) {
                            return (
                                <div>
                                    <ZeroData
                                        primaryText='Sorry, you do not have access yet.'
                                        secondaryText={
                                            <span>
                                                Please contact the DevOps team
                                                or your team lead for access to
                                                this extension.
                                            </span>
                                        }
                                        imageAltText='No Access'
                                        imagePath={'../static/notfound.png'}
                                    />
                                </div>
                            );
                        }
                        return (
                            <div>
                                <ZeroData
                                    primaryText='Getting things ready...'
                                    secondaryText={
                                        <span>
                                            Please wait, Sprintly is starting
                                            up...
                                        </span>
                                    }
                                    imageAltText='Starting up...'
                                    imagePath={'../static/notfound.png'}
                                />
                            </div>
                        );
                    }}
                </Observer>
            </Page>
        );
    }
}

function renderTabBar(): JSX.Element {
    return (
        <TabBar
            onSelectedTabChanged={onSelectedTabChanged}
            selectedTabId={selectedTabIdObservable}
            tabSize={TabSize.Tall}
        >
            <Tab name={sprintlyPageTabName} id={sprintlyPageTabKey} />
            <Tab name={sprintlyInReleaseTabName} id={sprintlyInReleaseTabKey} />
            <Tab
                name={sprintlyPostReleaseTabName}
                id={sprintlyPostReleaseTabKey}
            />
            <Tab
                name={sprintlyBranchCreatorsTabName}
                id={sprintlyBranchCreatorsTabKey}
            />
            <Tab
                name={sprintlyBranchNameSearchTabName}
                id={sprintlyBranchNameSearchTabKey}
            />
            <Tab
                name={sprintlyEnvironmentVariableViewerTabName}
                id={sprintlyEnvironmentVariableViewerTabKey}
            />
            <Tab name={sprintlySettingsTabName} id={sprintlySettingsTabKey} />
        </TabBar>
    );
}

function onSelectedTabChanged(newTabId: string): void {
    selectedTabIdObservable.value = newTabId;
    localStorage.setItem(
        loggedInUserDescriptorObservable.value + '-' + selectedTabKey,
        newTabId
    );
}

function getUserSelectedTab(): string {
    return (
        localStorage.getItem(
            loggedInUserDescriptorObservable.value + '-' + selectedTabKey
        ) ?? 'sprintly-page'
    );
}

showRootComponent(<FoundationSprintly />);
