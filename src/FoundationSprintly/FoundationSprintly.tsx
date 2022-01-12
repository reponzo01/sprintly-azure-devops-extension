import * as React from 'react';
import axios from 'axios';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    IExtensionDataManager,
    IGlobalMessagesService,
} from 'azure-devops-extension-api';
import { ReleaseDefinition } from 'azure-devops-extension-api/Release';
import { BuildDefinition } from 'azure-devops-extension-api/Build';
import { TeamProjectReference } from 'azure-devops-extension-api/Core';

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
import SprintlyBranchSearch from './SprintlyBranchSearch';
import * as Common from './SprintlyCommon';

import { showRootComponent } from '../Common';

const selectedTabKey: string = 'selected-tab';
const userSettingsDataManagerKey: string = 'user-settings';
const systemSettingsDataManagerKey: string = 'system-settings';
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
const sprintlyBranchSearchTabKey: string = 'sprintly-branch-search';
const sprintlyBranchSearchTabName: string = 'Branch Search';

const selectedTabIdObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const userIsAllowedObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const loggedInUserDescriptorObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const organizationNameObservable: ObservableValue<string> =
    new ObservableValue<string>('');

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

    private alwaysAllowedGroups: Common.IAllowedEntity[] = [
        {
            displayName: 'Dev Team Leads',
            originId: '841aee2f-860d-45a1-91a5-779aa4dca78c',
            descriptor:
                'vssgp.Uy0xLTktMTU1MTM3NDI0NS00MjgyNjUyNjEyLTI3NDUxOTk2OTMtMjk1ODAyODI0OS0yMTc4MDQ3MTU1LTEtNjQxMDY2NzIxLTg5MzE2MjA2MS0yNzg1NjUwNzE5LTE3MTcxNTU1MDk',
        },
        {
            displayName: 'DevOps',
            originId: 'b2620fb7-f672-4162-a15f-940b1ec78efe',
            descriptor:
                'vssgp.Uy0xLTktMTU1MTM3NDI0NS0xODk1NzMzMjY1LTQ3ODY0Mzg0LTMwMjU3MjkyMzQtOTM5ODg1NzU0LTEtMzA1NDcxNjM4Mi0zNjc1OTA4OTI5LTI3MjY5NzI4MTctMzczODgxNDI4NQ',
        },
        // {
        //     displayName: 'Sample Project Team', // fsllc
        //     originId: 'fccefee4-a7a9-432a-a7a2-fc6d3d8bc45d',
        //     descriptor:
        //         'vssgp.Uy0xLTktMTU1MTM3NDI0NS0zMTEzMzAyODctMzI5MTIzMzA5NC0zMTI4MjY0MTg3LTQwMTUzMTUzOTYtMS0xNTY5MTY5Mjc5LTIzODYzODU5OTQtMjU1MDU2OTgzMi02NDQyOTAwODc',
        // },
        // {
        //     displayName: 'Sample Project Team', // reponzo01
        //     originId: '221ca28d-8d55-4229-aeee-d96b619d8bf9',
        //     descriptor:
        //         'vssgp.Uy0xLTktMTU1MTM3NDI0NS0zNTI2OTIzMzAwLTE2ODEyODk1MzctMjE5OTc3MDkxOC0yNDEwMzk4MTQ4LTEtODgxNTgyODM0LTIyMjg0NjE4OTgtMzA0NDA1NzUwOC03NTYzNzk0ODA',
        // },
    ];

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
        loggedInUserDescriptorObservable.value = SDK.getUser().descriptor;
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
                userSettingsDataManagerKey
            );
        const systemSettings: Common.ISystemSettings | undefined =
            await Common.getSystemSettings(
                this.dataManager,
                systemSettingsDataManagerKey
            );

        const filteredProjects: TeamProjectReference[] =
            await Common.getFilteredProjects();
        this.releaseDefinitions = await Common.getReleaseDefinitions(
            filteredProjects,
            organizationNameObservable.value,
            this.accessToken
        );
        this.buildDefinitions = await Common.getBuildDefinitions(
            filteredProjects,
            organizationNameObservable.value,
            this.accessToken
        );

        this.setState({
            userSettings,
            systemSettings,
        });

        await this.loadAllowedUserGroupsUsers();
        this.loadAllowedUsers();
    }

    private async loadAllowedUserGroupsUsers(): Promise<void> {
        let userGroups: Common.IAllowedEntity[] | undefined =
            this.state.systemSettings?.allowedUserGroups;
        if (!userGroups) {
            userGroups = this.alwaysAllowedGroups;
        } else {
            userGroups = userGroups.concat(this.alwaysAllowedGroups);
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
            userSettingsDataManagerKey
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
                userSettingsDataManagerKey,
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
                        dataManager={this.dataManager}
                        globalMessagesSvc={this.globalMessagesSvc}
                    />
                );
            case sprintlySettingsTabKey:
                return (
                    <SprintlySettings
                        organizationName={organizationNameObservable.value}
                        globalMessagesSvc={this.globalMessagesSvc}
                        dataManager={this.dataManager}
                    />
                );
            case sprintlyInReleaseTabKey:
                return (
                    <SprintlyInRelease
                        organizationName={organizationNameObservable.value}
                        globalMessagesSvc={this.globalMessagesSvc}
                        dataManager={this.dataManager}
                        releaseDefinitions={this.releaseDefinitions}
                        buildDefinitions={this.buildDefinitions}
                    />
                );
            case sprintlyPostReleaseTabKey:
                return (
                    <SprintlyPostRelease
                        organizationName={organizationNameObservable.value}
                        globalMessagesSvc={this.globalMessagesSvc}
                        dataManager={this.dataManager}
                        releaseDefinitions={this.releaseDefinitions}
                        buildDefinitions={this.buildDefinitions}
                    />
                );
            case sprintlyBranchCreatorsTabKey:
                return (
                    <SprintlyBranchCreators dataManager={this.dataManager} />
                );
            case sprintlyBranchSearchTabKey:
                return (
                    <SprintlyBranchSearch />
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
                >
                    {(props: {
                        selectedTabIdObservable: string;
                        userIsAllowedObservable: boolean;
                        refreshDataObservable: boolean;
                    }) => {
                        if (userIsAllowedObservable.value) {
                            return this.renderSelectedTabPage();
                        }
                        return (
                            <div>
                                <ZeroData
                                    primaryText='Sorry, you do not have access yet.'
                                    secondaryText={
                                        <span>
                                            Please contact the DevOps team or
                                            your team lead for access to this
                                            extension.
                                        </span>
                                    }
                                    imageAltText='No Access'
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
                name={sprintlyBranchSearchTabName}
                id={sprintlyBranchSearchTabKey}
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
