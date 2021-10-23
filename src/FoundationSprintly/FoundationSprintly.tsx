import * as React from 'react';
import axios from 'axios';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    IExtensionDataManager,
    IGlobalMessagesService,
} from 'azure-devops-extension-api';

import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { Tab, TabBar, TabSize } from 'azure-devops-ui/Tabs';
import { Page } from 'azure-devops-ui/Page';
import { Header, TitleSize } from 'azure-devops-ui/Header';
import { IHeaderCommandBarItem } from 'azure-devops-ui/HeaderCommandBar';
import { ZeroData } from 'azure-devops-ui/ZeroData';

import SprintlyPage from './SprintlyPage';
import SprintlyInRelease from './SprintlyInRelease';
import SprintlyPostRelease from './SprintlyPostRelease';
import SprintlySettings from './SprintlySettings';
import * as Common from './SprintlyCommon';
import { showRootComponent } from '../Common';

const selectedTabKey: string = 'selected-tab';
const allowedUserGroupsKey: string = 'allowed-user-groups';
const allowedUsersKey: string = 'allowed-users';
const sprintlyPageTabKey: string = 'sprintly-page';
const sprintlyPageTabName: string = 'Sprintly';
const sprintlyInReleaseTabKey: string = 'sprintly-in-release';
const sprintlyInReleaseTabName: string = 'In-Release (QA)';
const sprintlyPostReleaseTabKey: string = 'sprintly-post-release';
const sprintlyPostReleaseTabName: string = 'Post Release';
const sprintlySettingsTabKey: string = 'sprintly-settings';
const sprintlySettingsTabName: string = 'Settings';

const selectedTabIdObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const userIsAllowedObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
const loggedInUserDescriptorObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const organizationNameObservable: ObservableValue<string> =
    new ObservableValue<string>('');

export interface IFoundationSprintlyState {
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

    private alwaysAllowedGroups: Common.IAllowedEntity[] = [
        /*{
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
        },*/
        {
            displayName: 'ample Project Team',
            originId: 'fccefee4-a7a9-432a-a7a2-fc6d3d8bc45d',
            descriptor:
                'vssgp.Uy0xLTktMTU1MTM3NDI0NS0zMTEzMzAyODctMzI5MTIzMzA5NC0zMTI4MjY0MTg3LTQwMTUzMTUzOTYtMS0xNTY5MTY5Mjc5LTIzODYzODU5OTQtMjU1MDU2OTgzMi02NDQyOTAwODc',
        },
    ];

    constructor(props: {}) {
        super(props);
        this.state = {
            allAllowedUsersDescriptors: [],
        };
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

        this.loadAllowedUserGroupsUsers();
        this.loadAllowedUsers();
    }

    private loadAllowedUserGroupsUsers(): void {
        this.dataManager!.getValue<Common.IAllowedEntity[]>(
            allowedUserGroupsKey
        ).then(
            async (userGroups: Common.IAllowedEntity[]) => {
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
                                const allAllowedUsersDescriptors: string[] =
                                    res.data['members'].map(
                                        (item: any) =>
                                            item['user']['descriptor']
                                    );
                                this.setState({
                                    allAllowedUsersDescriptors:
                                        allAllowedUsersDescriptors.concat(
                                            this.state
                                                .allAllowedUsersDescriptors
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
            },
            () => {
                this.setState({
                    allAllowedUsersDescriptors: [],
                });
            }
        );
    }

    private loadAllowedUsers(): void {
        this.dataManager!.getValue<Common.IAllowedEntity[]>(
            allowedUsersKey
        ).then(
            (users: Common.IAllowedEntity[]) => {
                if (users) {
                    const allAllowedUsersDescriptors: string[] = users.map(
                        (user: Common.IAllowedEntity) => user.descriptor || ''
                    );
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
                }
            },
            () => {
                this.setState({
                    allAllowedUsersDescriptors: [],
                });
            }
        );
    }

    private getCommandBarItems(): IHeaderCommandBarItem[] {
        return [this.refreshButtonCommanBarItem()];
    }

    private refreshButtonCommanBarItem(): IHeaderCommandBarItem {
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
                        dataManager={this.dataManager}
                    />
                );
            case sprintlyInReleaseTabKey:
                return (
                    <SprintlyInRelease
                        organizationName={organizationNameObservable.value}
                        dataManager={this.dataManager}
                    />
                );
            case sprintlyPostReleaseTabKey:
                return (
                    <SprintlyPostRelease
                        organizationName={organizationNameObservable.value}
                        globalMessagesSvc={this.globalMessagesSvc}
                        dataManager={this.dataManager}
                    />
                );
            default:
                return <div></div>;
        }
    }

    public render(): JSX.Element {
        return (
            <Page className='flex-grow foundation-sprintly'>
                <Header
                    title='Foundation Sprintly'
                    commandBarItems={this.getCommandBarItems()}
                    titleSize={TitleSize.Large}
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
