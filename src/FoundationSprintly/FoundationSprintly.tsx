import * as React from 'react';
import axios from 'axios';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    IExtensionDataManager,
    IExtensionDataService,
} from 'azure-devops-extension-api';

import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { Tab, TabBar, TabSize } from 'azure-devops-ui/Tabs';
import { Page } from 'azure-devops-ui/Page';
import { Header, TitleSize } from 'azure-devops-ui/Header';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import { Button } from 'azure-devops-ui/Button';

import { SprintlyPage } from './SprintlyPage';
import SprintlyPostRelease from './SprintlyPostRelease';
import SprintlySettings from './SprintlySettings';
import { showRootComponent } from '../Common';

const selectedTabKey: string = 'selected-tab';
const allowedUserGroupsKey: string = 'allowed-user-groups';
const allowedUsersKey: string = 'allowed-users';
const sprintlyPageTab: string = 'sprintly-page';
const sprintlyPostReleaseTab: string = 'sprintly-post-release';
const sprintlySettingsTab: string = 'sprintly-settings';

const selectedTabId: ObservableValue<string> = new ObservableValue<string>('');
const userIsAllowed: ObservableValue<boolean> = new ObservableValue<boolean>(
    false
);
const loggedInUserDescriptorObservable: ObservableValue<string> =
    new ObservableValue<string>('');
const organizationNameObservable: ObservableValue<string> =
    new ObservableValue<string>('');

export interface AllowedEntity {
    displayName: string;
    originId: string;
    descriptor?: string;
}

export interface IFoundationSprintlyState {
    allAllowedUsersDescriptors: string[];
}

export default class FoundationSprintly extends React.Component<
    {},
    IFoundationSprintlyState
> {
    private _dataManager?: IExtensionDataManager;
    private accessToken: string = '';

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
        this._dataManager = await this.initializeDataManager();

        selectedTabId.value = getUserSelectedTab();

        this.loadAllowedUserGroupsUsers();
        this.loadAllowedUsers();
    }

    private async initializeDataManager(): Promise<IExtensionDataManager> {
        const extDataService: IExtensionDataService =
            await SDK.getService<IExtensionDataService>(
                CommonServiceIds.ExtensionDataService
            );
        return await extDataService.getExtensionDataManager(
            SDK.getExtensionContext().id,
            this.accessToken
        );
    }

    private loadAllowedUserGroupsUsers(): void {
        this._dataManager!.getValue<AllowedEntity[]>(allowedUserGroupsKey).then(
            (userGroups: AllowedEntity[]) => {
                for (const group of userGroups) {
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
                                    (item: any) => item['user']['descriptor']
                                );
                            this.setState({
                                allAllowedUsersDescriptors:
                                    allAllowedUsersDescriptors.concat(
                                        this.state.allAllowedUsersDescriptors
                                    ),
                            });
                            userIsAllowed.value =
                                this.state.allAllowedUsersDescriptors.includes(
                                    loggedInUserDescriptorObservable.value
                                );
                        })
                        .catch((error: any) => {
                            console.error(error);
                        });
                }
            }
        );
    }

    private loadAllowedUsers(): void {
        this._dataManager!.getValue<AllowedEntity[]>(allowedUsersKey).then(
            (users: AllowedEntity[]) => {
                const allAllowedUsersDescriptors: string[] = users.map(
                    (user: AllowedEntity) => user.descriptor || ''
                );
                this.setState({
                    allAllowedUsersDescriptors:
                        allAllowedUsersDescriptors.concat(
                            this.state.allAllowedUsersDescriptors
                        ),
                });
                userIsAllowed.value =
                    this.state.allAllowedUsersDescriptors.includes(
                        loggedInUserDescriptorObservable.value
                    );
            }
        );
    }

    public render(): JSX.Element {
        return (
            /* tslint:disable */
            <Page className="flex-grow foundation-sprintly">
                <Header
                    title="Foundation Sprintly"
                    titleSize={TitleSize.Large}
                />
                <div className="page-content page-content-top flex-column rhythm-vertical-16">
                    <Button
                        text="Refresh Data"
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={() => window.location.reload()}
                    />
                </div>
                <Observer userIsAllowed={userIsAllowed}>
                    {(props: { userIsAllowed: boolean }) => {
                        if (userIsAllowed.value) {
                            return (
                                <TabBar
                                    onSelectedTabChanged={onSelectedTabChanged}
                                    selectedTabId={selectedTabId}
                                    tabSize={TabSize.Tall}
                                >
                                    <Tab name="Sprintly" id={sprintlyPageTab} />
                                    <Tab
                                        name="Post Release"
                                        id={sprintlyPostReleaseTab}
                                    />
                                    <Tab
                                        name="Settings"
                                        id={sprintlySettingsTab}
                                    />
                                </TabBar>
                            );
                        }
                        return <div></div>;
                    }}
                </Observer>

                <Observer
                    selectedTabId={selectedTabId}
                    userIsAllowed={userIsAllowed}
                >
                    {(props: {
                        selectedTabId: string;
                        userIsAllowed: boolean;
                    }) => {
                        if (userIsAllowed.value) {
                            switch (selectedTabId.value) {
                                case sprintlyPageTab:
                                case '':
                                    return <SprintlyPage />;
                                case sprintlySettingsTab:
                                    return (
                                        <SprintlySettings
                                            organizationName={
                                                organizationNameObservable.value
                                            }
                                        />
                                    );
                                case sprintlyPostReleaseTab:
                                    return <SprintlyPostRelease />;
                                default:
                                    return <div></div>;
                            }
                        }

                        return (
                            <div>
                                <ZeroData
                                    primaryText="Sorry, you don't have access yet."
                                    secondaryText={
                                        <span>
                                            Please contact the DevOps team or{' '}
                                            your team lead for access to this{' '}
                                            extension.
                                        </span>
                                    }
                                    imageAltText="No Access"
                                    imagePath={'../static/notfound.png'}
                                />
                            </div>
                        );
                    }}
                </Observer>
            </Page>
            /* tslint:disable */
        );
    }
}

function onSelectedTabChanged(newTabId: string) {
    console.log('setting tab to ', loggedInUserDescriptorObservable.value);
    selectedTabId.value = newTabId;
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
