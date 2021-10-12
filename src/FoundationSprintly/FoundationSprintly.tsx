import * as React from 'react';

import * as SDK from 'azure-devops-extension-sdk';

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
import {
    CommonServiceIds,
    IExtensionDataManager,
    IExtensionDataService,
} from 'azure-devops-extension-api';
import axios from 'axios';

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
    descriptor: string;
}

export interface IFoundationSprintlyState {
    // TODO: These "persisted" properties may not be needed
    persistedAllowedUserGroups?: AllowedEntity[];
    persistedAllowedUsers?: AllowedEntity[];
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

    public async componentDidMount() {
        await SDK.init();
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        await SDK.ready();
        loggedInUserDescriptorObservable.value = SDK.getUser().descriptor;
        organizationNameObservable.value = SDK.getHost().name;

        this.accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>(
            CommonServiceIds.ExtensionDataService
        );
        this._dataManager = await extDataService.getExtensionDataManager(
            SDK.getExtensionContext().id,
            this.accessToken
        );

        selectedTabId.value =
            localStorage.getItem(
                loggedInUserDescriptorObservable.value.replace('.', '-') +
                    '-selected-tab'
            ) ?? 'sprintly-page';
        console.log('saved tab ', selectedTabId.value);

        this._dataManager.getValue<AllowedEntity[]>('allowed-user-groups').then(
            (userGroups) => {
                console.log('data is this ', userGroups);
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
                            console.log(
                                this.state.allAllowedUsersDescriptors,
                                userIsAllowed.value
                            );
                        })
                        .catch((error) => {
                            console.error(error);
                        });
                }
                this.setState({
                    persistedAllowedUserGroups: userGroups,
                });
            },
            () => {
                this.setState({
                    persistedAllowedUserGroups: [],
                });
            }
        );

        this._dataManager.getValue<AllowedEntity[]>('allowed-users').then(
            (users) => {
                const allAllowedUsersDescriptors = users.map(
                    (user) => user.descriptor
                );
                this.setState({
                    persistedAllowedUsers: users,
                    allAllowedUsersDescriptors:
                        allAllowedUsersDescriptors.concat(
                            this.state.allAllowedUsersDescriptors
                        ),
                });
                userIsAllowed.value =
                    this.state.allAllowedUsersDescriptors.includes(
                        loggedInUserDescriptorObservable.value
                    );
                console.log(
                    this.state.allAllowedUsersDescriptors,
                    userIsAllowed.value
                );
            },
            () => {
                this.setState({
                    persistedAllowedUsers: [],
                });
            }
        );
    }

    public render() {
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
                                    <Tab name="Sprintly" id="sprintly-page" />
                                    <Tab
                                        name="Post Release"
                                        id="sprintly-post-release"
                                    />
                                    <Tab
                                        name="Settings"
                                        id="sprintly-settings"
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
                            console.log('user is allowed', selectedTabId.value);

                            switch (selectedTabId.value) {
                                case 'sprintly-page':
                                case '':
                                    return (
                                        <SprintlyPage
                                            loggedInUserDescriptor={
                                                loggedInUserDescriptorObservable.value
                                            }
                                        />
                                    );
                                case 'sprintly-settings':
                                    return (
                                        <SprintlySettings
                                            sampleProp={selectedTabId.value}
                                            loggedInUserDescriptor={
                                                loggedInUserDescriptorObservable.value
                                            }
                                            organizationName={
                                                organizationNameObservable.value
                                            }
                                        />
                                    );
                                case 'sprintly-post-release':
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
        loggedInUserDescriptorObservable.value.replace('.', '-') +
            '-selected-tab',
        newTabId
    );
}

showRootComponent(<FoundationSprintly />);
