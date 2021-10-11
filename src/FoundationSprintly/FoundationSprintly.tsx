import * as React from 'react';

import * as SDK from 'azure-devops-extension-sdk';

import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { Tab, TabBar, TabSize } from 'azure-devops-ui/Tabs';
import { Page } from 'azure-devops-ui/Page';
import { Header, TitleSize } from 'azure-devops-ui/Header';
import { ZeroData } from 'azure-devops-ui/ZeroData';

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

export interface AllowedEntity {
    displayName: string;
    originId: string;
    descriptor: string;
}

export interface IFoundationSprintlyState {
    // TODO: These "persisted" properties may not be needed
    persistedAllowedUserGroups?: AllowedEntity[];
    persistedAllowedUsers?: AllowedEntity[];
    // TODO: Passed logged in user as a prop to sprintly settings
    loggedInUserDescriptor: string;
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
            loggedInUserDescriptor: '',
            allAllowedUsersDescriptors: [],
        };
    }

    public async componentDidMount() {
        await SDK.init();
        this.initializeState();
    }

    private async initializeState(): Promise<void> {
        await SDK.ready();
        const loggedInUserDescriptor: string = SDK.getUser().descriptor;
        loggedInUserDescriptorObservable.value = loggedInUserDescriptor;
        this.setState({ loggedInUserDescriptor: loggedInUserDescriptor });

        this.accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>(
            CommonServiceIds.ExtensionDataService
        );
        this._dataManager = await extDataService.getExtensionDataManager(
            SDK.getExtensionContext().id,
            this.accessToken
        );

        this._dataManager.getValue<AllowedEntity[]>('allowed-user-groups').then(
            (userGroups) => {
                console.log('data is this ', userGroups);
                for (const group of userGroups) {
                    axios
                        .get(
                            `https://vsaex.dev.azure.com/reponzo01/_apis/GroupEntitlements/${group.originId}/members`,
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
                                    this.state.loggedInUserDescriptor
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
                        this.state.loggedInUserDescriptor
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

                <Observer userIsAllowed={userIsAllowed}>
                    {(props: { userIsAllowed: boolean }) => {
                        if (userIsAllowed.value) {
                            selectedTabId.value = 'sprintly-page';
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

                <Observer selectedTabId={selectedTabId}>
                    {(props: { selectedTabId: string }) => {
                        if (selectedTabId.value === 'sprintly-page') {
                            return (
                                <SprintlyPage
                                    loggedInUserDescriptor={
                                        loggedInUserDescriptorObservable.value
                                    }
                                />
                            );
                        } else if (
                            selectedTabId.value === 'sprintly-settings'
                        ) {
                            return (
                                <SprintlySettings
                                    sampleProp={selectedTabId.value}
                                    loggedInUserDescriptor={
                                        loggedInUserDescriptorObservable.value
                                    }
                                />
                            );
                        } else if (
                            selectedTabId.value === 'sprintly-post-release'
                        ) {
                            return <SprintlyPostRelease />;
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
    selectedTabId.value = newTabId;
}

showRootComponent(<FoundationSprintly />);
