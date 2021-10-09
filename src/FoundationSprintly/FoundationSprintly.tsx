import * as React from 'react';

import * as SDK from 'azure-devops-extension-sdk';

import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { Tab, TabBar, TabSize } from 'azure-devops-ui/Tabs';
import { Page } from 'azure-devops-ui/Page';
import { Header, TitleSize } from 'azure-devops-ui/Header';

import { SprintlyPage } from './SprintlyPage';
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

export interface AllowedEntity {
    displayName: string;
    originId: string;
    descriptor: string;
}

export interface IExtensionDataState {
    persistedAllowedUserGroups?: AllowedEntity[];
    persistedAllowedUsers?: AllowedEntity[];
    loggedInUserDescriptor: string;
    allAllowedUsersDescriptors: string[];
}

export default class FoundationSprintly extends React.Component<
    {},
    IExtensionDataState
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
        this.setState({ loggedInUserDescriptor: SDK.getUser().descriptor });

        this.accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>(
            CommonServiceIds.ExtensionDataService
        );
        this._dataManager = await extDataService.getExtensionDataManager(
            SDK.getExtensionContext().id,
            this.accessToken
        );

        this._dataManager.getValue<AllowedEntity[]>('allowed-user-groups').then(
            (data) => {
                console.log('data is this ', data);
                for (const group of data) {
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
                    persistedAllowedUserGroups: data,
                });
            },
            () => {
                this.setState({
                    persistedAllowedUserGroups: [],
                });
            }
        );

        this._dataManager.getValue<AllowedEntity[]>('allowed-users').then(
            (data) => {
                const allAllowedUsersDescriptors = data.map(
                    (item) => item.descriptor
                );
                this.setState({
                    persistedAllowedUsers: data,
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
                                        name="Settings"
                                        id="sprintly-settings"
                                    />
                                </TabBar>
                            );
                        }
                        return <div>No access</div>;
                    }}
                </Observer>

                <Observer selectedTabId={selectedTabId}>
                    {(props: { selectedTabId: string }) => {
                        if (selectedTabId.value === 'sprintly-page') {
                            return <SprintlyPage />;
                        } else if (
                            selectedTabId.value === 'sprintly-settings'
                        ) {
                            return (
                                <SprintlySettings
                                    sampleProp={selectedTabId.value}
                                />
                            );
                        }
                        return <div>No access</div>;
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
