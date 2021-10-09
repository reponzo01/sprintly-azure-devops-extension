import * as React from 'react';

import axios from 'axios';

import * as SDK from 'azure-devops-extension-sdk';
import {
    CommonServiceIds,
    getClient,
    IExtensionDataManager,
    IExtensionDataService,
} from 'azure-devops-extension-api';
import {
    GraphRestClient,
    GraphSubject,
    GraphSubjectQuery,
} from 'azure-devops-extension-api/Graph';

import { Button } from 'azure-devops-ui/Button';
import { TextField } from 'azure-devops-ui/TextField';
import { Dropdown } from 'azure-devops-ui/Dropdown';
import { Observer } from 'azure-devops-ui/Observer';
import { DropdownMultiSelection } from 'azure-devops-ui/Utilities/DropdownSelection';
import { ISelectionRange } from 'azure-devops-ui/Utilities/Selection';
import { CoreRestClient } from 'azure-devops-extension-api/Core';
import { resolveTypeReferenceDirective } from 'typescript';

// TODO: Remove all this extension data from here and pass in as prop from Parent page
export interface AllowedEntity {
    displayName: string;
    originId: string;
    descriptor: string;
}

export interface IExtensionDataState {
    dataAllowedUserGroups?: AllowedEntity[];
    dataAllowedUsers?: AllowedEntity[];
    persistedAllowedUserGroups?: AllowedEntity[];
    persistedAllowedUsers?: AllowedEntity[];
    ready?: boolean;
}

export default class SprintlySettings extends React.Component<
    { sampleProp: string },
    IExtensionDataState
> {
    private userGroupsSelection = new DropdownMultiSelection();
    private usersSelection = new DropdownMultiSelection();
    private userGroups: AllowedEntity[] = [];
    private users: AllowedEntity[] = [];
    private _dataManager?: IExtensionDataManager;
    private sampleProp: string;

    constructor(props: { sampleProp: string }) {
        super(props);
        this.state = {};
        this.sampleProp = props.sampleProp;
    }

    public async componentDidMount() {
        this.initializeState();
        console.log(SDK.getUser());
    }

    private async initializeState(): Promise<void> {
        await SDK.ready();
        const accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>(
            CommonServiceIds.ExtensionDataService
        );
        this._dataManager = await extDataService.getExtensionDataManager(
            SDK.getExtensionContext().id,
            accessToken
        );

        await this.getGroups(accessToken);
        await this.getUsers(accessToken);

        this.setState({ ready: true });

        this._dataManager.getValue<AllowedEntity[]>('allowed-user-groups').then(
            (data) => {
                console.log('data is this ', data);
                this.userGroupsSelection.clear();
                for (const selectedUserGroup of data) {
                    console.log(
                        'searching the user gruops for ',
                        selectedUserGroup
                    );
                    const idx = this.userGroups.findIndex(
                        (item) => item.originId === selectedUserGroup.originId
                    );
                    if (idx >= 0) {
                        this.userGroupsSelection.select(idx);
                    }
                    console.log('would have selected gruop ', idx);
                }
                this.setState({
                    dataAllowedUserGroups: data,
                    persistedAllowedUserGroups: data,
                    ready: true,
                });
            },
            () => {
                this.setState({
                    dataAllowedUserGroups: [],
                    ready: true,
                });
            }
        );

        this._dataManager.getValue<AllowedEntity[]>('allowed-users').then(
            (data) => {
                this.usersSelection.clear();
                for (const selectedUser of data) {
                    const idx = this.users.findIndex(
                        (item) => item.originId === selectedUser.originId
                    );
                    if (idx >= 0) {
                        this.usersSelection.select(idx);
                    }
                    console.log('would have selected user ', idx);
                }
                this.setState({
                    dataAllowedUsers: data,
                    persistedAllowedUsers: data,
                    ready: true,
                });
            },
            () => {
                this.setState({
                    dataAllowedUsers: [],
                    ready: true,
                });
            }
        );
    }

    private async getGraphResource(
        resouce: string,
        accessToken: string,
        callback: (data: any) => void
    ): Promise<void> {
        // TODO: extract the organization name globally
        axios
            .get(
                `https://vssps.dev.azure.com/reponzo01/_apis/graph/${resouce}`,
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                    },
                }
            )
            .then((res) => {
                console.log(res.data);
                callback(res.data);
            })
            .catch((error) => {
                console.error(error);
                throw error;
            });
    }

    private async getGroups(accessToken: string): Promise<void> {
        return new Promise((resolve) => {
            this.getGraphResource('groups', accessToken, (data: any) => {
                this.userGroups = [];
                for (const group in data.value) {
                    this.userGroups.push({
                        displayName: data.value[group].displayName,
                        originId: data.value[group].originId,
                        descriptor: data.value[group].descriptor,
                    });
                }
                console.log('resolving getGroups with ', this.userGroups);
                resolve();
            });
        });
    }

    private async getUsers(accessToken: string): Promise<void> {
        return new Promise((resolve) => {
            this.getGraphResource('users', accessToken, (data: any) => {
                this.users = [];
                for (const user in data.value) {
                    this.users.push({
                        displayName: data.value[user].displayName,
                        originId: data.value[user].originId,
                        descriptor: data.value[user].descriptor,
                    });
                }
                console.log('resolving getUsers with ', this.users);
                resolve();
            });
        });
    }

    public render() {
        const {
            dataAllowedUserGroups,
            dataAllowedUsers,
            ready,
            persistedAllowedUserGroups,
            persistedAllowedUsers,
        } = this.state;

        console.log('returning a render');
        return (
            <div className="page-content page-content-top flex-column rhythm-vertical-16">
                <div>
                    By default the Azure groups{' '}
                    <u>
                        <code>Dev Team Leads</code>
                    </u>{' '}
                    and{' '}
                    <u>
                        <code>DevOps</code>
                    </u>{' '}
                    have access to this extension. Use the dropdowns to add more{' '}
                    groups or individual users.
                </div>
                <div className="flex-column">
                    <Observer selection={this.userGroupsSelection}>
                        {() => {
                            return (
                                <Dropdown
                                    ariaLabel="Multiselect"
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
                                    className="example-dropdown flex-column"
                                    items={this.userGroups.map((item) => item.displayName)}
                                    selection={this.userGroupsSelection}
                                    placeholder="Select User Groups"
                                    showFilterBox={true}
                                />
                            );
                        }}
                    </Observer>
                </div>
                <div className="flex-column">
                    <Observer selection={this.usersSelection}>
                        {() => {
                            return (
                                <Dropdown
                                    ariaLabel="Multiselect"
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
                                    className="example-dropdown flex-column"
                                    items={this.users.map((item) => item.displayName)}
                                    selection={this.usersSelection}
                                    placeholder="Select Individual Users"
                                    showFilterBox={true}
                                />
                            );
                        }}
                    </Observer>
                </div>
                <div className="bolt-button-group flex-row rhythm-horizontal-8">
                    <Button
                        text="Save"
                        primary={true}
                        onClick={this.onSaveData}
                        disabled={!ready}
                    />
                </div>
            </div>
        );
    }

    private onSaveData = (): void => {
        const {
            dataAllowedUserGroups,
            dataAllowedUsers,
            ready,
            persistedAllowedUserGroups,
            persistedAllowedUsers,
        } = this.state;

        this.setState({ ready: false });

        const userGroupsSelectedArray: AllowedEntity[] = this.setSelectionRange(
            this.userGroupsSelection.value,
            this.userGroups
        );
        const usersSelectedArray: AllowedEntity[] = this.setSelectionRange(
            this.usersSelection.value,
            this.users
        );

        this._dataManager!.setValue<AllowedEntity[]>(
            'allowed-user-groups',
            userGroupsSelectedArray || []
        ).then(() => {
            this.setState({
                ready: true,
                persistedAllowedUserGroups: userGroupsSelectedArray,
            });
        });

        this._dataManager!.setValue<AllowedEntity[]>(
            'allowed-users',
            usersSelectedArray || []
        ).then(() => {
            this.setState({
                ready: true,
                persistedAllowedUsers: usersSelectedArray,
            });
        });
    };

    private setSelectionRange(
        selectionRange: ISelectionRange[],
        dataArray: AllowedEntity[]
    ): AllowedEntity[] {
        const selectedArray: AllowedEntity[] = [];
        for (const rng of selectionRange) {
            var sliced = dataArray.slice(rng.beginIndex, rng.endIndex + 1);
            for (const slic of sliced) {
                selectedArray.push(slic);
            }
        }
        return selectedArray;
    }
}
