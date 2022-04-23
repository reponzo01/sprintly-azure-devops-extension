import * as React from 'react';
import axios, { AxiosResponse } from 'axios';

import * as SDK from 'azure-devops-extension-sdk';
import { Card } from 'azure-devops-ui/Card';
import {
    ObservableArray,
    ObservableValue,
} from 'azure-devops-ui/Core/Observable';
import {
    ColumnSorting,
    ITableColumn,
    SimpleTableCell,
    sortItems,
    SortOrder,
    Table,
} from 'azure-devops-ui/Table';
import * as Common from './SprintlyCommon';
import { IExtensionDataManager } from 'azure-devops-extension-api';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { Observer } from 'azure-devops-ui/Observer';
import { Tooltip } from 'azure-devops-ui/TooltipEx';
import { ButtonGroup } from 'azure-devops-ui/ButtonGroup';
import { Button } from 'azure-devops-ui/Button';

export interface ISprintlyEnvironmentVariableViewerState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    environmentVariablesObservable: ObservableArray<ISearchResultEnvironmentVariableItem>;
}

export interface ISearchResultEnvironmentVariableValue {
    environmentName: string;
    value: string;
}

export interface ISearchResultEnvironmentVariableItem {
    name: string;
    values: ISearchResultEnvironmentVariableValue[];
}

//#region "Observables"
//#endregion "Observables"

const rawTableItems: ISearchResultEnvironmentVariableItem[] = [
    {
        name: 't1',
        values: [
            {
                environmentName: 'Dev',
                value: 'var1',
            },
        ],
    },
];
const tableItems = new ObservableArray<ISearchResultEnvironmentVariableItem>(
    rawTableItems
);

export default class SprintlyEnvironmentVariableViewer extends React.Component<
    {
        dataManager: IExtensionDataManager;
        organizationName: string;
    },
    ISprintlyEnvironmentVariableViewerState
> {
    private dataManager: IExtensionDataManager;
    private organizationName: string;
    private accessToken: string = '';
    private environmentVariablesResponse: any;
    private environmentVariablesExclusionFilter: Set<string> = new Set();
    private columns: ITableColumn<ISearchResultEnvironmentVariableItem>[] = [];
    private sortingBehavior: ColumnSorting<ISearchResultEnvironmentVariableItem> =
        new ColumnSorting<ISearchResultEnvironmentVariableItem>(
            (
                columnIndex: number,
                proposedSortOrder: SortOrder,
                event:
                    | React.KeyboardEvent<HTMLElement>
                    | React.MouseEvent<HTMLElement>
            ) => {
                this.state.environmentVariablesObservable.splice(
                    0,
                    this.state.environmentVariablesObservable.length,
                    ...sortItems<ISearchResultEnvironmentVariableItem>(
                        columnIndex,
                        proposedSortOrder,
                        this.sortFunctions,
                        this.columns,
                        this.state.environmentVariablesObservable.value
                    )
                );
            }
        );
    private sortFunctions: any = [
        (
            a: ISearchResultEnvironmentVariableItem,
            b: ISearchResultEnvironmentVariableItem
        ): number => {
            return a.name.localeCompare(b.name);
        },
    ];

    constructor(props: {
        dataManager: IExtensionDataManager;
        organizationName: string;
    }) {
        super(props);

        this.onSize = this.onSize.bind(this);

        this.columns = [
            {
                id: 'environmentVariableName',
                name: 'Environment Variable',
                onSize: this.onSize,
                renderCell: this.renderEnvironmentVariableNameCell,
                sortProps: {
                    ariaLabelAscending: 'Sorted A to Z',
                    ariaLabelDescending: 'Sorted Z to A',
                },
                width: new ObservableValue<number>(-30),
            },
        ];

        this.state = {
            environmentVariablesObservable:
                new ObservableArray<ISearchResultEnvironmentVariableItem>([]),
        };

        this.dataManager = props.dataManager;
        this.organizationName = props.organizationName;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        this.accessToken = await SDK.getAccessToken();

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

        this.setState({
            userSettings,
            systemSettings,
        });

        await this.loadEnvironmentVariables();

        // repositoriesToProcess = Common.getSavedRepositoriesToView(
        //     this.state.userSettings,
        //     this.state.systemSettings
        // );

        // totalRepositoriesToProcessObservable.value =
        //     repositoriesToProcess.length;
        // if (repositoriesToProcess.length > 0) {
        //     const filteredProjects: TeamProjectReference[] =
        //         await Common.getFilteredProjects();
        //     await this.loadRepositoriesDisplayState(filteredProjects);
        // }
    }

    private async loadEnvironmentVariables(): Promise<void> {
        const currentProject = await Common.getCurrentProject();
        if (currentProject !== undefined) {
            let environmentVariableGroupIds: string = '';
            for (const groupId of Common.ALLOWED_ENVIRONMENT_VARIABLE_GROUP_IDS) {
                environmentVariableGroupIds += `${groupId.toString()},`;
            }
            const url: string = `https://dev.azure.com/${this.organizationName}/${currentProject.id}/_apis/distributedtask/variablegroups?groupIds=${environmentVariableGroupIds}`;
            this.accessToken = await Common.getOrRefreshToken(this.accessToken);
            const response: AxiosResponse<never> = await axios
                .get(url, {
                    headers: {
                        Authorization: `Bearer ${this.accessToken}`,
                    },
                })
                .catch((error: any) => {
                    console.error(error);
                    throw error;
                });
            this.environmentVariablesResponse = response.data; //No defined type exists in the api

            this.resetEnvironmentVariablesColumns();
        }
    }

    private resetEnvironmentVariablesColumns(): void {
        const resultEnvironmentVariables: ISearchResultEnvironmentVariableItem[] =
            [];
        this.columns.splice(1, this.columns.length - 1);
        for (const environmentVariableGroup of this.environmentVariablesResponse
            .value) {
            if (
                !this.environmentVariablesExclusionFilter.has(
                    environmentVariableGroup.name
                )
            ) {
                this.columns.push({
                    id: `environment${environmentVariableGroup.name}`,
                    name: environmentVariableGroup.name,
                    onSize: this.onSize,
                    renderCell: this.renderEnvironmentVariableValueCell,
                    width: new ObservableValue<number>(-30),
                });
                for (const [
                    environmentVariableName,
                    environmentVariableValue,
                ] of Object.entries(environmentVariableGroup.variables)) {
                    let variableIsSaved: boolean = false;
                    for (const environmentVariable of resultEnvironmentVariables) {
                        if (
                            environmentVariableName === environmentVariable.name
                        ) {
                            variableIsSaved = true;
                            environmentVariable.values.push({
                                environmentName: environmentVariableGroup.name,
                                value: (environmentVariableValue as any).value,
                            });
                        }
                    }
                    if (!variableIsSaved) {
                        resultEnvironmentVariables.push({
                            name: environmentVariableName,
                            values: [
                                {
                                    environmentName:
                                        environmentVariableGroup.name,
                                    value: (environmentVariableValue as any)
                                        .value,
                                },
                            ],
                        });
                    }
                }
            }
        }
        this.setState({
            environmentVariablesObservable:
                new ObservableArray<ISearchResultEnvironmentVariableItem>(
                    //TODO: Sort by variable name
                    resultEnvironmentVariables
                ),
        });
    }

    private updateEnvironmentVariablesExcludeFilter(
        environmentName: string,
        show: boolean
    ): void {
        if (show) {
            this.environmentVariablesExclusionFilter.delete(environmentName);
        } else {
            this.environmentVariablesExclusionFilter.add(environmentName);
        }
        this.resetEnvironmentVariablesColumns();
    }

    private renderEnvironmentVariablesExcludeFilterCheckboxes(): JSX.Element {
        if (this.environmentVariablesResponse !== undefined) {
            return (
                <>
                    {this.environmentVariablesResponse.value.map(
                        (environment: any) => (
                            <Checkbox
                                key={environment.name}
                                onChange={(event, checked) =>
                                    this.updateEnvironmentVariablesExcludeFilter(
                                        environment.name,
                                        checked
                                    )
                                }
                                checked={
                                    !this.environmentVariablesExclusionFilter.has(
                                        environment.name
                                    )
                                }
                                label={`Show ${environment.name}`}
                            />
                        )
                    )}
                </>
            );
        }
        return <></>;
    }

    private renderEnvironmentVariableNameCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<ISearchResultEnvironmentVariableItem>,
        tableItem: ISearchResultEnvironmentVariableItem
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
                children={<>{tableItem.name}</>}
            ></SimpleTableCell>
        );
    }

    private renderEnvironmentVariableValueCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<ISearchResultEnvironmentVariableItem>,
        tableItem: ISearchResultEnvironmentVariableItem
    ): JSX.Element {
        let itemValue: String = '';
        for (const value of tableItem.values) {
            if (value.environmentName === tableColumn.name) {
                itemValue = value.value;
            }
        }
        return (
            <SimpleTableCell
                key={'col-' + columnIndex}
                columnIndex={columnIndex}
                tableColumn={tableColumn}
            >
                <div className='flex-row scroll-hidden'>
                    <Tooltip overflowOnly={true}>
                        <span className='text-ellipsis'>{itemValue}</span>
                    </Tooltip>
                </div>
            </SimpleTableCell>
        );
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (this.columns[index].width as ObservableValue<number>).value = width;
    }

    public render(): JSX.Element {
        return (
            <Observer
                environmentVariables={this.state.environmentVariablesObservable}
            >
                {(props: {
                    environmentVariables: ISearchResultEnvironmentVariableItem[];
                }) => (
                    <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                        <ButtonGroup>
                            <Button
                                text='Show all environment variables'
                                primary={true}
                                onClick={() => {}}
                            />
                            <Button
                                text='Repository specific variables'
                                primary={true}
                                onClick={() => {}}
                            />
                        </ButtonGroup>
                        <div className='rhythm-horizontal-8 flex-row'>
                            {this.renderEnvironmentVariablesExcludeFilterCheckboxes()}
                        </div>
                        <Card className='bolt-table-card bolt-card-white'>
                            <Table
                                columns={this.columns}
                                behaviors={[this.sortingBehavior]}
                                selectableText={true}
                                itemProvider={
                                    this.state.environmentVariablesObservable
                                }
                            />
                        </Card>
                    </div>
                )}
            </Observer>
        );
    }
}
