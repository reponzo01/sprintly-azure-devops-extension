import * as React from 'react';
import axios, { AxiosResponse } from 'axios';

import * as SDK from 'azure-devops-extension-sdk';
import { Card } from 'azure-devops-ui/Card';
import {
    ObservableArray,
    ObservableValue,
} from 'azure-devops-ui/Core/Observable';
import { ITableColumn, SimpleTableCell, Table } from 'azure-devops-ui/Table';
import * as Common from './SprintlyCommon';
import { IExtensionDataManager } from 'azure-devops-extension-api';

export interface ISprintlyEnvironmentVariableViewerState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    environmentVariablesObservable: ObservableArray<ISearchResultEnvironmentVariableItem>;
}

export interface ISearchResultEnvironmentVariableItem {
    test1: string;
    test2?: string;
    test3: string;
}

//#region "Observables"
const test1ColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const test2ColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const test3ColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-40);
//#endregion "Observables"

const rawTableItems: ISearchResultEnvironmentVariableItem[] = [
    {
        test1: 't1',
        test2: 't2',
        test3: 't3',
    },
    {
        test1: 't1',
        test3: 't3',
    },
    {
        test1: 't1',
        test2: 't2',
        test3: 't3',
    },
];
const tableItems = new ObservableArray<ISearchResultEnvironmentVariableItem>(
    rawTableItems
);

export default class SprintlyEnvironmentVariableViewer extends React.Component<
    {
        dataManager: IExtensionDataManager;
    },
    ISprintlyEnvironmentVariableViewerState
> {
    private dataManager: IExtensionDataManager;
    private accessToken: string = '';
    private columns: any = [];

    constructor(props: { dataManager: IExtensionDataManager }) {
        super(props);

        this.columns = [
            {
                id: 'name',
                name: 'Test1',
                onSize: this.onSize,
                renderCell: this.renderTest1,
                sortProps: {
                    ariaLabelAscending: 'Sorted A to Z',
                    ariaLabelDescending: 'Sorted Z to A',
                },
                width: test1ColumnWidthObservable,
            },
            {
                id: 'name',
                name: 'Test2',
                onSize: this.onSize,
                renderCell: this.renderTest2,
                sortProps: {
                    ariaLabelAscending: 'Sorted A to Z',
                    ariaLabelDescending: 'Sorted Z to A',
                },
                width: test2ColumnWidthObservable,
            },
            {
                id: 'name',
                name: 'Test3',
                onSize: this.onSize,
                renderCell: this.renderTest3,
                sortProps: {
                    ariaLabelAscending: 'Sorted A to Z',
                    ariaLabelDescending: 'Sorted Z to A',
                },
                width: test3ColumnWidthObservable,
            },
        ];

        this.state = {
            environmentVariablesObservable:
                new ObservableArray<ISearchResultEnvironmentVariableItem>([]),
        };

        this.dataManager = props.dataManager;
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
        const project = await Common.getCurrentProject();
        console.log(project);
        //const url: string = `https://vsrm.dev.azure.com/${this.organizationName}/${clickedDeployProjectReferenceObservable.value.id}/_apis/Release/releases/${clickedDeployReleaseIdObservable.value}/environments/${clickedDeployEnvironmentObservable.value.id}?api-version=5.0-preview.6`;
    }

    private renderTest1(
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
                children={<>{tableItem.test1}</>}
            ></SimpleTableCell>
        );
    }

    private renderTest2(
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
                children={<>{tableItem.test2}</>}
            ></SimpleTableCell>
        );
    }

    private renderTest3(
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
                children={<>{tableItem.test3}</>}
            ></SimpleTableCell>
        );
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (this.columns[index].width as ObservableValue<number>).value = width;
    }

    public render(): JSX.Element {
        return (
            <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                <Card className='bolt-table-card bolt-card-white'>
                    <Table
                        columns={this.columns}
                        //behaviors={[this.sortingBehavior]}
                        //itemProvider={this.state.repositoryBranchesObservable}
                        itemProvider={tableItems}
                    />
                </Card>
            </div>
        );
    }
}
