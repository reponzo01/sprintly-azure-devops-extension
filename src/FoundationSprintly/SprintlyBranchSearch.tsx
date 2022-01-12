import * as React from 'react';

import { getClient, IExtensionDataManager } from 'azure-devops-extension-api';
import { GitRef, GitRestClient } from 'azure-devops-extension-api/Git';

import {
    ObservableArray,
    ObservableValue,
} from 'azure-devops-ui/Core/Observable';
import { TeamProjectReference } from 'azure-devops-extension-api/Core';

import { Button } from 'azure-devops-ui/Button';
import { ButtonGroup } from 'azure-devops-ui/ButtonGroup';
import { Card } from 'azure-devops-ui/Card';
import { Page } from 'azure-devops-ui/Page';
import { TextField, TextFieldWidth } from 'azure-devops-ui/TextField';
import {
    ColumnSorting,
    ITableColumn,
    sortItems,
    SortOrder,
    Table,
} from 'azure-devops-ui/Table';

import * as Common from './SprintlyCommon';

//#region "Observables"
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const searchObservable = new ObservableValue<string>('');
const nameColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const repositoryColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-40);
//#endregion "Observables"

const userSettingsDataManagerKey: string = 'user-settings';
const systemSettingsDataManagerKey: string = 'system-settings';

let repositoriesToProcess: string[] = [];

export interface ISprintlyBranchSearchPageState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    searchResultBranches: ObservableArray<GitRef>;
}

export default class SprintlyBranchSearchPage extends React.Component<
    { dataManager: IExtensionDataManager },
    ISprintlyBranchSearchPageState
> {
    private dataManager: IExtensionDataManager;
    private columns: any = [];
    private sortingBehavior: ColumnSorting<GitRef> = new ColumnSorting<GitRef>(
        (
            columnIndex: number,
            proposedSortOrder: SortOrder,
            event:
                | React.KeyboardEvent<HTMLElement>
                | React.MouseEvent<HTMLElement>
        ) => {
            this.state.searchResultBranches.splice(
                0,
                this.state.searchResultBranches.length,
                ...sortItems<GitRef>(
                    columnIndex,
                    proposedSortOrder,
                    this.sortFunctions,
                    this.columns,
                    this.state.searchResultBranches.value
                )
            );
        }
    );
    private sortFunctions: any = [
        (a: GitRef, b: GitRef): number => {
            return a.name.localeCompare(b.name);
        },
        null,
        (a: GitRef, b: GitRef): number => {
            return a.creator.displayName.localeCompare(b.creator.displayName);
        },
    ];

    constructor(props: { dataManager: IExtensionDataManager }) {
        super(props);

        this.onSize = this.onSize.bind(this);
        this.renderNameCell = this.renderNameCell.bind(this);
        this.renderRepositoryCell = this.renderRepositoryCell.bind(this);

        this.columns = [
            {
                id: 'name',
                name: 'Branch',
                onSize: this.onSize,
                renderCell: this.renderNameCell,
                sortProps: {
                    ariaLabelAscending: 'Sorted A to Z',
                    ariaLabelDescending: 'Sorted Z to A',
                },
                width: nameColumnWidthObservable,
            },
            {
                id: 'repository',
                name: 'Repository',
                onSize: this.onSize,
                renderCell: this.renderRepositoryCell,
                sortProps: {
                    ariaLabelAscending: 'Sorted A to Z',
                    ariaLabelDescending: 'Sorted Z to A',
                },
                width: repositoryColumnWidthObservable,
            },
        ];

        this.state = {
            searchResultBranches: new ObservableArray<GitRef>([]),
        };

        this.dataManager = props.dataManager;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
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

        this.setState({
            userSettings,
            systemSettings,
        });

        repositoriesToProcess = Common.getSavedRepositoriesToView(
            this.state.userSettings,
            this.state.systemSettings
        );

        totalRepositoriesToProcessObservable.value =
            repositoriesToProcess.length;
    }

    private renderNameCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<GitRef>,
        tableItem: GitRef
    ): JSX.Element {
        return <></>;
    }

    private renderRepositoryCell(
        rowIndex: number,
        columnIndex: number,
        tableColumn: ITableColumn<GitRef>,
        tableItem: GitRef
    ): JSX.Element {
        return <></>;
    }

    private async searchAction(): Promise<void> {
        const searchTerm: string = searchObservable.value.trim();
        if (searchTerm && totalRepositoriesToProcessObservable.value > 0) {
            for (const repositoryId of repositoriesToProcess) {
                const repositoryBranches: GitRef[] = await getClient(
                    GitRestClient
                ).getRefs(
                    repositoryId,
                    undefined,
                    Common.repositoryHeadsFilter,
                    undefined,
                    undefined,
                    undefined,
                    undefined,
                    undefined,
                    'feature/'
                );
                for (const branch of repositoryBranches) {
                    console.log(branch);
                }
            }
        }
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (this.columns[index].width as ObservableValue<number>).value = width;
    }

    public render(): JSX.Element {
        return (
            <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                <ButtonGroup>
                    <TextField
                        prefixIconProps={{ iconName: 'Search' }}
                        value={searchObservable}
                        onChange={(e, newValue) =>
                            (searchObservable.value = newValue)
                        }
                        placeholder='Search Branches'
                        width={TextFieldWidth.standard}
                    />
                    <Button
                        text='Search'
                        primary={true}
                        onClick={async () => await this.searchAction()}
                    />
                </ButtonGroup>
                <Page>
                    <Card className='bolt-table-card bolt-card-white'>
                        <Table
                            columns={this.columns}
                            behaviors={[this.sortingBehavior]}
                            itemProvider={this.state.searchResultBranches}
                        />
                    </Card>
                </Page>
            </div>
        );
    }
}
