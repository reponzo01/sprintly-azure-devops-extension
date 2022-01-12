import * as React from 'react';

import { getClient, IExtensionDataManager } from 'azure-devops-extension-api';
import { GitRef } from 'azure-devops-extension-api/Git';

import {
    ObservableArray,
    ObservableValue,
} from 'azure-devops-ui/Core/Observable';

import { Button } from 'azure-devops-ui/Button';
import { ButtonGroup } from 'azure-devops-ui/ButtonGroup';
import { Card } from 'azure-devops-ui/Card';
import { Page } from 'azure-devops-ui/Page';
import { TextField, TextFieldWidth } from 'azure-devops-ui/TextField';
import { ColumnSorting, ITableColumn, sortItems, SortOrder, Table } from 'azure-devops-ui/Table';

//#region "Observables"
const searchObservable = new ObservableValue<string>('');
const nameColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-30);
const repositoryColumnWidthObservable: ObservableValue<number> =
    new ObservableValue<number>(-40);
//#endregion "Observables"

export interface ISprintlyBranchSearchPageState {
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
                        onClick={() => alert(searchObservable.value)}
                    />
                </ButtonGroup>
                <Page>
                    <Card className='bolt-table-card bolt-card-white'>
                        <Table
                            columns={this.columns}
                            behaviors={[
                                this.sortingBehavior,
                            ]}
                            itemProvider={this.state.searchResultBranches}
                        />
                    </Card>
                </Page>
            </div>
        );
    }
}
