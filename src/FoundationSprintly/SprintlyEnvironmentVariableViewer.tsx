import * as React from 'react';
import axios, { AxiosResponse } from 'axios';

import * as Common from './SprintlyCommon';
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
import {
    IExtensionDataManager,
    IProjectInfo,
} from 'azure-devops-extension-api';
import { FilterBar } from 'azure-devops-ui/FilterBar';
import { KeywordFilterBarItem } from 'azure-devops-ui/TextFilterBarItem';
import {
    Filter,
    FILTER_CHANGE_EVENT,
    IFilterState,
} from 'azure-devops-ui/Utilities/Filter';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { Observer } from 'azure-devops-ui/Observer';
import { Tooltip } from 'azure-devops-ui/TooltipEx';
import { ButtonGroup } from 'azure-devops-ui/ButtonGroup';
import { Button } from 'azure-devops-ui/Button';
import { GitRepository } from 'azure-devops-extension-api/Git';
import {
    IListItemDetails,
    List,
    ListItem,
    ListSelection,
} from 'azure-devops-ui/List';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import { bindSelectionToObservable } from 'azure-devops-ui/MasterDetailsContext';
import { Spinner } from 'azure-devops-ui/Spinner';
import {
    Splitter,
    SplitterDirection,
    SplitterElementPosition,
} from 'azure-devops-ui/Splitter';

export interface ISprintlyEnvironmentVariableViewerState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    environmentVariablesObservable: ObservableArray<ISearchResultEnvironmentVariableItem>;
    repositories: ArrayItemProvider<GitRepository>;
    repositoryListSelection: ListSelection;
    repositoryListSelectedItemObservable: ObservableValue<GitRepository>;
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
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const showAllEnvironmentVariablesObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(true);
const environmentVariableSearchFilterCurrentState: ObservableValue<IFilterState> =
    new ObservableValue<any>({});
//#endregion "Observables"

const environmentVariableNameFilterKey: string =
    'environmentVariableNameFilterKey';
const environmentVariableValueFilterKey: string =
    'environmentVariableValueFilterKey';

let repositoriesToProcess: string[] = [];

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
    private environmentVariableNameSearchFilter: Filter;
    private environmentVariableValueSearchFilter: Filter;
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
        this.renderRepositoryMasterPageList =
            this.renderRepositoryMasterPageList.bind(this);
        this.renderDetailPageContent = this.renderDetailPageContent.bind(this);

        this.justATest = this.justATest.bind(this);
        this.justATest2 = this.justATest2.bind(this);

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

        this.environmentVariableNameSearchFilter = new Filter();
        this.environmentVariableValueSearchFilter = new Filter();
        this.environmentVariableNameSearchFilter.subscribe(() => {
            this.redrawEnvironmentVariablesSearchResult();
        }, FILTER_CHANGE_EVENT);
        this.environmentVariableValueSearchFilter.subscribe(() => {
            this.redrawEnvironmentVariablesSearchResult();
        }, FILTER_CHANGE_EVENT);

        this.state = {
            environmentVariablesObservable:
                new ObservableArray<ISearchResultEnvironmentVariableItem>([]),
            repositories: new ArrayItemProvider<GitRepository>([]),
            repositoryListSelection: new ListSelection({
                selectOnFocus: false,
            }),
            repositoryListSelectedItemObservable: new ObservableValue<any>({}),
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

        repositoriesToProcess = Common.getSavedRepositoriesToView(
            this.state.userSettings,
            this.state.systemSettings
        );

        totalRepositoriesToProcessObservable.value =
            repositoriesToProcess.length;
        if (repositoriesToProcess.length > 0) {
            const currentProject: IProjectInfo | undefined =
                await Common.getCurrentProject();
            await this.loadRepositoriesDisplayState(currentProject);
        }
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

            this.redrawEnvironmentVariablesSearchResult();
        }
    }

    private justATest(): void {
        console.log(this.environmentVariableNameSearchFilter.getState());
        var s: IFilterState =
            this.environmentVariableNameSearchFilter.getState();
        if (s[environmentVariableNameFilterKey] === undefined) {
            console.log('empty filter state');
        } else {
            console.log(s[environmentVariableNameFilterKey]!.value);
            console.log(s[environmentVariableNameFilterKey]!.value.length);
        }
    }
    private justATest2(): void {
        console.log(this.environmentVariableValueSearchFilter.getState());
    }

    private redrawEnvironmentVariablesSearchResult(): void {
        const resultEnvironmentVariables: ISearchResultEnvironmentVariableItem[] =
            [];
        let environmentVariableNameSearchFilterString: string = '';
        let environmentVariableValueSearchFilterString: string = '';

        var environmentVariableNameSearchFilterState: IFilterState =
            this.environmentVariableNameSearchFilter.getState();
        var environmentVariableValueSearchFilterState: IFilterState =
            this.environmentVariableValueSearchFilter.getState();

        if (
            environmentVariableNameSearchFilterState[
                environmentVariableNameFilterKey
            ] !== undefined
        ) {
            environmentVariableNameSearchFilterString =
                environmentVariableNameSearchFilterState[
                    environmentVariableNameFilterKey
                ]!.value;
        }

        if (
            environmentVariableValueSearchFilterState[
                environmentVariableValueFilterKey
            ] !== undefined
        ) {
            environmentVariableValueSearchFilterString =
                environmentVariableValueSearchFilterState[
                    environmentVariableValueFilterKey
                ]!.value;
        }

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
                            if (
                                environmentVariableValueSearchFilterString.length ===
                                    0 ||
                                (environmentVariableValue as any).value
                                    .toLowerCase()
                                    .includes(
                                        environmentVariableValueSearchFilterString.toLowerCase()
                                    )
                            ) {
                                environmentVariable.values.push({
                                    environmentName:
                                        environmentVariableGroup.name,
                                    value: (environmentVariableValue as any)
                                        .value,
                                });
                            }
                        }
                    }
                    if (!variableIsSaved) {
                        if (
                            environmentVariableNameSearchFilterString.length ===
                                0 ||
                            environmentVariableName
                                .toLowerCase()
                                .includes(
                                    environmentVariableNameSearchFilterString.toLowerCase()
                                )
                        ) {
                            if (
                                environmentVariableValueSearchFilterString.length ===
                                    0 ||
                                (environmentVariableValue as any).value
                                    .toLowerCase()
                                    .includes(
                                        environmentVariableValueSearchFilterString.toLowerCase()
                                    )
                            ) {
                                resultEnvironmentVariables.push({
                                    name: environmentVariableName,
                                    values: [
                                        {
                                            environmentName:
                                                environmentVariableGroup.name,
                                            value: (
                                                environmentVariableValue as any
                                            ).value,
                                        },
                                    ],
                                });
                            }
                        }
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
        this.redrawEnvironmentVariablesSearchResult();
    }

    private async loadRepositoriesDisplayState(
        currentProject: IProjectInfo | undefined
    ): Promise<void> {
        let repos: GitRepository[] = [];
        totalRepositoriesToProcessObservable.value = 0;
        if (currentProject !== undefined) {
            const filteredRepos: GitRepository[] =
                await Common.getFilteredProjectRepositories(
                    currentProject.id,
                    repositoriesToProcess
                );

            totalRepositoriesToProcessObservable.value += filteredRepos.length;
            repos = repos.concat(filteredRepos);
        }

        this.setState({
            repositories: new ArrayItemProvider(
                Common.sortRepositoryList(repos)
            ),
        });

        bindSelectionToObservable(
            this.state.repositoryListSelection,
            this.state.repositories,
            this.state
                .repositoryListSelectedItemObservable as ObservableValue<GitRepository>
        );
    }

    private renderRepositoryMasterPageList(): JSX.Element {
        return !this.state.repositories ||
            this.state.repositories.length === 0 ? (
            <div className='page-content-top'>
                <Spinner label='loading' />
            </div>
        ) : (
            <List
                ariaLabel={'Repositories'}
                itemProvider={this.state.repositories}
                selection={this.state.repositoryListSelection}
                renderRow={this.renderRepositoryListItem}
                width='100%'
                singleClickActivation={true}
                onSelect={async () => {
                    await this.selectRepository();
                }}
            />
        );
    }

    private renderDetailPageContent(): JSX.Element {
        //TODO: get release defs to be passed in from page 1
        //get single release for repo
        //get the inline transform variables and parse them out
        //build a tree table nesting each environment under each variable (root)
        //show appsettings transform at root level, second column, and environment value for nested children.
        return <></>;
    }

    private renderRepositoryListItem(
        index: number,
        item: GitRepository,
        details: IListItemDetails<GitRepository>,
        key?: string
    ): JSX.Element {
        return (
            <ListItem
                className='master-row border-bottom'
                key={key || 'list-item' + index}
                index={index}
                details={details}
            >
                <div className='master-row-content flex-row flex-center h-scroll-hidden'>
                    <div className='flex-column text-ellipsis'>
                        <Tooltip overflowOnly={true}>
                            <div className='primary-text text-ellipsis'>
                                {Common.repositoryLinkJsxElement(
                                    item.webUrl,
                                    'font-size-1',
                                    item.name
                                )}
                            </div>
                        </Tooltip>
                    </div>
                </div>
            </ListItem>
        );
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

    private async selectRepository(): Promise<void> {}

    private onSize(event: MouseEvent, index: number, width: number): void {
        (this.columns[index].width as ObservableValue<number>).value = width;
    }

    public render(): JSX.Element {
        return (
            <Observer
                environmentVariables={this.state.environmentVariablesObservable}
                showAllEnvironmentVariables={
                    showAllEnvironmentVariablesObservable
                }
            >
                {(props: {
                    environmentVariables: ISearchResultEnvironmentVariableItem[];
                }) => (
                    <div className='page-content page-content-top flex-column rhythm-vertical-16'>
                        <ButtonGroup>
                            <Button
                                text='Show all environment variables'
                                primary={true}
                                onClick={() => {
                                    showAllEnvironmentVariablesObservable.value =
                                        true;
                                }}
                            />
                            <Button
                                text='Repository specific variables'
                                primary={true}
                                onClick={() => {
                                    showAllEnvironmentVariablesObservable.value =
                                        false;
                                }}
                            />
                        </ButtonGroup>
                        {showAllEnvironmentVariablesObservable.value && (
                            <>
                                <div className='rhythm-horizontal-8 flex-row'>
                                    {this.renderEnvironmentVariablesExcludeFilterCheckboxes()}
                                </div>
                                <div className='rhythm-horizontal-8 flex-row'>
                                    <div className='flex-grow'>
                                        <FilterBar
                                            filter={
                                                this
                                                    .environmentVariableNameSearchFilter
                                            }
                                        >
                                            <KeywordFilterBarItem
                                                placeholder='Filter by variable name'
                                                filterItemKey={
                                                    environmentVariableNameFilterKey
                                                }
                                            />
                                        </FilterBar>
                                    </div>
                                    <div className='flex-grow sprintly-margin-right-auto'>
                                        <FilterBar
                                            filter={
                                                this
                                                    .environmentVariableValueSearchFilter
                                            }
                                        >
                                            <KeywordFilterBarItem
                                                placeholder='Filter by value'
                                                filterItemKey={
                                                    environmentVariableValueFilterKey
                                                }
                                            />
                                        </FilterBar>
                                    </div>
                                </div>
                                <Card className='bolt-table-card bolt-card-white'>
                                    <Table
                                        columns={this.columns}
                                        behaviors={[this.sortingBehavior]}
                                        selectableText={true}
                                        itemProvider={
                                            this.state
                                                .environmentVariablesObservable
                                        }
                                    />
                                </Card>
                            </>
                        )}
                        {!showAllEnvironmentVariablesObservable.value && (
                            <>
                                <Splitter
                                    fixedElement={SplitterElementPosition.Near}
                                    splitterDirection={
                                        SplitterDirection.Vertical
                                    }
                                    initialFixedSize={450}
                                    minFixedSize={100}
                                    nearElementClassName='v-scroll-auto custom-scrollbar'
                                    farElementClassName='v-scroll-auto custom-scrollbar'
                                    onRenderNearElement={
                                        this.renderRepositoryMasterPageList
                                    }
                                    onRenderFarElement={
                                        this.renderDetailPageContent
                                    }
                                />
                            </>
                        )}
                    </div>
                )}
            </Observer>
        );
    }
}
