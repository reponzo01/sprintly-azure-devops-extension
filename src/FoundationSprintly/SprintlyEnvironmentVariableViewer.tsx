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
    getClient,
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
import { Page } from 'azure-devops-ui/Page';
import { ITreeColumn, Tree } from 'azure-devops-ui/TreeEx';
import {
    ITreeItem,
    ITreeItemEx,
    ITreeItemProvider,
    TreeItemProvider,
} from 'azure-devops-ui/Utilities/TreeItemProvider';
import {
    ReleaseDefinition,
    ReleaseRestClient,
} from 'azure-devops-extension-api/Release';
import { BuildDefinition } from 'azure-devops-extension-api/Build';
import { Icon } from 'azure-devops-ui/Icon';

export interface ISprintlyEnvironmentVariableViewerState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    globalEnvironmentVariablesObservable: ObservableArray<ISearchResultEnvironmentVariableItem>;
    repositories: ArrayItemProvider<GitRepository>;
    repositoryListSelection: ListSelection;
    repositoryListSelectedItemObservable: ObservableValue<GitRepository>;
    repositoryEnvironmentVariablesItemProvider: ITreeItemProvider<ISearchResultRepositoryEnvironmentVariableItem>;
}

export interface ISearchResultEnvironmentVariableValue {
    environmentName: string;
    value: string;
}

export interface ISearchResultEnvironmentVariableItem {
    name: string;
    values: ISearchResultEnvironmentVariableValue[];
}

export interface ISearchResultRepositoryEnvironmentVariableItem {
    name: string;
    value: string;
    isRootItem: boolean;
}

//#region "Observables"
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const showAllEnvironmentVariablesObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(true);
//#endregion "Observables"

const environmentVariableNameFilterKey: string =
    'environmentVariableNameFilterKey';
const environmentVariableValueFilterKey: string =
    'environmentVariableValueFilterKey';
const localStorageShowAllVariablesKey: string =
    'show-all-environment-variables';

let repositoriesToProcess: string[] = [];

export default class SprintlyEnvironmentVariableViewer extends React.Component<
    {
        dataManager: IExtensionDataManager;
        organizationName: string;
        userDescriptor: string;
        releaseDefinitions: ReleaseDefinition[];
        buildDefinitions: BuildDefinition[];
    },
    ISprintlyEnvironmentVariableViewerState
> {
    private dataManager: IExtensionDataManager;
    private organizationName: string;
    private userDescriptor: string;
    private releaseDefinitions: ReleaseDefinition[];
    private buildDefinitions: BuildDefinition[];
    private accessToken: string = '';
    private currentProject: IProjectInfo | undefined;

    private environmentVariablesResponse: any;
    private environmentVariablesExclusionFilter: Set<string> = new Set();
    private environmentVariableNameSearchFilter: Filter;
    private environmentVariableValueSearchFilter: Filter;

    private columns: ITableColumn<ISearchResultEnvironmentVariableItem>[] = [];
    private repositoryTreeColumns: ITreeColumn<ISearchResultRepositoryEnvironmentVariableItem>[] =
        [];
    private sortingBehavior: ColumnSorting<ISearchResultEnvironmentVariableItem> =
        new ColumnSorting<ISearchResultEnvironmentVariableItem>(
            (
                columnIndex: number,
                proposedSortOrder: SortOrder,
                event:
                    | React.KeyboardEvent<HTMLElement>
                    | React.MouseEvent<HTMLElement>
            ) => {
                this.state.globalEnvironmentVariablesObservable.splice(
                    0,
                    this.state.globalEnvironmentVariablesObservable.length,
                    ...sortItems<ISearchResultEnvironmentVariableItem>(
                        columnIndex,
                        proposedSortOrder,
                        this.sortFunctions,
                        this.columns,
                        this.state.globalEnvironmentVariablesObservable.value
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
        userDescriptor: string;
        releaseDefinitions: ReleaseDefinition[];
        buildDefinitions: BuildDefinition[];
    }) {
        super(props);

        this.onSize = this.onSize.bind(this);
        this.renderRepositoryMasterPageList =
            this.renderRepositoryMasterPageList.bind(this);
        this.renderDetailPageContent = this.renderDetailPageContent.bind(this);
        this.renderRepositoryEnvironmentVariableAppSettingsElementCell =
            this.renderRepositoryEnvironmentVariableAppSettingsElementCell.bind(
                this
            );
        this.renderRepositoryEnvironmentVariableTransformValueCell =
            this.renderRepositoryEnvironmentVariableTransformValueCell.bind(
                this
            );

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

        this.repositoryTreeColumns = [
            {
                id: 'appsettingsElementName',
                name: 'Appsettings Field',
                onSize: this.onSizeTreeColumn,
                renderCell:
                    this
                        .renderRepositoryEnvironmentVariableAppSettingsElementCell,
                width: new ObservableValue<number>(-30),
            },
            {
                id: 'transformValue',
                name: 'Transform Value',
                onSize: this.onSizeTreeColumn,
                renderCell:
                    this.renderRepositoryEnvironmentVariableTransformValueCell,
                width: new ObservableValue<number>(-80),
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
            globalEnvironmentVariablesObservable:
                new ObservableArray<ISearchResultEnvironmentVariableItem>([]),
            repositories: new ArrayItemProvider<GitRepository>([]),
            repositoryListSelection: new ListSelection({
                selectOnFocus: false,
            }),
            repositoryListSelectedItemObservable: new ObservableValue<any>({}),
            repositoryEnvironmentVariablesItemProvider:
                new TreeItemProvider<ISearchResultRepositoryEnvironmentVariableItem>(
                    []
                ),
        };

        this.dataManager = props.dataManager;
        this.organizationName = props.organizationName;
        this.userDescriptor = props.userDescriptor;
        this.releaseDefinitions = props.releaseDefinitions;
        this.buildDefinitions = props.buildDefinitions;
    }

    public async componentDidMount(): Promise<void> {
        await this.initializeComponent();
    }

    private async initializeComponent(): Promise<void> {
        this.accessToken = await SDK.getAccessToken();
        this.currentProject = await Common.getCurrentProject();

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

        showAllEnvironmentVariablesObservable.value = JSON.parse(
            localStorage.getItem(
                `${this.userDescriptor}-${localStorageShowAllVariablesKey}`
            ) ?? true.toString()
        );

        if (showAllEnvironmentVariablesObservable.value) {
            await this.loadEnvironmentVariables();
        }

        repositoriesToProcess = Common.getSavedRepositoriesToView(
            this.state.userSettings,
            this.state.systemSettings
        );

        totalRepositoriesToProcessObservable.value =
            repositoriesToProcess.length;
        if (repositoriesToProcess.length > 0) {
            await this.loadRepositoriesDisplayState(this.currentProject);
        }
    }

    private async loadEnvironmentVariables(): Promise<void> {
        if (this.currentProject !== undefined) {
            let environmentVariableGroupIds: string = '';
            for (const groupId of Common.ALLOWED_ENVIRONMENT_VARIABLE_GROUP_IDS) {
                environmentVariableGroupIds += `${groupId.toString()},`;
            }
            const url: string = `https://dev.azure.com/${this.organizationName}/${this.currentProject.id}/_apis/distributedtask/variablegroups?groupIds=${environmentVariableGroupIds}`;
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
            globalEnvironmentVariablesObservable:
                new ObservableArray<ISearchResultEnvironmentVariableItem>(
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
        return (
            <Observer
                selectedItem={this.state.repositoryListSelectedItemObservable}
            >
                {(observerProps: { selectedItem: GitRepository }) => (
                    <Page className='flex-grow single-layer-details'>
                        {this.state.repositoryListSelection.selectedCount ===
                            0 && (
                            <span className='single-layer-details-contents'>
                                Select a repository on the right to see its
                                environment variable transforms.
                            </span>
                        )}
                        {this.state
                            .repositoryEnvironmentVariablesItemProvider &&
                            this.state.repositoryListSelection.selectedCount !==
                                0 && (
                                <Page>
                                    <div className='page-content'>
                                        <Card className='bolt-table-card bolt-card-white'>
                                            <Tree<ISearchResultRepositoryEnvironmentVariableItem>
                                                ariaLabel='Basic tree'
                                                columns={
                                                    this.repositoryTreeColumns
                                                }
                                                itemProvider={
                                                    this.state
                                                        .repositoryEnvironmentVariablesItemProvider
                                                }
                                                onToggle={(
                                                    event: React.SyntheticEvent<
                                                        HTMLElement,
                                                        Event
                                                    >,
                                                    treeItem: ITreeItemEx<ISearchResultRepositoryEnvironmentVariableItem>
                                                ) => {
                                                    this.state.repositoryEnvironmentVariablesItemProvider.toggle(
                                                        treeItem.underlyingItem
                                                    );
                                                }}
                                                selectableText={true}
                                                scrollable={true}
                                            />
                                        </Card>
                                    </div>
                                </Page>
                            )}
                    </Page>
                )}
            </Observer>
        );
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

    private renderRepositoryEnvironmentVariableAppSettingsElementCell(
        rowIndex: number,
        columnIndex: number,
        treeColumn: ITreeColumn<ISearchResultRepositoryEnvironmentVariableItem>,
        treeItem: ITreeItemEx<ISearchResultRepositoryEnvironmentVariableItem>
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={`col-${columnIndex}-row-${rowIndex}`}
                columnIndex={columnIndex}
                tableColumn={treeColumn}
                children={
                    <>
                        {treeItem.depth === 0 ? (
                            <>
                                <Icon
                                    iconName={
                                        treeItem.underlyingItem.expanded
                                            ? 'ChevronDownMed'
                                            : 'ChevronRightMed'
                                    }
                                    className='bolt-tree-expand-button cursor-pointer'
                                ></Icon>
                                <Icon
                                    iconName='Settings'
                                    className='icon-margin'
                                    style={{ color: '#0081E3' }}
                                ></Icon>
                                <span className='icon-margin'>
                                    {treeItem.underlyingItem.data.name}
                                </span>
                            </>
                        ) : (
                            <>
                                <Icon
                                    iconName='ChevronRightMed'
                                    className='invisible'
                                    style={{
                                        marginLeft: `${treeItem.depth * 16}px`,
                                    }}
                                ></Icon>
                                {treeItem.underlyingItem.data.name}
                            </>
                        )}
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private renderRepositoryEnvironmentVariableTransformValueCell(
        rowIndex: number,
        columnIndex: number,
        treeColumn: ITreeColumn<ISearchResultRepositoryEnvironmentVariableItem>,
        treeItem: ITreeItemEx<ISearchResultRepositoryEnvironmentVariableItem>
    ): JSX.Element {
        const regex = /(\$\([^\)]+\))/g;
        const variableHighlightSplit =
            treeItem.underlyingItem.data.value.split(regex);
        return (
            <SimpleTableCell
                key={`col-${columnIndex}-row-${rowIndex}`}
                columnIndex={columnIndex}
                tableColumn={treeColumn}
                children={
                    <>
                        {treeItem.depth === 0 ? (
                            <>
                                {variableHighlightSplit.map((substring) => {
                                    if (substring.startsWith('$(')) {
                                        return (
                                            <>
                                                <b>{substring}</b>
                                            </>
                                        );
                                    } else {
                                        return <>{substring}</>;
                                    }
                                })}
                            </>
                        ) : (
                            <>{treeItem.underlyingItem.data.value}</>
                        )}
                    </>
                }
            ></SimpleTableCell>
        );
    }

    private async selectRepository(): Promise<void> {
        //TODO: get release defs to be passed in from page 1
        //get single release for repo
        //get the inline transform variables and parse them out
        //build a tree table nesting each environment under each variable (root)
        //show appsettings transform at root level, second column, and environment value for nested children.
        console.log(this.releaseDefinitions);
        console.log(this.state.repositoryListSelectedItemObservable.value);
        const releaseDefinitionIdForRepo =
            Common.getRepositoryReleaseDefinitionId(
                this.buildDefinitions,
                this.releaseDefinitions,
                this.state.repositoryListSelectedItemObservable.value.id
            );
        const repositoryEnvironmentVariablesRootItems: Array<
            ITreeItem<ISearchResultRepositoryEnvironmentVariableItem>
        > = [];

        if (
            releaseDefinitionIdForRepo > -1 &&
            this.currentProject !== undefined
        ) {
            const environmentVariableRegex = /(\$\([^\)]+\))/g;

            console.log(releaseDefinitionIdForRepo);

            const releaseDefinition: ReleaseDefinition = await getClient(
                ReleaseRestClient
            ).getReleaseDefinition(
                this.currentProject.id,
                releaseDefinitionIdForRepo
            );
            console.log(releaseDefinition.variables);
            if (releaseDefinition.variables['inlineTransforms'] !== undefined) {
                if (this.environmentVariablesResponse === undefined) {
                    await this.loadEnvironmentVariables();
                }
                const inlineTransforms: any = JSON.parse(
                    releaseDefinition.variables['inlineTransforms'].value
                );
                for (const [appsetting, transform] of Object.entries(
                    inlineTransforms
                )) {
                    const transformValue: string = (
                        transform as any
                    ).toString();
                    const repositoryEnvironmentVariablesTableItem: ITreeItem<ISearchResultRepositoryEnvironmentVariableItem> =
                        {
                            childItems: [],
                            data: {
                                name: appsetting,
                                value: transformValue,
                                isRootItem: true,
                            },
                            expanded: true,
                        };

                    const foundEnvironmentVariables: RegExpMatchArray | null =
                        transformValue.match(environmentVariableRegex);
                    if (
                        foundEnvironmentVariables !== undefined &&
                        foundEnvironmentVariables !== null
                    ) {
                        for (const foundEnvironmentVariable of foundEnvironmentVariables) {
                            const cleanEnvironmentVariable: string =
                                foundEnvironmentVariable.substring(
                                    2,
                                    foundEnvironmentVariable.length - 1
                                );

                            for (const environmentVariableGroup of this
                                .environmentVariablesResponse.value) {
                                for (const [
                                    environmentVariableName,
                                    environmentVariableValue,
                                ] of Object.entries(
                                    environmentVariableGroup.variables
                                )) {
                                    if (
                                        environmentVariableName ===
                                        cleanEnvironmentVariable
                                    ) {
                                        repositoryEnvironmentVariablesTableItem.childItems!.push(
                                            {
                                                data: {
                                                    isRootItem: false,
                                                    name: environmentVariableGroup.name,
                                                    value: (
                                                        environmentVariableValue as any
                                                    ).value,
                                                },
                                            }
                                        );
                                    }
                                }
                            }
                        }
                    }

                    repositoryEnvironmentVariablesRootItems.push(
                        repositoryEnvironmentVariablesTableItem
                    );
                }
            }
        }
        this.setState({
            repositoryEnvironmentVariablesItemProvider: new TreeItemProvider(
                repositoryEnvironmentVariablesRootItems
            ),
        });
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (this.columns[index].width as ObservableValue<number>).value = width;
    }

    private onSizeTreeColumn(
        event: MouseEvent,
        index: number,
        width: number
    ): void {
        (
            this.repositoryTreeColumns[index].width as ObservableValue<number>
        ).value = width;
    }

    public render(): JSX.Element {
        return (
            <Observer
                environmentVariables={
                    this.state.globalEnvironmentVariablesObservable
                }
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
                                    localStorage.setItem(
                                        `${this.userDescriptor}-${localStorageShowAllVariablesKey}`,
                                        showAllEnvironmentVariablesObservable.value.toString()
                                    );
                                    if (
                                        this.environmentVariablesResponse ===
                                        undefined
                                    ) {
                                        this.loadEnvironmentVariables();
                                    }
                                }}
                            />
                            <Button
                                text='Repository specific variables'
                                primary={true}
                                onClick={() => {
                                    showAllEnvironmentVariablesObservable.value =
                                        false;
                                    localStorage.setItem(
                                        `${this.userDescriptor}-${localStorageShowAllVariablesKey}`,
                                        showAllEnvironmentVariablesObservable.value.toString()
                                    );
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
                                                .globalEnvironmentVariablesObservable
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
