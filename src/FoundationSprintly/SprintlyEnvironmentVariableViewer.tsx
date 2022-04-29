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
import {
    GitItem,
    GitRef,
    GitRepository,
    GitRestClient,
    GitVersionDescriptor,
    GitVersionOptions,
    GitVersionType,
} from 'azure-devops-extension-api/Git';
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
import { Icon } from 'azure-devops-ui/Icon';
import { Dropdown } from 'azure-devops-ui/Dropdown';
import { DropdownSelection } from 'azure-devops-ui/Utilities/DropdownSelection';
import { IListBoxItem } from 'azure-devops-ui/ListBox';
import {
    Artifact,
    ConfigurationVariableValue,
    Release,
    ReleaseDefinition,
    ReleaseRestClient,
    VariableGroup,
} from 'azure-devops-extension-api/Release';
import { BuildDefinition } from 'azure-devops-extension-api/Build';
import { ZeroData } from 'azure-devops-ui/ZeroData';

export interface ISprintlyEnvironmentVariableViewerState {
    userSettings?: Common.IUserSettings;
    systemSettings?: Common.ISystemSettings;
    globalEnvironmentVariablesObservable: ObservableArray<ISearchResultEnvironmentVariableItem>;
    repositories: ArrayItemProvider<GitRepository>;
    repositoryListSelection: ListSelection;
    repositoryListSelectedItemObservable: ObservableValue<GitRepository>;
    repositoryEnvironmentVariablesFromAppSettingsItemProvider: ITreeItemProvider<ISearchResultRepositoryEnvironmentVariableItem>;
    repositoryEnvironmentVariablesFromConfigSettingsItemProvider: ITreeItemProvider<ISearchResultRepositoryEnvironmentVariableItem>;
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
    transformValueFromCode: string;
    transformValueFromPipeline: string;
    isRootItem: boolean;
    hasDiscrepancy: boolean;
}

export enum TransformsTypeEnum {
    AppSettings = 0,
    ConfigSettings = 1,
}

//#region "Observables"
const totalRepositoriesToProcessObservable: ObservableValue<number> =
    new ObservableValue<number>(0);
const showAllEnvironmentVariablesObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(true);
const loadingRepositoryObservable: ObservableValue<boolean> =
    new ObservableValue<boolean>(false);
//#endregion "Observables"

const environmentVariableNameFilterKey: string =
    'environmentVariableNameFilterKey';
const environmentVariableValueFilterKey: string =
    'environmentVariableValueFilterKey';
const localStorageShowAllVariablesKey: string =
    'show-all-environment-variables';
const appSettingsTransformsFileName: string = 'transforms.json';
const configSettingsTransformsFileName: string = 'ENVTransforms.json';
const appSettingsTransformsPipelineVariableName: string = 'inlineTransforms';
const configSettingsTransformsPipelineVariableName: string = 'appSettings';
const secretValueText: string = '*********';

let repositoriesToProcess: string[] = [];

export default class SprintlyEnvironmentVariableViewer extends React.Component<
    {
        organizationName: string;
        userDescriptor: string;
        releaseDefinitions: ReleaseDefinition[];
        buildDefinitions: BuildDefinition[];
    },
    ISprintlyEnvironmentVariableViewerState
> {
    private dataManager!: IExtensionDataManager;
    private organizationName: string;
    private userDescriptor: string;
    private releaseDefinitions: ReleaseDefinition[];
    private buildDefinitions: BuildDefinition[];
    private accessToken: string = '';
    private currentProject: IProjectInfo | undefined;
    private repositoryBranchSelection: DropdownSelection =
        new DropdownSelection();
    private selectedRepositoryBranchesInfo:
        | Common.IRepositoryBranchInfo
        | undefined;

    private globalEnvironmentVariables: VariableGroup[] = [];
    private globalEnvironmentVariablesExclusionFilter: Set<string> = new Set();
    private globalEnvironmentVariableNameSearchFilter: Filter;
    private globalEnvironmentVariableValueSearchFilter: Filter;

    private repositoryBranchBuildArtifactEnvironmentVariables: VariableGroup[] =
        [];
    private repositoryBranchBuildArtifactBaseCommit: string = '';

    private columns: Array<ITableColumn<ISearchResultEnvironmentVariableItem>> =
        [];
    private repositoryTreeColumns: Array<
        ITreeColumn<ISearchResultRepositoryEnvironmentVariableItem>
    > = [];
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
        organizationName: string;
        userDescriptor: string;
        releaseDefinitions: ReleaseDefinition[];
        buildDefinitions: BuildDefinition[];
    }) {
        super(props);

        this.onSize = this.onSize.bind(this);
        this.onSizeTreeColumn = this.onSizeTreeColumn.bind(this);
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
        this.loadInlineTransforms = this.loadInlineTransforms.bind(this);
        this.getTransformsFileFromCode =
            this.getTransformsFileFromCode.bind(this);
        this.getJsonTransformsFromPipeline =
            this.getJsonTransformsFromPipeline.bind(this);
        this.onSelectBranchAction = this.onSelectBranchAction.bind(this);
        this.loadBaseCommitAndEnvironmentVariablesFromSelectedBuildArtifact =
            this.loadBaseCommitAndEnvironmentVariablesFromSelectedBuildArtifact.bind(
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
                id: 'transformValueFromCode',
                name: 'Transform from code base',
                onSize: this.onSizeTreeColumn,
                renderCell:
                    this.renderRepositoryEnvironmentVariableTransformValueCell,
                width: new ObservableValue<number>(-30),
            },
            {
                id: 'transformValueFromPipeline',
                name: 'Transform Value From Pipeline',
                onSize: this.onSizeTreeColumn,
                renderCell:
                    this.renderRepositoryEnvironmentVariableTransformValueCell,
                width: new ObservableValue<number>(-40),
            },
        ];

        this.globalEnvironmentVariableNameSearchFilter = new Filter();
        this.globalEnvironmentVariableValueSearchFilter = new Filter();
        this.globalEnvironmentVariableNameSearchFilter.subscribe(() => {
            this.redrawEnvironmentVariablesSearchResult();
        }, FILTER_CHANGE_EVENT);
        this.globalEnvironmentVariableValueSearchFilter.subscribe(() => {
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
            repositoryEnvironmentVariablesFromAppSettingsItemProvider:
                new TreeItemProvider<ISearchResultRepositoryEnvironmentVariableItem>(
                    []
                ),
            repositoryEnvironmentVariablesFromConfigSettingsItemProvider:
                new TreeItemProvider<ISearchResultRepositoryEnvironmentVariableItem>(
                    []
                ),
        };

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
        this.dataManager = await Common.initializeDataManager(this.accessToken);
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
            await this.loadAllEnvironmentVariables();
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

    private async loadAllEnvironmentVariables(): Promise<void> {
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
            this.globalEnvironmentVariables = (response.data as any).value; // No defined type exists in the api

            this.redrawEnvironmentVariablesSearchResult();
        }
    }

    //#region Global Environment Variables

    private redrawEnvironmentVariablesSearchResult(): void {
        const resultEnvironmentVariables: ISearchResultEnvironmentVariableItem[] =
            [];
        let environmentVariableNameSearchFilterString: string = '';
        let environmentVariableValueSearchFilterString: string = '';

        const environmentVariableNameSearchFilterState: IFilterState =
            this.globalEnvironmentVariableNameSearchFilter.getState();
        const environmentVariableValueSearchFilterState: IFilterState =
            this.globalEnvironmentVariableValueSearchFilter.getState();

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
        for (const environmentVariableGroup of this
            .globalEnvironmentVariables) {
            if (
                !this.globalEnvironmentVariablesExclusionFilter.has(
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
                            const environmentVariableValueString:
                                | string
                                | undefined = environmentVariableValue.isSecret
                                ? secretValueText
                                : environmentVariableValue.value;
                            if (
                                environmentVariableValue === undefined ||
                                environmentVariableValueString === undefined ||
                                environmentVariableValueSearchFilterString.length ===
                                    0 ||
                                    environmentVariableValueString
                                    .toLowerCase()
                                    .includes(
                                        environmentVariableValueSearchFilterString.toLowerCase()
                                    )
                            ) {
                                environmentVariable.values.push({
                                    environmentName:
                                        environmentVariableGroup.name,
                                    value: environmentVariableValueString,
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
                            const environmentVariableValueString:
                                | string
                                | undefined = environmentVariableValue.isSecret
                                ? secretValueText
                                : environmentVariableValue.value;
                            if (
                                environmentVariableValue === undefined ||
                                environmentVariableValueString === undefined ||
                                environmentVariableValueSearchFilterString.length ===
                                    0 ||
                                    environmentVariableValueString
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
                                            value: environmentVariableValueString,
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
                    this.sortEnvironmentVariableSearchResult(
                        resultEnvironmentVariables
                    )
                ),
        });
    }

    private sortEnvironmentVariableSearchResult(
        environmentVariableList: ISearchResultEnvironmentVariableItem[]
    ): ISearchResultEnvironmentVariableItem[] {
        if (environmentVariableList.length > 0) {
            return environmentVariableList.sort(
                (
                    a: ISearchResultEnvironmentVariableItem,
                    b: ISearchResultEnvironmentVariableItem
                ) => {
                    return a.name.localeCompare(b.name);
                }
            );
        }
        return environmentVariableList;
    }

    private updateEnvironmentVariablesExcludeFilter(
        environmentName: string,
        show: boolean
    ): void {
        if (show) {
            this.globalEnvironmentVariablesExclusionFilter.delete(
                environmentName
            );
        } else {
            this.globalEnvironmentVariablesExclusionFilter.add(environmentName);
        }
        this.redrawEnvironmentVariablesSearchResult();
    }

    private renderEnvironmentVariablesExcludeFilterCheckboxes(): JSX.Element {
        if (this.globalEnvironmentVariables !== undefined) {
            return (
                <>
                    {this.globalEnvironmentVariables.map(
                        (environment: VariableGroup) => (
                            <Checkbox
                                key={environment.name}
                                onChange={(
                                    event:
                                        | React.MouseEvent<
                                              HTMLElement,
                                              MouseEvent
                                          >
                                        | React.KeyboardEvent<HTMLElement>,
                                    checked: boolean
                                ) =>
                                    this.updateEnvironmentVariablesExcludeFilter(
                                        environment.name,
                                        checked
                                    )
                                }
                                checked={
                                    !this.globalEnvironmentVariablesExclusionFilter.has(
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
        let itemValue: string = '';
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

    private renderGlobalEnvironmentVariablesPage(): JSX.Element {
        return (
            <>
                <div className='rhythm-horizontal-8 flex-row'>
                    {this.renderEnvironmentVariablesExcludeFilterCheckboxes()}
                </div>
                <div className='rhythm-horizontal-8 flex-row'>
                    <div className='flex-grow'>
                        <FilterBar
                            filter={
                                this.globalEnvironmentVariableNameSearchFilter
                            }
                        >
                            <KeywordFilterBarItem
                                placeholder='Filter by variable name'
                                filterItemKey={environmentVariableNameFilterKey}
                            />
                        </FilterBar>
                    </div>
                    <div className='flex-grow sprintly-margin-right-auto'>
                        <FilterBar
                            filter={
                                this.globalEnvironmentVariableValueSearchFilter
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
                            this.state.globalEnvironmentVariablesObservable
                        }
                    />
                </Card>
            </>
        );
    }

    private onSize(event: MouseEvent, index: number, width: number): void {
        (this.columns[index].width as ObservableValue<number>).value = width;
    }

    //#endregion

    //#region Repository Environment Variable Transforms
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
                loadingRepository={loadingRepositoryObservable}
            >
                {(observerProps: {
                    selectedItem: GitRepository;
                    loadingRepository: boolean;
                }) => (
                    <Page className='flex-grow single-layer-details'>
                        {loadingRepositoryObservable.value && (
                            <div className='page-content-top'>
                                <Spinner label='loading' />
                            </div>
                        )}
                        {this.state.repositoryListSelection.selectedCount ===
                            0 && (
                            <Page>
                                <div className='page-content'>
                                    Select a repository on the right to see its
                                    environment variable transforms.
                                </div>
                            </Page>
                        )}
                        {this.renderTreeTablePage(
                            TransformsTypeEnum.AppSettings
                        )}
                        {this.renderTreeTablePage(
                            TransformsTypeEnum.ConfigSettings
                        )}
                    </Page>
                )}
            </Observer>
        );
    }

    private renderTreeTablePage(
        transformsType: TransformsTypeEnum
    ): JSX.Element {
        const codeTransformsFileName: string =
            transformsType === TransformsTypeEnum.AppSettings
                ? appSettingsTransformsFileName
                : configSettingsTransformsFileName;
        const pipelineTransformsVariableName: string =
            transformsType === TransformsTypeEnum.AppSettings
                ? appSettingsTransformsPipelineVariableName
                : configSettingsTransformsPipelineVariableName;
        const treeItemProvider: ITreeItemProvider<ISearchResultRepositoryEnvironmentVariableItem> =
            transformsType === TransformsTypeEnum.AppSettings
                ? this.state
                      .repositoryEnvironmentVariablesFromAppSettingsItemProvider
                : this.state
                      .repositoryEnvironmentVariablesFromConfigSettingsItemProvider;

        const title: string =
            transformsType === TransformsTypeEnum.AppSettings
                ? 'Transforms for appsettings.json'
                : 'Transforms for resource-level application configuration settings';

        return (
            <>
                {!loadingRepositoryObservable.value &&
                    this.state.repositoryListSelection.selectedCount !== 0 &&
                    treeItemProvider.length === 0 && (
                        <Page>
                            <div className='page-content'>
                                Please ensure there is a{' '}
                                <code>{codeTransformsFileName}</code> file on
                                this branch and that a build artifact has been
                                created.
                            </div>
                        </Page>
                    )}
                {!loadingRepositoryObservable.value &&
                    this.state.repositoryListSelection.selectedCount !== 0 && (
                        <Page>
                            <div className='page-content'>
                                <Card
                                    titleProps={{
                                        text: title,
                                    }}
                                    headerDescriptionProps={{
                                        text: (
                                            <div>
                                                These transforms are sourced
                                                from the{' '}
                                                <code>
                                                    /{codeTransformsFileName}
                                                </code>{' '}
                                                file on the selected branch
                                                compared with the{' '}
                                                <code>
                                                    $(
                                                    {
                                                        pipelineTransformsVariableName
                                                    }
                                                    )
                                                </code>{' '}
                                                variable on the release
                                                pipeline.
                                            </div>
                                        ),
                                    }}
                                    className='bolt-table-card bolt-card-white'
                                >
                                    <div className='master-row-content'>
                                        <div className='flex-row'>
                                            <div
                                                style={{
                                                    color: '#FF3E3E',
                                                    fontWeight: 'bold',
                                                }}
                                            >
                                                $(VariableName)
                                            </div>
                                            <div>
                                                &nbsp; = This means this
                                                variable does not exist in this
                                                environment
                                            </div>
                                        </div>
                                        <div className='flex-row'>
                                            <div
                                                style={{
                                                    border: 'red 1px solid',
                                                }}
                                            >
                                                TransformValue
                                            </div>
                                            <div>
                                                &nbsp; = This shows a
                                                discrepancy between{' '}
                                                <code>
                                                    {codeTransformsFileName}
                                                </code>{' '}
                                                and the{' '}
                                                <code>
                                                    $(
                                                    {
                                                        pipelineTransformsVariableName
                                                    }
                                                    )
                                                </code>{' '}
                                                variable from the release
                                                pipeline.
                                            </div>
                                        </div>
                                        {transformsType ===
                                            TransformsTypeEnum.AppSettings && (
                                            <>
                                                <div className='page-content-top'>
                                                    Select a branch to use its
                                                    latest build artifact.
                                                </div>
                                                <Dropdown
                                                    className='page-content-top'
                                                    ariaLabel='Button Dropdown'
                                                    placeholder='Select a branch'
                                                    selection={
                                                        this
                                                            .repositoryBranchSelection
                                                    }
                                                    items={this.selectedRepositoryBranchesInfo!.allBranchesAndTags.map(
                                                        (branchInfo: GitRef) =>
                                                            Common.getBranchShortName(
                                                                branchInfo.name
                                                            )
                                                    )}
                                                    onSelect={
                                                        this
                                                            .onSelectBranchAction
                                                    }
                                                />
                                            </>
                                        )}

                                        <Tree<ISearchResultRepositoryEnvironmentVariableItem>
                                            columns={this.repositoryTreeColumns}
                                            itemProvider={treeItemProvider}
                                            onToggle={(
                                                event: React.SyntheticEvent<
                                                    HTMLElement,
                                                    Event
                                                >,
                                                treeItem: ITreeItemEx<ISearchResultRepositoryEnvironmentVariableItem>
                                            ) => {
                                                treeItemProvider.toggle(
                                                    treeItem.underlyingItem
                                                );
                                            }}
                                            selectableText={true}
                                            scrollable={true}
                                        />
                                    </div>
                                </Card>
                            </div>
                        </Page>
                    )}
            </>
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

    private renderRepositoryEnvironmentVariableAppSettingsElementCell(
        rowIndex: number,
        columnIndex: number,
        treeColumn: ITreeColumn<ISearchResultRepositoryEnvironmentVariableItem>,
        treeItem: ITreeItemEx<ISearchResultRepositoryEnvironmentVariableItem>
    ): JSX.Element {
        return (
            <SimpleTableCell
                key={`appsettings-col-${columnIndex}-row-${rowIndex}`}
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
                                <Tooltip overflowOnly={true}>
                                    <span className='icon-margin text-ellipsis'>
                                        {treeItem.underlyingItem.data.name}
                                    </span>
                                </Tooltip>
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
                                <Tooltip overflowOnly={true}>
                                    <div className='text-ellipsis'>
                                        {treeItem.underlyingItem.data.name}
                                    </div>
                                </Tooltip>
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
        const regex: RegExp = /(\$\([^\)]+\))/g;
        const transformValueFromCodeColumnIndex: number = 1;
        const transformValueFromPipelineColumnIndex: number = 2;

        const variableHighlightSplit: string[] =
            columnIndex === transformValueFromCodeColumnIndex
                ? treeItem.underlyingItem.data.transformValueFromCode.split(
                      regex
                  )
                : columnIndex === transformValueFromPipelineColumnIndex
                ? treeItem.underlyingItem.data.transformValueFromPipeline.split(
                      regex
                  )
                : [];
        return (
            <SimpleTableCell
                key={`transform-col-${columnIndex}-row-${rowIndex}`}
                columnIndex={columnIndex}
                tableColumn={treeColumn}
                children={
                    <Tooltip overflowOnly={true}>
                        <div
                            className='text-ellipsis flex-row'
                            style={{
                                border: treeItem.underlyingItem.data
                                    .hasDiscrepancy
                                    ? 'red 1px solid'
                                    : '',
                            }}
                        >
                            {treeItem.depth === 0 ? (
                                <>
                                    {variableHighlightSplit.map(
                                        (
                                            valueSubstring: string,
                                            index: number
                                        ) => {
                                            if (
                                                valueSubstring.startsWith('$(')
                                            ) {
                                                return (
                                                    <div
                                                        key={`transform-val-${columnIndex}-row-${rowIndex}-substr-${index}`}
                                                    >
                                                        <b>{valueSubstring}</b>
                                                    </div>
                                                );
                                            } else {
                                                return (
                                                    <div
                                                        key={`transform-val-${columnIndex}-row-${rowIndex}-substr-${index}`}
                                                    >
                                                        {valueSubstring}
                                                    </div>
                                                );
                                            }
                                        }
                                    )}
                                </>
                            ) : (
                                <>
                                    {variableHighlightSplit.map(
                                        (
                                            valueSubstring: string,
                                            index: number
                                        ) => {
                                            if (
                                                valueSubstring.startsWith('$(')
                                            ) {
                                                return (
                                                    <div
                                                        key={`transform-envval-${columnIndex}-row-${rowIndex}-substr-${index}`}
                                                    >
                                                        <b
                                                            style={{
                                                                color: '#FF3E3E',
                                                            }}
                                                        >
                                                            {valueSubstring}
                                                        </b>
                                                    </div>
                                                );
                                            } else {
                                                return (
                                                    <div
                                                        key={`transform-envval-${columnIndex}-row-${rowIndex}-substr-${index}`}
                                                    >
                                                        {valueSubstring}
                                                    </div>
                                                );
                                            }
                                        }
                                    )}
                                </>
                            )}
                        </div>
                    </Tooltip>
                }
            ></SimpleTableCell>
        );
    }

    private async selectRepository(): Promise<void> {
        loadingRepositoryObservable.value = true;
        this.selectedRepositoryBranchesInfo =
            await Common.getRepositoryBranchesInfo(
                this.state.repositoryListSelectedItemObservable.value.id,
                Common.repositoryHeadsFilter
            );

        this.repositoryBranchSelection.select(
            this.selectedRepositoryBranchesInfo.allBranchesAndTags.findIndex(
                (branch: GitRef) =>
                    Common.getBranchShortName(branch.name) === Common.DEVELOP
            )
        );

        await this.loadBaseCommitAndEnvironmentVariablesFromSelectedBuildArtifact(
            Common.DEVELOP
        );
        await this.loadInlineTransforms(TransformsTypeEnum.AppSettings);
        await this.loadInlineTransforms(TransformsTypeEnum.ConfigSettings);

        loadingRepositoryObservable.value = false;
    }

    private async loadInlineTransforms(
        transformsType: TransformsTypeEnum
    ): Promise<void> {
        const codeTransformsFileName: string =
            transformsType === TransformsTypeEnum.AppSettings
                ? appSettingsTransformsFileName
                : configSettingsTransformsFileName;
        const pipelineTransformsVariableName: string =
            transformsType === TransformsTypeEnum.AppSettings
                ? appSettingsTransformsPipelineVariableName
                : configSettingsTransformsPipelineVariableName;

        const repositoryEnvironmentVariablesRootItems: Array<
            ITreeItem<ISearchResultRepositoryEnvironmentVariableItem>
        > = [];

        if (this.currentProject !== undefined) {
            const inlineTransformsFromCodeParsed: any | undefined =
                await this.getTransformsFileFromCode(codeTransformsFileName);
            const inlineTransformsFromPipelineParsed: any | undefined =
                await this.getJsonTransformsFromPipeline(
                    pipelineTransformsVariableName
                );

            try {
                if (inlineTransformsFromCodeParsed !== undefined) {
                    for (const [appsetting, transform] of Object.entries(
                        inlineTransformsFromCodeParsed
                    )) {
                        const transformFromCodeValue: string = (
                            transform as any
                        ).toString();
                        const transformFromPipelineValue: string =
                            inlineTransformsFromPipelineParsed[appsetting] ??
                            '';

                        const repositoryEnvironmentVariablesTableItem: ITreeItem<ISearchResultRepositoryEnvironmentVariableItem> =
                            this.buildSingleTransformTreeItem(
                                appsetting,
                                transformFromCodeValue,
                                transformFromPipelineValue
                            );

                        repositoryEnvironmentVariablesRootItems.push(
                            repositoryEnvironmentVariablesTableItem
                        );
                    }

                    for (const [appsetting, transform] of Object.entries(
                        inlineTransformsFromPipelineParsed
                    )) {
                        const transformFromPipelineValue: string = (
                            transform as any
                        ).toString();

                        const transformFromCodeValue: string =
                            inlineTransformsFromCodeParsed[appsetting] ?? '';

                        if (transformFromCodeValue === '') {
                            const repositoryEnvironmentVariablesTableItem: ITreeItem<ISearchResultRepositoryEnvironmentVariableItem> =
                                this.buildSingleTransformTreeItem(
                                    appsetting,
                                    transformFromCodeValue,
                                    transformFromPipelineValue
                                );

                            repositoryEnvironmentVariablesRootItems.push(
                                repositoryEnvironmentVariablesTableItem
                            );
                        }
                    }
                }
            } catch (err) {
                console.error('Error retrieving transforms: ', err);
            }
        }
        if (transformsType === TransformsTypeEnum.AppSettings) {
            this.setState({
                repositoryEnvironmentVariablesFromAppSettingsItemProvider:
                    new TreeItemProvider(
                        repositoryEnvironmentVariablesRootItems
                    ),
            });
        } else {
            this.setState({
                repositoryEnvironmentVariablesFromConfigSettingsItemProvider:
                    new TreeItemProvider(
                        repositoryEnvironmentVariablesRootItems
                    ),
            });
        }
    }

    private buildSingleTransformTreeItem(
        appsetting: string,
        transformFromCodeValue: string,
        transformFromPipelineValue: string
    ): ITreeItem<ISearchResultRepositoryEnvironmentVariableItem> {
        const environmentVariableRegex: RegExp = /(\$\([^\)]+\))/g;
        const repositoryEnvironmentVariablesTableItem: ITreeItem<ISearchResultRepositoryEnvironmentVariableItem> =
            {
                childItems: [],
                data: {
                    name: appsetting,
                    transformValueFromCode: transformFromCodeValue,
                    transformValueFromPipeline: transformFromPipelineValue,
                    isRootItem: true,
                    hasDiscrepancy:
                        transformFromCodeValue !== transformFromPipelineValue,
                },
                expanded: true,
            };

        for (const environmentVariableGroup of this
            .repositoryBranchBuildArtifactEnvironmentVariables) {
            let environmentTransformFromCodeValue: string =
                transformFromCodeValue;
            let environmentTransformFromPipelineValue: string =
                transformFromPipelineValue;

            environmentTransformFromCodeValue =
                this.findReplaceEnvironmentVariables(
                    environmentVariableGroup,
                    environmentTransformFromCodeValue,
                    environmentVariableRegex
                );
            environmentTransformFromPipelineValue =
                this.findReplaceEnvironmentVariables(
                    environmentVariableGroup,
                    environmentTransformFromPipelineValue,
                    environmentVariableRegex
                );

            repositoryEnvironmentVariablesTableItem.childItems!.push({
                data: {
                    isRootItem: false,
                    name: environmentVariableGroup.name,
                    transformValueFromCode: environmentTransformFromCodeValue,
                    transformValueFromPipeline:
                        environmentTransformFromPipelineValue,
                    hasDiscrepancy:
                        environmentTransformFromCodeValue !==
                        environmentTransformFromPipelineValue,
                },
            });
        }

        return repositoryEnvironmentVariablesTableItem;
    }

    private findReplaceEnvironmentVariables(
        environment: VariableGroup,
        environmentTransformValue: string,
        environmentVariableRegex: RegExp
    ): string {
        let returnEnvironmentTransformedValue: string =
            environmentTransformValue;
        const environmentVariablesInTransformValue: RegExpMatchArray | null =
            environmentTransformValue.match(environmentVariableRegex);
        if (
            environmentVariablesInTransformValue !== undefined &&
            environmentVariablesInTransformValue !== null
        ) {
            for (const foundEnvironmentVariable of environmentVariablesInTransformValue) {
                const cleanEnvironmentVariable: string =
                    foundEnvironmentVariable.substring(
                        2,
                        foundEnvironmentVariable.length - 1
                    );
                for (const [
                    environmentVariableName,
                    environmentVariableValue,
                ] of Object.entries(environment.variables)) {
                    if (environmentVariableName === cleanEnvironmentVariable) {
                        const customRegex: RegExp = new RegExp(
                            `\\$\\(${cleanEnvironmentVariable}\\)`,
                            'g'
                        );
                        returnEnvironmentTransformedValue =
                            returnEnvironmentTransformedValue.replace(
                                customRegex,
                                environmentVariableValue.isSecret
                                    ? secretValueText
                                    : environmentVariableValue.value
                            );

                        break;
                    }
                }
            }
        }
        return returnEnvironmentTransformedValue;
    }

    private async getTransformsFileFromCode(
        fileName: string
    ): Promise<any | undefined> {
        if (this.currentProject !== undefined) {
            if (
                this.repositoryBranchBuildArtifactBaseCommit !== undefined &&
                this.repositoryBranchBuildArtifactBaseCommit !== ''
            ) {
                const versionDescriptor: GitVersionDescriptor = {
                    version: this.repositoryBranchBuildArtifactBaseCommit,
                    versionOptions: GitVersionOptions.None,
                    versionType: GitVersionType.Commit,
                };
                try {
                    const inlineTransformResponse: GitItem = await getClient(
                        GitRestClient
                    ).getItem(
                        this.state.repositoryListSelectedItemObservable.value
                            .id,
                        `/${fileName}`,
                        this.currentProject.id,
                        undefined,
                        undefined,
                        undefined,
                        undefined,
                        undefined,
                        versionDescriptor,
                        true,
                        undefined
                    );
                    if (inlineTransformResponse !== undefined) {
                        try {
                            return JSON.parse(
                                inlineTransformResponse.content.trim()
                            );
                        } catch (err) {
                            console.error('Error parsing JSON file: ', err);
                        }
                    }
                } catch (err) {
                    console.error('Error retrieving transforms: ', err);
                }
            }
        }

        return undefined;
    }

    private async getJsonTransformsFromPipeline(
        variableName: string
    ): Promise<any | undefined> {
        if (this.currentProject !== undefined) {
            const releaseDefinitionIdForRepo: number =
                Common.getRepositoryReleaseDefinitionId(
                    this.buildDefinitions,
                    this.releaseDefinitions,
                    this.state.repositoryListSelectedItemObservable.value.id
                );

            if (releaseDefinitionIdForRepo > -1) {
                const releaseDefinition: ReleaseDefinition = await getClient(
                    ReleaseRestClient
                ).getReleaseDefinition(
                    this.currentProject.id,
                    releaseDefinitionIdForRepo
                );

                let pipelineVariableValues: ConfigurationVariableValue =
                    releaseDefinition.variables[variableName];
                if (pipelineVariableValues === undefined) {
                    const tryLowerCaseVariableName: string =
                        variableName.toLowerCase();
                    pipelineVariableValues =
                        releaseDefinition.variables[tryLowerCaseVariableName];
                }

                if (pipelineVariableValues !== undefined) {
                    const variableText: string =
                        pipelineVariableValues.value.trim();
                    if (variableText[0] === '-') {
                        const variableTextClean: string =
                            variableText.substring(1, variableText.length);
                        const variableTextSplit: string[] =
                            variableTextClean.split(/\s+-/);
                        const returnObject: any = {};
                        for (const transformVariable of variableTextSplit) {
                            const transformVariableSplit: string[] =
                                transformVariable.split(/(?<=^\S+)\s/);
                            if (transformVariableSplit[1] === undefined) {
                                returnObject[transformVariableSplit[0]] = '';
                            } else {
                                let cleanValue: string =
                                    transformVariableSplit[1].trim();
                                if (cleanValue.length > 1) {
                                    cleanValue = cleanValue.substring(
                                        1,
                                        cleanValue.length - 1
                                    );
                                }
                                returnObject[transformVariableSplit[0]] =
                                    cleanValue;
                            }
                        }
                        return returnObject;
                    } else {
                        try {
                            return JSON.parse(variableText);
                        } catch (err) {
                            console.error('Error parsing JSON file: ', err);
                        }
                    }

                    return releaseDefinition.variables[variableName].value;
                }
            }
        }

        return undefined;
    }

    private async onSelectBranchAction(
        event: React.SyntheticEvent<HTMLElement>,
        item: IListBoxItem<{}>
    ): Promise<void> {
        loadingRepositoryObservable.value = true;
        await this.loadBaseCommitAndEnvironmentVariablesFromSelectedBuildArtifact(
            item.text!
        );
        await this.loadInlineTransforms(TransformsTypeEnum.AppSettings);
        await this.loadInlineTransforms(TransformsTypeEnum.ConfigSettings);
        loadingRepositoryObservable.value = false;
    }

    private async loadBaseCommitAndEnvironmentVariablesFromSelectedBuildArtifact(
        branchName: string
    ): Promise<void> {
        const releaseDefinitionForRepo: ReleaseDefinition | undefined =
            Common.getReleaseDefinitionForRepo(
                this.buildDefinitions,
                this.releaseDefinitions,
                this.state.repositoryListSelectedItemObservable.value.id
            );

        this.repositoryBranchBuildArtifactBaseCommit = '';
        this.repositoryBranchBuildArtifactEnvironmentVariables = [];

        if (
            this.currentProject !== undefined &&
            releaseDefinitionForRepo !== undefined
        ) {
            const mostRecentReleasesForBranch: Release[] =
                await Common.getTopReleasesForBranch(
                    this.currentProject.id,
                    releaseDefinitionForRepo.id,
                    1,
                    `refs/heads/${branchName}`
                );
            if (
                mostRecentReleasesForBranch !== undefined &&
                mostRecentReleasesForBranch.length > 0
            ) {
                const mostRecentReleaseForBranch: Release =
                    await Common.getReleaseInfoData(
                        this.currentProject.id,
                        mostRecentReleasesForBranch[0].id
                    );
                if (mostRecentReleaseForBranch.artifacts.length > 0) {
                    const buildArtifact: Artifact | undefined =
                        mostRecentReleaseForBranch.artifacts.find(
                            (artifact: Artifact) =>
                                artifact.type.toLowerCase() === 'build'
                        );
                    if (buildArtifact !== undefined) {
                        this.repositoryBranchBuildArtifactBaseCommit =
                            buildArtifact.definitionReference[
                                'sourceVersion'
                            ].id;
                    }
                }
                for (const environment of mostRecentReleaseForBranch.environments) {
                    for (const variableGroup of environment.variableGroups) {
                        if (
                            Common.ALLOWED_ENVIRONMENT_VARIABLE_GROUP_IDS.includes(
                                variableGroup.id
                            )
                        ) {
                            this.repositoryBranchBuildArtifactEnvironmentVariables.push(
                                variableGroup
                            );
                        }
                    }
                }
            }
        }
    }

    private renderRepositoryEnvironmentVariablesTransformsPage(): JSX.Element {
        if (totalRepositoriesToProcessObservable.value === 0) {
            return (
                <ZeroData
                    primaryText='No repositories.'
                    secondaryText={
                        <span>
                            Please select valid repositories from the Settings
                            page.
                        </span>
                    }
                    imageAltText='No repositories.'
                    imagePath={'../static/notfound.png'}
                />
            );
        }
        return (
            <>
                <Splitter
                    fixedElement={SplitterElementPosition.Near}
                    splitterDirection={SplitterDirection.Vertical}
                    initialFixedSize={450}
                    minFixedSize={100}
                    nearElementClassName='v-scroll-auto custom-scrollbar'
                    farElementClassName='v-scroll-auto custom-scrollbar'
                    onRenderNearElement={this.renderRepositoryMasterPageList}
                    onRenderFarElement={this.renderDetailPageContent}
                />
            </>
        );
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

    //#endregion

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
                                        this.globalEnvironmentVariables ===
                                            undefined ||
                                        this.globalEnvironmentVariables
                                            .length === 0
                                    ) {
                                        this.loadAllEnvironmentVariables();
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
                        {showAllEnvironmentVariablesObservable.value &&
                            this.renderGlobalEnvironmentVariablesPage()}
                        {!showAllEnvironmentVariablesObservable.value &&
                            this.renderRepositoryEnvironmentVariablesTransformsPage()}
                    </div>
                )}
            </Observer>
        );
    }
}
