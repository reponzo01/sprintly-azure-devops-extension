import { SingleLayerMasterPanelHeader } from 'azure-devops-ui/Components/SingleLayerMasterPanel/SingleLayerMasterPanel';
import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import {
    IListItemDetails,
    List,
    ListItem,
    ListSelection,
} from 'azure-devops-ui/List';
import {
    SplitterElementPosition,
    Splitter,
    SplitterDirection,
} from 'azure-devops-ui/Splitter';
import { SingleLayerMasterPanel } from 'azure-devops-ui/MasterDetails';
import { bindSelectionToObservable } from 'azure-devops-ui/MasterDetailsContext';
import { Observer } from 'azure-devops-ui/Observer';
import {
    ArrayItemProvider,
    IItemProvider,
} from 'azure-devops-ui/Utilities/Provider';
import { Page } from 'azure-devops-ui/Page';
import { Tooltip } from 'azure-devops-ui/TooltipEx';
import * as React from 'react';

const sampleDate: string[] = [
    'Added GitHub aliases',
    'Remove reference to Design System components',
    'Using new design/pattern/components',
    'Fixing bug with spacing',
    'Added some new components',
    'Setting up theme variables',
    'Updating Button focus behavior',
    'Remove reference to Design System components',
    'Using new design/pattern/components',
    'Fixing bug with spacing',
    'Added some new components',
    'Setting up theme variables',
    'Updating Button focus behavior',
    'Remove reference to Design System components',
    'Using new design/pattern/components',
    'Fixing bug with spacing',
    'Added some new components',
    'Setting up theme variables',
    'Updating Button focus behavior',
    'Remove reference to Design System components',
    'Using new design/pattern/components',
    'Fixing bug with spacing',
    'Added some new components',
    'Setting up theme variables',
    'Updating Button focus behavior',
    'Remove reference to Design System components',
    'Using new design/pattern/components',
    'Fixing bug with spacing',
    'Added some new components',
    'Setting up theme variables',
    'Updating Button focus behavior',
];

export interface ISprintlyPostReleaseState {
    selection: ListSelection;
    itemProvider: ArrayItemProvider<string>;
    selectedItemObservable: ObservableValue<string>;
}

export default class SprintlyPostRelease extends React.Component<
    {
    },
    ISprintlyPostReleaseState
> {

    constructor(props: {}) {
        super(props);

        this.state = {
            itemProvider: new ArrayItemProvider(sampleDate),
            selectedItemObservable: new ObservableValue<string>(sampleDate[0]),
            selection: new ListSelection({ selectOnFocus: false })
        };
    }

    public componentDidMount(): void {
        bindSelectionToObservable(
            this.state.selection,
            this.state.itemProvider,
            this.state.selectedItemObservable
        );
    }

    public render(): JSX.Element {
        return (<div style={{ height: '85%', width: '100%', display: 'flex' }}>
        <Splitter
            fixedElement={SplitterElementPosition.Near}
            splitterDirection={SplitterDirection.Vertical}
            initialFixedSize={450}
            minFixedSize={100}
            nearElementClassName="v-scroll-auto custom-scrollbar light-grey"
            farElementClassName="v-scroll-auto custom-scrollbar"
            onRenderNearElement={() => (<List
                ariaLabel={'Commits Master Table'}
                itemProvider={this.state.itemProvider}
                selection={this.state.selection}
                renderRow={renderListItem}
                width="100%"
                singleClickActivation={true}
            />)}
            onRenderFarElement={() => (
                <Observer selectedItem={this.state.selectedItemObservable}>
                    {(observerProps: { selectedItem: string }) => (
                        <Page className="flex-grow single-layer-details">
                            {observerProps.selectedItem && (
                                <Tooltip
                                    text={observerProps.selectedItem}
                                    overflowOnly={true}
                                >
                                    <span className="single-layer-details-contents">
                                        {observerProps.selectedItem} This is
                                        the Detail Page
                                    </span>
                                </Tooltip>
                            )}
                        </Page>
                    )}
                </Observer>
            )}
        />
    </div>);
    }
}

function renderListItem(
    index: number,
    item: string,
    details: IListItemDetails<string>,
    key?: string
): JSX.Element {
    return (
        <ListItem
            className="master-row"
            key={key || 'list-item' + index}
            index={index}
            details={details}
        >
            <div className="master-row-content flex-row flex-center h-scroll-hidden">
                <Tooltip overflowOnly={true}>
                    <div className="primary-text text-ellipsis">{item}xxxx</div>
                </Tooltip>
            </div>
        </ListItem>
    );
};
