import * as React from "react";

import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Observer } from 'azure-devops-ui/Observer';
import { Tab, TabBar, TabSize } from 'azure-devops-ui/Tabs';
import { Page } from "azure-devops-ui/Page";
import { Header, TitleSize } from "azure-devops-ui/Header";

import { SprintlyPage } from './SprintlyPage';
import SprintlySettings from './SprintlySettings';
import { showRootComponent } from "../Common";

const selectedTabId: ObservableValue<string> = new ObservableValue<string>('sprintly-page');

export default class FoundationSprintly extends React.Component<{}> {
    constructor(props: {}) {
        super(props);
    }

    public render() {
        return (
            /* tslint:disable */
            <Page className="flex-grow foundation-sprintly">
                <Header title="Foundation Sprintly"
                    titleSize={TitleSize.Large} />

                <TabBar
                    onSelectedTabChanged={onSelectedTabChanged}
                    selectedTabId={selectedTabId}
                    tabSize={TabSize.Tall}
                >
                    <Tab name="Sprintly" id="sprintly-page" />
                    <Tab name="Settings" id="sprintly-settings" />
                </TabBar>
                <Observer selectedTabId={selectedTabId}>
                    {(props: { selectedTabId: string }) => {
                        if (selectedTabId.value === 'sprintly-page') {
                            return <SprintlyPage />;
                        }
                        return <SprintlySettings sampleProp={selectedTabId.value} />;
                    }}
                </Observer>
            </Page>
            /* tslint:disable */
        )
    }
}

function onSelectedTabChanged(newTabId: string) {
    selectedTabId.value = newTabId;
}

showRootComponent(<FoundationSprintly />);
