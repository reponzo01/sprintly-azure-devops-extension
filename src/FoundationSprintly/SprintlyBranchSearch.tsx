import * as React from 'react';
import { Button } from 'azure-devops-ui/Button';
import { ButtonGroup } from 'azure-devops-ui/ButtonGroup';
import { Card } from 'azure-devops-ui/Card';
import { ObservableValue } from 'azure-devops-ui/Core/Observable';
import { Page } from 'azure-devops-ui/Page';
import { TextField, TextFieldWidth } from 'azure-devops-ui/TextField';

const searchObservable = new ObservableValue<string>('');

export interface ISprintlyBranchSearchPageState {}

export default class SprintlyBranchSearchPage extends React.Component<
    {},
    ISprintlyBranchSearchPageState
> {
    constructor(props: {}) {
        super(props);

        this.state = {};
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
                    onClick={() => alert('Primary button clicked!')}
                /></ButtonGroup>
                <Page>
                <Card className='bolt-table-card bolt-card-white'>test</Card>
                </Page>
            </div>
        );
    }
}
