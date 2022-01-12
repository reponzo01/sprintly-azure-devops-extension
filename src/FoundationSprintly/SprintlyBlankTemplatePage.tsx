import { ZeroData } from 'azure-devops-ui/ZeroData';
import * as React from 'react';

export interface ISprintlyBlankTemplatePageState {}

export default class SprintlyBlankTemplatePage extends React.Component<
    {},
    ISprintlyBlankTemplatePageState
> {
    constructor(props: {}) {
        super(props);

        this.state = {};
    }

    public render(): JSX.Element {
        return (
            <div>
                <ZeroData
                    primaryText='Nothing to see here.'
                    secondaryText={<span>Add some content.</span>}
                    imageAltText='Nothing Here'
                    imagePath={'../static/notfound.png'}
                />
            </div>
        );
    }
}
