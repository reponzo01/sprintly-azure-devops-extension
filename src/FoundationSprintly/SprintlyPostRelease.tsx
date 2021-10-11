import { ZeroData } from 'azure-devops-ui/ZeroData';
import * as React from 'react';

export default class SprintlyPostRelease extends React.Component<{}> {
    constructor(props: {}) {
        super(props);
    }

    public render() {
        return (
            <div>
                <ZeroData
                    primaryText="Coming Soon!"
                    secondaryText={
                        <span>Post release functionality is coming soon!</span>
                    }
                    imageAltText="Coming Soon"
                    imagePath={'../static/notfound.png'}
                />
            </div>
        );
    }
}
