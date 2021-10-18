import { IExtensionDataManager } from 'azure-devops-extension-api';
import { ZeroData } from 'azure-devops-ui/ZeroData';
import * as React from 'react';

export default class SprintlyInRelease extends React.Component<{ organizationName: string; dataManager: IExtensionDataManager }> {
    private dataManager: IExtensionDataManager;
    private organizationName: string;

    constructor(props: {
        organizationName: string;
        dataManager: IExtensionDataManager;
    }) {
        super(props);

        this.organizationName = props.organizationName;
        this.dataManager = props.dataManager;
    }

    public render() {
        return (
            <div>
                <ZeroData
                    primaryText="Coming Soon!"
                    secondaryText={
                        <span>In-release (QA) functionality is coming soon!</span>
                    }
                    imageAltText="Coming Soon"
                    imagePath={'../static/notfound.png'}
                />
            </div>
        );
    }
}
