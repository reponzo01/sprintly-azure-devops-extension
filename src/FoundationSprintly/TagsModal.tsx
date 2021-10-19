import { GitRef } from 'azure-devops-extension-api/Git';
import { Dialog } from 'azure-devops-ui/Dialog';
import { SimpleList } from 'azure-devops-ui/List';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import * as React from 'react';

export interface ITagsModalContent {
    modalName: string;
    modalValues: string[];
}

export class TagsModal extends React.Component<{
    isTagsDialogOpen: boolean;
    tagsRepoName: string;
    tags: string[];
    closeMe: () => void;
}> {
    constructor(props: {
        isTagsDialogOpen: boolean;
        tagsRepoName: string;
        tags: string[];
        closeMe: () => void;
    }) {
        super(props);
    }

    public render() {
        return this.props.isTagsDialogOpen ? (
            <Dialog
                titleProps={{
                    text: this.props.tagsRepoName,
                }}
                footerButtonProps={[
                    {
                        text: 'Close',
                        onClick: this.props.closeMe,
                    },
                ]}
                onDismiss={this.props.closeMe}
            >
                <SimpleList
                    itemProvider={
                        new ArrayItemProvider<string>(this.props.tags)
                    }
                />
            </Dialog>
        ) : null;
    }
}

export function getTagsModalContent(
    repositoryName: string,
    branchesAndTags: GitRef[]
): ITagsModalContent {
    const modalName: string = `${repositoryName} Tags`;
    const modalValues: string[] = [];
    for (const branch of branchesAndTags) {
        if (branch.name.includes('refs/tags')) {
            modalValues.push(branch.name);
        }
    }
    return { modalName, modalValues };
}
