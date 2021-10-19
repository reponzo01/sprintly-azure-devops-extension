import { GitRef } from 'azure-devops-extension-api/Git';
import { Dialog } from 'azure-devops-ui/Dialog';
import { SimpleList } from 'azure-devops-ui/List';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import * as React from 'react';

export interface ITagsModalContent {
    modalName: string;
    modalValues: string[];
}

export interface ITagsModalState {
    isTagsDialogOpen: boolean;
}

export class TagsModal extends React.Component<
    {
        isTagsDialogOpen: boolean;
        tagsRepoName: string;
        tags: string[];
    },
    ITagsModalState
> {
    constructor(props: {
        isTagsDialogOpen: boolean;
        tagsRepoName: string;
        tags: string[];
    }) {
        super(props);

        this.state = {
            isTagsDialogOpen: props.isTagsDialogOpen,
        };
    }

    public render() {
        const onDismiss: () => void = () => {
            this.setState({
                isTagsDialogOpen: false,
            });
        };
        return this.state.isTagsDialogOpen ? (
            <Dialog
                titleProps={{
                    text: this.props.tagsRepoName,
                }}
                footerButtonProps={[
                    {
                        text: 'Close',
                        onClick: onDismiss,
                    },
                ]}
                onDismiss={onDismiss}
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
