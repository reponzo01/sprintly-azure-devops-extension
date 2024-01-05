import { GitRef } from 'azure-devops-extension-api/Git';
import { Dialog } from 'azure-devops-ui/Dialog';
import { SimpleList } from 'azure-devops-ui/List';
import { ArrayItemProvider } from 'azure-devops-ui/Utilities/Provider';
import * as React from 'react';

import { repositoryTagsFilter, getRepositoryInfo } from './SprintlyCommon';

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

    public render(): JSX.Element | null {
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
                    scrollable={true}
                />
            </Dialog>
        ) : null;
    }
}

export async function getTagsModalContent(
    repositoryName: string,
    repositoryId: string
): Promise<ITagsModalContent> {
    const modalName: string = `${repositoryName} Tags`;
    const modalValues: string[] = [];

    const tags: GitRef[] = await getRepositoryInfo(
        repositoryId,
        repositoryTagsFilter
    );
    if (!tags || tags.length <= 0) {
        modalValues.push('No tags found.');
    } else {
        for (const tag of tags) {
            modalValues.push(tag.name);
        }
    }

    return { modalName, modalValues };
}
