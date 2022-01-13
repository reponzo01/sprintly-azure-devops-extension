import * as SDK from 'azure-devops-extension-sdk';

import {
    GitBaseVersionDescriptor,
    GitCommitDiffs,
    GitPullRequest,
    GitPullRequestSearchCriteria,
    GitRef,
    GitRepository,
    GitRestClient,
    GitTargetVersionDescriptor,
    PullRequestStatus,
} from 'azure-devops-extension-api/Git';
import {
    ArtifactSourceReference,
    EnvironmentStatus,
    Release,
    ReleaseDefinition,
    ReleaseEnvironment,
} from 'azure-devops-extension-api/Release';
import {
    CommonServiceIds,
    getClient,
    IColor,
    IExtensionDataManager,
    IExtensionDataService,
} from 'azure-devops-extension-api';
import {
    CoreRestClient,
    TeamProjectReference,
} from 'azure-devops-extension-api/Core';
import axios, { AxiosResponse } from 'axios';
import { BuildDefinition } from 'azure-devops-extension-api/Build';
import { ObservableArray } from 'azure-devops-ui/Core/Observable';
import React from 'react';
import { Link } from 'azure-devops-ui/Link';
import { Icon, IconSize } from 'azure-devops-ui/Icon';
import { Pill, PillSize, PillVariant } from 'azure-devops-ui/Pill';
import { Status, Statuses, StatusSize } from 'azure-devops-ui/Status';
import { DropdownMultiSelection } from 'azure-devops-ui/Utilities/DropdownSelection';

export const primaryColor: IColor = {
    red: 0,
    green: 120,
    blue: 114,
};

export const primaryColorShade30: IColor = {
    red: 0,
    green: 69,
    blue: 120,
};

export const redColor: IColor = {
    red: 191,
    green: 65,
    blue: 65,
};

export const greenColor: IColor = {
    red: 109,
    green: 210,
    blue: 109,
};

export const orangeColor: IColor = {
    red: 225,
    green: 172,
    blue: 74,
};

export const successColor: IColor = {
    red: 47,
    green: 92,
    blue: 55,
};

export const failedColor: IColor = {
    red: 205,
    green: 74,
    blue: 69,
};

export const warningColor: IColor = {
    red: 118,
    green: 90,
    blue: 37,
};

export const repositoryHeadsFilter: string = 'heads/';
export const repositoryTagsFilter: string = 'tags/';
export const DEVELOP: string = 'develop';
export const MASTER: string = 'master';
export const MAIN: string = 'main';

export interface IAllowedEntity {
    displayName: string;
    originId: string;
    descriptor?: string;
}

export interface IReleaseBranchInfo {
    targetBranch: GitRef;
    repositoryId: string;
    aheadOfDevelop?: boolean;
    aheadOfMasterMain?: boolean;
    developPR?: GitPullRequest;
    masterMainPR?: GitPullRequest;
}

export interface IRepositoryBranchInfo {
    repositoryId: string;
    allBranchesAndTags: GitRef[];
    releaseBranches: GitRef[];
    hasDevelopBranch: boolean;
    hasMasterBranch: boolean;
    hasMainBranch: boolean;
}

export interface IReleaseInfo {
    repositoryId: string;
    releaseBranch: IReleaseBranchInfo;
    releases: Release[];
}

export interface IGitRepositoryExtended extends GitRepository {
    hasExistingRelease: boolean;
    hasMainBranch: boolean;
    baseDevelopBranch?: GitRef;
    baseMasterMainBranch?: GitRef;
    existingReleaseBranches: IReleaseBranchInfo[];
    createRelease: boolean;
    branchesAndTags: GitRef[];
}

export interface IProjectRepositories {
    id: string;
    label: string;
    selections: DropdownMultiSelection;
    repositories: IAllowedEntity[];
}

export interface IUserSettings {
    myRepositories: IAllowedEntity[];
    projectRepositoriesId: string;
}

export interface ISystemSettings {
    projectRepositories: IProjectRepositories[];
    allowedUserGroups: IAllowedEntity[];
    allowedUsers: IAllowedEntity[];
}

export async function getOrRefreshToken(token: string): Promise<string> {
    const base64Url: string = token.split('.')[1];
    const base64: string = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const jsonPayload: string = decodeURIComponent(
        atob(base64)
            .split('')
            .map((c: string) => {
                return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
            })
            .join('')
    );

    const decodedToken: any = JSON.parse(jsonPayload);
    const tokenDate: Date = new Date(parseInt(decodedToken.exp) * 1000);
    const now: Date = new Date();
    if (tokenDate <= now) {
        return await SDK.getAccessToken();
    }
    return token;
}

export async function initializeDataManager(
    accessToken: string
): Promise<IExtensionDataManager> {
    const extDataService: IExtensionDataService =
        await SDK.getService<IExtensionDataService>(
            CommonServiceIds.ExtensionDataService
        );
    return await extDataService.getExtensionDataManager(
        SDK.getExtensionContext().id,
        accessToken
    );
}

export async function getUserSettings(
    dataManager: IExtensionDataManager,
    userSettingsDataManagerKey: string
): Promise<IUserSettings | undefined> {
    const userSettings: IUserSettings =
        await dataManager!.getValue<IUserSettings>(userSettingsDataManagerKey, {
            scopeType: 'User',
        });
    return userSettings;
}

export async function getSystemSettings(
    dataManager: IExtensionDataManager,
    systemSettingsDataManagerKey: string
): Promise<ISystemSettings | undefined> {
    const systemSettings: ISystemSettings =
        await dataManager!.getValue<ISystemSettings>(
            systemSettingsDataManagerKey
        );
    return systemSettings;
}

export function getSavedRepositoriesToView(
    userSettings?: IUserSettings,
    systemSettings?: ISystemSettings
): string[] {
    let allowedRepositories: IAllowedEntity[] = [];
    if (!userSettings) {
        return [];
    } else {
        if (userSettings.projectRepositoriesId.trim() === '') {
            allowedRepositories = userSettings.myRepositories;
        } else {
            if (!systemSettings?.projectRepositories) {
                allowedRepositories = userSettings.myRepositories;
            } else {
                const projectRepoIdx: number =
                    systemSettings.projectRepositories.findIndex(
                        (item: IProjectRepositories) =>
                            item.id === userSettings.projectRepositoriesId
                    );
                if (projectRepoIdx > -1) {
                    allowedRepositories =
                        systemSettings.projectRepositories[projectRepoIdx]
                            .repositories;
                } else {
                    allowedRepositories = userSettings.myRepositories;
                }
            }
        }
    }

    return allowedRepositories.map((item: IAllowedEntity) => item.originId);
}

export async function getFilteredProjects(): Promise<TeamProjectReference[]> {
    const projects: TeamProjectReference[] = await getClient(
        CoreRestClient
    ).getProjects();

    const filteredProjects: TeamProjectReference[] = projects.filter(
        (project: TeamProjectReference) => {
            return (
                project.name === 'Portfolio' ||
                project.name === 'Sample Project'
            );
        }
    );
    return filteredProjects;
}

export async function getFilteredProjectRepositories(
    projectId: string,
    savedRepos: string[]
): Promise<GitRepository[]> {
    const repos: GitRepository[] = await getClient(
        GitRestClient
    ).getRepositories(projectId);
    let filteredRepos: GitRepository[] = repos;
    filteredRepos = repos.filter((repo: GitRepository) =>
        savedRepos.includes(repo.id)
    );
    return filteredRepos;
}

export async function isBranchAheadOfDevelop(
    branchName: string,
    repositoryId: string
): Promise<boolean> {
    const developBranchDescriptor: GitBaseVersionDescriptor = {
        baseVersion: 'develop',
        baseVersionOptions: 0,
        baseVersionType: 0,
        version: 'develop',
        versionOptions: 0,
        versionType: 0,
    };
    const releaseBranchDescriptor: GitTargetVersionDescriptor = {
        targetVersion: branchName,
        targetVersionOptions: 0,
        targetVersionType: 0,
        version: branchName,
        versionOptions: 0,
        versionType: 0,
    };

    const developCommitsDiff: GitCommitDiffs = await getCommitDiffs(
        repositoryId,
        developBranchDescriptor,
        releaseBranchDescriptor
    );

    return codeChangesInCommitDiffs(developCommitsDiff);
}

export async function isBranchAheadOMasterMain(
    repositoryBranchInfo: IRepositoryBranchInfo,
    branchName: string,
    repositoryId: string
): Promise<boolean> {
    const masterMainBranchDescriptor: GitBaseVersionDescriptor = {
        baseVersion: repositoryBranchInfo.hasMasterBranch ? 'master' : 'main',
        baseVersionOptions: 0,
        baseVersionType: 0,
        version: repositoryBranchInfo.hasMasterBranch ? 'master' : 'main',
        versionOptions: 0,
        versionType: 0,
    };
    const releaseBranchDescriptor: GitTargetVersionDescriptor = {
        targetVersion: branchName,
        targetVersionOptions: 0,
        targetVersionType: 0,
        version: branchName,
        versionOptions: 0,
        versionType: 0,
    };

    const masterMainCommitsDiff: GitCommitDiffs = await getCommitDiffs(
        repositoryId,
        masterMainBranchDescriptor,
        releaseBranchDescriptor
    );

    return codeChangesInCommitDiffs(masterMainCommitsDiff);
}

export async function getCommitDiffs(
    repoId: string,
    baseVersion: GitBaseVersionDescriptor,
    targetVersion: GitTargetVersionDescriptor
): Promise<GitCommitDiffs> {
    return await getClient(GitRestClient).getCommitDiffs(
        repoId,
        undefined,
        undefined,
        1000,
        0,
        baseVersion,
        targetVersion
    );
}

export function codeChangesInCommitDiffs(commitsDiff: GitCommitDiffs): boolean {
    return (
        Object.keys(commitsDiff.changeCounts).length > 0 ||
        commitsDiff.changes.length > 0
    );
}

export async function getRepositoryInfo(
    repoId: string,
    filter?: string
): Promise<GitRef[]> {
    return await getClient(GitRestClient).getRefs(
        repoId,
        undefined,
        filter ? filter : undefined,
        false,
        false,
        undefined,
        true,
        false,
        undefined
    );
}

export async function getRepositoryBranchesInfo(
    repositoryId: string,
    filter?: string
): Promise<IRepositoryBranchInfo> {
    let hasDevelopBranch: boolean = false;
    let hasMasterBranch: boolean = false;
    let hasMainBranch: boolean = false;

    const allBranchesAndTags: GitRef[] = await getRepositoryInfo(
        repositoryId,
        filter
    );
    const releaseBranches: GitRef[] = [];

    for (const branch of allBranchesAndTags) {
        if (branch.name.includes('heads/develop')) {
            hasDevelopBranch = true;
        } else if (branch.name.includes('heads/master')) {
            hasMasterBranch = true;
        } else if (branch.name.includes('heads/main')) {
            hasMainBranch = true;
        } else if (branch.name.includes('heads/release')) {
            releaseBranches.push(branch);
        }
    }

    return {
        repositoryId,
        allBranchesAndTags,
        releaseBranches,
        hasDevelopBranch,
        hasMasterBranch,
        hasMainBranch,
    };
}

export function sortRepositoryList(
    repositoryList: GitRepository[]
): GitRepository[] {
    if (repositoryList.length > 0) {
        return repositoryList.sort(
            (a: GitRepository, b: GitRepository) => {
                return a.name.localeCompare(b.name);
            }
        );
    }
    return repositoryList;
}

export function sortRepositoryExtendedList(
    repositoryList: IGitRepositoryExtended[]
): IGitRepositoryExtended[] {
    if (repositoryList.length > 0) {
        return repositoryList.sort(
            (a: IGitRepositoryExtended, b: IGitRepositoryExtended) => {
                return a.name.localeCompare(b.name);
            }
        );
    }
    return repositoryList;
}

export function sortAllowedEntityList(
    allowedEntityList: IAllowedEntity[]
): IAllowedEntity[] {
    if (allowedEntityList.length > 0) {
        return allowedEntityList.sort(
            (a: IAllowedEntity, b: IAllowedEntity) => {
                return a.displayName.localeCompare(b.displayName);
            }
        );
    }
    return allowedEntityList;
}

export function sortBranchesList(
    branchesList: GitRef[]
): GitRef[] {
    if (branchesList.length > 0) {
        return branchesList.sort(
            (a: GitRef, b: GitRef) => {
                return a.name.localeCompare(b.name);
            }
        );
    }
    return branchesList;
}

export async function fetchAndStoreBranchReleaseInfoIntoObservable(
    releaseInfoObservable: ObservableArray<IReleaseInfo>,
    buildDefinitionForRepo: BuildDefinition,
    releaseDefinitions: ReleaseDefinition[],
    releaseBranch: IReleaseBranchInfo,
    projectId: string,
    repositoryId: string,
    organizationName: string,
    accessToken: string
): Promise<void> {
    const buildDefinitionId: number = buildDefinitionForRepo.id;
    let releaseDefinitionId: number = -1;
    for (const releaseDefinition of releaseDefinitions) {
        for (const artifact of releaseDefinition.artifacts) {
            if (artifact.isPrimary) {
                const releaseDefBuildDef: ArtifactSourceReference =
                    artifact.definitionReference['definition'];
                if (releaseDefBuildDef) {
                    if (
                        releaseDefBuildDef.id === buildDefinitionId.toString()
                    ) {
                        releaseDefinitionId = releaseDefinition.id;
                        break;
                    }
                }
            }
        }
        if (releaseDefinitionId > -1) break;
    }
    if (releaseDefinitionId > -1) {
        await getReleasesForReleaseBranch(
            releaseInfoObservable,
            releaseBranch,
            releaseDefinitionId,
            projectId,
            repositoryId,
            organizationName,
            accessToken
        );
    }
}

export async function getReleasesForReleaseBranch(
    releaseInfoObservable: ObservableArray<IReleaseInfo>,
    releaseBranch: IReleaseBranchInfo,
    releaseDefinitionId: number,
    projectId: string,
    repositoryId: string,
    organizationName: string,
    accessToken: string
): Promise<void> {
    accessToken = await getOrRefreshToken(accessToken);
    const response: AxiosResponse<never> = await axios
        .get(
            `https://vsrm.dev.azure.com/${organizationName}/${projectId}/_apis/release/releases?$expand=environments&sourceBranchFilter=${releaseBranch.targetBranch.name}&definitionId=${releaseDefinitionId}&api-version=6.0`,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                },
            }
        )
        .catch((error: any) => {
            console.error(error);
            throw error;
        });

    const data: { count: number; value: Release[] } = response.data;
    if (data && data.count > 0) {
        const existingIndex: number = releaseInfoObservable.value.findIndex(
            (item: IReleaseInfo) =>
                item.releaseBranch.targetBranch.name ===
                    releaseBranch.targetBranch.name &&
                item.repositoryId === releaseBranch.repositoryId
        );
        const releaseInfo: IReleaseInfo = {
            repositoryId,
            releaseBranch,
            releases: data.value,
        };
        if (existingIndex < 0) {
            releaseInfoObservable.push(releaseInfo);
        } else {
            releaseInfoObservable.change(existingIndex, releaseInfo);
        }
    }
}

export async function getReleaseDefinitions(
    projects: TeamProjectReference[],
    organizationName: string,
    accessToken: string
): Promise<ReleaseDefinition[]> {
    for (const project of projects) {
        accessToken = await getOrRefreshToken(accessToken);
        const response: AxiosResponse<never> = await axios
            .get(
                `https://vsrm.dev.azure.com/${organizationName}/${project.id}/_apis/release/definitions?$expand=artifacts&api-version=6.0`,
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                    },
                }
            )
            .catch((error: any) => {
                console.error(error);
                throw error;
            });

        const data: { count: number; value: ReleaseDefinition[] } =
            response.data;
        if (data && data.count > 0) {
            return data.value;
        }
    }
    return [];
}

export async function getBuildDefinitions(
    projects: TeamProjectReference[],
    organizationName: string,
    accessToken: string
): Promise<BuildDefinition[]> {
    for (const project of projects) {
        accessToken = await getOrRefreshToken(accessToken);
        const response: AxiosResponse<never> = await axios
            .get(
                `https://dev.azure.com/${organizationName}/${project.id}/_apis/build/definitions?includeAllProperties=true&api-version=6.0`,
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                    },
                }
            )
            .catch((error: any) => {
                console.error(error);
                throw error;
            });

        const data: { count: number; value: BuildDefinition[] } = response.data;
        if (data && data.count > 0) {
            return data.value;
        }
    }
    return [];
}

export async function getPullRequests(
    projects: TeamProjectReference[]
): Promise<GitPullRequest[]> {
    // Statuses:
    // 1 = Queued, 2 = Conflicts, 3 = Premerge Succeeded, 4 = RejectedByPolicy, 5 = Failure
    let pullRequests: GitPullRequest[] = [];
    const pullRequestCriteria: GitPullRequestSearchCriteria = {
        includeLinks: false,
        creatorId: '',
        repositoryId: '',
        reviewerId: '',
        sourceRefName: '',
        sourceRepositoryId: '',
        status: PullRequestStatus.Active,
        targetRefName: '',
    };
    for (const project of projects) {
        const pullRequestsResponse: GitPullRequest[] = await getClient(
            GitRestClient
        ).getPullRequestsByProject(project.id, pullRequestCriteria);
        pullRequests = pullRequests.concat(pullRequestsResponse);
    }
    return pullRequests;
}

export function getMostRecentReleaseForBranch(
    releaseBranchInfo: IReleaseBranchInfo | undefined,
    releaseInfoForAllBranches: IReleaseInfo[]
): Release | undefined {
    if (!releaseBranchInfo) return;
    const releaseInfoForBranch: IReleaseInfo | undefined =
        releaseInfoForAllBranches.find(
            (ri: IReleaseInfo) =>
                ri.releaseBranch.targetBranch.name ===
                    releaseBranchInfo.targetBranch.name &&
                ri.repositoryId === releaseBranchInfo.repositoryId
        );

    if (releaseInfoForBranch && releaseInfoForBranch.releases.length > 0) {
        const sortedReleases: Release[] = releaseInfoForBranch.releases.sort(
            (a: Release, b: Release) => {
                return (
                    new Date(b.createdOn.toString()).getTime() -
                    new Date(a.createdOn.toString()).getTime()
                );
            }
        );
        if (sortedReleases.length > 0) {
            return sortedReleases[0];
        }
    }
}

export function getBranchShortName(branchRealName: string): string {
    if (branchRealName.includes('refs/heads/')) {
        return branchRealName.split('refs/heads/')[1];
    }
    return branchRealName;
}

export function repositoryLinkJsxElement(
    webUrl: string,
    className: string,
    repositoryName: string
): JSX.Element {
    return (
        <Link
            excludeTabStop
            href={webUrl + '/branches'}
            subtle={true}
            target='_blank'
            className={className}
        >
            <u>{repositoryName}</u>
        </Link>
    );
}

export function branchLinkJsxElement(
    key: string,
    webUrl: string,
    branchName: string,
    className: string,
    isReleaseLink: boolean = false
): JSX.Element {
    return (
        <Link
            key={key}
            excludeTabStop
            href={
                webUrl +
                (isReleaseLink ? '' : '?version=GB' + encodeURI(branchName))
            }
            target='_blank'
            className={className}
        >
            {branchName}
        </Link>
    );
}

export function noReleaseExistsPillJsxElement(): JSX.Element {
    return (
        <Pill
            color={warningColor}
            size={PillSize.regular}
            variant={PillVariant.outlined}
            className='bolt-list-overlay sprintly-environment-status'
        >
            <div className='sprintly-text-white'>
                <Icon
                    ariaLabel='No Release Exists'
                    iconName='Warning'
                    size={IconSize.small}
                />{' '}
                No Release
            </div>
        </Pill>
    );
}

export function getAllEnvironmentStatusPillJsxElements(
    environments: ReleaseEnvironment[],
    onClickAction?: () => void
): JSX.Element[] {
    const environmentStatuses: JSX.Element[] = [];
    for (const environment of environments) {
        environmentStatuses.push(
            getSingleEnvironmentStatusPillJsxElement(environment, onClickAction)
        );
    }
    return environmentStatuses;
}

export function getSingleEnvironmentStatusPillJsxElement(
    environment: ReleaseEnvironment,
    onClickAction?: () => void
): JSX.Element {
    let envStatusEnumValue: string = '';
    let statusIconName: string = 'Cancel';
    let divTextClassName: string = 'sprintly-text-white';
    for (const idx in EnvironmentStatus) {
        if (idx.toLowerCase() === environment.status.toString().toLowerCase()) {
            envStatusEnumValue = EnvironmentStatus[idx];
        }
    }
    switch (parseInt(envStatusEnumValue)) {
        case EnvironmentStatus.NotStarted:
            statusIconName = 'CircleRing';
            divTextClassName = '';
            break;
        case EnvironmentStatus.InProgress:
            statusIconName = 'UseRunningStatus';
            divTextClassName = '';
            break;
        case EnvironmentStatus.Queued:
            statusIconName = 'UseRunningStatus';
            divTextClassName = '';
            break;
        case EnvironmentStatus.Scheduled:
            statusIconName = 'UseRunningStatus';
            divTextClassName = '';
            break;
        case EnvironmentStatus.Succeeded:
            statusIconName = 'Accept';
            break;
    }
    return environmentStatusPillJsxElement(
        environment,
        envStatusEnumValue,
        divTextClassName,
        statusIconName,
        onClickAction
    );
}

export function environmentStatusPillJsxElement(
    environment: ReleaseEnvironment,
    envStatusEnumString: string,
    divTextClassName: string,
    statusIconName: string,
    onClickAction?: () => void
): JSX.Element {
    return (
        <Pill
            onClick={onClickAction}
            key={environment.id}
            color={
                parseInt(envStatusEnumString) ===
                parseInt(EnvironmentStatus.Succeeded.toString())
                    ? successColor
                    : parseInt(envStatusEnumString) ===
                          parseInt(EnvironmentStatus.Undefined.toString()) ||
                      parseInt(envStatusEnumString) ===
                          parseInt(EnvironmentStatus.Canceled.toString()) ||
                      parseInt(envStatusEnumString) ===
                          parseInt(EnvironmentStatus.Rejected.toString()) ||
                      parseInt(envStatusEnumString) ===
                          parseInt(
                              EnvironmentStatus.PartiallySucceeded.toString()
                          )
                    ? failedColor
                    : undefined
            }
            size={PillSize.regular}
            variant={PillVariant.outlined}
            className={
                'sprintly-environment-status' +
                (onClickAction === undefined ? ' bolt-list-overlay' : '')
            }
        >
            <div className={divTextClassName}>
                {statusIconName === 'UseRunningStatus' ? (
                    <Status
                        {...Statuses.Running}
                        key='running'
                        size={StatusSize.m}
                        className='sprintly-vertical-align-bottom'
                    />
                ) : (
                    <Icon iconName={statusIconName} size={IconSize.small} />
                )}{' '}
                {environment.name}
            </div>
        </Pill>
    );
}
