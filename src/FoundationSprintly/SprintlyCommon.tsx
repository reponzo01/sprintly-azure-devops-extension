import * as SDK from 'azure-devops-extension-sdk';

import {
    GitBaseVersionDescriptor,
    GitCommitDiffs,
    GitPullRequest,
    GitRef,
    GitRepository,
    GitRestClient,
    GitTargetVersionDescriptor,
} from 'azure-devops-extension-api/Git';
import { Release } from 'azure-devops-extension-api/Release';
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

export const primaryColor: IColor = {
    red: 0,
    green: 120,
    blue: 114,
}

export const primaryColorShade30: IColor = {
    red: 0,
    green: 69,
    blue: 120,
}

export const redColor: IColor = {
    red: 191,
    green: 65,
    blue: 65,
}

export const greenColor: IColor = {
    red: 109,
    green: 210,
    blue: 109,
}

export const orangeColor: IColor = {
    red: 225,
    green: 172,
    blue: 74,
}

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
    branchesAndTags: GitRef[];
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
    existingReleaseBranches: IReleaseBranchInfo[];
    createRelease: boolean;
    branchesAndTags: GitRef[];
}

export async function getOrRefreshToken(token: string): Promise<string> {
    const base64Url = token.split('.')[1];
    const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const jsonPayload = decodeURIComponent(
        atob(base64)
            .split('')
            .map((c) => {
                return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
            })
            .join('')
    );

    const decodedToken = JSON.parse(jsonPayload);
    const tokenDate = new Date(parseInt(decodedToken.exp) * 1000);
    const now = new Date();
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

export async function getSavedRepositoriesToProcess(
    dataManager: IExtensionDataManager,
    repositoriesToProcessKey: string
): Promise<IAllowedEntity[]> {
    let repositoriesToProcess: IAllowedEntity[] = [];
    const savedRepositories: IAllowedEntity[] = await dataManager!.getValue<
        IAllowedEntity[]
    >(repositoriesToProcessKey, {
        scopeType: 'User',
    });
    if (savedRepositories) {
        repositoriesToProcess = repositoriesToProcess.concat(savedRepositories);
    }
    return repositoriesToProcess;
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

export async function getRepositoryInfo(repoId: string): Promise<GitRef[]> {
    return await getClient(GitRestClient).getRefs(
        repoId,
        undefined,
        undefined,
        false,
        false,
        undefined,
        true,
        false,
        undefined
    );
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

export async function getRepositoryBranchInfo(
    repositoryId: string
): Promise<IRepositoryBranchInfo> {
    let hasDevelopBranch: boolean = false;
    let hasMasterBranch: boolean = false;
    let hasMainBranch: boolean = false;

    const branchesAndTags: GitRef[] = await getRepositoryInfo(repositoryId);

    for (const ref of branchesAndTags) {
        if (ref.name.includes('heads/develop')) {
            hasDevelopBranch = true;
        } else if (ref.name.includes('heads/master')) {
            hasMasterBranch = true;
        } else if (ref.name.includes('heads/main')) {
            hasMainBranch = true;
        }
    }

    return {
        repositoryId,
        branchesAndTags,
        hasDevelopBranch,
        hasMasterBranch,
        hasMainBranch
    }
}

export function sortRepositoryList(repositoryList: IGitRepositoryExtended[]): IGitRepositoryExtended[] {
    if (repositoryList.length > 0) {
        return repositoryList.sort(
            (
                a: IGitRepositoryExtended,
                b: IGitRepositoryExtended
            ) => {
                return a.name.localeCompare(b.name);
            }
        );
    }
    return repositoryList;
}
