import { cutPrefix } from "./lib";
import * as ClientAPI from 'azure-devops-extension-api/Common/Client';
import * as BuildClientAPI from 'azure-devops-extension-api/Build/BuildClient';
import * as GetClientAPI from 'azure-devops-extension-api/Git/GitClient';
import * as ExtMgmtAPI from 'azure-devops-extension-api/ExtensionManagement/ExtensionManagementClient'
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService } from 'azure-devops-extension-api/Common';
import { GitAsyncOperationStatus, GitMerge, GitMergeParameters } from "azure-devops-extension-api/Git/Git";
import { BuildResult, BuildStatus } from "azure-devops-extension-api/Build/Build";

export function getBuildClient() { return ClientAPI.getClient(BuildClientAPI.BuildRestClient); }
export function getRunClient() { return ClientAPI.getClient(BuildClientAPI.BuildRestClient); }
export function getGitClient() { return ClientAPI.getClient(GetClientAPI.GitRestClient); }
export function getExtensionManagementClient() { return ClientAPI.getClient(ExtMgmtAPI.ExtensionManagementRestClient); }

export interface TenantInfo {
    organization: string;
    project: string;
}

export interface AuthorInfo {
    displayName: string;
    imageUrl?: string;
}

export async function getAzdoInfo(): Promise<TenantInfo | undefined> {
    let host = SDK.getHost()
    let organization = host.name;

    const projectInfoService = await SDK.getService<IProjectPageService>(
        "ms.vss-tfs-web.tfs-page-data-service" // TODO: CommonServiceIds.ProjectPageService
    );
    const proj = await projectInfoService.getProject();
    let project = proj?.name;

    if (!organization || !project) {
        return undefined;
    }

    return {
        organization: organization,
        project: project
    };
}

export async function getRefCommitId(gitClient: GetClientAPI.GitRestClient, projectId: string, repoId: string, refName: string): Promise<string | undefined> {
    const refHead = cutPrefix(refName, "refs/");
    let resultRefs = (await gitClient.getRefs(repoId, projectId, refHead)).filter((ref) => ref.name === refName);
    if (resultRefs.length === 0) {
        return undefined;
    }
    if (resultRefs.length !== 1) {
        console.error(`Expected 1 ref, but got ${resultRefs.length}`, resultRefs);
        return undefined;
    }
    let repoRef = resultRefs[0];
    let repoObjectId = repoRef.objectId;
    return repoObjectId;
}

export async function getDefaultBranchCommitId(gitClient: GetClientAPI.GitRestClient, projectId: string, repoId: string): Promise<string | undefined> {
    const gitRepo = await gitClient.getRepository(repoId);
    const defaultRef = gitRepo.defaultBranch;
    return await getRefCommitId(gitClient, projectId, repoId, defaultRef);
}

export async function mergeCommits(gitClient: GetClientAPI.GitRestClient, projectId: string, repoId: string, sourceCommitId: string, targetCommitId: string, comment: string): Promise<GitMerge | undefined> {
    let mergeRequestParams: GitMergeParameters = {
        parents: [sourceCommitId, targetCommitId],
        comment: comment,
    };
    let mergeRequest = await gitClient.createMergeRequest(mergeRequestParams, projectId, repoId);
    if (!mergeRequest) {
        return undefined;
    }
    var didTimeout = false;
    var timeout = setTimeout(() => {
        didTimeout = true;
    }, 10000);
    while (!didTimeout) {
        console.log("Checking merge request status...");
        let mergeRequest2 = await gitClient.getMergeRequest(projectId, repoId, mergeRequest.mergeOperationId);
        if (mergeRequest2) {
            let newStatus = mergeRequest2.status;
            if (newStatus == GitAsyncOperationStatus.InProgress) { }
            else if (newStatus == GitAsyncOperationStatus.Queued) { }
            else if (newStatus === GitAsyncOperationStatus.Completed) {
                clearTimeout(timeout);
                return mergeRequest2;
            }
            else if (mergeRequest2.status === GitAsyncOperationStatus.Failed) {
                // TODO:
                clearTimeout(timeout);
                return mergeRequest2;
            }
            else if (mergeRequest2.status === GitAsyncOperationStatus.Abandoned) {
                clearTimeout(timeout);
                return undefined;
            }
            else { // unexpected
                clearTimeout(timeout);
                return undefined;
            }
            await new Promise(resolve => setTimeout(resolve, 1000)); // delay
        }
    }
    return undefined;
}

export interface PullRequestVotingResult {
    status: "none" | "approved" | "suggestions" | "waiting" | "rejected" | "unknown";
    count: number;
}

export function summarizeVotes(reviewers: any): PullRequestVotingResult {
    let finalVote: number | undefined = undefined;
    let finalCount = 0;
    let voters = 0;
    for (let reviewer of reviewers) {
        if (!reviewer) { continue; }
        if (!reviewer.vote) { continue; }

        let vote = reviewer.vote;
        if (!vote || typeof vote !== "number") { continue; }
        // if (!reviewer.isRequired) { continue; } // skip non-required reviewers with no vote

        vote = vote > 10 ? 10 : vote;
        vote = vote < -10 ? -10 : vote;
        if (!reviewer.isContainer) {
            // count only non-container reviewers
            voters++;

            if (finalVote === undefined || finalVote > vote) {
                finalVote = vote;
                if (vote == 5) { finalCount++; }
                else { finalCount = 1; }
            }
            else if (vote == finalVote) {
                finalCount++;
            }
            else {
                continue;
            }
        }
    }
    switch (finalVote || 0) {
        case 10: return { status: "approved", count: finalCount }
        case 5: return { status: "suggestions", count: finalCount }
        case 0: return { status: "none", count: 0 }
        case -5: return { status: "waiting", count: finalCount }
        case -10: return { status: "rejected", count: finalCount }
        default: return { status: "unknown", count: 0 };
    }
}

export function getRunStatus(result: BuildStatus): string {
    switch (result) {
        case BuildStatus.Cancelling: return "cancelling";
        case BuildStatus.Completed: return "completed";
        case BuildStatus.InProgress: return "inProgress";
        case BuildStatus.NotStarted: return "notStarted";
        case BuildStatus.Postponed: return "postponed";
        case BuildStatus.Abandoned: return "abandoned";
        default: return "unknown";
    }
}

export function getRunResult(result: BuildResult): string {
    switch (result) {
        case BuildResult.Succeeded: return "succeeded";
        case BuildResult.Failed: return "failed";
        case BuildResult.Canceled: return "canceled";
        case BuildResult.PartiallySucceeded: return "partiallySucceeded";
        default: return "unknown";
    }
}