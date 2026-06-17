import { cutPrefix } from "./lib";
import * as ClientAPI from 'azure-devops-extension-api/Common/Client';
import * as BuildClientAPI from 'azure-devops-extension-api/Build/BuildClient';
import * as GetClientAPI from 'azure-devops-extension-api/Git/GitClient';
import * as ExtMgmtAPI from 'azure-devops-extension-api/ExtensionManagement/ExtensionManagementClient'
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService } from 'azure-devops-extension-api/Common';
import { GitAsyncOperationStatus, GitMerge, GitMergeParameters } from "azure-devops-extension-api/Git/Git";

export function getBuildClient() { return ClientAPI.getClient(BuildClientAPI.BuildRestClient); }
export function getGitClient() { return ClientAPI.getClient(GetClientAPI.GitRestClient); }
export function getExtensionManagementClient() { return ClientAPI.getClient(ExtMgmtAPI.ExtensionManagementRestClient); }

export interface TenantInfo {
    organization: string;
    project: string;
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
    const defaultHead = cutPrefix(refName, "refs/");
    const resultRefs = await gitClient.getRefs(repoId, projectId, defaultHead);
    if (resultRefs.length !== 1) {
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

export async function mergeCommits(gitClient: GetClientAPI.GitRestClient, projectId: string, repoId: string, sourceCommitId: string, targetCommitId: string): Promise<GitMerge | undefined> {
    let mergeRequestParams: GitMergeParameters = {
        parents: [sourceCommitId, targetCommitId],
        comment: `Merge Queue: ${sourceCommitId} into ${targetCommitId}`
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