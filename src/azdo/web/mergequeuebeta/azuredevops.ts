import { cutPrefix } from "./lib";
import * as ClientAPI from 'azure-devops-extension-api/Common/Client';
import * as BuildClientAPI from 'azure-devops-extension-api/Build/BuildClient';
import * as GetClientAPI from 'azure-devops-extension-api/Git/GitClient';
import * as ExtMgmtAPI from 'azure-devops-extension-api/ExtensionManagement/ExtensionManagementClient'
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService } from 'azure-devops-extension-api/Common';

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

export async function getDefaultBranchCommitId(gitClient:GetClientAPI.GitRestClient,projectId: string, repoId: string): Promise<string | undefined> {
    const gitRepo = await gitClient.getRepository(repoId);
    const defaultRef = gitRepo.defaultBranch;
    return await getRefCommitId(gitClient, projectId, repoId, defaultRef);
}
