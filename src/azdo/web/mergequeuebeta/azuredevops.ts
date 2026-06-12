import * as ClientAPI from 'azure-devops-extension-api/Common/Client';
import * as BuildClientAPI from 'azure-devops-extension-api/Build/BuildClient';
import * as GetClientAPI from 'azure-devops-extension-api/Git/GitClient';
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService } from 'azure-devops-extension-api/Common';

export function getBuildClient() { return ClientAPI.getClient(BuildClientAPI.BuildRestClient); }
export function getGitClient() { return ClientAPI.getClient(GetClientAPI.GitRestClient); }

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

