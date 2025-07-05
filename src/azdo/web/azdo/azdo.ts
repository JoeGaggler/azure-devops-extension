import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService } from "azure-devops-extension-api";
import { type IExtensionDataService } from 'azure-devops-extension-api';

export async function getAzdo(url: string, bearertoken: string): Promise<any> {
    const response = await fetch(url, {
        method: "GET",
        headers: {
            "Authorization": `Bearer ${bearertoken}`,
            "Content-Type": "application/json",
            "X-Bearer": bearertoken,
        }
    });
    if (response.ok) {
        const data = await response.json();
        // console.log(data);
        return data;
    } else {
        console.log("Error fetching azdo data", response.statusText);
    }
}

export async function deleteAzdo(url: string, bearertoken: string): Promise<any> {
    console.log("deleteapi", url);
    const response = await fetch(url, {
        method: "DELETE",
        headers: {
            "Authorization": `Bearer ${bearertoken}`,
            "Content-Type": "application/json"
        }
    });
    if (response.ok) {
        // const data = await response.json();
        // console.log(data);
        // return data;
    } else {
        console.log("Error deleting azdo data", response.statusText);
    }
}

export async function patchAzdo(url: string, body: any, bearertoken: string): Promise<any> {
    console.log("patchapi", url, body);
    const response = await fetch(url, {
        method: "PATCH",
        headers: {
            "Authorization": `Bearer ${bearertoken}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });
    console.log("patchapi response", response);
    if (response.ok) {
        const data = await response.json();
        // console.log(data);
        return data;
    } else {
        console.log("Error patching azdo data", response.statusText);
    }
}

export async function postAzdo(url: string, body: any, bearertoken: string): Promise<any> {
    console.log("postapi", url, body);
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Authorization": `Bearer ${bearertoken}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });
    if (response.ok) {
        const data = await response.json();
        // console.log(data);
        return data;
    } else {
        console.log("Error posting azdo data", response.statusText);
    }
}

export async function getAllPullRequests(tenantInfo: TenantInfo): Promise<Array<PullRequest>> {
    let bearer = await SDK.getAccessToken()
    let json = await getAzdo(`https://dev.azure.com/${tenantInfo.organization}/${tenantInfo.project}/_apis/git/pullrequests?api-version=7.2-preview.2`, bearer as string);
    return json.value
}

export async function approveRun(approvalId: string, status: string, comment: string, bearertoken: string): Promise<any> {
    const response = await patchAzdo(`https://dev.azure.com/org/proj/_apis/pipelines/approvals?api-version=7.2-preview.2`, [{
        approvalId: approvalId,
        comment: comment,
        status: status
    }], bearertoken as string);
    console.log("approval response: ", response);
    return response;
}

export async function deleteSharedDocument(colId: string, docId: string): Promise<boolean> {
    const accessToken = await SDK.getAccessToken();
    const extDataService = await SDK.getService<IExtensionDataService>("ms.vss-features.extension-data-service");
    const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

    try {
        await dataManager.deleteDocument(colId, docId);
        console.log("Delete Shared Document: ", colId, docId);
        return true;
    } catch {
        return false;
    }
}

export async function getOrCreateSharedDocument(colId: string, docId: string, document: any): Promise<any> {
    const accessToken = await SDK.getAccessToken();
    const extDataService = await SDK.getService<IExtensionDataService>("ms.vss-features.extension-data-service");
    const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

    let userFiltersDoc = document
    try {
        userFiltersDoc = await dataManager.getDocument(colId, docId);
        console.log("Get Shared Document: ", colId, docId, userFiltersDoc);
    } catch {
        userFiltersDoc.id = docId;
        userFiltersDoc = await dataManager.createDocument(colId, userFiltersDoc);
        console.log("New Shared Document: ", colId, docId, userFiltersDoc);
    }

    return userFiltersDoc;
}

export async function trySaveSharedDocument(colId: string, docId: string, document: any): Promise<boolean> {
    const accessToken = await SDK.getAccessToken();
    const extDataService = await SDK.getService<IExtensionDataService>("ms.vss-features.extension-data-service");
    const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

    document.id = docId;

    try {
        console.log("Updating Shared Document: ", colId, docId, document);
        let next_document = await dataManager.setDocument(colId, document);
        console.log("Update Shared Document: ", colId, docId, next_document);
        return true;
    } catch (err) {
        console.error("Error updating shared document", err);
        return false;
    }
}

export async function getOrCreateUserDocument(colId: string, docId: string, document: any): Promise<any> {
    const accessToken = await SDK.getAccessToken();
    const extDataService = await SDK.getService<IExtensionDataService>("ms.vss-features.extension-data-service");
    const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

    let userFiltersDoc = document
    try {
        userFiltersDoc = await dataManager.getDocument(colId, docId, { scopeType: "User" });
        console.log("Get User Document: ", colId, docId, userFiltersDoc);
    } catch {
        userFiltersDoc.id = docId;
        userFiltersDoc = await dataManager.createDocument(colId, userFiltersDoc, { scopeType: "User" });
        console.log("New User Document: ", colId, docId, userFiltersDoc);
    }

    return userFiltersDoc;
}

export async function trySaveUserDocument(colId: string, docId: string, document: any): Promise<boolean> {
    const accessToken = await SDK.getAccessToken();
    const extDataService = await SDK.getService<IExtensionDataService>("ms.vss-features.extension-data-service");
    const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

    document.id = docId;

    try {
        document = await dataManager.setDocument(colId, document, { scopeType: "User" });
        console.log("Update User Document: ", colId, docId, document);
        return true;
    } catch (err) {
        console.error("Error updating shared document", err);
        return false;
    }
}

export async function deleteUserDocument(colId: string, docId: string): Promise<boolean> {
    const accessToken = await SDK.getAccessToken();
    const extDataService = await SDK.getService<IExtensionDataService>("ms.vss-features.extension-data-service");
    const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

    try {
        await dataManager.deleteDocument(colId, docId, { scopeType: "User" });
        console.log("Delete User Document: ", colId, docId);
        return true;
    } catch {
        return false;
    }
}

export async function getAzdoInfo() {
    let newInfo: TenantInfo = {};

    let host = SDK.getHost()
    newInfo.organization = host.name;

    const projectInfoService = await SDK.getService<IProjectPageService>(
        "ms.vss-tfs-web.tfs-page-data-service" // TODO: CommonServiceIds.ProjectPageService
    );
    const proj = await projectInfoService.getProject();
    newInfo.project = proj?.name;

    return newInfo;
}

export async function postPullRequestStatus(
    bearerToken: string,
    tenantInfo: TenantInfo,
    repository: String,
    pullRequestId: number,
    state: string,
    description: string,
    targetUrl: string = "https://dev.azure.com/pingmint/pingmint-project/_apps/hub/pingmint.pingmint-extension.pingmint-pipeline-mergequeue"
): Promise<boolean> {
    try {
        let body = {
            "state": state,
            "description": description,
            "context": {
                "name": "merge-queue",
                "genre": "pingmint"
            },
            "targetUrl": targetUrl
        };
        let url = `https://dev.azure.com/${tenantInfo.organization}/${tenantInfo.project}/_apis/git/repositories/${repository}/pullRequests/${pullRequestId}/statuses?api-version=7.2-preview.2`;
        await postAzdo(url, body, bearerToken);
        return true;
    }
    catch (error) {
        console.error("Error posting pull request status:", error);
        return false;
    }
}

export async function getPullRequest(
    bearerToken: string,
    tenantInfo: TenantInfo,
    repository: String,
    pullRequestId: number
): Promise<any> {
    try {
        let url = `https://dev.azure.com/${tenantInfo.organization}/${tenantInfo.project}/_apis/git/repositories/${repository}/pullRequests/${pullRequestId}?api-version=7.2-preview.2`;
        let resp = await getAzdo(url, bearerToken);
        return resp;
    }
    catch (error) {
        console.error("Error fetching pull request:", error);
        return null;
    }
}

export async function getPullRequestStatuses(
    bearerToken: string,
    tenantInfo: TenantInfo,
    repository: String,
    pullRequestId: number
): Promise<any> {
    try {
        let url = `https://dev.azure.com/${tenantInfo.organization}/${tenantInfo.project}/_apis/git/repositories/${repository}/pullRequests/${pullRequestId}/statuses?api-version=7.2-preview.2`;
        let resp = await getAzdo(url, bearerToken);
        return resp;
    }
    catch (error) {
        console.error("Error fetching pull request status:", error);
        return null;
    }
}

export async function deletePullRequestStatus(
    bearerToken: string,
    tenantInfo: TenantInfo,
    repository: String,
    pullRequestId: number,
    statusId: number
): Promise<boolean> {
    try {
        let url = `https://dev.azure.com/${tenantInfo.organization}/${tenantInfo.project}/_apis/git/repositories/${repository}/pullRequests/${pullRequestId}/statuses/${statusId}?api-version=7.2-preview.2`;
        await deleteAzdo(url, bearerToken);
        return false;
    }
    catch (error) {
        console.error("Error deleting pull request status:", error);
        return false;
    }
}

export async function deletePullRequestStatuses(
    bearerToken: string,
    tenantInfo: TenantInfo,
    repository: String,
    pullRequestId: number) {
    let statuses = (await getPullRequestStatuses(bearerToken as string, tenantInfo, repository, pullRequestId))
        .value.filter((s: any) => s.context.genre == "pingmint" && s.context.name == "merge-queue");

    for (let status of statuses) {
        deletePullRequestStatus(
            bearerToken as string,
            tenantInfo,
            repository,
            pullRequestId,
            status.id)
    }
}

export interface TenantInfo {
    organization?: string;
    project?: string;
}

export interface PullRequest {
    pullRequestId?: number;
    repository?: any;
    targetRefName?: string;
    title?: string;
    isDraft?: boolean;
    creationDate?: string; // ISO date string
    autoCompleteSetBy?: boolean
}

export interface Repo {
    id?: string;
    name?: string;
    defaultBranch?: string;
}