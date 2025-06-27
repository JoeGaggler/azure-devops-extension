import * as SDK from 'azure-devops-extension-sdk';
import { type IExtensionDataService } from 'azure-devops-extension-api';

export async function getAzdo(url: string, bearertoken: string): Promise<any> {
    console.log("getapi", url);
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
        console.log(data);
        return data;
    } else {
        console.log("Error fetching azdo data", response.statusText);
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
        console.log(data);
        return data;
    } else {
        console.log("Error patching azdo data", response.statusText);
    }
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
        document = await dataManager.setDocument(colId, document);
        console.log("Update Shared Document: ", colId, docId, document);
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