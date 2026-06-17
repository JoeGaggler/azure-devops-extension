// TODO: dequeue must still update the final commit id
import React from "react";
import * as luxon from 'luxon'
import * as SDK from 'azure-devops-extension-sdk';

import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { Header } from "azure-devops-ui/Components/Header/Header";
import { IHeaderCommandBarItem } from "azure-devops-ui/Components/HeaderCommandBar/HeaderCommandBar.Props";
import { IListItemDetails, IListRow } from "azure-devops-ui/Components/List/List.Props";
import { ListItem, ScrollableList } from "azure-devops-ui/Components/List/List";
import { ListSelection } from "azure-devops-ui/Components/List/ListSelection";
import { Page } from "azure-devops-ui/Components/Page/Page";
import { TitleSize } from "azure-devops-ui/Components/Header/Header.Props";
import { Card } from "azure-devops-ui/Card";

import { getAzdoInfo, getGitClient, getExtensionManagementClient, TenantInfo, getDefaultBranchCommitId, getRefCommitId, mergeCommits } from "./azuredevops";

import { GitAsyncOperationStatus, GitPullRequestSearchCriteria, PullRequestStatus, PullRequestTimeRangeType } from "azure-devops-extension-api/Git/Git";
import { ExtensionManagementRestClient } from "azure-devops-extension-api/ExtensionManagement/ExtensionManagementClient";
import { Icon, IconSize } from "azure-devops-ui/Icon";
import { distinctBy } from "./lib";
import { GitRestClient } from "azure-devops-extension-api/Git/GitClient";
import { VssPersona } from "azure-devops-ui/Components/VssPersona/VssPersona";
import { IHostNavigationService } from "azure-devops-extension-api/Common/CommonServices";

const publisher = "pingmint";
const extensionName = "pingmint-extension";
const collectionId = "mergequeue";
const mergeQueueDocumentId = "mergequeue";
const activePullRequestsDocumentId = "activepullrequests";
const zeroCommitId = "0000000000000000000000000000000000000000";

export interface MergeQueueAppSingleton {
    bearerToken: string;
    appToken: string;
}

interface RepositoryInfo {
    id: string;
    name: string;
}

interface AuthorInfo {
    displayName: string;
    imageUrl?: string;
}

interface PullRequestInfo {
    id: number;
    title: string;
    repository: RepositoryInfo;
    author: AuthorInfo;
    createdTimestamp: number;
    sourceRefName: string;
}

interface MergeQueueItemInfo {
    id: number;
    title: string;
    repository: RepositoryInfo;
    author: AuthorInfo;
    status: MergeQueueStatus;
    createdTimestamp: number;
    sourceRefName: string;
    sourceCommitId: string;
    targetCommitId: string;
    mergedCommitId: string;
}

type MergeQueueStatus =
    "queued" | // requires recalculation
    "recalculating" | // currently being recalculated
    "valid"; // can be merged

interface PullRequestDocument {
    id: string;
    __etag: number;
    pullRequests: PullRequestInfo[];
}

interface MergeQueueDocument {
    id: string;
    __etag: number;
    mergeQueueItems: MergeQueueItemInfo[];
}

interface ReducerState {
    mergeQueueItems: MergeQueueItemInfo[];
    selectedMergeQueuePullRequestIds: number[];
    mergeQueuePullRequests: PullRequestInfo[];

    activePullRequests: PullRequestInfo[];
    selectedActivePullRequestIds: number[];
}

interface ReducerAction {
    // collection loading
    mergeQueueItems?: MergeQueueItemInfo[];
    activePullRequests?: PullRequestInfo[];

    // selection changes
    selectedMergeQueuePullRequestIds?: number[];
    selectedActivePullRequestIds?: number[];

    // actions
    enqueuePullRequests?: PullRequestInfo[];
}

function applySelections<T, TItem>(listSelection: ListSelection, all: T[], accessor: (t: T) => TItem, ids: TItem[]) {
    listSelection.clear();
    for (let id of ids) {
        let idx = all.findIndex((item) => accessor(item) === id);
        if (idx < 0) { continue; }
        listSelection.select(idx, 1, true, true);
    }
}

function reducer(state: ReducerState, action: ReducerAction): ReducerState {
    let next = { ...state };

    if (action.selectedMergeQueuePullRequestIds) {
        console.log("MQ: reducer -> updating selected merge queue pull request IDs", action.selectedMergeQueuePullRequestIds);
        next.selectedMergeQueuePullRequestIds = action.selectedMergeQueuePullRequestIds;
    }

    if (action.selectedActivePullRequestIds) {
        console.log("MQ: reducer -> updating selected active pull request IDs", action.selectedActivePullRequestIds);
        next.selectedActivePullRequestIds = action.selectedActivePullRequestIds;
    }

    if (action.mergeQueueItems) {
        console.log("MQ: reducer -> updating merge queue items", action.mergeQueueItems);
        next.mergeQueueItems = action.mergeQueueItems;
        next.mergeQueuePullRequests = action.mergeQueueItems.map((item): PullRequestInfo => {
            return {
                id: item.id,
                title: item.title,
                repository: item.repository,
                sourceRefName: item.sourceRefName,
                createdTimestamp: item.createdTimestamp,
                author: item.author,
            };
        });
        // TODO: confirm selections
    }

    if (action.activePullRequests) {
        console.log("MQ: reducer -> updating active pull requests", action.activePullRequests);
        next.activePullRequests = action.activePullRequests.sort((a, b) => b.createdTimestamp - a.createdTimestamp);
        // TODO: confirm selections
    }

    return next;
}

async function getDocument(client: ExtensionManagementRestClient, collectionId: string, documentId: string): Promise<any | undefined> {
    try {
        return await client.getDocumentByName(publisher, extensionName, "Default", "Current", collectionId, documentId);
    } catch (error) {
        console.error("MQ: getDocument -> error occurred", error);
        return undefined;
    }
}

async function updateDocument(client: ExtensionManagementRestClient, collectionId: string, documentId: string, document: any): Promise<any | undefined> {
    try {
        document.id = documentId;
        return await client.updateDocumentByName(document, publisher, extensionName, "Default", "Current", collectionId);
    } catch (error) {
        console.error("MQ: updateDocument -> error occurred", error);
        return undefined;
    }
}

async function createDocument(client: ExtensionManagementRestClient, collectionId: string, documentId: string, document: any): Promise<any | undefined> {
    try {
        document.id = documentId;
        return await client.createDocumentByName(document, publisher, extensionName, "Default", "Current", collectionId);
    } catch (error) {
        console.error("MQ: createDocument -> error occurred", error);
        return undefined;
    }
}

export function MergeQueueApp(p: { singleton: MergeQueueAppSingleton }) {
    let tenantInfo = React.useRef<TenantInfo>();
    let singleton = React.useRef(p.singleton);
    let gitClient = React.useRef(getGitClient());
    let extensionManagementClient = React.useRef(getExtensionManagementClient());
    let isMergeQueueRunning = React.useRef(false);

    const [state, dispatch] = React.useReducer<(state: ReducerState, action: ReducerAction) => ReducerState>(reducer, {
        mergeQueueItems: [],
        mergeQueuePullRequests: [],
        selectedMergeQueuePullRequestIds: [],

        activePullRequests: [],
        selectedActivePullRequestIds: [],
    })

    // initialize the app
    React.useEffect(() => { init() }, []);
    async function init() {
        try {
            console.log("MQ: init");
            let nextTenantInfo = await getAzdoInfo();
            if (!nextTenantInfo) {
                // TODO: lock the app
                console.error("Failed to get Azure DevOps info");
                return;
            }
            tenantInfo.current = nextTenantInfo;

            let em = getExtensionManagementClient();
            extensionManagementClient.current = em;

            // TODO: REMOVE THIS
            // await extensionManagementClient.current.deleteDocumentByName(
            //     publisher,
            //     extensionName,
            //     "Default",
            //     "Current",
            //     collectionId,
            //     mergeQueueDocumentId
            // );

            var mdoc = await getMergeQueueDocument();
            if (mdoc) {
                dispatch({ mergeQueueItems: mdoc.mergeQueueItems });
            }

            var adoc = await getActivePullRequestDocument();
            if (adoc) {
                dispatch({ activePullRequests: adoc.pullRequests });
            }

            console.log("MQ: init -> done");
        } catch (error) {
            console.error("MQ: init -> error occurred", error);
        }
    }

    // ticktock
    React.useEffect(() => {
        let id = setInterval(() => { ticktock(); }, 10000);
        return () => clearInterval(id);
    }, []);
    async function ticktock() {
        try {
            if (!singleton.current) {
                console.error("MQ: ticktock -> no singleton available");
                return;
            }

            let ti = tenantInfo.current;
            if (!ti) {
                console.error("MQ: ticktock -> no tenant info available");
                return;
            }

            let proj = ti.project;
            if (!proj) {
                console.error("MQ: ticktock -> no project available");
                return;
            }

            let git = gitClient.current;
            if (!git) {
                console.error("MQ: ticktock -> no git client available");
                return;
            }

            // TODO: run concurrently
            await refreshActivePullRequests(git, ti);
            await runMergeQueue(git, proj);
        } catch (error) {
            console.error("MQ: ticktock -> error occurred", error);
        }
    }

    async function refreshActivePullRequests(gitClient: GitRestClient, tenantInfo: TenantInfo): Promise<void> {
        let criteria: GitPullRequestSearchCriteria = {
            creatorId: undefined!,
            includeLinks: false,
            maxTime: undefined!,
            minTime: undefined!,
            queryTimeRangeType: PullRequestTimeRangeType.Created,
            repositoryId: undefined!,
            reviewerId: undefined!,
            sourceRefName: undefined!,
            sourceRepositoryId: undefined!,
            status: PullRequestStatus.Active,
            targetRefName: undefined!,
            title: undefined!
        };
        let gitPRs = await gitClient.getPullRequestsByProject(tenantInfo.project, criteria, undefined, undefined, undefined);
        let newPullRequests: PullRequestInfo[] = [];
        for (const gitPR of gitPRs) {
            const repoId = gitPR.repository?.id;
            if (!repoId) { continue; }

            const repoName = gitPR.repository?.name;
            if (!repoName) { continue; }

            const sourceRefName = gitPR.sourceRefName;
            if (!sourceRefName) { continue; }

            newPullRequests.push({
                id: gitPR.pullRequestId,
                title: gitPR.title,
                createdTimestamp: luxon.DateTime.fromJSDate(gitPR.creationDate).toUnixInteger(),
                sourceRefName: gitPR.sourceRefName,
                repository: {
                    id: gitPR.repository.id,
                    name: gitPR.repository.name
                },
                author: {
                    displayName: gitPR.createdBy?.displayName || "Unknown User",
                    imageUrl: gitPR.createdBy?.imageUrl
                }
            });
        }

        dispatch({ activePullRequests: newPullRequests });
        // console.log("MQ: refreshActivePullRequests -> updated pull requests", newPullRequests);

        let adoc = await getActivePullRequestDocument();
        if (adoc) {
            adoc.pullRequests = newPullRequests;
            let adoc2 = await updateActivePullRequestDocument(adoc);
            if (adoc2) {
                // console.log("MQ: refreshActivePullRequests -> updated active pull request document", adoc2);
            }
        }
    }

    async function runMergeQueue(gitClient: GitRestClient, project: string): Promise<void> {
        if (isMergeQueueRunning.current === true) {
            console.log("MQ: merge queue is already running");
            return;
        }
        console.log("MQ: running merge queue");
        try {
            isMergeQueueRunning.current = true;

            var old_mqdoc = await getMergeQueueDocument();
            if (!old_mqdoc) {
                console.error("MQ: runMergeQueue -> failed to get merge queue document");
                return;
            }

            // get unique repositories referenced by the merge queue items
            const old_mqitems = [...old_mqdoc.mergeQueueItems];
            dispatch({ mergeQueueItems: old_mqdoc.mergeQueueItems });
            const repos = old_mqitems.map(i => i.repository);
            const repoSet = distinctBy(repos, r => r.id);
            const gitRepos = await Promise.all(repoSet.map(r => gitClient.getRepository(r.id, project)));
            console.log("MQ: runMergeQueue -> retrieved git repositories", gitRepos);

            let repoBaseCommits: { repoId: string; baseCommitId: string }[] = [];
            for (const repo of gitRepos) {
                const commitId = await getDefaultBranchCommitId(gitClient, project, repo.id);
                if (!commitId) {
                    console.error("MQ: runMergeQueue -> failed to get default branch commit ID for repository", repo.id);
                    continue;
                }
                repoBaseCommits.push({ repoId: repo.id, baseCommitId: commitId });
            }

            async function syncDocument(doc: MergeQueueDocument, items: MergeQueueItemInfo[]): Promise<MergeQueueDocument | undefined> {
                doc.mergeQueueItems = items;
                var doc2 = await updateDocument(extensionManagementClient.current, collectionId, mergeQueueDocumentId, doc);
                if (!doc2 || !doc2.mergeQueueItems) {
                    return undefined;
                }
                // console.log("MQ: runMergeQueue -> updated merge queue document", doc2);
                dispatch({ mergeQueueItems: doc2.mergeQueueItems });
                return doc2;
            }

            let new_mqitems: MergeQueueItemInfo[] = [];
            for (const [index, old_mqitem] of old_mqitems.entries()) {
                // calculate target commit
                const repoid = old_mqitem.repository.id;
                const targetCommitEntry = repoBaseCommits.find(r => r.repoId === repoid);
                const targetCommitId = targetCommitEntry?.baseCommitId;
                if (!targetCommitId) {
                    console.error("MQ: runMergeQueue -> failed to find base commit for repository", repoid);
                    new_mqitems.push(old_mqitem);
                    continue;
                }

                // calculate source commit
                const sourceRefName = old_mqitem.sourceRefName;
                const sourceCommitId = await getRefCommitId(gitClient, project, repoid, sourceRefName);
                if (!sourceCommitId) {
                    console.error("MQ: runMergeQueue -> failed to get source commit ID for repository", repoid, sourceRefName);
                    new_mqitems.push(old_mqitem);
                    continue;
                }

                // skip item if no changes are needed
                const status = old_mqitem.status ?? 'queued';
                const isQueuedStatus = status === 'queued' || status === 'recalculating';
                const isSameSourceCommit = sourceCommitId === old_mqitem.sourceCommitId;
                const isSameTargetCommit = targetCommitId === old_mqitem.targetCommitId;
                const needsMergedCommit = (old_mqitem.mergedCommitId || zeroCommitId) === zeroCommitId;
                if (isQueuedStatus || needsMergedCommit) { console.log(`MQ: runMergeQueue -> ${index}: queued`); }
                else if (!isSameSourceCommit) { console.log(`MQ: runMergeQueue -> ${index}: new source commit`, sourceCommitId, old_mqitem.sourceCommitId); }
                else if (!isSameTargetCommit) { console.log(`MQ: runMergeQueue -> ${index}: new target commit`, targetCommitId, old_mqitem.targetCommitId); }
                else {
                    // skip!
                    console.log(`MQ: runMergeQueue -> ${index}: same commits`);
                    new_mqitems.push(old_mqitem);
                    targetCommitEntry.baseCommitId = old_mqitem.mergedCommitId;
                    continue;
                }

                // invalidate subsequent merge queue items in the same repository
                invalidateDependentPullRequests(old_mqitems, index, repoid);

                // checkout item
                console.log(`MQ: runMergeQueue -> ${index}: recalculating`, sourceCommitId, targetCommitId);
                const new_mqitem: MergeQueueItemInfo = {
                    ...old_mqitem,
                    status: "recalculating",
                    sourceCommitId: sourceCommitId,
                    targetCommitId: targetCommitId,
                };
                old_mqitems[index] = new_mqitem; // TODO: progressive update document
                old_mqdoc = await syncDocument(old_mqdoc, old_mqitems);
                if (!old_mqdoc) {
                    console.error(`MQ: runMergeQueue -> ${index}: failed to sync document`);
                    return;
                }

                // merge request
                let mergeResult = await mergeCommits(gitClient, project, repoid, sourceCommitId, targetCommitId);
                if (!mergeResult) {
                    console.error(`MQ: runMergeQueue -> ${index}: failed to merge commits`);
                    return;
                }
                let mergeResultStatus = mergeResult.status;
                if (mergeResultStatus === GitAsyncOperationStatus.Completed) {
                    let mergedCommitId = mergeResult.detailedStatus.mergeCommitId;
                    new_mqitem.status = 'valid'; // TODO: check for conflicts
                    new_mqitem.mergedCommitId = mergedCommitId;
                    targetCommitEntry.baseCommitId = mergedCommitId;
                } else {
                    invalidateDependentPullRequests(old_mqitems, index, repoid);
                    return;
                }

                // append and sync
                new_mqitems.push(new_mqitem);
                old_mqitems[index] = new_mqitem; // TODO: progressive update document
                old_mqdoc = await syncDocument(old_mqdoc, old_mqitems);
                if (!old_mqdoc) {
                    console.error(`MQ: runMergeQueue -> ${index}: failed to sync document`);
                    return;
                }
            }

            // dispatch({ mergeQueueItems: new_mqitems }); // complete update
            // old_mqdoc.mergeQueueItems = new_mqitems;
            // if (!(await syncDocument())) {
            //     console.error(`MQ: runMergeQueue -> end: failed to sync document`);
            //     return;
            // }
        } catch (error) {
            console.error("MQ: runMergeQueue -> error occurred", error);
        } finally {
            console.log("MQ: runMergeQueue -> finished");
            isMergeQueueRunning.current = false;
        }
    }

    async function getAppDocument<T>(appDocumentId: string): Promise<T | undefined> {
        let doc = await getDocument(extensionManagementClient.current, collectionId, appDocumentId);
        if (!doc) {
            doc = await createDocument(extensionManagementClient.current, collectionId, appDocumentId, {});
            if (!doc) {
                console.error("Failed to create app document");
                return undefined;
            }
        }
        if (doc.id !== appDocumentId) {
            console.error("Wrong document ID", doc.id);
            return undefined;
        }
        if (!doc.__etag) {
            console.error("Failed to get document etag");
            return undefined;
        }
        return doc as T;
    }

    async function getMergeQueueDocument(): Promise<MergeQueueDocument | undefined> {
        return await getAppDocument<MergeQueueDocument>(mergeQueueDocumentId);
    }

    async function getActivePullRequestDocument(): Promise<PullRequestDocument | undefined> {
        return await getAppDocument<PullRequestDocument>(activePullRequestsDocumentId);
    }

    async function updateActivePullRequestDocument(doc: PullRequestDocument): Promise<PullRequestDocument | undefined> {
        return await updateDocument(extensionManagementClient.current, collectionId, activePullRequestsDocumentId, doc);
    }

    async function updateMergeQueueDocument(doc: MergeQueueDocument): Promise<MergeQueueDocument | undefined> {
        return await updateDocument(extensionManagementClient.current, collectionId, mergeQueueDocumentId, doc);
    }

    function renderPageCommandBarItems(): IHeaderCommandBarItem[] {
        // TODO: return page command bar items
        return [];
    }

    function renderMergeQueueCommandBarItems(): IHeaderCommandBarItem[] {
        let hasSelection = state.selectedMergeQueuePullRequestIds.length === 1;
        if (hasSelection) {
        }
        return [
            {
                id: "promote",
                iconProps: {
                    iconName: "Up"
                },
                onActivate: () => { onPromotePullRequest(); },
                isPrimary: false,
                important: true,
                disabled: (state.selectedMergeQueuePullRequestIds.length === 0) // TODO: not at top
            },
            {
                id: "demote",
                iconProps: {
                    iconName: "Down"
                },
                onActivate: () => { onDemotePullRequest(); },
                isPrimary: false,
                important: true,
                disabled: (state.selectedMergeQueuePullRequestIds.length === 0) // TODO: not at bottom
            },
            {
                id: "dequeue",
                text: "Dequeue",
                onActivate: () => { onDequeuePullRequest(); },
                isPrimary: true,
                important: true,
                disabled: (state.selectedMergeQueuePullRequestIds.length === 0)
            }
        ];
    }

    function renderAllPullRequestsCommandBarItems(): IHeaderCommandBarItem[] {
        return [
            {
                id: "enqueue",
                text: "Enqueue",
                onActivate: () => { onEnqueuePullRequest(); },
                isPrimary: true,
                important: true,
                disabled: (state.selectedActivePullRequestIds.length === 0)
            }
        ];
    }

    function invalidateDependentPullRequests(array: MergeQueueItemInfo[], index: number, repoid: string) {
        for (let i2 = index + 1; i2 < array.length; i2++) {
            let itr_mqitem = array[i2];
            if (itr_mqitem.repository.id !== repoid) { continue; }

            console.log(`MQ: runMergeQueue -> ${i2}: invalidating`);
            array[i2] = {
                ...array[i2],
                status: 'queued',
            };
        }
    }

    async function onPromotePullRequest() {
        if (!state.selectedMergeQueuePullRequestIds || state.selectedMergeQueuePullRequestIds.length !== 1) {
            // TODO: toast error
            return;
        }
        let mq_items = state.mergeQueueItems;
        let selectedId = state.selectedMergeQueuePullRequestIds[0];
        let selectedIndex = mq_items.findIndex(m => m.id === selectedId);
        if (selectedIndex === -1) {
            // TODO: toast error
            return;
        }

        // get remote doc
        let mdoc = await getMergeQueueDocument();
        if (!mdoc) {
            console.error("Failed to get merge queue document");
            return;
        }
        mq_items = mdoc.mergeQueueItems || [];
        dispatch({ mergeQueueItems: mq_items });
        if (selectedIndex !== mq_items.findIndex(m => m.id === selectedId)) {
            // TODO: toast error
            return;
        }
        if (selectedIndex === 0) {
            // TODO: toast error
            return;
        }

        // swap
        let tmp = mq_items[selectedIndex];
        mq_items[selectedIndex] = mq_items[selectedIndex - 1];
        mq_items[selectedIndex - 1] = tmp;
        invalidateDependentPullRequests(mq_items, selectedIndex - 1, tmp.repository.id);

        // sync
        let updatedMdoc = await updateMergeQueueDocument(mdoc);
        if (!updatedMdoc) {
            console.error("MQ: onDequeuePullRequest -> failed to update merge queue document");
            return;
        }
        dispatch({ mergeQueueItems: updatedMdoc.mergeQueueItems });

        await ticktock(); // immediate refresh
    }

    async function onDemotePullRequest() {
        if (!state.selectedMergeQueuePullRequestIds || state.selectedMergeQueuePullRequestIds.length !== 1) {
            // TODO: toast error
            return;
        }
        let mq_items = state.mergeQueueItems;
        let selectedId = state.selectedMergeQueuePullRequestIds[0];
        let selectedIndex = mq_items.findIndex(m => m.id === selectedId);
        if (selectedIndex === -1) {
            // TODO: toast error
            return;
        }

        // get remote doc
        let mdoc = await getMergeQueueDocument();
        if (!mdoc) {
            console.error("Failed to get merge queue document");
            return;
        }
        mq_items = mdoc.mergeQueueItems || [];
        dispatch({ mergeQueueItems: mq_items });
        if (selectedIndex !== mq_items.findIndex(m => m.id === selectedId)) {
            // TODO: toast error
            return;
        }
        if (selectedIndex === mq_items.length - 1) {
            // TODO: toast error
            return;
        }

        // swap
        let tmp = mq_items[selectedIndex];
        mq_items[selectedIndex] = mq_items[selectedIndex + 1];
        mq_items[selectedIndex + 1] = tmp;
        invalidateDependentPullRequests(mq_items, selectedIndex, tmp.repository.id);

        // sync
        let updatedMdoc = await updateMergeQueueDocument(mdoc);
        if (!updatedMdoc) {
            console.error("MQ: onDequeuePullRequest -> failed to update merge queue document");
            return;
        }
        dispatch({ mergeQueueItems: updatedMdoc.mergeQueueItems });

        await ticktock(); // immediate refresh
    }

    async function onDequeuePullRequest() {
        let adoc = await getActivePullRequestDocument();
        if (!adoc) {
            console.error("Failed to get active pull request document");
            return;
        }
        let aprs = adoc.pullRequests || [];
        dispatch({ activePullRequests: aprs });

        let mdoc = await getMergeQueueDocument();
        if (!mdoc) {
            console.error("Failed to get merge queue document");
            return;
        }
        let mitems = mdoc.mergeQueueItems || [];
        dispatch({ mergeQueueItems: mitems });

        let oldids = state.selectedMergeQueuePullRequestIds;
        let newMergeQueueItems = mitems.filter(m => !oldids.includes(m.id));
        console.log("MQ: onDequeuePullRequest -> new pull requests", newMergeQueueItems);

        mdoc.mergeQueueItems = [...newMergeQueueItems];
        let updatedMdoc = await updateMergeQueueDocument(mdoc);
        if (!updatedMdoc) {
            console.error("MQ: onDequeuePullRequest -> failed to update merge queue document");
            return;
        }
        console.log("MQ: onDequeuePullRequest -> updated merge queue document", updatedMdoc);
        dispatch({ mergeQueueItems: updatedMdoc.mergeQueueItems });

        await ticktock(); // immediate refresh
    }

    async function onEnqueuePullRequest() {
        let adoc = await getActivePullRequestDocument();
        if (!adoc) {
            console.error("Failed to get active pull request document");
            return;
        }
        let aprs = adoc.pullRequests || [];
        dispatch({ activePullRequests: aprs });

        let mdoc = await getMergeQueueDocument();
        if (!mdoc) {
            console.error("Failed to get merge queue document");
            return;
        }
        let mitems = mdoc.mergeQueueItems || [];
        dispatch({ mergeQueueItems: mitems });

        let nextids = state.selectedActivePullRequestIds;
        let nextprs = aprs.filter(pr => nextids.includes(pr.id));
        console.log("MQ: onEnqueuePullRequest -> next pull requests", nextids, nextprs);

        // exclude nextprs that are already in the merge queue
        let filteredprs = nextprs.filter(pr => !mitems.some(mpr => mpr.id === pr.id));
        console.log("MQ: onEnqueuePullRequest -> filtered pull requests", filteredprs);
        if (filteredprs.length === 0) {
            console.warn("MQ: onEnqueuePullRequest -> no pull requests to enqueue");
            return;
        }

        var newMergeQueueItems = filteredprs.map((pr): MergeQueueItemInfo => {
            return {
                id: pr.id,
                title: pr.title,
                repository: pr.repository,
                author: pr.author,
                status: "queued",
                createdTimestamp: pr.createdTimestamp,
                sourceRefName: pr.sourceRefName,
                sourceCommitId: zeroCommitId,
                targetCommitId: zeroCommitId,
                mergedCommitId: zeroCommitId,
            };
        });

        mdoc.mergeQueueItems = [...mitems, ...newMergeQueueItems];
        let updatedMdoc = await updateMergeQueueDocument(mdoc);
        if (!updatedMdoc) {
            console.error("MQ: onEnqueuePullRequest -> failed to update merge queue document");
            return;
        }
        console.log("MQ: onEnqueuePullRequest -> updated merge queue document", updatedMdoc);
        dispatch({ mergeQueueItems: updatedMdoc.mergeQueueItems });

        await ticktock(); // immediate refresh
    }

    function onSelectMergeQueuePullRequestIds(ids: number[]) {
        dispatch({ selectedMergeQueuePullRequestIds: ids });
    }

    function onSelectActivePullRequestIds(ids: number[]) {
        dispatch({ selectedActivePullRequestIds: ids });
    }

    function onActivateMergeQueuePullRequest(id: number, repo: string) {
        activatePullRequest(id, repo);
    }

    function onActivateActivePullRequest(id: number, repo: string) {
        activatePullRequest(id, repo);
    }

    async function activatePullRequest(id: number, repo: string) {
        let ten = tenantInfo.current; if (!ten) { return; }
        let org = ten.organization; if (!org) { return; }
        let proj = ten.project; if (!proj) { return; }

        const navService = await SDK.getService<IHostNavigationService>("ms.vss-features.host-navigation-service");
        let url = `https://dev.azure.com/${org}/${proj}/_git/${repo}/pullrequest/${id}`;
        console.log("url: ", url);
        navService.openNewWindow(url, "");
    }

    function mapMergeQueueItemToPullRequestListItems(): PullRequestListItem[] {
        return state.mergeQueueItems.map((item): PullRequestListItem => {
            let status = item.status;
            let icon = "Starburst";
            if (status === "queued") {
                icon = "CircleRing";
            } else if (status === "recalculating") {
                icon = "WorkFlow";
            } else {
                icon = "Starburst";
            }

            let dateString = item.createdTimestamp ? luxon.DateTime.fromSeconds(item.createdTimestamp).toRelative() || undefined : undefined;

            return {
                icon: icon,
                pullRequestId: item.id,
                repository: item.repository.name,
                author: item.author,
                title: `${item.title}`,// - ${item.sourceCommitId} onto ${item.targetCommitId} is ${item.mergedCommitId}`,
                dateString: dateString,
            };
        });
    }

    function mapActivePullRequestsToPullRequestListItems(): PullRequestListItem[] {
        return state.activePullRequests.map((pr): PullRequestListItem => {
            let dateString = pr.createdTimestamp ? luxon.DateTime.fromSeconds(pr.createdTimestamp).toRelative() || undefined : undefined;

            return {
                icon: "CircleRing",
                pullRequestId: pr.id,
                repository: pr.repository.name,
                author: pr.author,
                title: pr.title,
                dateString: dateString
            };
        });
    }

    return (
        <Page className="">
            <Header
                title="Merge Queue Beta __MERGEQUEUEVERSION__ "
                titleSize={TitleSize.Large}
                commandBarItems={renderPageCommandBarItems()}
            />

            <Card
                className="padding-8 margin-8"
                contentProps={{ contentPadding: false }}
                titleProps={{ text: "Merge Queue", className: "", size: TitleSize.Medium }}
                headerClassName=""
                headerCommandBarItems={renderMergeQueueCommandBarItems()}
            >
                <PullRequestList
                    pullRequests={mapMergeQueueItemToPullRequestListItems()}
                    selectedIds={state.selectedMergeQueuePullRequestIds}
                    onSelectPullRequestIds={onSelectMergeQueuePullRequestIds}
                    onActivatePullRequest={onActivateMergeQueuePullRequest}
                />
            </Card>

            <Card
                className="padding-8 margin-8"
                contentProps={{ contentPadding: false }}
                titleProps={{ text: "All Pull Requests", className: "", size: TitleSize.Medium }}
                headerClassName=""
                headerCommandBarItems={renderAllPullRequestsCommandBarItems()}
            >
                <PullRequestList
                    pullRequests={mapActivePullRequestsToPullRequestListItems()}
                    selectedIds={state.selectedActivePullRequestIds}
                    onSelectPullRequestIds={onSelectActivePullRequestIds}
                    onActivatePullRequest={onActivateActivePullRequest}
                />
            </Card>

            <div className="text-neutral-30 flex-row padding-4">
                <div className="flex-grow"></div>
                <div>__MERGEQUEUEVERSION__</div>
            </div>

            <div>
                {AllIcons()}
            </div>
        </Page>
    );
}

export interface PullRequestListItem {
    pullRequestId: number;
    repository: string;
    title: string;
    icon: string;
    author: AuthorInfo;
    dateString?: string;
}

export interface PullRequestListProps {
    pullRequests: PullRequestListItem[];
    selectedIds: number[];
    onSelectPullRequestIds: (id: number[]) => void;
    onActivatePullRequest?: (id: number, repo: string) => void;
}

export function PullRequestList({ pullRequests, selectedIds, onSelectPullRequestIds, onActivatePullRequest }: PullRequestListProps) {
    let listSelection = new ListSelection(true);

    applySelections(listSelection, pullRequests, (pr) => pr.pullRequestId, selectedIds);

    function onSelectRow(row: IListRow<PullRequestListItem>) {
        console.log("NextRunTab -> targetPipelineSelect", row);
        onSelectPullRequestIds([row.data.pullRequestId]);
    }

    function onActivateRow(row: IListRow<PullRequestListItem>) {
        console.log("NextRunTab -> targetPipelineActivate", row);
        if (onActivatePullRequest) {
            onActivatePullRequest(row.data.pullRequestId, row.data.repository);
        }
    }

    function renderRow(
        index: number,
        pullRequest: PullRequestListItem,
        details: IListItemDetails<PullRequestListItem>,
        key?: string
    ): JSX.Element {
        if (!pullRequest) { return <></> }
        let extra = "";
        let className = `scroll-hidden flex-row flex-center rhythm-horizontal-8 flex-grow padding-4 ${extra}`;

        let initialsIdentityProvider = {
            getDisplayName() {
                return pullRequest.author?.displayName || "?";
            },
            getIdentityImageUrl(_size: number) {
                return pullRequest.author?.imageUrl || undefined;
            }
        }

        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}
            >
                <div className={className}>
                    <Icon iconName={pullRequest.icon} size={IconSize.medium} />
                    <div className="font-size-m flex-row flex-center flex-shrink">{pullRequest.pullRequestId}</div>
                    <VssPersona size={"extra-small"} identityDetailsProvider={initialsIdentityProvider} />
                    <div className="font-size-m">{pullRequest.repository}</div>
                    <div className="font-size-m italic text-neutral-70 text-ellipsis">{pullRequest.title}</div>
                    <div className="font-size-m flex-row flex-center flex-grow rhythm-horizontal-8">
                        <div className="flex-grow" />
                        <div>{pullRequest.dateString}</div>
                    </div>
                </div>
            </ListItem>
        )
    }

    return <>
        <div className="flex-column">
            <ScrollableList
                itemProvider={new ArrayItemProvider(pullRequests || [])}
                selection={listSelection}
                onSelect={(_evt, listRow) => { onSelectRow(listRow); }}
                onActivate={(_evt, listRow) => { onActivateRow(listRow); }}
                renderRow={renderRow}
                width="100%"
            />
        </div>
    </>
}

export function AllIcons() {
    return (
        <div className="flex-column rhythm-vertical-16">
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Accept"} size={IconSize.large} tooltipProps={{ text: "Accept" }} />
                <Icon iconName={"AccountManagement"} size={IconSize.large} tooltipProps={{ text: "AccountManagement" }} />
                <Icon iconName={"Accounts"} size={IconSize.large} tooltipProps={{ text: "Accounts" }} />
                <Icon iconName={"ActivateOrders"} size={IconSize.large} tooltipProps={{ text: "ActivateOrders" }} />
                <Icon iconName={"ActivityFeed"} size={IconSize.large} tooltipProps={{ text: "ActivityFeed" }} />
                <Icon iconName={"Add"} size={IconSize.large} tooltipProps={{ text: "Add" }} />
                <Icon iconName={"AddFriend"} size={IconSize.large} tooltipProps={{ text: "AddFriend" }} />
                <Icon iconName={"AddGroup"} size={IconSize.large} tooltipProps={{ text: "AddGroup" }} />
                <Icon iconName={"AddReaction"} size={IconSize.large} tooltipProps={{ text: "AddReaction" }} />
                <Icon iconName={"AddTo"} size={IconSize.large} tooltipProps={{ text: "AddTo" }} />
                <Icon iconName={"Airplane"} size={IconSize.large} tooltipProps={{ text: "Airplane" }} />
                <Icon iconName={"AirplaneSolid"} size={IconSize.large} tooltipProps={{ text: "AirplaneSolid" }} />
                <Icon iconName={"AlertSolid"} size={IconSize.large} tooltipProps={{ text: "AlertSolid" }} />
                <Icon iconName={"AlignJustify"} size={IconSize.large} tooltipProps={{ text: "AlignJustify" }} />
                <Icon iconName={"AnalyticsView"} size={IconSize.large} tooltipProps={{ text: "AnalyticsView" }} />
                <Icon iconName={"AppIconDefault"} size={IconSize.large} tooltipProps={{ text: "AppIconDefault" }} />
                <Icon iconName={"ArrowDownRightMirrored8"} size={IconSize.large} tooltipProps={{ text: "ArrowDownRightMirrored8" }} />
                <Icon iconName={"ArrowTallUpRight"} size={IconSize.large} tooltipProps={{ text: "ArrowTallUpRight" }} />
                <Icon iconName={"ArrowUpRight8"} size={IconSize.large} tooltipProps={{ text: "ArrowUpRight8" }} />
                <Icon iconName={"Ascending"} size={IconSize.large} tooltipProps={{ text: "Ascending" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Assign"} size={IconSize.large} tooltipProps={{ text: "Assign" }} />
                <Icon iconName={"AsteriskSolid"} size={IconSize.large} tooltipProps={{ text: "AsteriskSolid" }} />
                <Icon iconName={"Attach"} size={IconSize.large} tooltipProps={{ text: "Attach" }} />
                <Icon iconName={"AwayStatus"} size={IconSize.large} tooltipProps={{ text: "AwayStatus" }} />
                <Icon iconName={"Back"} size={IconSize.large} tooltipProps={{ text: "Back" }} />
                <Icon iconName={"Backlog"} size={IconSize.large} tooltipProps={{ text: "Backlog" }} />
                <Icon iconName={"BacklogBoard"} size={IconSize.large} tooltipProps={{ text: "BacklogBoard" }} />
                <Icon iconName={"BacklogList"} size={IconSize.large} tooltipProps={{ text: "BacklogList" }} />
                <Icon iconName={"BackToWindow"} size={IconSize.large} tooltipProps={{ text: "BackToWindow" }} />
                <Icon iconName={"BankSolid"} size={IconSize.large} tooltipProps={{ text: "BankSolid" }} />
                <Icon iconName={"BlockContact"} size={IconSize.large} tooltipProps={{ text: "BlockContact" }} />
                <Icon iconName={"Blocked"} size={IconSize.large} tooltipProps={{ text: "Blocked" }} />
                <Icon iconName={"Blocked2"} size={IconSize.large} tooltipProps={{ text: "Blocked2" }} />
                <Icon iconName={"Blocked2Solid"} size={IconSize.large} tooltipProps={{ text: "Blocked2Solid" }} />
                <Icon iconName={"BlockedSite"} size={IconSize.large} tooltipProps={{ text: "BlockedSite" }} />
                <Icon iconName={"BlockedSolid"} size={IconSize.large} tooltipProps={{ text: "BlockedSolid" }} />
                <Icon iconName={"Bold"} size={IconSize.large} tooltipProps={{ text: "Bold" }} />
                <Icon iconName={"BranchCompare"} size={IconSize.large} tooltipProps={{ text: "BranchCompare" }} />
                <Icon iconName={"BranchFork2"} size={IconSize.large} tooltipProps={{ text: "BranchFork2" }} />
                <Icon iconName={"BranchMerge"} size={IconSize.large} tooltipProps={{ text: "BranchMerge" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"BranchMerged"} size={IconSize.large} tooltipProps={{ text: "BranchMerged" }} />
                <Icon iconName={"BranchPullRequest"} size={IconSize.large} tooltipProps={{ text: "BranchPullRequest" }} />
                <Icon iconName={"BranchRequestClosed"} size={IconSize.large} tooltipProps={{ text: "BranchRequestClosed" }} />
                <Icon iconName={"BranchRequestDraft"} size={IconSize.large} tooltipProps={{ text: "BranchRequestDraft" }} />
                <Icon iconName={"BuildQueue"} size={IconSize.large} tooltipProps={{ text: "BuildQueue" }} />
                <Icon iconName={"BuildQueueNew"} size={IconSize.large} tooltipProps={{ text: "BuildQueueNew" }} />
                <Icon iconName={"BulletedList"} size={IconSize.large} tooltipProps={{ text: "BulletedList" }} />
                <Icon iconName={"CalculatorAddition"} size={IconSize.large} tooltipProps={{ text: "CalculatorAddition" }} />
                <Icon iconName={"Calendar"} size={IconSize.large} tooltipProps={{ text: "Calendar" }} />
                <Icon iconName={"Camera"} size={IconSize.large} tooltipProps={{ text: "Camera" }} />
                <Icon iconName={"Cancel"} size={IconSize.large} tooltipProps={{ text: "Cancel" }} />
                <Icon iconName={"CannedChat"} size={IconSize.large} tooltipProps={{ text: "CannedChat" }} />
                <Icon iconName={"Car"} size={IconSize.large} tooltipProps={{ text: "Car" }} />
                <Icon iconName={"CaretSolidDown"} size={IconSize.large} tooltipProps={{ text: "CaretSolidDown" }} />
                <Icon iconName={"Certificate"} size={IconSize.large} tooltipProps={{ text: "Certificate" }} />
                <Icon iconName={"Chart"} size={IconSize.large} tooltipProps={{ text: "Chart" }} />
                <Icon iconName={"ChartSeries"} size={IconSize.large} tooltipProps={{ text: "ChartSeries" }} />
                <Icon iconName={"Chat"} size={IconSize.large} tooltipProps={{ text: "Chat" }} />
                <Icon iconName={"ChatInviteFriend"} size={IconSize.large} tooltipProps={{ text: "ChatInviteFriend" }} />
                <Icon iconName={"CheckboxComposite"} size={IconSize.large} tooltipProps={{ text: "CheckboxComposite" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"CheckboxCompositeReversed"} size={IconSize.large} tooltipProps={{ text: "CheckboxCompositeReversed" }} />
                <Icon iconName={"CheckList"} size={IconSize.large} tooltipProps={{ text: "CheckList" }} />
                <Icon iconName={"CheckMark"} size={IconSize.large} tooltipProps={{ text: "CheckMark" }} />
                <Icon iconName={"ChevronDown"} size={IconSize.large} tooltipProps={{ text: "ChevronDown" }} />
                <Icon iconName={"ChevronFold10"} size={IconSize.large} tooltipProps={{ text: "ChevronFold10" }} />
                <Icon iconName={"ChevronLeft"} size={IconSize.large} tooltipProps={{ text: "ChevronLeft" }} />
                <Icon iconName={"ChevronRight"} size={IconSize.large} tooltipProps={{ text: "ChevronRight" }} />
                <Icon iconName={"ChevronUnfold10"} size={IconSize.large} tooltipProps={{ text: "ChevronUnfold10" }} />
                <Icon iconName={"ChevronUp"} size={IconSize.large} tooltipProps={{ text: "ChevronUp" }} />
                <Icon iconName={"ChromeClose"} size={IconSize.large} tooltipProps={{ text: "ChromeClose" }} />
                <Icon iconName={"CircleFill"} size={IconSize.large} tooltipProps={{ text: "CircleFill" }} />
                <Icon iconName={"CirclePause"} size={IconSize.large} tooltipProps={{ text: "CirclePause" }} />
                <Icon iconName={"CirclePauseSolid"} size={IconSize.large} tooltipProps={{ text: "CirclePauseSolid" }} />
                <Icon iconName={"CirclePlus"} size={IconSize.large} tooltipProps={{ text: "CirclePlus" }} />
                <Icon iconName={"CircleRing"} size={IconSize.large} tooltipProps={{ text: "CircleRing" }} />
                <Icon iconName={"CircleShapeSolid"} size={IconSize.large} tooltipProps={{ text: "CircleShapeSolid" }} />
                <Icon iconName={"CircleStop"} size={IconSize.large} tooltipProps={{ text: "CircleStop" }} />
                <Icon iconName={"CircleStopSolid"} size={IconSize.large} tooltipProps={{ text: "CircleStopSolid" }} />
                <Icon iconName={"CityNext"} size={IconSize.large} tooltipProps={{ text: "CityNext" }} />
                <Icon iconName={"Clear"} size={IconSize.large} tooltipProps={{ text: "Clear" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"ClearFilter"} size={IconSize.large} tooltipProps={{ text: "ClearFilter" }} />
                <Icon iconName={"ClearFormatting"} size={IconSize.large} tooltipProps={{ text: "ClearFormatting" }} />
                <Icon iconName={"Clicked"} size={IconSize.large} tooltipProps={{ text: "Clicked" }} />
                <Icon iconName={"ClipboardSolid"} size={IconSize.large} tooltipProps={{ text: "ClipboardSolid" }} />
                <Icon iconName={"Clock"} size={IconSize.large} tooltipProps={{ text: "Clock" }} />
                <Icon iconName={"CloneToDesktop"} size={IconSize.large} tooltipProps={{ text: "CloneToDesktop" }} />
                <Icon iconName={"ClosePane"} size={IconSize.large} tooltipProps={{ text: "ClosePane" }} />
                <Icon iconName={"CloudDownload"} size={IconSize.large} tooltipProps={{ text: "CloudDownload" }} />
                <Icon iconName={"CloudUpload"} size={IconSize.large} tooltipProps={{ text: "CloudUpload" }} />
                <Icon iconName={"CloudWeather"} size={IconSize.large} tooltipProps={{ text: "CloudWeather" }} />
                <Icon iconName={"Cloudy"} size={IconSize.large} tooltipProps={{ text: "Cloudy" }} />
                <Icon iconName={"Code"} size={IconSize.large} tooltipProps={{ text: "Code" }} />
                <Icon iconName={"CoffeeScript"} size={IconSize.large} tooltipProps={{ text: "CoffeeScript" }} />
                <Icon iconName={"CollegeFootball"} size={IconSize.large} tooltipProps={{ text: "CollegeFootball" }} />
                <Icon iconName={"Color"} size={IconSize.large} tooltipProps={{ text: "Color" }} />
                <Icon iconName={"ColorSolid"} size={IconSize.large} tooltipProps={{ text: "ColorSolid" }} />
                <Icon iconName={"Comment"} size={IconSize.large} tooltipProps={{ text: "Comment" }} />
                <Icon iconName={"CommentAdd"} size={IconSize.large} tooltipProps={{ text: "CommentAdd" }} />
                <Icon iconName={"Completed"} size={IconSize.large} tooltipProps={{ text: "Completed" }} />
                <Icon iconName={"CompletedSolid"} size={IconSize.large} tooltipProps={{ text: "CompletedSolid" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Configuration"} size={IconSize.large} tooltipProps={{ text: "Configuration" }} />
                <Icon iconName={"ConnectContacts"} size={IconSize.large} tooltipProps={{ text: "ConnectContacts" }} />
                <Icon iconName={"ConstructionConeSolid"} size={IconSize.large} tooltipProps={{ text: "ConstructionConeSolid" }} />
                <Icon iconName={"Contact"} size={IconSize.large} tooltipProps={{ text: "Contact" }} />
                <Icon iconName={"ContactCard"} size={IconSize.large} tooltipProps={{ text: "ContactCard" }} />
                <Icon iconName={"ContactInfo"} size={IconSize.large} tooltipProps={{ text: "ContactInfo" }} />
                <Icon iconName={"Copy"} size={IconSize.large} tooltipProps={{ text: "Copy" }} />
                <Icon iconName={"CPU"} size={IconSize.large} tooltipProps={{ text: "CPU" }} />
                <Icon iconName={"CrownSolid"} size={IconSize.large} tooltipProps={{ text: "CrownSolid" }} />
                <Icon iconName={"CSharpLanguage"} size={IconSize.large} tooltipProps={{ text: "CSharpLanguage" }} />
                <Icon iconName={"CustomList"} size={IconSize.large} tooltipProps={{ text: "CustomList" }} />
                <Icon iconName={"DashKey"} size={IconSize.large} tooltipProps={{ text: "DashKey" }} />
                <Icon iconName={"Database"} size={IconSize.large} tooltipProps={{ text: "Database" }} />
                <Icon iconName={"DateTime2"} size={IconSize.large} tooltipProps={{ text: "DateTime2" }} />
                <Icon iconName={"DecisionSolid"} size={IconSize.large} tooltipProps={{ text: "DecisionSolid" }} />
                <Icon iconName={"Delete"} size={IconSize.large} tooltipProps={{ text: "Delete" }} />
                <Icon iconName={"Descending"} size={IconSize.large} tooltipProps={{ text: "Descending" }} />
                <Icon iconName={"Diagnostic"} size={IconSize.large} tooltipProps={{ text: "Diagnostic" }} />
                <Icon iconName={"Diamond2Solid"} size={IconSize.large} tooltipProps={{ text: "Diamond2Solid" }} />
                <Icon iconName={"DiamondSolid"} size={IconSize.large} tooltipProps={{ text: "DiamondSolid" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Dictionary"} size={IconSize.large} tooltipProps={{ text: "Dictionary" }} />
                <Icon iconName={"DictionaryRemove"} size={IconSize.large} tooltipProps={{ text: "DictionaryRemove" }} />
                <Icon iconName={"DiffInline"} size={IconSize.large} tooltipProps={{ text: "DiffInline" }} />
                <Icon iconName={"DiffSideBySide"} size={IconSize.large} tooltipProps={{ text: "DiffSideBySide" }} />
                <Icon iconName={"Dislike"} size={IconSize.large} tooltipProps={{ text: "Dislike" }} />
                <Icon iconName={"DockRight"} size={IconSize.large} tooltipProps={{ text: "DockRight" }} />
                <Icon iconName={"Documentation"} size={IconSize.large} tooltipProps={{ text: "Documentation" }} />
                <Icon iconName={"DocumentSearch"} size={IconSize.large} tooltipProps={{ text: "DocumentSearch" }} />
                <Icon iconName={"DocumentSet"} size={IconSize.large} tooltipProps={{ text: "DocumentSet" }} />
                <Icon iconName={"DoubleChevronDown"} size={IconSize.large} tooltipProps={{ text: "DoubleChevronDown" }} />
                <Icon iconName={"DoubleChevronLeft"} size={IconSize.large} tooltipProps={{ text: "DoubleChevronLeft" }} />
                <Icon iconName={"DoubleChevronRight"} size={IconSize.large} tooltipProps={{ text: "DoubleChevronRight" }} />
                <Icon iconName={"DoubleChevronUp"} size={IconSize.large} tooltipProps={{ text: "DoubleChevronUp" }} />
                <Icon iconName={"Down"} size={IconSize.large} tooltipProps={{ text: "Down" }} />
                <Icon iconName={"Download"} size={IconSize.large} tooltipProps={{ text: "Download" }} />
                <Icon iconName={"DownloadDocument"} size={IconSize.large} tooltipProps={{ text: "DownloadDocument" }} />
                <Icon iconName={"EatDrink"} size={IconSize.large} tooltipProps={{ text: "EatDrink" }} />
                <Icon iconName={"Edit"} size={IconSize.large} tooltipProps={{ text: "Edit" }} />
                <Icon iconName={"EditNote"} size={IconSize.large} tooltipProps={{ text: "EditNote" }} />
                <Icon iconName={"EditStyle"} size={IconSize.large} tooltipProps={{ text: "EditStyle" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Embed"} size={IconSize.large} tooltipProps={{ text: "Embed" }} />
                <Icon iconName={"EMI"} size={IconSize.large} tooltipProps={{ text: "EMI" }} />
                <Icon iconName={"Emoji"} size={IconSize.large} tooltipProps={{ text: "Emoji" }} />
                <Icon iconName={"Emoji2"} size={IconSize.large} tooltipProps={{ text: "Emoji2" }} />
                <Icon iconName={"EntryView"} size={IconSize.large} tooltipProps={{ text: "EntryView" }} />
                <Icon iconName={"Equalizer"} size={IconSize.large} tooltipProps={{ text: "Equalizer" }} />
                <Icon iconName={"Error"} size={IconSize.large} tooltipProps={{ text: "Error" }} />
                <Icon iconName={"ErrorBadge"} size={IconSize.large} tooltipProps={{ text: "ErrorBadge" }} />
                <Icon iconName={"ExploreContent"} size={IconSize.large} tooltipProps={{ text: "ExploreContent" }} />
                <Icon iconName={"ExploreData"} size={IconSize.large} tooltipProps={{ text: "ExploreData" }} />
                <Icon iconName={"Export"} size={IconSize.large} tooltipProps={{ text: "Export" }} />
                <Icon iconName={"ExportMirrored"} size={IconSize.large} tooltipProps={{ text: "ExportMirrored" }} />
                <Icon iconName={"EyeHide"} size={IconSize.large} tooltipProps={{ text: "EyeHide" }} />
                <Icon iconName={"EyeShow"} size={IconSize.large} tooltipProps={{ text: "EyeShow" }} />
                <Icon iconName={"FabricFolder"} size={IconSize.large} tooltipProps={{ text: "FabricFolder" }} />
                <Icon iconName={"FabricFolderFill"} size={IconSize.large} tooltipProps={{ text: "FabricFolderFill" }} />
                <Icon iconName={"FabricNewFolder"} size={IconSize.large} tooltipProps={{ text: "FabricNewFolder" }} />
                <Icon iconName={"FabricTextHighlightComposite"} size={IconSize.large} tooltipProps={{ text: "FabricTextHighlightComposite" }} />
                <Icon iconName={"FangBody"} size={IconSize.large} tooltipProps={{ text: "FangBody" }} />
                <Icon iconName={"FavoriteList"} size={IconSize.large} tooltipProps={{ text: "FavoriteList" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"FavoriteStar"} size={IconSize.large} tooltipProps={{ text: "FavoriteStar" }} />
                <Icon iconName={"FavoriteStarFill"} size={IconSize.large} tooltipProps={{ text: "FavoriteStarFill" }} />
                <Icon iconName={"Feedback"} size={IconSize.large} tooltipProps={{ text: "Feedback" }} />
                <Icon iconName={"FeedbackRequestSolid"} size={IconSize.large} tooltipProps={{ text: "FeedbackRequestSolid" }} />
                <Icon iconName={"FileBug"} size={IconSize.large} tooltipProps={{ text: "FileBug" }} />
                <Icon iconName={"FileCode"} size={IconSize.large} tooltipProps={{ text: "FileCode" }} />
                <Icon iconName={"FileCSS"} size={IconSize.large} tooltipProps={{ text: "FileCSS" }} />
                <Icon iconName={"FileHTML"} size={IconSize.large} tooltipProps={{ text: "FileHTML" }} />
                <Icon iconName={"FileImage"} size={IconSize.large} tooltipProps={{ text: "FileImage" }} />
                <Icon iconName={"FileJAVA"} size={IconSize.large} tooltipProps={{ text: "FileJAVA" }} />
                <Icon iconName={"FilePDB"} size={IconSize.large} tooltipProps={{ text: "FilePDB" }} />
                <Icon iconName={"FileSass"} size={IconSize.large} tooltipProps={{ text: "FileSass" }} />
                <Icon iconName={"FileTemplate"} size={IconSize.large} tooltipProps={{ text: "FileTemplate" }} />
                <Icon iconName={"FileYML"} size={IconSize.large} tooltipProps={{ text: "FileYML" }} />
                <Icon iconName={"Filter"} size={IconSize.large} tooltipProps={{ text: "Filter" }} />
                <Icon iconName={"FilterSolid"} size={IconSize.large} tooltipProps={{ text: "FilterSolid" }} />
                <Icon iconName={"FiltersSolid"} size={IconSize.large} tooltipProps={{ text: "FiltersSolid" }} />
                <Icon iconName={"FinancialSolid"} size={IconSize.large} tooltipProps={{ text: "FinancialSolid" }} />
                <Icon iconName={"Fingerprint"} size={IconSize.large} tooltipProps={{ text: "Fingerprint" }} />
                <Icon iconName={"Flag"} size={IconSize.large} tooltipProps={{ text: "Flag" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"FlameSolid"} size={IconSize.large} tooltipProps={{ text: "FlameSolid" }} />
                <Icon iconName={"Flashlight"} size={IconSize.large} tooltipProps={{ text: "Flashlight" }} />
                <Icon iconName={"FlowChart"} size={IconSize.large} tooltipProps={{ text: "FlowChart" }} />
                <Icon iconName={"Folder"} size={IconSize.large} tooltipProps={{ text: "Folder" }} />
                <Icon iconName={"FolderArrowRight"} size={IconSize.large} tooltipProps={{ text: "FolderArrowRight" }} />
                <Icon iconName={"FolderHorizontal"} size={IconSize.large} tooltipProps={{ text: "FolderHorizontal" }} />
                <Icon iconName={"FolderList"} size={IconSize.large} tooltipProps={{ text: "FolderList" }} />
                <Icon iconName={"FolderQuery"} size={IconSize.large} tooltipProps={{ text: "FolderQuery" }} />
                <Icon iconName={"FontColor"} size={IconSize.large} tooltipProps={{ text: "FontColor" }} />
                <Icon iconName={"FontColorA"} size={IconSize.large} tooltipProps={{ text: "FontColorA" }} />
                <Icon iconName={"FontSize"} size={IconSize.large} tooltipProps={{ text: "FontSize" }} />
                <Icon iconName={"Forward"} size={IconSize.large} tooltipProps={{ text: "Forward" }} />
                <Icon iconName={"FSharpLanguage"} size={IconSize.large} tooltipProps={{ text: "FSharpLanguage" }} />
                <Icon iconName={"FullHistory"} size={IconSize.large} tooltipProps={{ text: "FullHistory" }} />
                <Icon iconName={"FullScreen"} size={IconSize.large} tooltipProps={{ text: "FullScreen" }} />
                <Icon iconName={"Giftbox"} size={IconSize.large} tooltipProps={{ text: "Giftbox" }} />
                <Icon iconName={"GiftBoxSolid"} size={IconSize.large} tooltipProps={{ text: "GiftBoxSolid" }} />
                <Icon iconName={"GlobalNavButton"} size={IconSize.large} tooltipProps={{ text: "GlobalNavButton" }} />
                <Icon iconName={"Globe"} size={IconSize.large} tooltipProps={{ text: "Globe" }} />
                <Icon iconName={"GridViewSmall"} size={IconSize.large} tooltipProps={{ text: "GridViewSmall" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"GripperDotsVertical"} size={IconSize.large} tooltipProps={{ text: "GripperDotsVertical" }} />
                <Icon iconName={"Group"} size={IconSize.large} tooltipProps={{ text: "Group" }} />
                <Icon iconName={"HeadsetSolid"} size={IconSize.large} tooltipProps={{ text: "HeadsetSolid" }} />
                <Icon iconName={"Heart"} size={IconSize.large} tooltipProps={{ text: "Heart" }} />
                <Icon iconName={"HeartFill"} size={IconSize.large} tooltipProps={{ text: "HeartFill" }} />
                <Icon iconName={"Help"} size={IconSize.large} tooltipProps={{ text: "Help" }} />
                <Icon iconName={"Hide2"} size={IconSize.large} tooltipProps={{ text: "Hide2" }} />
                <Icon iconName={"History"} size={IconSize.large} tooltipProps={{ text: "History" }} />
                <Icon iconName={"Home"} size={IconSize.large} tooltipProps={{ text: "Home" }} />
                <Icon iconName={"Import"} size={IconSize.large} tooltipProps={{ text: "Import" }} />
                <Icon iconName={"Inbox"} size={IconSize.large} tooltipProps={{ text: "Inbox" }} />
                <Icon iconName={"IncidentTriangle"} size={IconSize.large} tooltipProps={{ text: "IncidentTriangle" }} />
                <Icon iconName={"Info"} size={IconSize.large} tooltipProps={{ text: "Info" }} />
                <Icon iconName={"InfoSolid"} size={IconSize.large} tooltipProps={{ text: "InfoSolid" }} />
                <Icon iconName={"Insights"} size={IconSize.large} tooltipProps={{ text: "Insights" }} />
                <Icon iconName={"IssueSolid"} size={IconSize.large} tooltipProps={{ text: "IssueSolid" }} />
                <Icon iconName={"Italic"} size={IconSize.large} tooltipProps={{ text: "Italic" }} />
                <Icon iconName={"JavaScriptLanguage"} size={IconSize.large} tooltipProps={{ text: "JavaScriptLanguage" }} />
                <Icon iconName={"KeyboardClassic"} size={IconSize.large} tooltipProps={{ text: "KeyboardClassic" }} />
                <Icon iconName={"LaptopSecure"} size={IconSize.large} tooltipProps={{ text: "LaptopSecure" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Library"} size={IconSize.large} tooltipProps={{ text: "Library" }} />
                <Icon iconName={"Lightbulb"} size={IconSize.large} tooltipProps={{ text: "Lightbulb" }} />
                <Icon iconName={"LightningBolt"} size={IconSize.large} tooltipProps={{ text: "LightningBolt" }} />
                <Icon iconName={"Like"} size={IconSize.large} tooltipProps={{ text: "Like" }} />
                <Icon iconName={"LikeSolid"} size={IconSize.large} tooltipProps={{ text: "LikeSolid" }} />
                <Icon iconName={"Link"} size={IconSize.large} tooltipProps={{ text: "Link" }} />
                <Icon iconName={"List"} size={IconSize.large} tooltipProps={{ text: "List" }} />
                <Icon iconName={"LocationDot"} size={IconSize.large} tooltipProps={{ text: "LocationDot" }} />
                <Icon iconName={"Lock"} size={IconSize.large} tooltipProps={{ text: "Lock" }} />
                <Icon iconName={"LockSolid"} size={IconSize.large} tooltipProps={{ text: "LockSolid" }} />
                <Icon iconName={"Mail"} size={IconSize.large} tooltipProps={{ text: "Mail" }} />
                <Icon iconName={"MarkDownLanguage"} size={IconSize.large} tooltipProps={{ text: "MarkDownLanguage" }} />
                <Icon iconName={"MediaStorageTower"} size={IconSize.large} tooltipProps={{ text: "MediaStorageTower" }} />
                <Icon iconName={"Megaphone"} size={IconSize.large} tooltipProps={{ text: "Megaphone" }} />
                <Icon iconName={"MegaphoneSolid"} size={IconSize.large} tooltipProps={{ text: "MegaphoneSolid" }} />
                <Icon iconName={"MiniExpand"} size={IconSize.large} tooltipProps={{ text: "MiniExpand" }} />
                <Icon iconName={"More"} size={IconSize.large} tooltipProps={{ text: "More" }} />
                <Icon iconName={"MoreVertical"} size={IconSize.large} tooltipProps={{ text: "MoreVertical" }} />
                <Icon iconName={"MSNVideos"} size={IconSize.large} tooltipProps={{ text: "MSNVideos" }} />
                <Icon iconName={"MSNVideosSolid"} size={IconSize.large} tooltipProps={{ text: "MSNVideosSolid" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"MultiSelect"} size={IconSize.large} tooltipProps={{ text: "MultiSelect" }} />
                <Icon iconName={"MusicInCollectionFill"} size={IconSize.large} tooltipProps={{ text: "MusicInCollectionFill" }} />
                <Icon iconName={"MyMoviesTV"} size={IconSize.large} tooltipProps={{ text: "MyMoviesTV" }} />
                <Icon iconName={"NavigateExternalInline"} size={IconSize.large} tooltipProps={{ text: "NavigateExternalInline" }} />
                <Icon iconName={"NavigateForward"} size={IconSize.large} tooltipProps={{ text: "NavigateForward" }} />
                <Icon iconName={"Next"} size={IconSize.large} tooltipProps={{ text: "Next" }} />
                <Icon iconName={"NotExecuted"} size={IconSize.large} tooltipProps={{ text: "NotExecuted" }} />
                <Icon iconName={"NotImpactedSolid"} size={IconSize.large} tooltipProps={{ text: "NotImpactedSolid" }} />
                <Icon iconName={"NumberedList"} size={IconSize.large} tooltipProps={{ text: "NumberedList" }} />
                <Icon iconName={"NumberSymbol"} size={IconSize.large} tooltipProps={{ text: "NumberSymbol" }} />
                <Icon iconName={"OEM"} size={IconSize.large} tooltipProps={{ text: "OEM" }} />
                <Icon iconName={"OfflineStorageSolid"} size={IconSize.large} tooltipProps={{ text: "OfflineStorageSolid" }} />
                <Icon iconName={"OpenInNewTab"} size={IconSize.large} tooltipProps={{ text: "OpenInNewTab" }} />
                <Icon iconName={"OpenPane"} size={IconSize.large} tooltipProps={{ text: "OpenPane" }} />
                <Icon iconName={"OpenSource"} size={IconSize.large} tooltipProps={{ text: "OpenSource" }} />
                <Icon iconName={"Org"} size={IconSize.large} tooltipProps={{ text: "Org" }} />
                <Icon iconName={"Package"} size={IconSize.large} tooltipProps={{ text: "Package" }} />
                <Icon iconName={"Page"} size={IconSize.large} tooltipProps={{ text: "Page" }} />
                <Icon iconName={"PageAdd"} size={IconSize.large} tooltipProps={{ text: "PageAdd" }} />
                <Icon iconName={"PageArrowRight"} size={IconSize.large} tooltipProps={{ text: "PageArrowRight" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"PageEdit"} size={IconSize.large} tooltipProps={{ text: "PageEdit" }} />
                <Icon iconName={"PageListSolid"} size={IconSize.large} tooltipProps={{ text: "PageListSolid" }} />
                <Icon iconName={"ParkingSolid"} size={IconSize.large} tooltipProps={{ text: "ParkingSolid" }} />
                <Icon iconName={"PartyLeader"} size={IconSize.large} tooltipProps={{ text: "PartyLeader" }} />
                <Icon iconName={"Paste"} size={IconSize.large} tooltipProps={{ text: "Paste" }} />
                <Icon iconName={"PasteAsCode"} size={IconSize.large} tooltipProps={{ text: "PasteAsCode" }} />
                <Icon iconName={"Pause"} size={IconSize.large} tooltipProps={{ text: "Pause" }} />
                <Icon iconName={"PaymentCard"} size={IconSize.large} tooltipProps={{ text: "PaymentCard" }} />
                <Icon iconName={"PC1"} size={IconSize.large} tooltipProps={{ text: "PC1" }} />
                <Icon iconName={"PDF"} size={IconSize.large} tooltipProps={{ text: "PDF" }} />
                <Icon iconName={"People"} size={IconSize.large} tooltipProps={{ text: "People" }} />
                <Icon iconName={"PeopleAdd"} size={IconSize.large} tooltipProps={{ text: "PeopleAdd" }} />
                <Icon iconName={"PeopleSettings"} size={IconSize.large} tooltipProps={{ text: "PeopleSettings" }} />
                <Icon iconName={"Permissions"} size={IconSize.large} tooltipProps={{ text: "Permissions" }} />
                <Icon iconName={"PermissionsSolid"} size={IconSize.large} tooltipProps={{ text: "PermissionsSolid" }} />
                <Icon iconName={"Phone"} size={IconSize.large} tooltipProps={{ text: "Phone" }} />
                <Icon iconName={"Photo2"} size={IconSize.large} tooltipProps={{ text: "Photo2" }} />
                <Icon iconName={"Pin"} size={IconSize.large} tooltipProps={{ text: "Pin" }} />
                <Icon iconName={"Pinned"} size={IconSize.large} tooltipProps={{ text: "Pinned" }} />
                <Icon iconName={"PlanView"} size={IconSize.large} tooltipProps={{ text: "PlanView" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Play"} size={IconSize.large} tooltipProps={{ text: "Play" }} />
                <Icon iconName={"PlayerSettings"} size={IconSize.large} tooltipProps={{ text: "PlayerSettings" }} />
                <Icon iconName={"PlayResume"} size={IconSize.large} tooltipProps={{ text: "PlayResume" }} />
                <Icon iconName={"PlugConnected"} size={IconSize.large} tooltipProps={{ text: "PlugConnected" }} />
                <Icon iconName={"PlugDisconnected"} size={IconSize.large} tooltipProps={{ text: "PlugDisconnected" }} />
                <Icon iconName={"POI"} size={IconSize.large} tooltipProps={{ text: "POI" }} />
                <Icon iconName={"PreviewLink"} size={IconSize.large} tooltipProps={{ text: "PreviewLink" }} />
                <Icon iconName={"Previous"} size={IconSize.large} tooltipProps={{ text: "Previous" }} />
                <Icon iconName={"Print"} size={IconSize.large} tooltipProps={{ text: "Print" }} />
                <Icon iconName={"Processing"} size={IconSize.large} tooltipProps={{ text: "Processing" }} />
                <Icon iconName={"Product"} size={IconSize.large} tooltipProps={{ text: "Product" }} />
                <Icon iconName={"ProFootball"} size={IconSize.large} tooltipProps={{ text: "ProFootball" }} />
                <Icon iconName={"ProgressLoopOuter"} size={IconSize.large} tooltipProps={{ text: "ProgressLoopOuter" }} />
                <Icon iconName={"Prohibited"} size={IconSize.large} tooltipProps={{ text: "Prohibited" }} />
                <Icon iconName={"Project"} size={IconSize.large} tooltipProps={{ text: "Project" }} />
                <Icon iconName={"ProjectCollection"} size={IconSize.large} tooltipProps={{ text: "ProjectCollection" }} />
                <Icon iconName={"PublishContent"} size={IconSize.large} tooltipProps={{ text: "PublishContent" }} />
                <Icon iconName={"Puzzle"} size={IconSize.large} tooltipProps={{ text: "Puzzle" }} />
                <Icon iconName={"PythonLanguage"} size={IconSize.large} tooltipProps={{ text: "PythonLanguage" }} />
                <Icon iconName={"QueryList"} size={IconSize.large} tooltipProps={{ text: "QueryList" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"QuickNoteSolid"} size={IconSize.large} tooltipProps={{ text: "QuickNoteSolid" }} />
                <Icon iconName={"RadioBtnOff"} size={IconSize.large} tooltipProps={{ text: "RadioBtnOff" }} />
                <Icon iconName={"RadioBtnOn"} size={IconSize.large} tooltipProps={{ text: "RadioBtnOn" }} />
                <Icon iconName={"RawSource"} size={IconSize.large} tooltipProps={{ text: "RawSource" }} />
                <Icon iconName={"ReadingMode"} size={IconSize.large} tooltipProps={{ text: "ReadingMode" }} />
                <Icon iconName={"ReadingModeSolid"} size={IconSize.large} tooltipProps={{ text: "ReadingModeSolid" }} />
                <Icon iconName={"ReceiptCheck"} size={IconSize.large} tooltipProps={{ text: "ReceiptCheck" }} />
                <Icon iconName={"Recent"} size={IconSize.large} tooltipProps={{ text: "Recent" }} />
                <Icon iconName={"RedEye"} size={IconSize.large} tooltipProps={{ text: "RedEye" }} />
                <Icon iconName={"Refresh"} size={IconSize.large} tooltipProps={{ text: "Refresh" }} />
                <Icon iconName={"ReleaseGate"} size={IconSize.large} tooltipProps={{ text: "ReleaseGate" }} />
                <Icon iconName={"Remove"} size={IconSize.large} tooltipProps={{ text: "Remove" }} />
                <Icon iconName={"RemoveLink"} size={IconSize.large} tooltipProps={{ text: "RemoveLink" }} />
                <Icon iconName={"Rename"} size={IconSize.large} tooltipProps={{ text: "Rename" }} />
                <Icon iconName={"Repair"} size={IconSize.large} tooltipProps={{ text: "Repair" }} />
                <Icon iconName={"Reply"} size={IconSize.large} tooltipProps={{ text: "Reply" }} />
                <Icon iconName={"ReplyMirrored"} size={IconSize.large} tooltipProps={{ text: "ReplyMirrored" }} />
                <Icon iconName={"Repo"} size={IconSize.large} tooltipProps={{ text: "Repo" }} />
                <Icon iconName={"ReportHacked"} size={IconSize.large} tooltipProps={{ text: "ReportHacked" }} />
                <Icon iconName={"ReviewSolid"} size={IconSize.large} tooltipProps={{ text: "ReviewSolid" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"RevToggleKey"} size={IconSize.large} tooltipProps={{ text: "RevToggleKey" }} />
                <Icon iconName={"Rewind"} size={IconSize.large} tooltipProps={{ text: "Rewind" }} />
                <Icon iconName={"Ribbon"} size={IconSize.large} tooltipProps={{ text: "Ribbon" }} />
                <Icon iconName={"RibbonSolid"} size={IconSize.large} tooltipProps={{ text: "RibbonSolid" }} />
                <Icon iconName={"Ringer"} size={IconSize.large} tooltipProps={{ text: "Ringer" }} />
                <Icon iconName={"RingerOff"} size={IconSize.large} tooltipProps={{ text: "RingerOff" }} />
                <Icon iconName={"Rocket"} size={IconSize.large} tooltipProps={{ text: "Rocket" }} />
                <Icon iconName={"RowsGroup"} size={IconSize.large} tooltipProps={{ text: "RowsGroup" }} />
                <Icon iconName={"Sad"} size={IconSize.large} tooltipProps={{ text: "Sad" }} />
                <Icon iconName={"Save"} size={IconSize.large} tooltipProps={{ text: "Save" }} />
                <Icon iconName={"SaveAll"} size={IconSize.large} tooltipProps={{ text: "SaveAll" }} />
                <Icon iconName={"SaveAs"} size={IconSize.large} tooltipProps={{ text: "SaveAs" }} />
                <Icon iconName={"ScheduleEventAction"} size={IconSize.large} tooltipProps={{ text: "ScheduleEventAction" }} />
                <Icon iconName={"Script"} size={IconSize.large} tooltipProps={{ text: "Script" }} />
                <Icon iconName={"ScrollUpDown"} size={IconSize.large} tooltipProps={{ text: "ScrollUpDown" }} />
                <Icon iconName={"Search"} size={IconSize.large} tooltipProps={{ text: "Search" }} />
                <Icon iconName={"SearchAndApps"} size={IconSize.large} tooltipProps={{ text: "SearchAndApps" }} />
                <Icon iconName={"SecurityGroup"} size={IconSize.large} tooltipProps={{ text: "SecurityGroup" }} />
                <Icon iconName={"SemanticZoom"} size={IconSize.large} tooltipProps={{ text: "SemanticZoom" }} />
                <Icon iconName={"SemiboldWeight"} size={IconSize.large} tooltipProps={{ text: "SemiboldWeight" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Send"} size={IconSize.large} tooltipProps={{ text: "Send" }} />
                <Icon iconName={"Server"} size={IconSize.large} tooltipProps={{ text: "Server" }} />
                <Icon iconName={"ServerEnviroment"} size={IconSize.large} tooltipProps={{ text: "ServerEnviroment" }} />
                <Icon iconName={"ServerProcesses"} size={IconSize.large} tooltipProps={{ text: "ServerProcesses" }} />
                <Icon iconName={"Settings"} size={IconSize.large} tooltipProps={{ text: "Settings" }} />
                <Icon iconName={"SettingsApp"} size={IconSize.large} tooltipProps={{ text: "SettingsApp" }} />
                <Icon iconName={"Share"} size={IconSize.large} tooltipProps={{ text: "Share" }} />
                <Icon iconName={"Shield"} size={IconSize.large} tooltipProps={{ text: "Shield" }} />
                <Icon iconName={"ShieldSolid"} size={IconSize.large} tooltipProps={{ text: "ShieldSolid" }} />
                <Icon iconName={"Shop"} size={IconSize.large} tooltipProps={{ text: "Shop" }} />
                <Icon iconName={"ShoppingCart"} size={IconSize.large} tooltipProps={{ text: "ShoppingCart" }} />
                <Icon iconName={"ShopServer"} size={IconSize.large} tooltipProps={{ text: "ShopServer" }} />
                <Icon iconName={"ShowResults"} size={IconSize.large} tooltipProps={{ text: "ShowResults" }} />
                <Icon iconName={"Signin"} size={IconSize.large} tooltipProps={{ text: "Signin" }} />
                <Icon iconName={"SkypeCircleMinus"} size={IconSize.large} tooltipProps={{ text: "SkypeCircleMinus" }} />
                <Icon iconName={"Snowflake"} size={IconSize.large} tooltipProps={{ text: "Snowflake" }} />
                <Icon iconName={"Soccer"} size={IconSize.large} tooltipProps={{ text: "Soccer" }} />
                <Icon iconName={"SortDown"} size={IconSize.large} tooltipProps={{ text: "SortDown" }} />
                <Icon iconName={"SortLines"} size={IconSize.large} tooltipProps={{ text: "SortLines" }} />
                <Icon iconName={"SortUp"} size={IconSize.large} tooltipProps={{ text: "SortUp" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Sprint"} size={IconSize.large} tooltipProps={{ text: "Sprint" }} />
                <Icon iconName={"StackedBarChart"} size={IconSize.large} tooltipProps={{ text: "StackedBarChart" }} />
                <Icon iconName={"StackedLineChart"} size={IconSize.large} tooltipProps={{ text: "StackedLineChart" }} />
                <Icon iconName={"Starburst"} size={IconSize.large} tooltipProps={{ text: "Starburst" }} />
                <Icon iconName={"StarburstSolid"} size={IconSize.large} tooltipProps={{ text: "StarburstSolid" }} />
                <Icon iconName={"StatusCircleCheckmark"} size={IconSize.large} tooltipProps={{ text: "StatusCircleCheckmark" }} />
                <Icon iconName={"StatusCircleErrorX"} size={IconSize.large} tooltipProps={{ text: "StatusCircleErrorX" }} />
                <Icon iconName={"StatusCircleInner"} size={IconSize.large} tooltipProps={{ text: "StatusCircleInner" }} />
                <Icon iconName={"StatusCircleRing"} size={IconSize.large} tooltipProps={{ text: "StatusCircleRing" }} />
                <Icon iconName={"StatusErrorFull"} size={IconSize.large} tooltipProps={{ text: "StatusErrorFull" }} />
                <Icon iconName={"StockDown"} size={IconSize.large} tooltipProps={{ text: "StockDown" }} />
                <Icon iconName={"StockUp"} size={IconSize.large} tooltipProps={{ text: "StockUp" }} />
                <Icon iconName={"Stopwatch"} size={IconSize.large} tooltipProps={{ text: "Stopwatch" }} />
                <Icon iconName={"Streaming"} size={IconSize.large} tooltipProps={{ text: "Streaming" }} />
                <Icon iconName={"StreamingOff"} size={IconSize.large} tooltipProps={{ text: "StreamingOff" }} />
                <Icon iconName={"Strikethrough"} size={IconSize.large} tooltipProps={{ text: "Strikethrough" }} />
                <Icon iconName={"SurveyQuestions"} size={IconSize.large} tooltipProps={{ text: "SurveyQuestions" }} />
                <Icon iconName={"Switch"} size={IconSize.large} tooltipProps={{ text: "Switch" }} />
                <Icon iconName={"SwitcherStartEnd"} size={IconSize.large} tooltipProps={{ text: "SwitcherStartEnd" }} />
                <Icon iconName={"Table"} size={IconSize.large} tooltipProps={{ text: "Table" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Tag"} size={IconSize.large} tooltipProps={{ text: "Tag" }} />
                <Icon iconName={"TaskSolid"} size={IconSize.large} tooltipProps={{ text: "TaskSolid" }} />
                <Icon iconName={"TeamFavorite"} size={IconSize.large} tooltipProps={{ text: "TeamFavorite" }} />
                <Icon iconName={"Teamwork"} size={IconSize.large} tooltipProps={{ text: "Teamwork" }} />
                <Icon iconName={"TestAutoSolid"} size={IconSize.large} tooltipProps={{ text: "TestAutoSolid" }} />
                <Icon iconName={"TestBeaker"} size={IconSize.large} tooltipProps={{ text: "TestBeaker" }} />
                <Icon iconName={"TestBeakerSolid"} size={IconSize.large} tooltipProps={{ text: "TestBeakerSolid" }} />
                <Icon iconName={"TestPlan"} size={IconSize.large} tooltipProps={{ text: "TestPlan" }} />
                <Icon iconName={"TextDocument"} size={IconSize.large} tooltipProps={{ text: "TextDocument" }} />
                <Icon iconName={"TextField"} size={IconSize.large} tooltipProps={{ text: "TextField" }} />
                <Icon iconName={"Tiles"} size={IconSize.large} tooltipProps={{ text: "Tiles" }} />
                <Icon iconName={"TimeEntry"} size={IconSize.large} tooltipProps={{ text: "TimeEntry" }} />
                <Icon iconName={"Trending12"} size={IconSize.large} tooltipProps={{ text: "Trending12" }} />
                <Icon iconName={"TriangleRight12"} size={IconSize.large} tooltipProps={{ text: "TriangleRight12" }} />
                <Icon iconName={"TriangleSolidDown12"} size={IconSize.large} tooltipProps={{ text: "TriangleSolidDown12" }} />
                <Icon iconName={"TriangleSolidRight12"} size={IconSize.large} tooltipProps={{ text: "TriangleSolidRight12" }} />
                <Icon iconName={"TriangleSolidUp12"} size={IconSize.large} tooltipProps={{ text: "TriangleSolidUp12" }} />
                <Icon iconName={"TriggerApproval"} size={IconSize.large} tooltipProps={{ text: "TriggerApproval" }} />
                <Icon iconName={"TriggerAuto"} size={IconSize.large} tooltipProps={{ text: "TriggerAuto" }} />
                <Icon iconName={"TripleColumnEdit"} size={IconSize.large} tooltipProps={{ text: "TripleColumnEdit" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"Trophy2Solid"} size={IconSize.large} tooltipProps={{ text: "Trophy2Solid" }} />
                <Icon iconName={"TwoKeys"} size={IconSize.large} tooltipProps={{ text: "TwoKeys" }} />
                <Icon iconName={"TypeScriptLanguage"} size={IconSize.large} tooltipProps={{ text: "TypeScriptLanguage" }} />
                <Icon iconName={"Underline"} size={IconSize.large} tooltipProps={{ text: "Underline" }} />
                <Icon iconName={"Undo"} size={IconSize.large} tooltipProps={{ text: "Undo" }} />
                <Icon iconName={"Unknown"} size={IconSize.large} tooltipProps={{ text: "Unknown" }} />
                <Icon iconName={"Unlock"} size={IconSize.large} tooltipProps={{ text: "Unlock" }} />
                <Icon iconName={"UnlockSolid"} size={IconSize.large} tooltipProps={{ text: "UnlockSolid" }} />
                <Icon iconName={"Unpin"} size={IconSize.large} tooltipProps={{ text: "Unpin" }} />
                <Icon iconName={"Up"} size={IconSize.large} tooltipProps={{ text: "Up" }} />
                <Icon iconName={"Upload"} size={IconSize.large} tooltipProps={{ text: "Upload" }} />
                <Icon iconName={"UserFollowed"} size={IconSize.large} tooltipProps={{ text: "UserFollowed" }} />
                <Icon iconName={"UserRemove"} size={IconSize.large} tooltipProps={{ text: "UserRemove" }} />
                <Icon iconName={"Variable"} size={IconSize.large} tooltipProps={{ text: "Variable" }} />
                <Icon iconName={"VerifiedBrand"} size={IconSize.large} tooltipProps={{ text: "VerifiedBrand" }} />
                <Icon iconName={"VerifiedBrandSolid"} size={IconSize.large} tooltipProps={{ text: "VerifiedBrandSolid" }} />
                <Icon iconName={"Video"} size={IconSize.large} tooltipProps={{ text: "Video" }} />
                <Icon iconName={"View"} size={IconSize.large} tooltipProps={{ text: "View" }} />
                <Icon iconName={"ViewAll"} size={IconSize.large} tooltipProps={{ text: "ViewAll" }} />
                <Icon iconName={"ViewDashboard"} size={IconSize.large} tooltipProps={{ text: "ViewDashboard" }} />
            </div>
            <div className="flex-row rhythm-horizontal-16">
                <Icon iconName={"ViewList"} size={IconSize.large} tooltipProps={{ text: "ViewList" }} />
                <Icon iconName={"ViewListGroup"} size={IconSize.large} tooltipProps={{ text: "ViewListGroup" }} />
                <Icon iconName={"ViewListTree"} size={IconSize.large} tooltipProps={{ text: "ViewListTree" }} />
                <Icon iconName={"VisualBasicLanguage"} size={IconSize.large} tooltipProps={{ text: "VisualBasicLanguage" }} />
                <Icon iconName={"Waffle"} size={IconSize.large} tooltipProps={{ text: "Waffle" }} />
                <Icon iconName={"WaffleOffice365"} size={IconSize.large} tooltipProps={{ text: "WaffleOffice365" }} />
                <Icon iconName={"WaitlistConfirm"} size={IconSize.large} tooltipProps={{ text: "WaitlistConfirm" }} />
                <Icon iconName={"Warning"} size={IconSize.large} tooltipProps={{ text: "Warning" }} />
                <Icon iconName={"Work"} size={IconSize.large} tooltipProps={{ text: "Work" }} />
                <Icon iconName={"WorkFlow"} size={IconSize.large} tooltipProps={{ text: "WorkFlow" }} />
                <Icon iconName={"WorkItem"} size={IconSize.large} tooltipProps={{ text: "WorkItem" }} />
                <Icon iconName={"World"} size={IconSize.large} tooltipProps={{ text: "World" }} />
                <Icon iconName={"WorldClock"} size={IconSize.large} tooltipProps={{ text: "WorldClock" }} />
                <Icon iconName={"ZipFolder"} size={IconSize.large} tooltipProps={{ text: "ZipFolder" }} />
                <Icon iconName={"Zoom"} size={IconSize.large} tooltipProps={{ text: "Zoom" }} />
                <Icon iconName={"ZoomIn"} size={IconSize.large} tooltipProps={{ text: "ZoomIn" }} />
                <Icon iconName={"ZoomOut"} size={IconSize.large} tooltipProps={{ text: "ZoomOut" }} />
            </div>
        </div>
    )
}