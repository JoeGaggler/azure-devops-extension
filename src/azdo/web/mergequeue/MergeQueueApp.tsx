// TODO: special icon for status checks incomplete
// TODO: refresh current data after a failed concurrent update
// TODO: show different icon for dependent pull requests in same repo
// TODO: post merge queue status to pull request
// TODO: push default branch to merge queue when last PR leaves queue
// TODO: fix draft status not updated
// TODO: PR refresh should update the merge queue and active list
// TODO: tags on all pull requests tab to indicate which merge queue they are in
// TODO: defer processing when app version is newer than ours
// TODO: discretely poll each tracked pipeline run
import React from "react";
import * as luxon from 'luxon'
import * as SDK from 'azure-devops-extension-sdk';

import { Header } from "azure-devops-ui/Components/Header/Header";
import { IHeaderCommandBarItem } from "azure-devops-ui/Components/HeaderCommandBar/HeaderCommandBar.Props";
import { Page } from "azure-devops-ui/Components/Page/Page";
import { TitleSize } from "azure-devops-ui/Components/Header/Header.Props";
import { Card } from "azure-devops-ui/Card";

import { getAzdoInfo, getGitClient, getExtensionManagementClient, TenantInfo, getDefaultBranchCommitId, getRefCommitId, mergeCommits, AuthorInfo, summarizeVotes, PullRequestVotingResult, getRunClient, getRunStatus, getRunResult } from "./azuredevops";

import { GitAsyncOperationStatus, GitPullRequestSearchCriteria, GitRefUpdate, PullRequestAsyncStatus, PullRequestStatus, PullRequestTimeRangeType } from "azure-devops-extension-api/Git/Git";
import { ExtensionManagementRestClient } from "azure-devops-extension-api/ExtensionManagement/ExtensionManagementClient";
import { Icon, IconSize } from "azure-devops-ui/Icon";
import { distinctBy } from "./lib";
import { GitRestClient } from "azure-devops-extension-api/Git/GitClient";
import { IHostNavigationService } from "azure-devops-extension-api/Common/CommonServices";
import { Toggle } from "azure-devops-ui/Toggle";
import { PullRequestList, PullRequestListItem } from "./PullRequestList";
import { PipelineList, PipelineListItem } from "./PipelineList";
import { BuildQueryOrder, BuildStatus } from "azure-devops-extension-api/Build/Build";
import { Tab, TabBar, TabSize } from "azure-devops-ui/Tabs";
import { IMenuItem } from "azure-devops-ui/Components/Menu/Menu.Props";
import { AddMergeQueuePanel } from "./AddMergeQueuePanel";

const publisher = "pingmint";
const extensionName = "pingmint-extension";
const collectionId = "mergequeue";
const allMergeQueuesDocumentId = "allmergequeues";
const activePullRequestsDocumentId = "activepullrequests"; // TODO: MOVE THIS TO ITS OWN COLLECTION ID
const allPipelineRunsDocumentId = "allpipelineruns";
const zeroCommitId = "0000000000000000000000000000000000000000";
const refMergeQueue = "refs/heads/merge-queue";
const allPullRequestsTabId = "all";
const mainQueueTabId = "main";
const demoQueueTabId = "demo";
// const refDemoQueue = "refs/heads/merge-queue-demo";

export interface MergeQueueAppSingleton {
    bearerToken: string;
    appToken: string;
}

interface PullRequestFilters {
    drafts: boolean;
    blocked: boolean;
    queued: boolean;
    allBranches: boolean;
    repositories: string[];
}

interface RepositoryDetails {
    id: string;
    defaultBranch: string;
}

interface RepositoryInfo {
    id: string;
    name: string;
}

interface PullRequestInfo {
    id: number;
    title: string;
    repository: RepositoryInfo;
    author: AuthorInfo;
    createdTimestamp: number;
    isDraft: boolean;
    isAutoComplete: boolean;
    sourceRefName: string;
    targetRefName: string;
    voting: PullRequestVotingResult;
    mergeStatus: PullRequestAsyncStatus;
}

interface MergeQueueItemInfo {
    id: number;
    title: string;
    repository: RepositoryInfo;
    author: AuthorInfo;
    isAutoComplete: boolean;
    status: MergeQueueStatus;
    createdTimestamp: number;
    isDraft: boolean;
    sourceRefName: string;
    targetRefName: string;
    sourceCommitId: string;
    targetCommitId: string;
    mergedCommitId: string;
    voting: PullRequestVotingResult;
    mergeStatus: PullRequestAsyncStatus;
}

interface MergeQueueInfo {
    id: string;
    name: string;
    items: MergeQueueItemInfo[];
    targetRefName: string;
}

interface MergeCommitInfo {
    repositoryId: string;
    commitId: string;
}

type MergeQueueStatus =
    "queued" | // requires recalculation
    "recalculating" | // currently being recalculated
    "draft" | // is a draft pull request
    "conflict" | // has a merge conflict with another item in the queue
    "blocked" | // is blocked by another item in the queue
    "valid"; // can be merged

interface PullRequestDocument {
    id: string;
    viaVersion: string;
    __etag: number;
    pullRequests: PullRequestInfo[];
}

interface MergeQueueDocument {
    id: string;
    name: string;
    targetRefName: string;
    viaVersion: string;
    __etag: number;
    mergeQueueItems: MergeQueueItemInfo[];
}

interface AllMergeQueuesItem {
    id: string;
    name: string;
}

interface AllMergeQueuesDocument {
    id: string;
    viaVersion: string;
    __etag: number;
    mergeQueues: AllMergeQueuesItem[];
}

interface PipelineRunInfo {
    runId: number;
    runName: string;
    pipelineId: number;
    pipelineName: string;
    sourceVersion: string;
    status: string;
    result: string;
    createdAt: number;
}

interface PipelineRunsDocument {
    id: string;
    viaVersion: string;
    __etag: number;
    pipelineRuns: PipelineRunInfo[];
}

interface ReducerState {
    didInit: boolean;

    mergeQueues: MergeQueueInfo[];
    selectedPullRequestIds: number[];
    repositories: RepositoryDetails[];
    repoMergeCommits: MergeCommitInfo[];

    activePullRequests: PullRequestInfo[];
    filters: PullRequestFilters;
    filteredActivePullRequests: PullRequestInfo[]; // derived

    pipelineRuns: PipelineRunInfo[];
    selectedFilteredPipelineRunIds: number[];
    filteredPipelineRuns: PipelineRunInfo[]; // derived

    selectedQueueTabId: string;

    isShowingAddQueue: boolean;
}

interface ReducerAction {
    didInit?: boolean;

    // collection loading
    mergeQueues?: MergeQueueInfo[];
    singleMergeQueue?: MergeQueueInfo;
    repoMergeCommits?: MergeCommitInfo[];

    // active pull requests
    activePullRequests?: PullRequestInfo[];
    filteredActivePullRequests?: PullRequestInfo[];
    repositories?: RepositoryDetails[];

    // selection changes
    selectedPullRequestIds?: number[];

    // queue actions
    actionQueueId?: string;

    // actions
    filters?: PullRequestFilters;
    selectedQueueTabId?: string;

    // pipeline runs
    pipelineRuns?: PipelineRunInfo[];
    selectedPipelineRunIds?: number[];

    showAddQueue?: boolean;
}

function reducer(state: ReducerState, action: ReducerAction): ReducerState {
    let next = { ...state };

    if (action.didInit !== undefined) {
        next.didInit = action.didInit;
    }

    if (action.showAddQueue !== undefined) {
        next.isShowingAddQueue = action.showAddQueue;
    }

    if (action.singleMergeQueue !== undefined) {
        let id = action.singleMergeQueue.id;
        let name = action.singleMergeQueue.name;
        let items = action.singleMergeQueue.items || [];
        let targetRefName = action.singleMergeQueue.targetRefName || refMergeQueue; // TODO: remove default value

        let nextQueues = [...next.mergeQueues];
        let found = next.mergeQueues.find(i => i.id === id);
        if (found) {
            found.name = name;
            found.items = items;
            found.targetRefName = targetRefName;
        } else {
            found = {
                id: id,
                name: name,
                items: items,
                targetRefName: targetRefName,
            };
            nextQueues = [...next.mergeQueues, found]
        }
        next.mergeQueues = nextQueues;
    }

    if (action.mergeQueues !== undefined) {
        console.log("MQ: reducer -> updating merge queues", action.mergeQueues);
        next.mergeQueues = action.mergeQueues;
    }

    if (action.selectedQueueTabId !== undefined) {
        console.log("MQ: reducer -> updating selected queue tab ID", action.selectedQueueTabId);
        next.selectedQueueTabId = action.selectedQueueTabId;
    }

    if (action.repositories !== undefined) {
        console.log("MQ: reducer -> updating repositories", action.repositories);
        next.repositories = action.repositories;
    }

    if (action.repoMergeCommits !== undefined) {
        console.log("MQ: reducer -> updating repo merge commits", action.repoMergeCommits);
        next.repoMergeCommits = action.repoMergeCommits;
    }

    // selection changes (via pull requests or pipeline runs)
    if (action.selectedPullRequestIds !== undefined) {
        console.log("MQ: reducer -> updating selected pull request IDs", action.selectedPullRequestIds);
        next.selectedPullRequestIds = action.selectedPullRequestIds;
    }
    else if (action.selectedPipelineRunIds !== undefined) {
        let selIds = action.selectedPipelineRunIds;
        console.log("MQ: reducer -> updating selected pipeline runs", selIds);

        let runs = next.filteredPipelineRuns.filter(run => selIds.includes(run.runId))
        let runIds = runs.map(run => run.runId);
        let selectedCommitIds = runs.map(run => run.sourceVersion);

        next.selectedFilteredPipelineRunIds = runIds;
        let mq_items = next.mergeQueues.find(mq => mq.id === next.selectedQueueTabId)?.items ?? [];
        if (mq_items && mq_items.length > 0) {
            next.selectedPullRequestIds = (mq_items
                .filter(mqi => selectedCommitIds.includes(mqi.mergedCommitId))
                .map(mqi => mqi.id)
            );
        }
    }

    if (action.activePullRequests !== undefined) {
        console.log("MQ: reducer -> updating active pull requests", action.activePullRequests);
        next.activePullRequests = action.activePullRequests.sort((a, b) => b.createdTimestamp - a.createdTimestamp);
        // TODO: confirm selections

        // update merge queue items with latest pull request data
        next.mergeQueues.forEach(mq => {
            let mq_items = mq.items ?? [];
            mq_items.forEach(mqi => {
                let pr = next.activePullRequests.find(pr => pr.id === mqi.id);
                if (pr) {
                    mqi.title = pr.title;
                }
            });
        });
    }

    if (action.filters !== undefined) {
        console.log("MQ: reducer -> updating filters", action.filters);
        next.filters = action.filters;
    }

    // apply filters
    {
        let filt = [...next.activePullRequests];
        if (next.filters.allBranches === false) {
            filt = filt.flatMap(pr => {
                return next.repositories.some(r => r.id === pr.repository.id && r.defaultBranch === pr.targetRefName) ? pr : [];
            });
        }
        if (next.filters.drafts === false) {
            filt = filt.filter(pr => !pr.isDraft);
        }
        // TODO: should this filter be applied to all queues or just the main queue?
        let main_mq_items = next.mergeQueues.find(mq => mq.id === mainQueueTabId)?.items ?? [];
        if (next.filters.queued === false) {
            filt = filt.filter(pr => false === main_mq_items.some(mqi => mqi.id === pr.id));
        }
        if (next.filters.blocked === false) {
            filt = filt.filter(pr =>
                pr.voting.status !== "rejected" &&
                pr.voting.status !== "waiting" &&
                pr.voting.status !== "unknown"
            );
        }
        next.filteredActivePullRequests = filt;
    }

    // pipeline runs
    if (action.pipelineRuns !== undefined) {
        console.log("MQ: reducer -> updating pipeline runs", action.pipelineRuns);
        let allMergeCommitIds = next.repoMergeCommits.map(mci => mci.commitId);
        let didChange = false;
        for (const run of action.pipelineRuns) {
            let existing = next.pipelineRuns.find(r => r.runId === run.runId);
            if (existing) {
                // avoid triggering render if run does not change
                if (existing.pipelineName !== run.pipelineName ||
                    existing.status !== run.status ||
                    existing.result !== run.result ||
                    (existing.createdAt !== run.createdAt && run.createdAt !== 0) // HACK
                ) {
                    Object.assign(existing, run);
                    didChange = true;
                }
            } else if (allMergeCommitIds.includes(run.sourceVersion)) {
                next.pipelineRuns.push(run);
                didChange = true;
            }
        }
        let removedRunIds = next.pipelineRuns.filter(r => {
            for (const mq of next.mergeQueues) {
                for (const mqi of mq.items) {
                    if (mqi.mergedCommitId === r.sourceVersion) { return false; }
                }
            }
            console.log("MQ: reducer -> removing pipeline run", r.runId, r);
            return true;
        });
        if (removedRunIds.length > 0) {
            didChange = true;
            next.pipelineRuns = next.pipelineRuns.filter(r => !removedRunIds.includes(r));
        }
        if (didChange) {
            next.pipelineRuns = [...next.pipelineRuns.sort((a, b) => a.createdAt - b.createdAt)]; // sort by creation time, 
            console.log("MQ: reducer -> updated pipeline runs", next.pipelineRuns);
        }
    }

    // filter pipeline runs
    let selectedMergeQueue = next.mergeQueues.find(mq => mq.id === next.selectedQueueTabId);
    if (selectedMergeQueue !== undefined) {
        var mq_items = selectedMergeQueue.items ?? [];
        next.filteredPipelineRuns = [...next.pipelineRuns];
        next.filteredPipelineRuns = next.filteredPipelineRuns.filter(run => {
            return mq_items.some(mqi => mqi.mergedCommitId === run.sourceVersion);
        });

        // select pipeline runs via merge queue selection
        let selectedCommitIds: string[] = (mq_items
            .filter(mqi => next.selectedPullRequestIds.includes(mqi.id))
            .map(mqi => mqi.mergedCommitId)
        );
        let runIds: number[] = (next.filteredPipelineRuns
            .filter(run => selectedCommitIds.includes(run.sourceVersion))
            .map(run => run.runId)
        );

        next.selectedFilteredPipelineRunIds = runIds;
        console.log("MQ: reducer -> updating selected pipeline runs", next.selectedFilteredPipelineRunIds);
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
    let runClient = React.useRef(getRunClient());
    let extensionManagementClient = React.useRef(getExtensionManagementClient());
    let isTickTockRunning = React.useRef(false);
    let isMergeQueueRunning = React.useRef(false);
    let [showIcons, setShowIcons] = React.useState(false);

    const [state, dispatch] = React.useReducer<(state: ReducerState, action: ReducerAction) => ReducerState>(reducer, {
        didInit: false,

        mergeQueues: [],
        selectedPullRequestIds: [],
        repositories: [],
        repoMergeCommits: [],

        activePullRequests: [],
        filteredActivePullRequests: [],

        filters: {
            drafts: false,
            blocked: false,
            queued: false,
            allBranches: false,
            repositories: [],
        },

        pipelineRuns: [],
        filteredPipelineRuns: [],
        selectedFilteredPipelineRunIds: [],

        selectedQueueTabId: allPullRequestsTabId,

        isShowingAddQueue: false,
    })

    // initialize the app
    React.useEffect(() => { init() }, []);
    async function init() {
        let state = {} // shadow state to avoid stale closures
        if (state) { }

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

            await refreshPipelineRuns(tenantInfo.current.project);

            let gitRepos = await gitClient.current.getRepositories(tenantInfo.current.project);
            const repoDetails = gitRepos.map(r => ({ id: r.id, defaultBranch: r.defaultBranch }));
            dispatch({ repositories: repoDetails });

            var alldoc = await getAllMergeQueuesDocument();
            if (alldoc === undefined) {
                // TODO: lock the app
                console.error("MQ: init -> failed to get all merge queues document");
                return;
            }

            // initialize main queue if it does not exist
            if ((alldoc.mergeQueues === undefined) ||
                (alldoc.mergeQueues.length === 0) ||
                (undefined === alldoc.mergeQueues.find(i => i.id === mainQueueTabId))
            ) {
                var mdoc = await getMergeQueueDocument(mainQueueTabId);
                if (!mdoc) {
                    // TODO: lock the app
                    console.error("MQ: init -> failed to get merge queue document for main queue");
                } else {
                    const targetRefName = mdoc.targetRefName || refMergeQueue;
                    let mq: MergeQueueInfo = {
                        id: mainQueueTabId,
                        name: "Main",
                        items: mdoc.mergeQueueItems || [],
                        targetRefName: targetRefName
                    };
                    mdoc.id = mainQueueTabId; // TODO: for upgrade only
                    mdoc.targetRefName = targetRefName;
                    await updateMergeQueueDocument(mdoc);

                    alldoc.mergeQueues = [{
                        id: mq.id,
                        name: mq.name
                    }];
                }
            }

            // TODO: remove demo queue
            // let demoQueue = alldoc.mergeQueues.find(i => i.id === demoQueueTabId);
            // if (!demoQueue) {
            //     var ddoc = await getMergeQueueDocument(demoQueueTabId);
            //     if (!ddoc) {
            //         console.error("MQ: init -> failed to get merge queue document for demo queue");
            //     } else {
            //         const targetRefName = ddoc.targetRefName || refDemoQueue;
            //         let mq: MergeQueueInfo = {
            //             id: demoQueueTabId,
            //             name: ddoc.name,
            //             items: [],
            //             targetRefName: targetRefName,
            //         };
            //         ddoc.id = demoQueueTabId;
            //         ddoc.name = "Demo";
            //         ddoc.targetRefName = targetRefName;
            //         await updateMergeQueueDocument(ddoc);

            //         alldoc.mergeQueues.push(mq);
            //     }
            // } else {
            //     var ddoc = await getMergeQueueDocument(demoQueueTabId);
            //     if (!ddoc) {
            //     }
            //     else {
            //         const targetRefName = ddoc.targetRefName || refDemoQueue;
            //         ddoc.name = "Demo";
            //         ddoc.targetRefName = targetRefName;
            //         await updateMergeQueueDocument(ddoc);

            //         demoQueue.name = "Demo";
            //     }
            // }

            // populate all merge queues
            let nextMergeQueues: MergeQueueInfo[] = [];
            for (const mqref of alldoc.mergeQueues) {
                let mqid = mqref.id;

                var mdoc = await getMergeQueueDocument(mqid);
                if (!mdoc) {
                    console.error(`MQ: init -> failed to get merge queue document for queue: ${mqid}`);
                    continue;
                }
                let mq: MergeQueueInfo = {
                    id: mqid,
                    name: mdoc.name,
                    items: mdoc.mergeQueueItems || [],
                    targetRefName: mdoc.targetRefName || refMergeQueue // TODO: remove default value
                };
                nextMergeQueues.push(mq);
            }

            console.log("MQ: init -> updating merge queues", nextMergeQueues);
            dispatch({ mergeQueues: nextMergeQueues });

            // update the all merge queues document with the latest names
            alldoc.mergeQueues = nextMergeQueues.map((mq): AllMergeQueuesItem => ({
                id: mq.id,
                name: mq.name,
            }));
            let nextAllDoc = await updateAllMergeQueuesDocument(alldoc);
            if (!nextAllDoc) {
                console.error("MQ: init -> failed to update all merge queues document");
            }
            console.log("MQ: init -> updated all merge queues document", nextAllDoc);

            var adoc = await getActivePullRequestDocument();
            if (adoc) {
                dispatch({ activePullRequests: adoc.pullRequests });
            }

            var pdoc = await getAppDocument<PipelineRunsDocument>(allPipelineRunsDocumentId);
            if (!pdoc) {
                console.error("MQ: refreshPipelineRuns -> failed to get pipeline runs document");
                return;
            }
            let pipelineRuns = pdoc.pipelineRuns ?? [];
            dispatch({ pipelineRuns: pipelineRuns });
            console.warn("MQ: init -> fetched pipeline runs document", pdoc);

            console.log("MQ: init -> done");
            dispatch({ didInit: true });
        } catch (error) {
            console.error("MQ: init -> error occurred", error);
        }

        await ticktock(); // run all logic on first load
    }

    // ticktock
    React.useEffect(() => {
        let id = setInterval(() => { ticktock(); }, 10000);
        return () => clearInterval(id);
    }, []);
    async function ticktock() {
        let state = {} // shadow state to avoid stale closures
        if (state) { }

        if (isTickTockRunning.current === true) {
            return;
        }
        try {
            isTickTockRunning.current = true;

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

            await refreshActivePullRequests(git, ti);
            await refreshPipelineRuns(proj);

            var alldoc = await getAllMergeQueuesDocument();
            if (alldoc) {
                let allMergeQueues = alldoc.mergeQueues ?? [];
                if (allMergeQueues.length > 0) {
                    for (const mq of allMergeQueues) {
                        await runMergeQueue(git, proj, mq.id);
                    }
                }
                else {
                    console.error("MQ: ticktock -> no merge queues to process2", alldoc);
                }
            }
            else {
                console.error("MQ: ticktock -> no merge queues to process1", alldoc);
            }
        } catch (error) {
            console.error("MQ: ticktock -> error occurred", error);
        } finally {
            isTickTockRunning.current = false;
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
        if (!gitPRs) {
            console.error("MQ: refreshActivePullRequests -> failed to fetch pull requests");
            return;
        }
        let newPullRequests: PullRequestInfo[] = [];
        for (const gitPR of gitPRs) {
            const repoId = gitPR.repository?.id;
            if (!repoId) { continue; }

            const repoName = gitPR.repository?.name;
            if (!repoName) { continue; }

            const sourceRefName = gitPR.sourceRefName;
            if (!sourceRefName) { continue; }

            if (gitPR.status === PullRequestStatus.Completed || gitPR.status === PullRequestStatus.Abandoned) {
                continue;
            }

            newPullRequests.push({
                id: gitPR.pullRequestId,
                title: gitPR.title,
                createdTimestamp: luxon.DateTime.fromJSDate(gitPR.creationDate).toUnixInteger(),
                isDraft: gitPR.isDraft,
                isAutoComplete: gitPR.autoCompleteSetBy !== undefined,
                sourceRefName: gitPR.sourceRefName,
                targetRefName: gitPR.targetRefName,
                repository: {
                    id: gitPR.repository.id,
                    name: gitPR.repository.name
                },
                author: {
                    displayName: gitPR.createdBy?.displayName || "Unknown User",
                    imageUrl: gitPR.createdBy?.imageUrl,
                    descriptor: gitPR.createdBy?.descriptor
                },
                voting: summarizeVotes(gitPR.reviewers),
                mergeStatus: gitPR.mergeStatus,
            });
        }

        dispatch({ activePullRequests: newPullRequests });

        // TODO: avoid redundant updates
        let adoc = await getActivePullRequestDocument();
        if (adoc) {
            adoc.pullRequests = newPullRequests;
            let adoc2 = await updateActivePullRequestDocument(adoc);
            if (!adoc2) {
                console.error("MQ: refreshActivePullRequests -> failed to update active pull request document");
            }
        }
    }

    async function runMergeQueue(gitClient: GitRestClient, project: string, queueId: string): Promise<void> {
        if (isMergeQueueRunning.current === true) {
            console.log("MQ: merge queue is already running");
            return;
        }
        console.log("MQ: running merge queue");
        try {
            isMergeQueueRunning.current = true;

            var sync_doc = await getMergeQueueDocument(queueId);
            if (!sync_doc) {
                console.error(`MQ: runMergeQueue ${queueId} -> failed to get merge queue document`);
                return;
            }
            if (!sync_doc.mergeQueueItems) {
                console.error(`MQ: runMergeQueue ${queueId} -> merge queue document has no items`, sync_doc);
                return;
            }
            const refForThisMergeQueue = sync_doc.targetRefName || "refs/error"; // TODO: remove default value
            console.log(`MQ: runMergeQueue ${queueId} -> starting`);

            // get unique repositories referenced by the merge queue items
            const sync_list = [...sync_doc.mergeQueueItems];
            const repos = sync_list.map(i => i.repository);
            const repoSet = distinctBy(repos, r => r.id);
            const gitRepos = await Promise.all(repoSet.map(r => gitClient.getRepository(r.id, project)));
            console.log(`MQ: runMergeQueue ${queueId} -> retrieved git repositories`, gitRepos);

            let repoBaseCommits: { repoId: string; baseCommitId: string, blockRepo: boolean }[] = [];
            for (const repo of gitRepos) {
                const commitId = await getDefaultBranchCommitId(gitClient, project, repo.id);
                if (!commitId) {
                    console.error(`MQ: runMergeQueue ${queueId} -> failed to get default branch commit ID for repository`, repo.id);
                    continue;
                }
                repoBaseCommits.push({ repoId: repo.id, baseCommitId: commitId, blockRepo: false });
            }

            let mustSync = false;
            let removed_indexes: number[] = [];
            for (const [index, sync_item] of sync_list.entries()) {
                if (mustSync) {
                    sync_doc.mergeQueueItems = sync_list;
                    sync_doc = await syncMergeQueueDocument(sync_doc);
                    if (!sync_doc) {
                        console.error(`MQ: runMergeQueue ${queueId} -> ${index}: failed to sync document`);
                        return;
                    }
                    mustSync = false;
                }
                // get pull request info
                const pr = await gitClient.getPullRequestById(sync_item.id, project);
                if (!pr) {
                    console.error(`MQ: runMergeQueue ${queueId} -> failed to get pull request`, sync_item.id);
                    sync_item.status = "blocked"; // TODO: error icon
                    continue;
                }
                if (pr.isDraft !== sync_item.isDraft) {
                    sync_item.isDraft = pr.isDraft;
                    mustSync = true;
                }
                if (pr.mergeStatus !== sync_item.mergeStatus) {
                    sync_item.mergeStatus = pr.mergeStatus;
                    mustSync = true;
                }
                let newVoting = summarizeVotes(pr.reviewers);
                if (newVoting.status !== sync_item.voting.status || newVoting.count !== sync_item.voting.count) {
                    sync_item.voting = newVoting;
                    mustSync = true;
                }

                if (pr.status !== PullRequestStatus.Active) {
                    console.log(`MQ: runMergeQueue ${queueId} -> pull request is abandoned or completed`, sync_item.id);
                    sync_item.status = "blocked"; // TODO: REMOVE
                    removed_indexes.push(index);
                    mustSync = true;
                    continue;
                }

                // calculate target commit
                const repoid = sync_item.repository.id;
                const targetCommitEntry = repoBaseCommits.find(r => r.repoId === repoid);
                const targetCommitId = targetCommitEntry?.baseCommitId;
                if (!targetCommitId) {
                    console.error(`MQ: runMergeQueue ${queueId} -> failed to find base commit for repository`, repoid);
                    continue;
                }

                // calculate source commit
                const sourceRefName = sync_item.sourceRefName;
                const sourceCommitId = await getRefCommitId(gitClient, project, repoid, sourceRefName);
                if (!sourceCommitId) {
                    console.error(`MQ: runMergeQueue ${queueId} -> failed to get source commit ID for repository`, repoid, sourceRefName);
                    continue;
                }

                // skip item if no changes are needed
                const status = sync_item.status ?? 'queued';
                const isQueuedStatus = status === 'queued';
                const isRecalculatingStatus = status === 'recalculating';
                const isConflictStatus = status === 'conflict';
                const isBlockedStatus = status === 'blocked';
                const isSameSourceCommit = sourceCommitId === sync_item.sourceCommitId;
                const isSameTargetCommit = targetCommitId === sync_item.targetCommitId;
                if (isQueuedStatus) { console.log(`MQ: runMergeQueue ${queueId} -> ${index}: queued`); }
                else if (isRecalculatingStatus) { console.log(`MQ: runMergeQueue ${queueId} -> ${index}: recalculating`); }
                else if (!isSameSourceCommit) { console.log(`MQ: runMergeQueue ${queueId} -> ${index}: new source commit`, sourceCommitId, sync_item.sourceCommitId); }
                else if (!isSameTargetCommit) { console.log(`MQ: runMergeQueue ${queueId} -> ${index}: new target commit`, targetCommitId, sync_item.targetCommitId); }
                else if (isConflictStatus) {
                    console.log(`MQ: runMergeQueue ${queueId} -> ${index}: same commits, but in conflict`);
                    sync_item.status = 'conflict';
                    resetStatusOfDependentPullRequests(sync_list, index, repoid, 'blocked');
                    targetCommitEntry.blockRepo = true;
                    continue;
                }
                else if (isBlockedStatus || targetCommitEntry.blockRepo) {
                    console.log(`MQ: runMergeQueue ${queueId} -> ${index}: same commits, but blocked`);
                    sync_item.status = 'blocked';
                    resetStatusOfDependentPullRequests(sync_list, index, repoid, 'blocked');
                    targetCommitEntry.blockRepo = true;
                    continue;
                }
                else {
                    // skip!
                    console.log(`MQ: runMergeQueue ${queueId} -> ${index}: same commits`, status, sync_item.title);
                    targetCommitEntry.baseCommitId = sync_item.mergedCommitId;
                    continue;
                }

                // invalidate subsequent merge queue items in the same repository
                resetStatusOfDependentPullRequests(sync_list, index, repoid, "queued");

                // checkout item
                console.log(`MQ: runMergeQueue ${queueId} -> ${index}: recalculating`, sourceCommitId, targetCommitId);
                sync_item.status = "recalculating";
                sync_item.sourceCommitId = sourceCommitId;
                sync_item.targetCommitId = targetCommitId;
                mustSync = true;

                // merge request
                let comment = `Merge Queue: !${sync_item.id} into ${targetCommitId}`;
                let mergeResult = await mergeCommits(gitClient, project, repoid, sourceCommitId, targetCommitId, comment);
                if (!mergeResult) {
                    console.error(`MQ: runMergeQueue ${queueId} -> ${index}: failed to merge commits 1`);
                    sync_item.status = 'blocked'; // TODO: error icon
                    resetStatusOfDependentPullRequests(sync_list, index, repoid, 'blocked');
                    targetCommitEntry.blockRepo = true;
                    mustSync = true;
                    continue;
                } else if (mergeResult.status === GitAsyncOperationStatus.Completed) {
                    let mergedCommitId = mergeResult.detailedStatus.mergeCommitId;
                    sync_item.status = 'valid';
                    sync_item.mergedCommitId = mergedCommitId;
                    targetCommitEntry.baseCommitId = mergedCommitId;
                } else {
                    console.error(`MQ: runMergeQueue ${queueId} -> ${index}: failed to merge commits 2`, mergeResult);
                    sync_item.status = 'conflict'; // TODO: check for conflicts
                    sync_item.mergedCommitId = zeroCommitId;
                    targetCommitEntry.blockRepo = true;
                    resetStatusOfDependentPullRequests(sync_list, index, repoid, 'blocked');
                }

                mustSync = true;
            }

            // removals, sort in reverse order to avoid index shifting
            if (removed_indexes.length > 0) {
                for (const index of removed_indexes.sort((a, b) => b - a)) {
                    sync_list.splice(index, 1);
                }

                mustSync = true;
            }

            // final sync
            if (mustSync) {
                sync_doc.mergeQueueItems = sync_list;
                sync_doc = await syncMergeQueueDocument(sync_doc);
                if (!sync_doc) {
                    console.error(`MQ: runMergeQueue ${queueId} -> final sync: failed to sync document`);
                    return;
                }
                mustSync = false;
            }

            // push to the merge-queue branch
            for (const repo of gitRepos) {
                let oldCommitId = await getRefCommitId(gitClient, project, repo.id, refForThisMergeQueue);
                if (!oldCommitId) {
                    oldCommitId = zeroCommitId;
                }

                const newCommitId = repoBaseCommits.find(r => r.repoId === repo.id)?.baseCommitId;
                if (!newCommitId) {
                    console.error(`MQ: runMergeQueue ${queueId} -> failed to find final commit for repository`, repo.id);
                    continue;
                }

                if (oldCommitId === newCommitId) {
                    console.log(`MQ: runMergeQueue ${queueId} -> no changes needed for repository`, repo.id, repo.name, newCommitId);
                    continue;
                }

                console.log(`MQ: runMergeQueue ${queueId} -> updating merge queue ref`, repo.id, repo.name, refForThisMergeQueue, oldCommitId, newCommitId);

                const refUpdate: GitRefUpdate = {
                    name: refForThisMergeQueue,
                    repositoryId: repo.id,
                    oldObjectId: oldCommitId,
                    newObjectId: newCommitId,
                    isLocked: undefined!, // HACK: API does not require this
                };
                let refResult = await gitClient.updateRefs([refUpdate], repo.id, project, project);
                console.log(`MQ: runMergeQueue ${queueId} -> updateRefs result`, refResult);
            }

            // TODO: derive this in the reducer instead?
            let mergeCommitIds = repoBaseCommits.map((i): MergeCommitInfo => ({ repositoryId: i.repoId, commitId: i.baseCommitId }));
            dispatch({ repoMergeCommits: mergeCommitIds });
        } catch (error) {
            console.error(`MQ: runMergeQueue ${queueId} -> error occurred`, error);
        } finally {
            console.log(`MQ: runMergeQueue ${queueId} -> finished`, queueId);
            isMergeQueueRunning.current = false;
        }
    }

    async function refreshPipelineRuns(project: string) {
        console.log("MQ: refreshPipelineRuns -> starting");
        try {
            let allBuilds = await runClient.current.getBuilds(
                project, // project
                undefined, // definitions
                undefined, // queues
                undefined, // buildNumber
                undefined, // minTime
                undefined, // maxTime
                undefined, // requestedFor
                undefined, // reasonFilter
                BuildStatus.All, // statusFilter
                undefined, // resultFilter
                undefined, // tagFilters
                undefined, // properties
                100, // top
                undefined, // continuationToken
                undefined, // maxBuildsPerDefinition
                undefined, // deletedFilter
                BuildQueryOrder.QueueTimeDescending, // queryOrder
                undefined, // branchName
                undefined, // buildIds
                undefined, // repositoryId
                undefined  // repositoryType
            );

            let pipelineRuns = allBuilds.map((run): PipelineRunInfo => ({
                runId: run.id || 0,
                pipelineId: run.definition.id || 0,
                runName: run.buildNumber,
                pipelineName: run.definition.name || "Unknown",
                sourceVersion: run.sourceVersion,
                status: getRunStatus(run.status),
                result: getRunResult(run.result),
                createdAt: run.queueTime ? luxon.DateTime.fromJSDate(run.queueTime || new Date()).toSeconds() : 0,
            }));

            dispatch({ pipelineRuns: pipelineRuns });
        } catch (error) {
            console.error("MQ: refreshPipelineRuns -> error occurred", error);
        }
    }

    // push new pipeline runs document
    React.useEffect(() => {
        if (state.didInit === false) {
            return;
        }
        let statePipelineRuns = state.pipelineRuns;
        console.log("MQ: refreshPipelineRuns -> updating pipeline runs document", statePipelineRuns, allPipelineRunsDocumentId);
        (async () => {
            console.log("MQ: refreshPipelineRuns -> updating pipeline runs document", statePipelineRuns);
            var doc = await getAppDocument<PipelineRunsDocument>(allPipelineRunsDocumentId);
            if (!doc) {
                console.error("MQ: refreshPipelineRuns -> failed to get pipeline runs document");
                return;
            }
            let old = doc.pipelineRuns ?? [];
            doc.pipelineRuns = statePipelineRuns;
            let doc2 = await updateDocument(extensionManagementClient.current, collectionId, allPipelineRunsDocumentId, doc);
            if (!doc2) {
                console.error("MQ: refreshPipelineRuns -> failed to update pipeline runs document");
                return;
            }
            console.log("MQ: refreshPipelineRuns -> updated pipeline runs document", old, doc2);
        })();
    }, [state.didInit, state.pipelineRuns]);

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

    async function getMergeQueueDocument(id: string): Promise<MergeQueueDocument | undefined> {
        let doc = await getAppDocument<MergeQueueDocument>(id);

        if (doc && doc.id && doc.name && doc.mergeQueueItems) {
            dispatch({
                singleMergeQueue: {
                    id: doc.id,
                    name: doc.name,
                    items: doc.mergeQueueItems,
                    targetRefName: doc.targetRefName || refMergeQueue // TODO: remove default value
                }
            });
        }

        return doc;
    }

    async function updateMergeQueueDocument(doc: MergeQueueDocument): Promise<MergeQueueDocument | undefined> {
        doc.viaVersion = "__MERGEQUEUEVERSION__";
        return await updateDocument(extensionManagementClient.current, collectionId, doc.id, doc);
    }

    async function getAllMergeQueuesDocument(): Promise<AllMergeQueuesDocument | undefined> {
        return await getAppDocument<AllMergeQueuesDocument>(allMergeQueuesDocumentId);
    }

    async function updateAllMergeQueuesDocument(doc: AllMergeQueuesDocument): Promise<AllMergeQueuesDocument | undefined> {
        doc.viaVersion = "__MERGEQUEUEVERSION__";
        return await updateDocument(extensionManagementClient.current, collectionId, doc.id, doc);
    }

    async function getActivePullRequestDocument(): Promise<PullRequestDocument | undefined> {
        return await getAppDocument<PullRequestDocument>(activePullRequestsDocumentId);
    }

    async function updateActivePullRequestDocument(doc: PullRequestDocument): Promise<PullRequestDocument | undefined> {
        doc.viaVersion = "__MERGEQUEUEVERSION__";
        return await updateDocument(extensionManagementClient.current, collectionId, activePullRequestsDocumentId, doc);
    }

    async function syncMergeQueueDocument(doc: MergeQueueDocument): Promise<MergeQueueDocument | undefined> {
        let doc2 = await updateMergeQueueDocument(doc);
        if (!doc2 || !doc2.mergeQueueItems) {
            return undefined;
        }
        // console.warn("MQ: runMergeQueue -> updated merge queue document", doc2);
        dispatch({
            singleMergeQueue: {
                id: doc2.id,
                name: doc2.name,
                items: doc2.mergeQueueItems,
                targetRefName: doc2.targetRefName || refMergeQueue // TODO: remove default value
            }
        })

        return doc2;
    }

    function renderPageCommandBarItems(): IHeaderCommandBarItem[] {
        return [{
            id: "addQueue",
            text: "New Queue",
            iconProps: {
                iconName: "Add",
            },
            onActivate: () => { onAddQueue(); },
            isPrimary: false,
            important: false,
            disabled: false
        }];
    }

    function renderMergeQueueCommandBarItems(): IHeaderCommandBarItem[] {
        let hasSelection = state.selectedPullRequestIds.length === 1;
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
                disabled: (state.selectedPullRequestIds.length === 0) // TODO: not at top
            },
            {
                id: "demote",
                iconProps: {
                    iconName: "Down"
                },
                onActivate: () => { onDemotePullRequest(); },
                isPrimary: false,
                important: true,
                disabled: (state.selectedPullRequestIds.length === 0) // TODO: not at bottom
            },
            {
                id: "dequeue",
                text: "Dequeue",
                onActivate: () => { onDequeuePullRequest(); },
                isPrimary: true,
                important: true,
                disabled: (state.selectedPullRequestIds.length === 0)
            }
        ];
    }

    async function onAddQueue() {
        dispatch({ showAddQueue: true })
    }

    function renderAllPullRequestsCommandBarItems(): IHeaderCommandBarItem[] {
        return [
            {
                id: "enqueue",
                text: "Enqueue",
                // onActivate: () => { onEnqueuePullRequest(); },
                isPrimary: true,
                important: true,
                subMenuProps: {
                    id: "enqueueSubMenu",
                    items: state.mergeQueues.map((mq): IMenuItem => ({
                        id: mq.id,
                        text: mq.name || mq.id, // TODO: WHY IS NAME NOT HERE?
                        onActivate: () => { onEnqueuePullRequest(mq.id); }
                    }))
                },
                disabled: (
                    state.selectedQueueTabId === allPullRequestsTabId &&
                    state.selectedPullRequestIds.length === 0
                ) // TODO: and not already in queue
            }
        ];
    }

    function resetStatusOfDependentPullRequests(array: MergeQueueItemInfo[], index: number, repoid: string, status: MergeQueueStatus) {
        for (let i2 = index + 1; i2 < array.length; i2++) {
            let itr_mqitem = array[i2];
            if (itr_mqitem.repository.id !== repoid) { continue; }

            console.log(`MQ: runMergeQueue -> ${i2}: invalidating`);
            array[i2] = {
                ...array[i2],
                status: status,
            };
        }
    }

    async function onPromotePullRequest() {
        if (!state.selectedPullRequestIds || state.selectedPullRequestIds.length !== 1) {
            // TODO: toast error
            return;
        }
        let queueId = state.selectedQueueTabId;
        let mq = state.mergeQueues.find(mq => mq.id === queueId);
        if (!mq) {
            // TODO: toast error
            return;
        }
        let mq_items = mq.items;
        let selectedId = state.selectedPullRequestIds[0];
        let selectedIndex = mq_items.findIndex(m => m.id === selectedId);
        if (selectedIndex === -1) {
            // TODO: toast error
            return;
        }

        // get remote doc
        let mdoc = await getMergeQueueDocument(queueId);
        if (!mdoc) {
            console.error("Failed to get merge queue document");
            return;
        }
        mq_items = mdoc.mergeQueueItems || [];
        if (selectedIndex !== mq_items.findIndex(m => m.id === selectedId)) {
            // TODO: toast error
            return;
        }
        if (selectedIndex === 0) {
            // TODO: toast error
            return;
        }

        // swap
        let up = mq_items[selectedIndex];
        let dn = mq_items[selectedIndex - 1];
        mq_items[selectedIndex] = dn;
        mq_items[selectedIndex - 1] = up;

        // requeue when items in the same repository change positions
        if (up.repository.id === dn.repository.id) {
            up.status = 'queued';
            resetStatusOfDependentPullRequests(mq_items, selectedIndex - 1, up.repository.id, 'queued');
        }

        // sync
        let updatedMdoc = await syncMergeQueueDocument(mdoc);
        if (!updatedMdoc) {
            console.error("MQ: onDequeuePullRequest -> failed to update merge queue document");
            return;
        }

        await ticktock(); // immediate refresh
    }

    async function onDemotePullRequest() {
        if (!state.selectedPullRequestIds || state.selectedPullRequestIds.length !== 1) {
            // TODO: toast error
            return;
        }
        let queueId = state.selectedQueueTabId;
        let mq = state.mergeQueues.find(mq => mq.id === queueId);
        if (!mq) {
            // TODO: toast error
            return;
        }
        let mq_items = mq.items;
        let selectedId = state.selectedPullRequestIds[0];
        let selectedIndex = mq_items.findIndex(m => m.id === selectedId);
        if (selectedIndex === -1) {
            // TODO: toast error
            return;
        }

        // get remote doc
        let mdoc = await getMergeQueueDocument(queueId);
        if (!mdoc) {
            console.error("Failed to get merge queue document");
            return;
        }
        mq_items = mdoc.mergeQueueItems || [];
        if (selectedIndex !== mq_items.findIndex(m => m.id === selectedId)) {
            // TODO: toast error
            return;
        }
        if (selectedIndex === mq_items.length - 1) {
            // TODO: toast error
            return;
        }

        // swap
        let dn = mq_items[selectedIndex];
        let up = mq_items[selectedIndex + 1];
        mq_items[selectedIndex] = up;
        mq_items[selectedIndex + 1] = dn;

        // requeue when items in the same repository change positions
        if (up.repository.id === dn.repository.id) {
            up.status = 'queued';
            resetStatusOfDependentPullRequests(mq_items, selectedIndex, dn.repository.id, 'queued');
        }

        // sync
        let updatedMdoc = await syncMergeQueueDocument(mdoc);
        if (!updatedMdoc) {
            console.error("MQ: onDemotePullRequest -> failed to update merge queue document");
            return;
        }

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

        let queueId = state.selectedQueueTabId;
        let mdoc = await getMergeQueueDocument(queueId);
        if (!mdoc) {
            console.error("Failed to get merge queue document");
            return;
        }
        let mitems = mdoc.mergeQueueItems || [];

        let oldids = state.selectedPullRequestIds;
        let newMergeQueueItems = mitems.filter(m => !oldids.includes(m.id));
        console.log("MQ: onDequeuePullRequest -> new pull requests", newMergeQueueItems);

        mdoc.mergeQueueItems = [...newMergeQueueItems];
        let updatedMdoc = await syncMergeQueueDocument(mdoc);
        if (!updatedMdoc) {
            console.error("MQ: onDequeuePullRequest -> failed to update merge queue document");
            return;
        }
        console.log("MQ: onDequeuePullRequest -> updated merge queue document", updatedMdoc);

        await ticktock(); // immediate refresh
    }

    async function onEnqueuePullRequest(queueId: string) {
        console.log("MQ: onEnqueuePullRequest -> starting");
        let adoc = await getActivePullRequestDocument();
        if (!adoc) {
            console.error("Failed to get active pull request document");
            return;
        }
        let aprs = adoc.pullRequests || [];
        dispatch({ activePullRequests: aprs });

        let alldoc = await getAllMergeQueuesDocument();
        if (!alldoc) {
            console.error("Failed to get all merge queues document");
            return;
        }
        if (alldoc.mergeQueues.find(mq => mq.id === queueId) === undefined) {
            console.error("Failed to find merge queue in all merge queues document", queueId);
            return;
        }

        let sync_doc = await getMergeQueueDocument(queueId);
        if (!sync_doc) {
            console.error("Failed to get merge queue document");
            return;
        }
        let sync_items = sync_doc.mergeQueueItems || [];

        let nextids = state.selectedPullRequestIds;
        let nextprs = aprs.filter(pr => nextids.includes(pr.id));
        console.log("MQ: onEnqueuePullRequest -> next pull requests", nextids, nextprs);

        // block if any prs are not going to the default branch
        if (!nextprs.every(pr => {
            let repo = state.repositories.find(r => r.id === pr.repository.id);
            return repo && repo.defaultBranch === pr.targetRefName;
        })) {
            console.error("MQ: onEnqueuePullRequest -> some pull requests are not going to the default branch");
            return;
        }

        // exclude nextprs that are already in the merge queue
        let filteredprs = nextprs.filter(pr => !sync_items.some(mpr => mpr.id === pr.id));
        console.log("MQ: onEnqueuePullRequest -> filtered pull requests", filteredprs);
        if (filteredprs.length === 0) {
            console.warn("MQ: onEnqueuePullRequest -> no pull requests to enqueue", sync_items);
            return;
        }

        // create new merge queue items
        var new_sync_items = filteredprs.map((pr): MergeQueueItemInfo => {
            return {
                id: pr.id,
                title: pr.title,
                repository: pr.repository,
                author: pr.author,
                status: "queued",
                createdTimestamp: pr.createdTimestamp,
                isDraft: pr.isDraft,
                isAutoComplete: pr.isAutoComplete,
                sourceRefName: pr.sourceRefName,
                targetRefName: pr.targetRefName,
                sourceCommitId: zeroCommitId,
                targetCommitId: zeroCommitId,
                mergedCommitId: zeroCommitId,
                voting: pr.voting,
                mergeStatus: pr.mergeStatus
            };
        });

        sync_doc.mergeQueueItems = [...sync_items, ...new_sync_items];
        sync_doc = await syncMergeQueueDocument(sync_doc);
        if (!sync_doc) {
            console.error("MQ: onEnqueuePullRequest -> failed to update merge queue document");
            return;
        }
        console.log("MQ: onEnqueuePullRequest -> updated merge queue document", sync_doc);

        await ticktock(); // immediate refresh
    }

    function onSelectMergeQueuePullRequestIds(ids: number[]) {
        dispatch({ selectedPullRequestIds: ids });
    }

    function onSelectActivePullRequestIds(ids: number[]) {
        dispatch({ selectedPullRequestIds: ids });
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
        let mq = state.mergeQueues.find(mq => mq.id === state.selectedQueueTabId);
        if (!mq) {
            return [];
        }
        return mq.items.map((item): PullRequestListItem => {
            let status = item.status;
            let icon = "Starburst";
            let iconClassName: string | undefined = undefined;
            if (status === "queued") {
                icon = "CircleRing";
            } else if (status === "conflict") {
                icon = "BlockedSolid";
                iconClassName = "color-red";
            } else if (status === "draft") {
                icon = "Edit";
                iconClassName = "color-pencil";
            } else if (status === "recalculating") {
                icon = "WorkFlow";
            } else if (status === "blocked") {
                icon = "ProgressLoopOuter";
                iconClassName = "color-red";
            } else {
                icon = "Starburst";
            }

            let dateString = item.createdTimestamp ? luxon.DateTime.fromSeconds(item.createdTimestamp).toRelative() || undefined : undefined;

            let isDefaultBranch = (state.repositories.find((r) => r.id === item.repository.id)?.defaultBranch || "") == item.targetRefName;

            return {
                icon: icon,
                iconClassName: iconClassName,
                pullRequestId: item.id,
                repository: item.repository.name,
                author: item.author,
                title: `${item.title}`,// - ${item.sourceCommitId} onto ${item.targetCommitId} is ${item.mergedCommitId}`,
                dateString: dateString,
                isDraft: item.isDraft,
                isAutoComplete: item.isAutoComplete,
                voting: item.voting || { status: "none", count: 0 },
                nonDefaultTargetBranch: isDefaultBranch ? null : item.targetRefName,
                mergeStatus: item.mergeStatus,
            };
        });
    }

    function mapActivePullRequestsToPullRequestListItems(): PullRequestListItem[] {
        return state.filteredActivePullRequests.map((pr): PullRequestListItem => {
            let dateString = pr.createdTimestamp ? luxon.DateTime.fromSeconds(pr.createdTimestamp).toRelative() || undefined : undefined;

            let isDefaultBranch = (state.repositories.find((r) => r.id === pr.repository.id)?.defaultBranch || "") == pr.targetRefName;

            return {
                icon: "CircleRing",
                pullRequestId: pr.id,
                repository: pr.repository.name,
                author: pr.author,
                title: pr.title,
                dateString: dateString,
                isDraft: pr.isDraft,
                isAutoComplete: pr.isAutoComplete,
                voting: pr.voting || { status: "none", count: 0 },
                nonDefaultTargetBranch: isDefaultBranch ? null : pr.targetRefName,
                mergeStatus: pr.mergeStatus,
            };
        });
    }

    async function saveUserFilters(value: PullRequestFilters) {
        dispatch({ filters: value });

        // let userFiltersDoc = { ...value };
        // userFiltersDoc = await Azdo.getOrCreateUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc);

        // userFiltersDoc.drafts = value.drafts;
        // userFiltersDoc.blocked = value.blocked;
        // userFiltersDoc.queued = value.queued;
        // userFiltersDoc.allBranches = value.allBranches;
        // userFiltersDoc.repositories = value.repositories;

        // Azdo.trySaveUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc);
    }

    function mapPipelinesForMergeQueue(): PipelineListItem[] {
        return state.filteredPipelineRuns.map((run) => ({
            runId: run.runId,
            runName: run.runName,
            pipelineId: run.pipelineId,
            pipelineName: run.pipelineName,
            sourceVersion: run.sourceVersion,
            status: run.status,
            result: run.result,
            createdAt: run.createdAt
        }));
    }

    async function onActivatePipeline(runId: number, pipelineId: number) {
        console.log("MQ: activating pipeline", runId, pipelineId);
        let run = await runClient.current.getBuild(tenantInfo?.current?.project || "", runId);
        if (run) {
            console.log("MQ: retrieved build", run, run._links.web.href);
            let url = run._links.web.href;
            if (url) {
                const navService = await SDK.getService<IHostNavigationService>("ms.vss-features.host-navigation-service");
                console.log("url: ", url);
                navService.openNewWindow(url, "");
            }
        }
    }

    function onSelectPipelines(runIds: number[]) {
        console.log("MQ: selecting pipelines", runIds);
        dispatch({ selectedPipelineRunIds: runIds });
    }

    function onSelectedQueueTabChanged(tabId: string) {
        console.log("MQ: selected tab changed", tabId);
        dispatch({ selectedQueueTabId: tabId });
    }

    async function onCommitAddMergeQueue() {
        dispatch({ showAddQueue: false });
    }

    async function onCancelAddMergeQueue() {
        dispatch({ showAddQueue: false });
    }

    function getAuthorImageUrl(author: AuthorInfo, _size: number): string | undefined {
        let org = tenantInfo.current?.organization;
        let desc = author.descriptor;
        if (!org || !desc) { return author.imageUrl; }
        return `https://vssps.dev.azure.com/${org}/_apis/graph/Subjects/${encodeURIComponent(desc)}/avatars?api-version=7.2-preview.1&format=png`
    }

    function getBadgeCountForTab(tabId: string): number | undefined {
        let mq = state.mergeQueues.find(mq => mq.id === tabId);
        if (!mq) {
            return undefined;
        }
        return mq.items.length;
    }

    return (
        <Page className="">
            <Header
                title="Merge Queue"
                titleSize={TitleSize.Large}
                commandBarItems={renderPageCommandBarItems()}
            />

            <TabBar
                tabSize={TabSize.Tall}
                selectedTabId={state.selectedQueueTabId}
                onSelectedTabChanged={onSelectedQueueTabChanged}
            >
                <Tab name="All" id={allPullRequestsTabId} />
                <Tab name="Main" id={mainQueueTabId} badgeCount={getBadgeCountForTab(mainQueueTabId)} />
                <Tab name="Demo" id={demoQueueTabId} badgeCount={getBadgeCountForTab(demoQueueTabId)} />
            </TabBar>

            {
                state.selectedQueueTabId !== allPullRequestsTabId && <>
                    <Card
                        className="padding-8 margin-8"
                        contentProps={{ contentPadding: false }}
                        titleProps={{ text: "Pull Requests", className: "", size: TitleSize.Medium }}
                        headerClassName=""
                        headerCommandBarItems={renderMergeQueueCommandBarItems()}
                    >
                        <PullRequestList
                            pullRequests={mapMergeQueueItemToPullRequestListItems()}
                            selectedIds={state.selectedPullRequestIds}
                            onSelectPullRequestIds={onSelectMergeQueuePullRequestIds}
                            onActivatePullRequest={onActivateMergeQueuePullRequest}
                            getAuthorImageUrl={getAuthorImageUrl}
                        />
                    </Card>

                    <Card
                        className="padding-8 margin-8"
                        contentProps={{ contentPadding: false }}
                        titleProps={{ text: "Pipelines", className: "", size: TitleSize.Medium }}
                        headerClassName=""
                        headerCommandBarItems={[]}
                    >
                        <PipelineList
                            pipelines={mapPipelinesForMergeQueue()}
                            selectedIds={state.selectedFilteredPipelineRunIds}
                            onSelectPipelines={onSelectPipelines}
                            onActivatePipeline={onActivatePipeline}
                        />
                    </Card>
                </>
            }

            {
                state.selectedQueueTabId === allPullRequestsTabId && <>

                    <div className="padding-8 flex-row rhythm-horizontal-8 flex-grow">
                        <div className="flex-grow" />
                        <Toggle
                            id="allBranches"
                            text="All Branches"
                            checked={state.filters.allBranches}
                            onChange={(_event, value) => { saveUserFilters({ ...state.filters, allBranches: value }) }}
                        />
                        <Toggle
                            id="drafts"
                            text="Drafts"
                            checked={state.filters.drafts}
                            onChange={(_event, value) => { saveUserFilters({ ...state.filters, drafts: value }) }}
                        />
                        <Toggle
                            id="blocked"
                            text="Blocked"
                            checked={state.filters.blocked}
                            onChange={(_event, value) => { saveUserFilters({ ...state.filters, blocked: value }) }}
                        />
                        <Toggle
                            id="Queued"
                            text="Queued"
                            checked={state.filters.queued}
                            onChange={(_event, value) => { saveUserFilters({ ...state.filters, queued: value }) }}
                        />
                    </div>

                    <Card
                        className="padding-8 margin-8"
                        contentProps={{ contentPadding: false }}
                        titleProps={{ text: "Pull Requests", className: "", size: TitleSize.Medium }}
                        headerClassName=""
                        headerCommandBarItems={renderAllPullRequestsCommandBarItems()}
                    >
                        <PullRequestList
                            pullRequests={mapActivePullRequestsToPullRequestListItems()}
                            selectedIds={state.selectedPullRequestIds}
                            onSelectPullRequestIds={onSelectActivePullRequestIds}
                            onActivatePullRequest={onActivateActivePullRequest}
                            getAuthorImageUrl={getAuthorImageUrl}
                        />
                    </Card>
                </>
            }

            <div className="text-neutral-30 flex-row padding-4">
                <div className="flex-grow"></div>
                <div onClick={() => { setShowIcons(!showIcons); }}>__MERGEQUEUEVERSION__</div>
            </div>

            <div>
                {AllIcons(showIcons)}
            </div>

            {
                state.isShowingAddQueue && (
                    <AddMergeQueuePanel
                        onCancel={() => { onCancelAddMergeQueue(); }}
                        onCommit={() => { onCommitAddMergeQueue(); }}
                    />
                )
            }
        </Page>
    );
}



export function AllIcons(showIcons: boolean) {
    return (
        showIcons && <div className="flex-column rhythm-vertical-16">
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