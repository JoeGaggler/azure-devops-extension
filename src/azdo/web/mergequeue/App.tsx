import React from "react";
import * as joe from "../lib.ts";
import * as SDK from 'azure-devops-extension-sdk';
import * as Azdo from '../azdo/azdo.ts';
import * as luxon from 'luxon'
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { Button } from "azure-devops-ui/Button";
import { Card } from "azure-devops-ui/Card";
import { Dropdown } from "azure-devops-ui/Dropdown";
import { DropdownMultiSelection } from "azure-devops-ui/Utilities/DropdownSelection";
import { Icon, IconSize } from "azure-devops-ui/Icon";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { ListSelection } from "azure-devops-ui/List";
import { Page } from "azure-devops-ui/Page";
import { Pill, PillVariant, PillSize } from "azure-devops-ui/Pill";
import { PillGroup } from "azure-devops-ui/PillGroup";
import { ScrollableList, IListItemDetails, ListItem } from "azure-devops-ui/List";
import { Toast } from "azure-devops-ui/Toast";
import { Toggle } from "azure-devops-ui/Toggle";
import { type IHostNavigationService } from 'azure-devops-extension-api';

interface AppSingleton {
    repositoryFilterDropdownMultiSelection: DropdownMultiSelection;
}

interface AppProps {
    bearerToken: string;
    appToken: string;
    singleton: AppSingleton;
}

interface PullRequestFilters {
    drafts: boolean;
    allBranches: boolean;
    repositories: string[];
}

interface MergeQueueList {
    queues: Array<MergeQueue>;
}

interface MergeQueue {
    pullRequests: Array<MergeQueuePullRequest>;
}

interface MergeQueuePullRequest extends SomePullRequest {
    ready: boolean
}

interface SomePullRequest {
    pullRequestId: number;
    status: string;
    repositoryName: string;
    title: string;
    targetRefName: string;
    isDraft: boolean;
    creationDate: string; // ISO date string
    autoComplete: boolean;
    mergeStatus: string;
    voteStatus: string;
}

interface AllPullRequests {
    pullRequests: Array<SomePullRequest>;
}

interface ToastState {
    message: string;
    visible: boolean;
    ref: React.RefObject<Toast>;
}

function App(p: AppProps) {
    // console.log("AppProps:", p);

    // Extension Document IDs
    let mergeQueueDocumentCollectionId = "mergeQueue";
    let mergeQueueListDocumentId = "mergeQueueList";
    let repoCacheDocumentId = "repoCache";
    let userPullRequestFiltersDocumentId = "userPullRequestFilters";

    // ui state
    const [selectedIds, setSelectedIds] = React.useState<number[]>([]);
    const [toastState, setToastState] = React.useState<ToastState>({ message: "Hi!", visible: false, ref: React.createRef() });

    // cached state
    const [tenantInfo, setTenantInfo] = React.useState<Azdo.TenantInfo>({});
    const [filters, setFilters] = React.useState<PullRequestFilters>({ drafts: false, allBranches: false, repositories: [] });
    const [repoMap, setRepoMap] = React.useState<Record<string, Azdo.Repo>>({});

    // state if the lists
    const [mergeQueueList, setMergeQueueList] = React.useState<MergeQueueList>({ queues: [{ pullRequests: [] }] }); // TODO use this
    const [allPullRequests, setAllPullRequests] = React.useState<AllPullRequests>({ pullRequests: [] });

    // HACK: force rerendering for server sync
    const [pollHack, setPollHack] = React.useState(Math.random());
    React.useEffect(() => { poll(); }, [pollHack]);
    function Resync() { setPollHack(Math.random()); }

    let selectedRepos: string[] = filters.repositories || [];

    // dropdown filter
    let filterRepoItems: IListBoxItem[] = [];
    for (let repoName in repoMap) {
        let repo = repoMap[repoName];
        if (repo) {
            filterRepoItems.push({
                id: repoName,
                text: repoName,
                iconProps: { iconName: "Code" },
                enforceSingleSelect: false,
            });
        }
    };
    joe.sortByString(filterRepoItems, (i: any) => i.text);
    joe.applySelections(p.singleton.repositoryFilterDropdownMultiSelection, filterRepoItems, (i: any) => i.id, selectedRepos);

    // rendering the all pull requests list
    let allFilteredPullRequests = allPullRequests.pullRequests.flatMap((pr) => {
        let repo = repoMap[pr.repositoryName];
        if (!repo) { return [] }
        if (pr.isDraft && !filters.drafts) { return [] } // filter out drafts if not enabled
        if (selectedRepos.length > 0 && !selectedRepos.some((item) => item == pr.repositoryName)) { return [] }
        let isDefaultBranch = ((pr.targetRefName == repo.defaultBranch) as boolean)
        if (!isDefaultBranch && !filters.allBranches) { return [] } // filter out non-default branches if not enabled
        return pr
    })
    joe.sortByNumber(allFilteredPullRequests, i => -i.pullRequestId || 0);

    // rebuild selections from state
    let primaryQueueItems: MergeQueuePullRequest[] = (mergeQueueList?.queues[0]?.pullRequests || []);
    let primaryQueueSelection = new ListSelection(true);
    let allSelection = new ListSelection(true);
    joe.applySelections(allSelection, allFilteredPullRequests, i => i.pullRequestId, selectedIds);
    joe.applySelections(primaryQueueSelection, primaryQueueItems, i => i.pullRequestId, selectedIds);

    function summarizeVotes(reviewers: any): string {
        let finalVote = 20;
        let voters = 0;
        for (let reviewer of reviewers) {
            if (reviewer?.vote) {
                let vote = reviewer?.vote || 0;
                if (!vote && !reviewer.isRequired) { continue; } // skip non-required reviewers with no vote
                finalVote = reviewer?.vote > 10 ? 10 : reviewer?.vote;
                if (!reviewer.isContainer) { voters++; } // count only non-container reviewers
            }
        }
        switch (finalVote) {
            case 20: return "none"
            case 10: return voters > 1 ? "approved" : "none" // require at least 2 voters to be tagged "approved"
            case 5: return voters > 1 ? "suggestions" : "none" // require at least 2 voters to be tagged "approved with suggestions"
            case 0: return "none"
            case -5: return "waiting"
            case -10: return "rejected"
            default: return "unknown";
        }
    }

    async function poll() {
        try {
            console.log("Polling...");

            if (!tenantInfo.organization || !tenantInfo.project) {
                console.warn("Tenant info not set, skipping poll.");
                return;
            }

            // UPDATE ALL PULL REQUESTS
            let pullRequests = await Azdo.getAllPullRequests(tenantInfo);
            if (!pullRequests) {
                console.warn("Unable to fetch all pull requests.");
                return;
            }

            console.log("Pull Requests value:", pullRequests);
            await refreshRepos(pullRequests);

            let pullRequestsArray: SomePullRequest[] = pullRequests
                .flatMap((p): (SomePullRequest | readonly SomePullRequest[]) => {
                    if (p.pullRequestId && p.repository && p.repository.name) {
                        return {
                            pullRequestId: p.pullRequestId,
                            repositoryName: p.repository.name,
                            title: p.title || "Untitled Pull Request",
                            status: p.status || "active",
                            targetRefName: p.targetRefName || "",
                            isDraft: p.isDraft || false,
                            creationDate: p.creationDate || "2000-01-01T01:01:01Z", // ISO date string
                            autoComplete: p.autoCompleteSetBy || false,
                            mergeStatus: p.mergeStatus || "notSet",
                            voteStatus: summarizeVotes(p.reviewers || [])
                        };
                    } else {
                        return [];
                    }
                });
            setAllPullRequests({
                ...allPullRequests,
                pullRequests: pullRequestsArray
            })

            // UPDATE MERGE QUEUE
            let didChange = false;
            let oldMergeQueueList = await downloadMergeQueuePullRequests();
            let queue = oldMergeQueueList.queues[0]; // TODO: support multiple queues
            let pullRequestList: MergeQueuePullRequest[] = (queue?.pullRequests || []);
            console.log("Old merge queue list:", oldMergeQueueList);

            let repoVisitedSet: Set<string> = new Set();
            let position = 1;
            let removedPullRequests: number[] = [];
            for (let pr of pullRequestList) {
                if (!pr.pullRequestId || !pr.repositoryName) {
                    console.warn("Invalid pull request:", pr);
                    continue;
                }

                // refresh the pull request details
                let pr2 = await Azdo.getPullRequest(p.bearerToken, tenantInfo, pr.repositoryName, pr.pullRequestId);
                if (pr2) {
                    console.log("Pull Request details:", pr2);
                    if (pr2.status == "completed" || pr2.status == "abandoned") {
                        console.log("Pull request is completed, removing from queue:", pr2);
                        removedPullRequests.push(pr.pullRequestId);
                        continue; // skip completed pull requests
                    }
                }
                if (pr2?.title && pr2.title !== pr.title) { pr.title = pr2.title; didChange = true; }
                if (pr2?.status && pr2.status !== pr.status) { pr.status = pr2.status; didChange = true; }
                if (pr2?.repositoryName && pr2.repositoryName !== pr.repositoryName) { pr.repositoryName = pr2.repositoryName; didChange = true; }
                if (pr2?.targetRefName && pr2.targetRefName !== pr.targetRefName) { pr.targetRefName = pr2.targetRefName; didChange = true; }
                if (pr2?.creationDate && pr2.creationDate !== pr.creationDate) { pr.creationDate = pr2.creationDate; /* didChange = true; */ } // HACK: fix precision issues
                if (pr2?.mergeStatus && pr2.mergeStatus !== pr.mergeStatus) { pr.mergeStatus = pr2.mergeStatus; didChange = true; }
                if (pr2?.voteStatus && pr2.voteStatus !== pr.voteStatus) { pr.voteStatus = pr2.voteStatus; didChange = true; }
                if (typeof pr2?.isDraft !== "undefined" && pr2.isDraft !== pr.isDraft) { pr.isDraft = pr2.isDraft; didChange = true; }
                if ((typeof pr2?.autoCompleteSetBy !== "undefined") !== pr.autoComplete) { pr.autoComplete = (typeof pr2?.autoCompleteSetBy !== "undefined"); didChange = true; }

                let isFirst = false
                if (!repoVisitedSet.has(pr.repositoryName)) {
                    repoVisitedSet.add(pr.repositoryName);
                    isFirst = true;
                }

                // update pull request statuses
                let statuses = (await Azdo.getPullRequestStatuses(p.bearerToken, tenantInfo, pr.repositoryName, pr.pullRequestId))
                    .value.filter((s: any) => s.context.genre == "pingmint" && s.context.name == "merge-queue");
                let status0 = statuses.length > 0 ? statuses[0] : null;

                let targetUrl = `https://dev.azure.com/${tenantInfo.organization}/${tenantInfo.project}/_apps/hub/pingmint.pingmint-extension.pingmint-pipeline-mergequeue#pr-${pr.pullRequestId}` // TODO: hashtag pr-id
                if (isFirst) {
                    pr.ready = true
                    let statusText = `Merge queue: ready at #${position}`;
                    if (pr2.status != "active") {
                        console.warn("Pull request is not active:", pr2);
                        continue; // skip non-active pull requests
                    }
                    else if (status0?.state != "succeeded" || status0.description != statusText) {
                        didChange = true;
                        Azdo.postPullRequestStatus(
                            p.bearerToken,
                            tenantInfo,
                            pr.repositoryName,
                            pr.pullRequestId,
                            "succeeded",
                            statusText,
                            targetUrl
                        );

                        for (let status of statuses) {
                            console.log("Deleting old status for pull request:", pr.pullRequestId, status.id);
                            Azdo.deletePullRequestStatus(
                                p.bearerToken,
                                tenantInfo,
                                pr.repositoryName,
                                pr.pullRequestId,
                                status.id)
                        }
                        console.log("Posting success status for pull request:", pr.pullRequestId, status0);
                    }
                } else {
                    pr.ready = false
                    let statusText = `Merge queue: waiting at #${position}`;
                    if (status0?.state != "pending" || status0.description != statusText) {
                        didChange = true;

                        Azdo.postPullRequestStatus(
                            p.bearerToken,
                            tenantInfo,
                            pr.repositoryName,
                            pr.pullRequestId,
                            "pending",
                            statusText,
                            targetUrl
                        );

                        for (let status of statuses) {
                            console.log("Deleting old status for pull request:", pr.pullRequestId, status.id);
                            Azdo.deletePullRequestStatus(
                                p.bearerToken,
                                tenantInfo,
                                pr.repositoryName,
                                pr.pullRequestId,
                                status.id)
                        }
                        console.log("Posting pending status for pull request:", pr.pullRequestId, status0);
                    }
                }

                isFirst = false;
                position++;
            }

            // remove completed pull requests
            if (pullRequestList.length != (pullRequestList = pullRequestList.filter(pr => !removedPullRequests.includes(pr.pullRequestId))).length) {
                didChange = true;
                console.log("Removed completed pull requests:", removedPullRequests);
            }

            // Update the primary queue items
            let newMergeQueueList: MergeQueueList = {
                ...oldMergeQueueList,
                queues: [{ ...queue, pullRequests: pullRequestList }] // TODO: support multiple queues
            }
            console.log("New merge queue list:", newMergeQueueList);

            // save
            if (!didChange) {
                return;
            } else if (await uploadMergeQueuePullRequests(newMergeQueueList)) {
                console.log("Merge queue list saved successfully");
                setMergeQueueList(newMergeQueueList);
            } else {
                console.warn("Failed to save the merge queue list");
            }
        } catch (err) {
            console.error("Error during polling:", err);
        }
    }

    // initialize the app
    React.useEffect(() => { init() }, []);
    async function init() {
        let info = await Azdo.getAzdoInfo();
        setTenantInfo(info)

        // setup merge queue list
        let newMergeQueueList = await downloadMergeQueuePullRequests();
        setMergeQueueList(newMergeQueueList);

        // setup user filters
        let userFiltersDoc: PullRequestFilters = {
            drafts: false,
            allBranches: false,
            repositories: []
        }
        userFiltersDoc = await Azdo.getOrCreateUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc)
        setFilters({ ...userFiltersDoc });

        // Refresh from server
        setInterval(() => { Resync(); }, 1000 * 20);
        Resync();
    }

    async function refreshRepos(value: Azdo.PullRequest[]) {
        let map = repoMap;

        let sharedMap = await Azdo.getOrCreateSharedDocument(mergeQueueDocumentCollectionId, repoCacheDocumentId, {});
        if (sharedMap) {
            console.log("Cached repositories:", sharedMap);
        }

        for (let pullRequest of value) {
            // TODO: cache expiration

            if (pullRequest.repository.name && map[pullRequest.repository.name]) {
                // already cached
                continue
            }

            if (sharedMap && sharedMap[pullRequest.repository.name]) {
                let repo = sharedMap[pullRequest.repository.name];
                if (repo && repo.id && repo.name && repo.defaultBranch) {
                    // copy cached repo to local map
                    map[pullRequest.repository.name] = repo;
                    continue
                }
            }

            if (pullRequest.repository && pullRequest.repository.name && pullRequest.repository.url) {
                let repo: Azdo.Repo = await Azdo.getAzdo(pullRequest.repository.url, p.bearerToken);
                console.log("Repo:", pullRequest.repository.name, repo);
                let newRepo: Azdo.Repo = {
                    id: repo.id,
                    name: repo.name,
                    defaultBranch: repo.defaultBranch,
                }
                map[pullRequest.repository.name] = newRepo;
                sharedMap[pullRequest.repository.name] = newRepo;
            } else {
                console.warn("No repository found for pull request:", pullRequest);
            }
        }

        setRepoMap(map);
        await Azdo.trySaveSharedDocument(mergeQueueDocumentCollectionId, repoCacheDocumentId, sharedMap);
    }

    async function saveUserFilters(value: PullRequestFilters) {
        setFilters(value);

        let userFiltersDoc = { ...value };
        userFiltersDoc = await Azdo.getOrCreateUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc);

        userFiltersDoc.drafts = value.drafts;
        userFiltersDoc.allBranches = value.allBranches;
        userFiltersDoc.repositories = value.repositories;

        Azdo.trySaveUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc);
    }

    function IsDefaultBranch(pullRequest: SomePullRequest): boolean {
        let repo = repoMap[pullRequest.repositoryName];
        if (!repo) {
            console.warn("No repository found for pull request:", pullRequest);
            return false;
        }
        return pullRequest.targetRefName === repo.defaultBranch;
    }

    function GetBranchName(pullRequest: SomePullRequest): string {
        return pullRequest?.targetRefName?.replace("refs/heads/", "") || "";
    }

    function GetIconForPullRequest(pullRequest: MergeQueuePullRequest): React.JSX.Element {
        if (pullRequest.ready) {
            return <Icon
                iconName="Starburst"
                size={IconSize.medium}
            />
        }
        return <div className="blank-icon-medium" />
    }

    function renderPills(pullRequest: SomePullRequest): React.JSX.Element {
        // TODO: approved by YOU!
        return <PillGroup className="padding-left-16 padding-right-16">
            {pullRequest.mergeStatus == "conflicts" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 192, green: 0, blue: 0 }}>Conflicts</Pill>)}
            {pullRequest.mergeStatus == "failure" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 192, green: 0, blue: 0 }}>Failure</Pill>)}
            {pullRequest.mergeStatus == "rejectedByPolicy" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 192, green: 0, blue: 0 }}>Policy</Pill>)}

            {pullRequest.voteStatus == "approved" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 64, green: 128, blue: 64 }}>Approved</Pill>)}
            {pullRequest.voteStatus == "suggestions" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 64, green: 64, blue: 128 }}>Suggestions</Pill>)}
            {pullRequest.voteStatus == "waiting" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 169, green: 154, blue: 60 }}>Waiting</Pill>)}
            {pullRequest.voteStatus == "rejected" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 192, green: 0, blue: 0 }}>Rejected</Pill>)}

            {pullRequest.isDraft && (<Pill size={PillSize.compact}>Draft</Pill>)}
            {pullRequest.autoComplete && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 92, green: 128, blue: 92 }}>Auto-Complete</Pill>)}

            {!IsDefaultBranch(pullRequest) && (<Pill size={PillSize.compact} variant={PillVariant.outlined}>{GetBranchName(pullRequest)}</Pill>)}
        </PillGroup>
    }

    function renderPullRequestRow(
        index: number,
        pullRequest: MergeQueuePullRequest,
        details: IListItemDetails<any>,
        key?: string
    ): React.JSX.Element {
        let extra = "";
        let className = `scroll-hidden flex-row flex-center rhythm-horizontal-8 flex-grow padding-4 ${extra}`;
        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}
            >
                <div className={className}>
                    {GetIconForPullRequest(pullRequest)}
                    <div className="font-size-m flex-row flex-center flex-shrink">
                        {index + 1}
                    </div>
                    <div className="font-size-m padding-left-8">{pullRequest.repositoryName}</div>
                    <div className="font-size-m italic text-neutral-70 text-ellipsis padding-left-8">{pullRequest.title}</div>
                    {renderPills(pullRequest)}
                    <div className="font-size-m flex-row flex-grow"><div className="flex-grow" />
                        <div>{(pullRequest.creationDate) ? (luxon.DateTime.fromISO(pullRequest.creationDate).toRelative()) : ""}</div>
                    </div>
                </div>
            </ListItem>
        );
    };

    function renderSomePullRequestRow(
        index: number,
        pullRequest: SomePullRequest,
        details: IListItemDetails<any>,
        key?: string
    ): React.JSX.Element {
        let extra = "";
        let className = `scroll-hidden flex-row flex-center rhythm-horizontal-8 flex-grow padding-4 ${extra}`;
        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}
            >
                <div className={className}>
                    <div className="font-size-m flex-row flex-center flex-shrink">
                        {index + 1}
                    </div>
                    <div className="font-size-m padding-left-8">{pullRequest.repositoryName}</div>
                    <div className="font-size-m italic text-neutral-70 text-ellipsis padding-left-8">{pullRequest.title}</div>
                    {renderPills(pullRequest)}
                    <div className="font-size-m flex-row flex-grow"><div className="flex-grow" />
                        <div>{(pullRequest.creationDate) ? (luxon.DateTime.fromISO(pullRequest.creationDate).toRelative()) : ""}</div>
                    </div>
                </div>
            </ListItem>
        );
    };

    async function activatePullRequest(_: any, evt: any) {
        console.log("activated pull request: ", evt);
        let idx = evt.index;
        let data: MergeQueuePullRequest = evt.data;
        console.log("activated pull request2: ", idx, data);
        const navService = await SDK.getService<IHostNavigationService>("ms.vss-features.host-navigation-service");
        let url = `https://dev.azure.com/${tenantInfo.organization}/${tenantInfo.project}/_git/${data.repositoryName}/pullrequest/${data.pullRequestId}`;
        console.log("url: ", url);
        navService.openNewWindow(url, "");
    }

    async function downloadMergeQueuePullRequests(): Promise<MergeQueueList> {
        console.log("Loading merge queue list...");
        let newMergeQueueList: MergeQueueList = {
            queues: [
                // maintain at least one queue
                {
                    pullRequests: []
                }
            ]
        }
        newMergeQueueList = await Azdo.getOrCreateSharedDocument(mergeQueueDocumentCollectionId, mergeQueueListDocumentId, newMergeQueueList)

        let queues: MergeQueue[]
        if (!newMergeQueueList.queues) {
            newMergeQueueList.queues = [];
        }
        queues = newMergeQueueList.queues;

        if (queues.length == 0) {
            queues = [
                {
                    pullRequests: []
                }
            ]
        }

        return newMergeQueueList;
    }

    async function uploadMergeQueuePullRequests(next: MergeQueueList): Promise<MergeQueueList | undefined> {
        console.log("Saving merge queue list:", next);
        return await Azdo.trySaveSharedDocument(mergeQueueDocumentCollectionId, mergeQueueListDocumentId, next)
    }

    async function enqueuePullRequests() {
        if (selectedIds.length == 0) {
            showToast("No pull requests selected to enqueue.");
            return;
        }
        if (selectedIds.length != 1) {
            showToast("Only one pull request can be enqueued at a time.");
            return;
        }
        let pullRequestId = selectedIds[0];
        console.log("Enqueuing pull request ID:", pullRequestId);

        let oldMergeQueueList = await downloadMergeQueuePullRequests();
        console.log("Old merge queue list:", oldMergeQueueList);

        let queue = oldMergeQueueList.queues[0]; // TODO: support multiple queues
        let pullRequests = queue.pullRequests
        if (pullRequests.findIndex(pr => pr.pullRequestId == pullRequestId) >= 0) {
            showToast(`Pull request ID: ${pullRequestId} is already in the queue.`);
            return;
        }

        let pullRequest = allFilteredPullRequests.find(pr => pr.pullRequestId == pullRequestId);
        if (!pullRequest) {
            showToast(`Pull request ID: ${pullRequestId} not found in all pull requests.`);
            return;
        }
        let repositoryName = pullRequest.repositoryName
        if (!repositoryName) {
            showToast(`Pull request ID: ${pullRequestId} has no repository name.`);
            return;
        }

        pullRequests.push({
            pullRequestId: pullRequestId,
            repositoryName: repositoryName,
            title: pullRequest.title,
            status: pullRequest.status,
            targetRefName: pullRequest.targetRefName,
            isDraft: pullRequest.isDraft,
            creationDate: pullRequest.creationDate || "",
            ready: false,
            autoComplete: pullRequest.autoComplete || false,
            mergeStatus: pullRequest.mergeStatus,
            voteStatus: pullRequest.voteStatus,
        })

        let newMergeQueueList: MergeQueueList = {
            ...oldMergeQueueList,
            queues: [{ ...queue, pullRequests: pullRequests }] // TODO: support multiple queues
        };
        console.log("New merge queue list:", newMergeQueueList);

        // save
        if (!await uploadMergeQueuePullRequests(newMergeQueueList)) {
            showToast("Failed to save the merge queue list.");
            return;
        }
        setMergeQueueList(newMergeQueueList);

        // post pull request status
        await Azdo.postPullRequestStatus(
            p.bearerToken,
            tenantInfo,
            pullRequest.repositoryName,
            pullRequestId,
            "pending",
            "Pending on merge queue", // TODO: position in queue
            `https://dev.azure.com/${tenantInfo.organization}/${tenantInfo.project}/_apps/hub/pingmint.pingmint-extension.pingmint-pipeline-mergequeue#pr-${pullRequestId}` // TODO: hashtag pr-id
        );

        console.log("Pull request status posted for ID:", pullRequestId);
    }

    async function removePullRequests() {
        if (selectedIds.length == 0) {
            showToast("No pull requests selected to remove.");
            return;
        }
        if (selectedIds.length != 1) {
            showToast("Only one pull request can be removed at a time.");
            return;
        }
        let pullRequestId = selectedIds[0];
        console.log("Removing pull request ID:", pullRequestId);

        let oldMergeQueueList = await downloadMergeQueuePullRequests();
        console.log("Old merge queue list:", oldMergeQueueList);

        let queue = oldMergeQueueList.queues[0]; // TODO: support multiple queues
        let pullRequests = queue.pullRequests
        let pullRequestIndex = pullRequests.findIndex(pr => pr.pullRequestId == pullRequestId);
        if (pullRequestIndex < 0) {
            showToast(`Pull request ID: ${pullRequestId} no longer in queue.`);
            return;
        }
        let pullRequestRepoName = pullRequests[pullRequestIndex].repositoryName;
        pullRequests.splice(pullRequestIndex, 1); // remove the pull request

        let newMergeQueueList: MergeQueueList = {
            ...oldMergeQueueList,
            queues: [{ ...queue, pullRequests: pullRequests }] // TODO: support multiple queues
        };
        console.log("New merge queue list:", newMergeQueueList);

        // save
        if (!await uploadMergeQueuePullRequests(newMergeQueueList)) {
            showToast("Failed to save the merge queue list.");
            return;
        }
        setMergeQueueList(newMergeQueueList);

        await Azdo.deletePullRequestStatuses(p.bearerToken, tenantInfo, pullRequestRepoName, pullRequestId);
    }

    function showToast(message: string) {
        if (!toastState.visible) {
            setToastState({ ...toastState, message: message, visible: true });
            setTimeout(() => {
                toastState.ref.current?.fadeOut();
                setTimeout(() => {
                    setToastState({ ...toastState, visible: false });
                }, 1000);
            }, 3000);
        }
    }

    function onSelectPullRequests(list: SomePullRequest[], listSelection: ListSelection) {
        if (list.length == 0) {
            setSelectedIds([]);
            return;
        }

        let pids: number[] = []
        for (let selRange of listSelection.value) {
            for (let i = selRange.beginIndex; i <= selRange.endIndex; i++) {
                let pr = list[i];
                if (pr && pr.pullRequestId) {
                    pids.push(pr.pullRequestId);
                }
            }
        }
        setSelectedIds(pids)
        console.log("Selected pull request IDs:", pids);
    }

    function onSelectRepository(list: IListBoxItem[], listSelection: ListSelection) {
        let newSelectedRepos: string[] = [];

        for (let i = 0; i < listSelection.value.length; i++) {
            let selRange = listSelection.value[i];
            for (let j = selRange.beginIndex; j <= selRange.endIndex; j++) {
                let item = list[j];
                if (item && item.id) {
                    newSelectedRepos.push(item.id);
                    console.log("Selected repository:", item.id);
                }
            }
        }

        console.log("New Selected repositories:", newSelectedRepos);
        saveUserFilters({ ...filters, repositories: newSelectedRepos })
    }

    async function onPromotePullRequest() {
        if (selectedIds.length == 0) {
            showToast("No pull requests selected to remove.");
            return;
        }
        if (selectedIds.length != 1) {
            showToast("Only one pull request can be removed at a time.");
            return;
        }
        let pullRequestId = selectedIds[0];
        console.log("Promoting pull request ID:", pullRequestId);

        let oldMergeQueueList = await downloadMergeQueuePullRequests();
        console.log("Old merge queue list:", oldMergeQueueList);

        let queue = oldMergeQueueList.queues[0]; // TODO: support multiple queues
        let pullRequests = queue.pullRequests
        let pullRequestIndex = pullRequests.findIndex(pr => pr.pullRequestId == pullRequestId);
        if (pullRequestIndex < 0) {
            showToast(`Pull request ID: ${pullRequestId} no longer in queue.`);
            return;
        }
        if (pullRequestIndex == 0) {
            showToast(`Pull request ID: ${pullRequestId} is already at the start of the queue.`);
            return;
        }
        let pullRequestRepoName = pullRequests[pullRequestIndex].repositoryName;
        let tmp = pullRequests[pullRequestIndex - 1];
        pullRequests[pullRequestIndex - 1] = pullRequests[pullRequestIndex];
        pullRequests[pullRequestIndex] = tmp;

        let newMergeQueueList: MergeQueueList = {
            ...oldMergeQueueList,
            queues: [{ ...queue, pullRequests: pullRequests }] // TODO: support multiple queues
        };
        console.log("New merge queue list:", newMergeQueueList);

        // save
        if (!await uploadMergeQueuePullRequests(newMergeQueueList)) {
            showToast("Failed to save the merge queue list.");
            return;
        }
        setMergeQueueList(newMergeQueueList);

        await Azdo.deletePullRequestStatuses(p.bearerToken, tenantInfo, pullRequestRepoName, pullRequestId);
    }

    async function onDemotePullRequest() {
        if (selectedIds.length == 0) {
            showToast("No pull requests selected to remove.");
            return;
        }
        if (selectedIds.length != 1) {
            showToast("Only one pull request can be removed at a time.");
            return;
        }
        let pullRequestId = selectedIds[0];
        console.log("Demoting pull request ID:", pullRequestId);

        let oldMergeQueueList = await downloadMergeQueuePullRequests();
        console.log("Old merge queue list:", oldMergeQueueList);

        let queue = oldMergeQueueList.queues[0]; // TODO: support multiple queues
        let pullRequests = queue.pullRequests
        let pullRequestIndex = pullRequests.findIndex(pr => pr.pullRequestId == pullRequestId);
        if (pullRequestIndex < 0) {
            showToast(`Pull request ID: ${pullRequestId} no longer in queue.`);
            return;
        }
        if (pullRequestIndex + 1 >= pullRequests.length) {
            showToast(`Pull request ID: ${pullRequestId} is already at the end of the queue.`);
            return;
        }
        let pullRequestRepoName = pullRequests[pullRequestIndex].repositoryName;
        let tmp = pullRequests[pullRequestIndex + 1];
        pullRequests[pullRequestIndex + 1] = pullRequests[pullRequestIndex];
        pullRequests[pullRequestIndex] = tmp;

        let newMergeQueueList: MergeQueueList = {
            ...oldMergeQueueList,
            queues: [{ ...queue, pullRequests: pullRequests }] // TODO: support multiple queues
        };
        console.log("New merge queue list:", newMergeQueueList);

        // save
        if (!await uploadMergeQueuePullRequests(newMergeQueueList)) {
            showToast("Failed to save the merge queue list.");
            return;
        }
        setMergeQueueList(newMergeQueueList);

        await Azdo.deletePullRequestStatuses(p.bearerToken, tenantInfo, pullRequestRepoName, pullRequestId);
    }

    return (
        <Page>
            {toastState.visible && <Toast message={toastState.message} ref={toastState.ref} />}
            <div className="padding-8 margin-8">

                <div className="padding-8 flex-row flex-baseline rhythm-horizontal-16">
                    <h2>Merge Queue</h2>
                    <div className="flex-grow"></div>
                    <div className="padding-8 flex-row flex-stretch rhythm-horizontal-16">
                        <Button
                            disabled={false} // TODO: validation
                            onClick={onPromotePullRequest}
                        >
                            <Icon iconName="Up" size={IconSize.medium} />
                        </Button>
                        <Button
                            disabled={false} // TODO: validation
                            onClick={onDemotePullRequest}
                        >
                            <Icon iconName="Down" size={IconSize.medium} />
                        </Button>
                        <Button
                            text="Remove"
                            danger={true}
                            disabled={false} // TODO: validation
                            onClick={removePullRequests}
                        />
                    </div>
                </div>
                <Card className="padding-8">
                    <div className="flex-column">
                        <ScrollableList
                            itemProvider={new ArrayItemProvider(primaryQueueItems)}
                            selection={primaryQueueSelection}
                            onSelect={(_evt, _listRow) => { onSelectPullRequests(primaryQueueItems, primaryQueueSelection); }}
                            onActivate={activatePullRequest}
                            renderRow={renderPullRequestRow}
                            width="100%"
                        />
                    </div>
                </Card>

                <br />

                <div className="padding-8 flex-row flex-center rhythm-horizontal-16">
                    <h2>All Pull Requests</h2>
                    <div className="flex-grow"></div>
                    <Dropdown
                        actions={[
                            {
                                className: "bolt-dropdown-action-right-button",
                                disabled: false, // TODO
                                iconProps: { iconName: "Select All" },
                                text: "All",
                                onClick: () => {
                                    saveUserFilters({ ...filters, repositories: filterRepoItems.map(item => item.id) })
                                }
                            },
                            {
                                className: "bolt-dropdown-action-right-button",
                                disabled: false, // TODO
                                iconProps: { iconName: "Clear" },
                                text: "Clear",
                                onClick: () => {
                                    saveUserFilters({ ...filters, repositories: [] })
                                }
                            }
                        ]}
                        className="repoDropdown"
                        items={filterRepoItems}
                        selection={p.singleton.repositoryFilterDropdownMultiSelection}
                        placeholder="Repositories"
                        showFilterBox={true}
                        onSelect={(_evt, _row) => { onSelectRepository(filterRepoItems, p.singleton.repositoryFilterDropdownMultiSelection); }}
                    />
                    <Toggle
                        offText={"All branches"}
                        onText={"All branches"}
                        checked={filters.allBranches}
                        onChange={(_event, value) => { saveUserFilters({ ...filters, allBranches: value }) }}
                    />
                    <Toggle
                        offText={"Drafts"}
                        onText={"Drafts"}
                        checked={filters.drafts}
                        onChange={(_event, value) => { saveUserFilters({ ...filters, drafts: value }) }}
                    />
                    <Button
                        text="Enqueue"
                        primary={true}
                        disabled={false} // TODO: validation
                        onClick={enqueuePullRequests}
                    />
                </div>
                <Card className="padding-8">
                    <div className="flex-column">
                        <ScrollableList
                            itemProvider={new ArrayItemProvider(allFilteredPullRequests)}
                            selection={allSelection}
                            onSelect={(_evt, _listRow) => { onSelectPullRequests(allFilteredPullRequests, allSelection); }}
                            onActivate={activatePullRequest}
                            renderRow={renderSomePullRequestRow}
                            width="100%"
                        />
                    </div>
                </Card>
            </div>
        </Page>
    )
}

export { App };
export type { AppProps, AppSingleton };
