import React from "react";
import * as SDK from 'azure-devops-extension-sdk';
import * as Azdo from '../azdo/azdo.ts';
import * as luxon from 'luxon'
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { Button } from "azure-devops-ui/Button";
import { Card } from "azure-devops-ui/Card";
// import { List } from "azure-devops-ui/List";
import { ListSelection } from "azure-devops-ui/List";
import { Pill, PillVariant } from "azure-devops-ui/Pill";
import { PillGroup } from "azure-devops-ui/PillGroup";
import { ScrollableList, IListItemDetails, ListItem } from "azure-devops-ui/List";
import { Status, Statuses, StatusSize } from "azure-devops-ui/Status";
import { Toast } from "azure-devops-ui/Toast";
import { Toggle } from "azure-devops-ui/Toggle";
import { type IHostNavigationService } from 'azure-devops-extension-api';

interface AppProps {
    bearerToken: string | null;
    appToken: string | null;
}

interface PullRequestFilters {
    drafts: boolean;
    allBranches: boolean;
}

interface MergeQueueList {
    queues: Array<MergeQueue>;
}

interface MergeQueue {
    pullRequests: Array<any>;
}

interface ToastState {
    message: string;
    visible: boolean;
    ref: React.RefObject<Toast>;
}

function App(p: AppProps) {
    console.log("AppProps:", p);

    // Extension Document IDs
    let mergeQueueDocumentCollectionId = "mergeQueue";
    let mergeQueueListDocumentId = "mergeQueueList";
    let repoCacheDocumentId = "repoCache";
    let userPullRequestFiltersDocumentId = "userPullRequestFilters";

    const [tenantInfo, setTenantInfo] = React.useState<Azdo.TenantInfo>({});
    const [allPullRequests, setAllPullRequests] = React.useState<Array<Azdo.PullRequest>>([]);
    const [filters, setFilters] = React.useState<PullRequestFilters>({ drafts: false, allBranches: false });
    const [selectedIds, setSelectedIds] = React.useState<number[]>([]); // TODO: use this

    const [repoMap, setRepoMap] = React.useState<Record<string, Azdo.Repo>>({});
    const [_mergeQueueList, setMergeQueueList] = React.useState<MergeQueueList>({ queues: [] }); // TODO use this
    const [toastState, setToastState] = React.useState<ToastState>({ message: "Hi!", visible: false, ref: React.createRef() });

    // rendering the all pull requests list
    // TODO: useMemo?
    let allFilteredPullRequests = allPullRequests.flatMap((pr) => {
        let repo = repoMap[pr.repository.name];
        if (!repo) { return [] }
        return {
            ...pr,
            isDefaultBranch: ((pr.targetRefName == repo.defaultBranch) as boolean),
            targetBranch: (pr.targetRefName ?? "").replace("refs/heads/", "")
        }
    }).filter(pr =>
        (pr.isDefaultBranch || filters.allBranches) &&
        (!pr.isDraft || (filters.drafts as boolean))
    );
    allFilteredPullRequests.sort((a, b) => {
        let x = a.pullRequestId || 0;
        let y = b.pullRequestId || 0;
        if (x < y) { return 1; }
        else if (x > y) { return -1; }
        else { return 0; }
    })

    // rebuild selections from state
    let allSelection = new ListSelection(true);
    for (let i = 0; i < allFilteredPullRequests.length; i++) {
        let pr = allFilteredPullRequests[i];
        if (selectedIds.includes(pr.pullRequestId || 0)) {
            allSelection.select(i, 1, true, true);
        }
    }
    // TODO: if changing the filter hides a selected pull request, we should update the selection to remove it

    // initialize the app
    React.useEffect(() => { init() }, []);
    async function init() {
        let info = await Azdo.getAzdoInfo();
        setTenantInfo(info)

        let pullRequests = await Azdo.getAllPullRequests(info);
        console.log("Pull Requests value:", pullRequests);

        refreshRepos(pullRequests); // not awaited

        setAllPullRequests(pullRequests);

        // setup merge queue list
        let newMergeQueueList: MergeQueueList = {
            queues: [
                // maintain at least one queue
                {
                    pullRequests: []
                }
            ]
        }
        newMergeQueueList = await Azdo.getOrCreateSharedDocument(mergeQueueDocumentCollectionId, mergeQueueListDocumentId, newMergeQueueList)
        setMergeQueueList(newMergeQueueList);
        // newMergeQueueList = { // TODO: REMOVE THIS
        //     queues: [
        //         {
        //             pullRequests: [
        //                 // HACK: for demo purposes, we will just take the first 5 pull requests
        //                 pullRequests.value[0],
        //                 pullRequests.value[1],
        //                 pullRequests.value[2],
        //                 pullRequests.value[3],
        //                 pullRequests.value[4]
        //             ]
        //         }
        //     ]
        // }

        // setup user filters
        let userFiltersDoc = {
            drafts: false,
            allBranches: false
        }
        userFiltersDoc = await Azdo.getOrCreateUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc)
        setFilters({ ...userFiltersDoc });
    }

    // TODO: FIX THIS
    // function getPrimaryPullRequests(): Array<any> {
    //     if (mergeQueueList && mergeQueueList.queues && mergeQueueList.queues.length > 0) {
    //         return mergeQueueList.queues[0].pullRequests;
    //     }
    //     return [];
    // }

    async function refreshRepos(value: any) {
        let map = repoMap;

        let sharedMap = await Azdo.getOrCreateSharedDocument(mergeQueueDocumentCollectionId, repoCacheDocumentId, {});
        if (sharedMap) {
            console.log("Shared repo map:", sharedMap);
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
                let repo: Azdo.Repo = await Azdo.getAzdo(pullRequest.repository.url, p.bearerToken as string);
                console.log("Repo:", pullRequest.repository.name, repo);
                let newRepo: Azdo.Repo = {
                    id: repo.id,
                    name: repo.name,
                    defaultBranch: repo.defaultBranch,
                }
                map[pullRequest.repository.name] = newRepo;
                sharedMap[pullRequest.repository.name] = newRepo;
            } else {
                // console.warn("No repository found for pull request:", pullRequest);
            }
        }

        setRepoMap(map);
        Azdo.trySaveSharedDocument(mergeQueueDocumentCollectionId, repoCacheDocumentId, sharedMap);

        console.log("Repo map:", map);
        console.log("Shared map:", sharedMap);
    }

    async function persistFilters(value: PullRequestFilters) {
        setFilters(value);

        let userFiltersDoc = { ...value };
        userFiltersDoc = await Azdo.getOrCreateUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc);

        userFiltersDoc.drafts = value.drafts;
        userFiltersDoc.allBranches = value.allBranches;

        Azdo.trySaveUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc);
    }

    function renderPullRequestRow(
        index: number,
        pullRequest: any,
        details: IListItemDetails<any>,
        key?: string
    ): React.JSX.Element {
        let extra = "";
        let className = `scroll-hidden flex-row flex-center flex-grow padding-4 ${extra}`;
        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}
            >
                <div className={className}>
                    <Status
                        {...(pullRequest.isDraft ? Statuses.Queued : Statuses.Information)}
                        key="information"
                        size={StatusSize.m}
                    />
                    <div className="font-size-m padding-left-8">{pullRequest.repository.name}</div>
                    <div className="font-size-m italic text-neutral-70 text-ellipsis padding-left-8">{pullRequest.title}</div>
                    <PillGroup className="padding-left-16 padding-right-16">
                        {
                            pullRequest.isDraft && (
                                <Pill>Draft</Pill>
                            )
                        }
                        {
                            !pullRequest.isDefaultBranch && pullRequest.targetBranch && (
                                <Pill variant={PillVariant.outlined}>{pullRequest.targetBranch}</Pill>
                            )
                        }
                    </PillGroup>
                    <div className="font-size-m flex-row flex-grow"><div className="flex-grow" />
                        <div>{luxon.DateTime.fromISO(pullRequest.creationDate).toRelative()}</div>
                    </div>
                </div>
            </ListItem>
        );
    };

    async function activatePullRequest(_: any, evt: any) {
        console.log("activated pull request: ", evt);
        let idx = evt.index;
        let data = evt.data;
        console.log("activated pull request2: ", idx, data);
        const navService = await SDK.getService<IHostNavigationService>("ms.vss-features.host-navigation-service");
        let url = `https://dev.azure.com/${tenantInfo.organization}/${tenantInfo.project}/_git/${data.repository.name}/pullrequest/${data.pullRequestId}`;
        console.log("url: ", url);
        navService.openNewWindow(url, "");
    }

    async function enqueuePullRequests() {
        let list = selectedIds
        console.log("Enqueue pull request:", list);
        showToast(`Enqueuing pull request: ${list.join(", ")}`);
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

    return (
        <>
            {toastState.visible && <Toast message={toastState.message} ref={toastState.ref} />}
            <div className="padding-8 margin-8">

                <h2>Merge Queue</h2>
                <Card className="padding-8">
                    {
                        /*
                        <ButtonGroup className="flex-wrap">
                            <Button
                                text="New Train"
                                onClick={() => alert("TODO: Create a new release train")}
                            />
                        </ButtonGroup> 
                        */
                    }

                    {
                        /*** TODO
                        <PullRequestList
                        pullRequests={getPrimaryPullRequests()}
                        organization={tenantInfo.organization}
                        project={tenantInfo.project}
                        selection={new ListSelection(true)} // TODO: THIS IS WRONG
                        />
                        */
                    }
                </Card>

                <br />

                <h2>All Pull Requests</h2>
                <Card className="padding-8">
                    <div className="flex-column">
                        <div className="padding-8 flex-row rhythm-horizontal-16">
                            <Toggle
                                offText={"All branches"}
                                onText={"All branches"}
                                checked={filters.allBranches}
                                onChange={(_event, value) => { persistFilters({ ...filters, allBranches: value }) }}
                            />
                            <Toggle
                                offText={"Drafts"}
                                onText={"Drafts"}
                                checked={filters.drafts}
                                onChange={(_event, value) => { persistFilters({ ...filters, drafts: value }) }}
                            />
                            <div className="flex-grow"></div>
                            <Button
                                text="Enqueue"
                                primary={true}
                                disabled={false} // TODO: validation
                                onClick={enqueuePullRequests}
                            />
                        </div>
                        <ScrollableList
                            itemProvider={new ArrayItemProvider(allFilteredPullRequests)}
                            selection={allSelection}
                            onSelect={(_evt, _listRow) => {
                                console.log("Event Selection changed:", _evt, _listRow);

                                let list = allFilteredPullRequests;
                                if (list.length == 0) {
                                    setSelectedIds([]);
                                    return;
                                }

                                let pids: number[] = []
                                for (let selRange of allSelection.value) {
                                    for (let i = selRange.beginIndex; i <= selRange.endIndex; i++) {
                                        let pr = list[i];
                                        if (pr && pr.pullRequestId) {
                                            pids.push(pr.pullRequestId);
                                        }
                                    }
                                }
                                setSelectedIds(pids)
                                console.log("Selected pull request IDs:", pids);
                            }}
                            onActivate={activatePullRequest}
                            renderRow={renderPullRequestRow}
                            width="100%"
                        />
                    </div>
                </Card>
            </div>
        </>
    )
}

export { App };
export type { AppProps };
