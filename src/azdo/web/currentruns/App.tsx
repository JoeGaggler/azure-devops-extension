import React from "react";
import * as joe from "../lib.ts";
import * as SDK from 'azure-devops-extension-sdk';
import * as Azdo from '../azdo/azdo.ts';
import * as luxon from 'luxon'
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
// import { Button } from "azure-devops-ui/Button";
import { Card } from "azure-devops-ui/Card";
// import { Dropdown } from "azure-devops-ui/Dropdown";
// import { DropdownMultiSelection } from "azure-devops-ui/Utilities/DropdownSelection";
// import { Icon, IconSize } from "azure-devops-ui/Icon";
// import { IListBoxItem } from "azure-devops-ui/ListBox";
import { ListSelection } from "azure-devops-ui/List";
import { Page } from "azure-devops-ui/Page";
// import { Pill, PillVariant, PillSize } from "azure-devops-ui/Pill";
// import { PillGroup } from "azure-devops-ui/PillGroup";
import { ScrollableList, IListItemDetails, ListItem } from "azure-devops-ui/List";
import { Tab, TabBar, TabSize } from "azure-devops-ui/Tabs";
// import { Toast } from "azure-devops-ui/Toast";
// import { Toggle } from "azure-devops-ui/Toggle";
// import { VssPersona } from "azure-devops-ui/VssPersona";
import { type IHostNavigationService } from 'azure-devops-extension-api';
import { Run, GetRunStatusType, StatusType } from "./Run.tsx";


interface AppSingleton {
    // repositoryFilterDropdownMultiSelection: DropdownMultiSelection;
}

interface AppProps {
    bearerToken: string;
    appToken: string;
    singleton: AppSingleton;
}

interface PipelineGroupsDocument {
    groups: PipelineGroup[];
}

interface PipelineGroup {
    name: string;
    pipelines: number[]; // pipeline IDs
}

interface CurrentRunsDocument {
    runs: CurrentRun[];
}

interface CurrentRun {
    pipelineId?: number;
    buildId?: number;
    buildNumber?: string;
    repositoryName?: string;
    definitionName?: string;
    status?: string;
    result?: string;
    webUrl?: string;
    queueTime?: string;

    queueTimestamp?: number;
    sortOrder?: number;
    comment?: string;
}

function App(p: AppProps) {
    console.log("App render", p);

    let collectionId = "currentPipelines";
    let currentRunsDocumentId = "currentRuns";
    let pipelineGroupsDocumentId = "pipelineGroups";

    const [tenantInfo, setTenantInfo] = React.useState<Azdo.TenantInfo>({});
    const [currentRuns, setCurrentRuns] = React.useState<CurrentRun[]>([]);
    const [pipelineGroups, setPipelineGroups] = React.useState<PipelineGroup[]>([]);
    const [allTopBuilds, setAllTopBuilds] = React.useState<Azdo.TopBuild[]>([]);
    const [selectedRunId, setSelectedRunId] = React.useState<number | null>(null);
    const [selectedPipelineId, setSelectedPipelineId] = React.useState<number | null>(null);
    const [selectedTabId, setSelectedTabId] = React.useState<string>("tabAll");

    // HACK: force rerendering for server sync
    const [pollHack, setPollHack] = React.useState(Math.random());
    React.useEffect(() => { poll(); }, [pollHack]);
    function Resync() { setPollHack(Math.random()); }

    // state
    let selectedGroup: PipelineGroup | null;
    switch (selectedTabId) {
        case "tabAll": selectedGroup = null; break;
        case "tabGroup-Builds": selectedGroup = pipelineGroups.find((g) => g.name === "Builds") || null; break;
        case "tabGroup-Releases": selectedGroup = pipelineGroups.find((g) => g.name === "Releases") || null; break;
        default: selectedGroup = null; break;
    };
    let currentRunsItems: CurrentRun[] = getRunsForGroup(selectedGroup)
    joe.sortByNumber(currentRunsItems, (i) => -(i.sortOrder || 0)); // newest first

    let currentRunsSelection = new ListSelection(true);
    let currentRunsIndex = (selectedRunId) ?
        currentRuns.findIndex((r: CurrentRun) => r.buildId === selectedRunId) :
        (-1);
    if (currentRunsIndex >= 0) {
        currentRunsSelection.select(currentRunsIndex);
    }

    // state
    let topBuildsQueueItems: Azdo.TopBuild[] = (allTopBuilds || []);
    let topBuildsSelection = new ListSelection(true);

    let idx = (selectedRunId) ?
        topBuildsQueueItems.findIndex((b: Azdo.TopBuild) => b.buildId === selectedRunId) :
        (-1);

    if (idx >= 0) {
        topBuildsSelection.select(idx);
    }

    // initialize the app
    React.useEffect(() => { init() }, []);
    async function init() {
        let info = await Azdo.getAzdoInfo();
        console.log("Tenant Info", info);
        setTenantInfo(info)

        // setup pipeline groups
        let pipelineGroupsDoc: PipelineGroupsDocument = {
            groups: [
                {
                    name: "Builds", pipelines: [
                        1264, // shadowauth
                        1649, // shadowsync
                        3407, // shadowscribe
                    ]
                },
                {
                    name: "Releases", pipelines: [
                        1652, // shadowsync prod
                        2544, // shadowauth prod
                        3304, // shadowscribe prod
                    ]
                }
            ]
        };
        pipelineGroupsDoc = await Azdo.getOrCreateUserDocument(collectionId, pipelineGroupsDocumentId, pipelineGroupsDoc)
        setPipelineGroups(pipelineGroupsDoc.groups);
        console.log("Init Pipeline Groups Document", pipelineGroupsDoc);

        // setup current runs
        let currentRunsDoc: CurrentRunsDocument = { runs: [] };
        currentRunsDoc = await Azdo.getOrCreateSharedDocument(collectionId, currentRunsDocumentId, currentRunsDoc);
        setCurrentRuns(currentRunsDoc.runs);
        console.log("Init Current Runs Document", currentRunsDoc);

        // // setup merge queue list
        // let newMergeQueueList = await downloadMergeQueuePullRequests();
        // setMergeQueueList(newMergeQueueList);

        // // setup user filters
        // let userFiltersDoc: PullRequestFilters = {
        //     drafts: false,
        //     allBranches: false,
        //     repositories: []
        // }
        // userFiltersDoc = await Azdo.getOrCreateUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc)
        // setFilters({ ...userFiltersDoc });

        // Refresh from server
        setInterval(() => { Resync(); }, 1000 * 10);
        Resync();
    }

    async function poll() {
        if (!tenantInfo.organization || !tenantInfo.project) {
            console.warn("Tenant info not set, skipping poll.");
            return;
        }

        let currentRunsDoc: CurrentRunsDocument = { runs: [] };
        currentRunsDoc = await Azdo.getOrCreateSharedDocument(collectionId, currentRunsDocumentId, currentRunsDoc);
        console.log("Current Runs Document", currentRunsDoc);

        let top = await getTopBuilds();
        let tempRuns: CurrentRun[] = currentRunsDoc.runs || [];
        for (let i = 0; i < top.length; i++) {
            let it = top[i];

            let queueDateTime = it.queueTime && luxon.DateTime.fromISO(it.queueTime);
            if (!queueDateTime || !queueDateTime.isValid) {
                console.warn("Invalid queue time for build:", it);
                continue;
            }

            let j = tempRuns.findIndex((r) => (r.buildId === it.buildId));
            if (j >= 0) {
                // TODO: this doesn't seem to update the status
                tempRuns[j] = {
                    ...tempRuns[j], // existing data
                    ...it, // new data
                    // TODO: update comment?
                }
            } else {
                let queueTimestamp = queueDateTime.toUnixInteger();
                // let pr = await Azdo.getPipelineRun(tenantInfo, it.pipelineId!, it.buildId!);
                // if (pr) {
                //     console.log("COMPARE", it, pr);
                // }
                tempRuns.push({
                    ...it,
                    sortOrder: queueTimestamp,
                    queueTimestamp: queueTimestamp,
                    comment: `${queueTimestamp} - ${luxon.DateTime.fromSeconds(queueTimestamp, { zone: "utc" }).toISO({})} - ${it.queueTime}`,
                });
            }
            // console.log("Current Run", tempRuns[i]);
        }

        tempRuns = cleanRuns(tempRuns)

        setCurrentRuns(tempRuns);
        const newRunsDoc: CurrentRunsDocument = {
            ...currentRunsDoc,
            runs: tempRuns
        };
        const nextRunsDoc = await Azdo.trySaveSharedDocument(collectionId, currentRunsDocumentId, newRunsDoc);
        if (!nextRunsDoc) {
            console.warn("Failed to save current runs document, will retry on next poll.", newRunsDoc);
        } else {
            console.log("Saved current runs document.", nextRunsDoc);
        }
    }

    function cleanRuns(runs: CurrentRun[]): CurrentRun[] {
        joe.sortByNumber(runs, (r) => -(r.sortOrder || 0)); // newest first

        for (let i = 0; i < runs.length; i++) {
            let it = runs[i];
            // let it_id = it.buildId;
            let it_p = it.pipelineId;
            let hasCompleted = false;
            let hasSucceeded = false;
            for (let j = 0; j < runs.length; /* NO INCREMENT */) {
                let jt = runs[j];

                // skip other pipelines
                let jt_p = jt.pipelineId;
                if (it_p !== jt_p) { j++; continue; }

                let s = jt.status;
                let r = jt.result;

                // remove invalid entries
                if (!s) {
                    console.warn("Removing invalid run:", jt);
                    runs.splice(j, 1); continue;
                }

                // keep unfinished runs
                if (s !== "completed") {
                    j++;
                    continue;
                }

                // done with this pipeline
                if (hasSucceeded) {
                    // console.warn("Removing run after succeeded:", jt);
                    runs.splice(j, 1); continue;
                }

                if (!hasCompleted && r) {
                    hasCompleted = true;
                    if (r === "succeeded" || r === "partiallySucceeded" || r === "succeededWithIssues") {
                        hasSucceeded = true;
                    }

                    j++;
                    continue;
                }

                // console.warn("Removing run after completed:", jt);
                runs.splice(j, 1);
                continue;
            }
        }

        return runs;
    }

    async function getTopBuilds(): Promise<Azdo.TopBuild[]> {
        const ti = tenantInfo;
        if (!ti.organization || !ti.project) {
            console.warn("Tenant info not set, skipping getTopBuilds.");
            return [];
        }
        let pullRequests = (await Azdo.getTopRecentBuilds(ti)) || [];
        console.log("Top Recent Builds", pullRequests);

        setAllTopBuilds(pullRequests);
        return pullRequests;
    }

    async function activateCurrentRun(_: any, evt: any) {
        console.log("activated: ", evt);
        let idx = evt.index;
        let data: CurrentRun = evt.data;
        console.log("activated: ", idx, data);
        const navService = await SDK.getService<IHostNavigationService>("ms.vss-features.host-navigation-service");
        let url = data.webUrl;
        console.log("url: ", url);
        if (!url) {
            return;
        }
        navService.openNewWindow(url, "");
    }

    //"Success" | "Failed" | "Warning" | "Information" | "Running" | "Waiting" | "Queued" | "Canceled" | "Skipped";
    function topBuildToStatus(topBuild: Azdo.TopBuild): StatusType {
        return GetRunStatusType(topBuild.status, topBuild.result);
    }

    function renderCurrentRunRow(
        index: number,
        run: CurrentRun,
        details: IListItemDetails<any>,
        key?: string
    ): React.JSX.Element {
        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}
            >
                <Run
                    name={run.buildNumber || "?"}
                    status={topBuildToStatus(run)}
                    comment={run.comment || ""}
                    started={run.sortOrder || null}
                    isAlternate={isAlternate(run)}
                />
            </ListItem>
        );
    };

    function isAlternate(topBuild: Azdo.TopBuild | undefined): boolean {
        if (!topBuild) { return false; }
        if (!selectedPipelineId || !selectedRunId) { return false; }
        if (topBuild.buildId === selectedRunId) { return false; }
        if (topBuild.pipelineId !== selectedPipelineId) { return false; }
        return true;
    }

    function onSelectCurrentRun(list: CurrentRun[], listSelection: ListSelection) {
        if (list.length == 0) {
            // setSelectedIds([]);
            setSelectedPipelineId(null);
            setSelectedRunId(null);
            return;
        }

        // let pids: number[] = []
        for (let selRange of listSelection.value) {
            for (let i = selRange.beginIndex; i <= selRange.endIndex; i++) {
                let b = list[i];
                if (b && b.buildId && b.pipelineId) {
                    console.log("Selected run ID:", b.buildId, "Pipeline ID:", b.pipelineId);
                    setSelectedRunId(b.buildId);
                    setSelectedPipelineId(b.pipelineId);
                }
                // if (pr && pr.pullRequestId) {
                //     pids.push(pr.pullRequestId);
                // }
            }
        }
        // setSelectedIds(pids)
        // console.log("Selected pull request IDs:", pids);
    }

    function getRunsForGroup(pg: PipelineGroup | null): CurrentRun[] {
        var runs = currentRuns || [];
        if (!pg) { return runs; }
        let matches: CurrentRun[] = [];
        for (let r of runs) {
            let rp = r.pipelineId;
            if (rp && pg.pipelines.indexOf(rp) >= 0) {
                matches.push(r);
            }
        }
        return matches;
    }

    function getTabBar() {
        // var runs = currentRuns || [];
        var pgs = pipelineGroups || [];
        joe.sortByString(pgs, (p) => p.name);
        let tabs: React.JSX.Element[] = [];
        for (let pg of pgs) {
            let matches = getRunsForGroup(pg);
            let tab = getTab(`tabGroup-${pg.name}`, pg.name, matches.length);
            tabs.push(tab);
        }

        function getTab(id: string, name: string, badgeCount: number) {
            return (
                <Tab
                    id={id}
                    name={name}
                    badgeCount={badgeCount}
                />);
        }

        return (
            <TabBar
                tabSize={TabSize.Tall}
                disableSticky={true}
                selectedTabId={selectedTabId}
                onSelectedTabChanged={onSelectedTabChanged}
                tabsClassName="run-tabbar"
            >
                {getTab("tab-All", "All", currentRuns.length)}
                {...tabs}
            </TabBar>
        );
    }

    function onSelectedTabChanged(newTabId: string) {
        console.log("Tab changed:", newTabId);
        setSelectedTabId(newTabId);
    }

    return (
        <Page>
            <div className="padding-8 margin-8">
                <div className="padding-8 flex-row flex-baseline rhythm-horizontal-16">
                    <h2>Current Pipelines</h2>
                    <div className="flex-grow"></div>
                </div>
                <Card className="padding-8">
                    <div className="flex-column">
                        {getTabBar()}

                        <ScrollableList
                            itemProvider={new ArrayItemProvider(currentRunsItems)}
                            selection={currentRunsSelection}
                            onSelect={(_evt, _listRow) => { onSelectCurrentRun(currentRunsItems, currentRunsSelection); }}
                            onActivate={activateCurrentRun}
                            renderRow={renderCurrentRunRow}
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
