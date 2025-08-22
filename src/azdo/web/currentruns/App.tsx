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

    sortOrder?: number;
    comment?: string;
}

function App(p: AppProps) {
    console.log("App render", p);

    const [tenantInfo, setTenantInfo] = React.useState<Azdo.TenantInfo>({});
    const [currentRuns, setCurrentRuns] = React.useState<CurrentRun[]>([]);
    const [allTopBuilds, setAllTopBuilds] = React.useState<Azdo.TopBuild[]>([]);
    const [selectedRunId, setSelectedRunId] = React.useState<number | null>(null);
    const [selectedPipelineId, setSelectedPipelineId] = React.useState<number | null>(null);

    // HACK: force rerendering for server sync
    const [pollHack, setPollHack] = React.useState(Math.random());
    React.useEffect(() => { poll(); }, [pollHack]);
    function Resync() { setPollHack(Math.random()); }

    // state
    let currentRunsItems: CurrentRun[] = (currentRuns || []);
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
        setInterval(() => { Resync(); }, 1000 * 20);
        Resync();
    }

    async function poll() {
        if (!tenantInfo.organization || !tenantInfo.project) {
            console.warn("Tenant info not set, skipping poll.");
            return;
        }

        let top = await getTopBuilds();
        let tempRuns: CurrentRun[] = currentRuns.slice();
        for (let i = 0; i < top.length; i++) {
            let it = top[i];

            let queueDateTime = it.queueTime && luxon.DateTime.fromISO(it.queueTime);
            if (!queueDateTime || !queueDateTime.isValid) {
                console.warn("Invalid queue time for build:", it);
                continue;
            }

            let j = tempRuns.findIndex((r) => (r.buildId === it.buildId));
            if (j >= 0) {
                // TODO: does this work?
                top[i] = {
                    ...tempRuns[j], // existing data
                    ...it, // new data
                    // TODO: update comment?
                }
            } else {
                let sortOrder = queueDateTime.toUnixInteger();
                tempRuns.push({
                    ...it,
                    sortOrder: sortOrder,
                    comment: `${sortOrder} - ${luxon.DateTime.fromSeconds(sortOrder, { zone: "utc" }).toISO({})} - ${it.queueTime}`,
                });
            }
            console.log("Current Run", tempRuns[i]);
        }

        tempRuns = cleanRuns(tempRuns)

        setCurrentRuns(tempRuns);
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
                    console.warn("Removing run after succeeded:", jt);
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

                console.warn("Removing run after completed:", jt);
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
        let url = data.webUrl //`https://dev.azure.com/${tenantInfo.organization}/${tenantInfo.project}/_git/${data.repositoryName}/pullrequest/${data.pullRequestId}`;
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
                    started={null}
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

    return (
        <Page>
            <div className="padding-8 margin-8">
                <Card className="padding-8">
                    <div className="flex-column">
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
