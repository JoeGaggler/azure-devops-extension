import React from "react";
import * as joe from "../lib.ts";
import * as SDK from 'azure-devops-extension-sdk';
import * as Azdo from '../azdo/azdo.ts';
import * as luxon from 'luxon'
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { Button } from "azure-devops-ui/Button";
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
import { TextField, TextFieldWidth } from "azure-devops-ui/TextField";
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
    groups: PipelineGroup[];
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

interface MapTabIdToGroup {
    [key: string]: MappingTarget;
}

interface MappingTarget {
    id: string;
    path: string[];
    pipelineGroup: PipelineGroup;
}

function App(p: AppProps) {
    console.log("App render", p);

    let collectionId = "currentPipelines";
    let currentRunsDocumentId = "currentRuns";
    let pipelineGroupsDocumentId = "pipelineGroups";

    const [tenantInfo, setTenantInfo] = React.useState<Azdo.TenantInfo>({});
    const [currentRuns, setCurrentRuns] = React.useState<CurrentRun[]>([]);
    const [pipelineGroups, setPipelineGroups] = React.useState<PipelineGroup[]>([]);
    const [selectedRunId, setSelectedRunId] = React.useState<number | null>(null);
    const [selectedPipelineId, setSelectedPipelineId] = React.useState<number | null>(null);
    const [selectedPipelineName, setSelectedPipelineName] = React.useState<string>("");
    const [selectedTabId, setSelectedTabId] = React.useState<string | null>(null);
    const [tabIdPath, setTabIdPath] = React.useState<string[]>([]);
    const [newSubgroupName, setNewSubgroupName] = React.useState<string>("");

    // HACK: force rerendering for server sync
    const [pollHack, setPollHack] = React.useState(Math.random());
    React.useEffect(() => { poll(); }, [pollHack]);
    function Resync() { setPollHack(Math.random()); }

    let mapTabsToGroups: MapTabIdToGroup = {};
    let tabBarItems: JSX.Element[] = [
        <Tab id="tabGroup-ROOT" name="All" badgeCount={currentRuns.length} />
    ];
    function buildTabMappings(prefix: String, pgs: PipelineGroup[], tip: string[], tip2: string[], depth: number) {
        if (depth > 5) { return; }
        console.log("buildTabMappings", prefix, pgs, tip, tip2, depth);

        if (!pgs || pgs.length == 0) {
            console.log("buildTabMappings: no groups", prefix, pgs, tip, tip2, depth);
            return;
        }

        let icon: any | undefined = { iconName: "ChevronRight" };
        if (tip.length > 0) {
            let chop = [...tip];
            let last = chop.splice(0, 1)[0];
            console.log(`buildTabMappings: chopped ${last}`, prefix, pgs, tip, tip2, depth);
            let g = pgs.find((g) => g.name === last);
            if (!g) {
                console.warn("Failed to find pipeline group for tab path:", tip, last);
                return;
            }

            let newTip2 = [...tip2, g.name];
            let newPrefix = `${prefix}-${g.name}`;
            mapTabsToGroups[newPrefix] = {
                id: newPrefix,
                pipelineGroup: g,
                path: newTip2,
            };
            tabBarItems.push(<Tab id={newPrefix} name={g.name} badgeCount={getRunsForGroup(g).length} iconProps={icon} />);

            buildTabMappings(newPrefix, g.groups, chop, newTip2, depth + 1);
        } else {
            console.log("buildTabMappings: final", prefix, pgs, tip, tip2, depth);
            let pgs2 = [...pgs];
            joe.sortByString(pgs2, (p) => p.name);
            for (let g of pgs2) {
                let newTip2: string[] = [...tip2, g.name];
                let newPrefix = `${prefix}-${g.name}`;
                mapTabsToGroups[newPrefix] = {
                    id: newPrefix,
                    pipelineGroup: g,
                    path: newTip2,
                };
                tabBarItems.push(<Tab id={newPrefix} name={g.name} badgeCount={getRunsForGroup(g).length} iconProps={icon} />);
                icon = undefined; // only first one gets icon
            }
            // TODO: ADD "OTHER" TAB HERE!
        }
    }
    buildTabMappings("tabGroupId", pipelineGroups, tabIdPath, [], 0);
    console.log("mapTabsToGroups", mapTabsToGroups);

    function getPipelineGroupForPath(groups: PipelineGroup[], path: string[]): PipelineGroup | undefined {
        if (path.length == 0) { return undefined; }
        let gg = groups;
        let g: PipelineGroup | undefined = undefined;
        let p = [...path];
        console.log("*** getPipelineGroupForPath", p);

        while (p.length > 0) {
            let n = p.splice(0, 1)[0];
            g = gg.find((i) => i.name === n);
            if (!g) {
                console.warn("Failed to find pipeline group for path:", path, n);
                return undefined;
            }
            gg = g.groups
        }
        console.log("*** getPipelineGroupForPath found", p, g);
        if (g && !g.groups) { g.groups = []; }
        return g;
    }

    // state
    let selectedGroup = getPipelineGroupForPath(pipelineGroups, tabIdPath);
    console.log("*** selectedGroup", selectedGroup, tabIdPath);
    let currentRunsItems: CurrentRun[] = getRunsForGroup(selectedGroup || null)
    joe.sortByNumber(currentRunsItems, (i) => -(i.sortOrder || 0)); // newest first

    let currentRunsSelection = new ListSelection(true);
    let currentRunsIndex = (selectedRunId) ?
        currentRunsItems.findIndex((r: CurrentRun) => r.buildId === selectedRunId) :
        (-1);
    if (currentRunsIndex >= 0) {
        currentRunsSelection.select(currentRunsIndex);
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
                // {
                //     name: "Builds", pipelines: [
                //         1264, // shadowauth
                //         1649, // shadowsync
                //         3407, // shadowscribe
                //     ]
                // },
                // {
                //     name: "Releases", pipelines: [
                //         1652, // shadowsync prod
                //         2544, // shadowauth prod
                //         3304, // shadowscribe prod
                //     ]
                // }
            ]
        };
        pipelineGroupsDoc = await Azdo.getOrCreateSharedDocument(collectionId, pipelineGroupsDocumentId, pipelineGroupsDoc)
        setPipelineGroups(pipelineGroupsDoc.groups);
        console.log("Init Pipeline Groups Document", pipelineGroupsDoc);

        // setup current runs
        let currentRunsDoc: CurrentRunsDocument = { runs: [] };
        currentRunsDoc = await Azdo.getOrCreateSharedDocument(collectionId, currentRunsDocumentId, currentRunsDoc);
        setCurrentRuns(currentRunsDoc.runs);
        console.log("Init Current Runs Document", currentRunsDoc);

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
                    comment: ``, // TODO: COMMENT
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
                    comment: ``, // TODO: `${queueTimestamp} - ${luxon.DateTime.fromSeconds(queueTimestamp, { zone: "utc" }).toISO({})} - ${it.queueTime}`,
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
                if (s !== "completed" && s !== "cancelling") {
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

    function isKnown(pipelineId: number | undefined): boolean {
        if (!pipelineId) { return false; }
        for (let pg of pipelineGroups || []) {
            if (getRecursivePipelineIds(pg).indexOf(pipelineId) >= 0) { return true; }
        }
        return false;
    }

    function knownTagsForRun(pipelineId: number | undefined): string[] {
        let tags: string[] = [];
        if (!pipelineId) { return tags; }
        for (let pg of pipelineGroups || []) {
            rec_knownTagsForRun(pg, pipelineId, tags, []);
        }
        return tags;
    }

    function rec_knownTagsForRun(pg: PipelineGroup, pipelineId: number, tags: string[], path: string[]) {
        let newPath = [...path, pg.name];
        if (pg.pipelines.indexOf(pipelineId) >= 0) {
            tags.push(newPath.join(" / "));
        }
        let ggg = pg.groups || [];
        for (let g of ggg) {
            rec_knownTagsForRun(g, pipelineId, tags, newPath);
        }
    }

    // function rec_knownTagsForRun(pg: PipelineGroup, pipelineId: number, tags: string[]) {
    //     if (pg.pipelines.indexOf(pipelineId) >= 0) {
    //         tags.push(pg.name);
    //     }
    //     let ggg = pg.groups || [];
    //     for (let g of ggg) {
    //         rec_knownTagsForRun(g, pipelineId, tags);
    //     }
    // }

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
                    definitionName={run.definitionName || "?"}
                    status={topBuildToStatus(run)}
                    comment={run.comment || ""}
                    started={run.sortOrder || null}
                    isAlternate={isAlternate(run)}
                    isKnown={isKnown(run.pipelineId)}
                    knownTags={knownTagsForRun(run.pipelineId)}
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
                    setSelectedPipelineName(b.definitionName || "");
                }
                // if (pr && pr.pullRequestId) {
                //     pids.push(pr.pullRequestId);
                // }
            }
        }
        // setSelectedIds(pids)
        // console.log("Selected pull request IDs:", pids);
    }

    function getRecursivePipelineIds(pg: PipelineGroup | null | undefined): number[] {
        if (!pg) { return []; }
        let pids: number[] = [...pg.pipelines];
        let ggg = pg.groups || [];
        for (let g of ggg) {
            let pids2 = getRecursivePipelineIds(g);
            for (let p of pids2) {
                if (pids.indexOf(p) < 0) {
                    pids.push(p);
                }
            }
        }
        return pids;
    }

    function getRunsForGroup(pg: PipelineGroup | null | undefined): CurrentRun[] {
        var runs = currentRuns || [];
        if (!pg) { return runs; }
        let pids = getRecursivePipelineIds(pg);
        let matches: CurrentRun[] = [];
        for (let r of runs) {
            let rp = r.pipelineId;
            if (rp && pids.indexOf(rp) >= 0) {
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
                selectedTabId={selectedTabId || "tabGroup-ROOT"}
                onSelectedTabChanged={onSelectedTabChanged}
                tabsClassName="run-tabbar"
            >
                {tabBarItems}
            </TabBar>
        );
    }

    function onSelectedTabChanged(newTabId: string) {
        console.log("Tab changed:", newTabId);
        let ttt = mapTabsToGroups[newTabId];
        if (!ttt) {
            console.warn("Failed to find mapping for tab ID:", newTabId);
            setSelectedTabId(null)
            setTabIdPath([]);
            return
        }
        console.log("Found mapping for tab ID:", newTabId, ttt);

        setSelectedTabId(newTabId);
        setTabIdPath(ttt.path);
    }

    async function onAddSubgroup() {
        let pgd: PipelineGroupsDocument = await Azdo.getOrCreateSharedDocument(collectionId, pipelineGroupsDocumentId, { groups: [] });
        if (!newSubgroupName || newSubgroupName.trim().length == 0) {
            alert("Please enter a valid subgroup name.");
            return;
        }

        let ggg = getPipelineGroupForPath(pgd.groups, [...tabIdPath]);
        let ggg2 = (ggg ? ggg.groups : pgd.groups)

        if (ggg2.find((g) => g.name === newSubgroupName)) {
            alert("A subgroup with that name already exists.");
            return;
        }

        ggg2.push({ name: newSubgroupName, pipelines: [], groups: [] });
        let r = await Azdo.trySaveSharedDocument(collectionId, pipelineGroupsDocumentId, pgd)
        if (!r) {
            alert("Failed to save new subgroup, please try again.");
            return;
        }
        setPipelineGroups(pgd.groups);
        setNewSubgroupName("");
    }

    async function onRemoveSubgroup() {
        if (tabIdPath.length == 0) {
            alert("Cannot remove the 'All' group.");
            return;
        }
        let pgd: PipelineGroupsDocument = await Azdo.getOrCreateSharedDocument(collectionId, pipelineGroupsDocumentId, { groups: [] });

        let ggg = getPipelineGroupForPath(pgd.groups, [...tabIdPath].splice(0, tabIdPath.length - 1));
        // if (!ggg) {
        //     alert("Failed to find the selected subgroup parent.");
        //     return;
        // }

        let ggg2 = (ggg ? ggg.groups : pgd.groups)
        let idx = ggg2.findIndex((g) => g.name === tabIdPath[tabIdPath.length - 1]);
        if (idx < 0) {
            alert("Failed to find the selected subgroup.");
            return;
        }
        let ggg3 = ggg2[idx];

        // let confirmed = confirm(`Are you sure you want to remove the subgroup '${pgd.groups[idx].name}'? This will not delete any pipelines or runs, it will just remove the grouping.`);
        let confirmed = confirm(`Are you sure you want to remove the subgroup '${ggg3.name}'? This will not delete any pipelines or runs, it will just remove the grouping.`);
        if (!confirmed) { return; }
        ggg2.splice(idx, 1);
        let r = await Azdo.trySaveSharedDocument(collectionId, pipelineGroupsDocumentId, pgd)
        if (!r) {
            alert("Failed to remove subgroup, please try again.");
            return;
        }
        setPipelineGroups(pgd.groups);
        setSelectedTabId("tabGroup-ROOT");
    }

    async function onAddPipelineToGroup() {
        if (!selectedPipelineId) {
            alert("Please select a pipeline to add to the group.");
            return;
        }
        if (tabIdPath.length == 0) {
            alert("Please select a subgroup to add the pipeline to.");
            return;
        }
        let pgd: PipelineGroupsDocument = await Azdo.getOrCreateSharedDocument(collectionId, pipelineGroupsDocumentId, { groups: [] });
        let ggg = getPipelineGroupForPath(pgd.groups, [...tabIdPath]);
        if (!ggg) {
            alert("Failed to find the selected subgroup.");
            return;
        }

        if (ggg.pipelines.indexOf(selectedPipelineId) >= 0) {
            alert("That pipeline is already in the selected group.");
            return;
        }
        ggg.pipelines.push(selectedPipelineId);
        let r = await Azdo.trySaveSharedDocument(collectionId, pipelineGroupsDocumentId, pgd)
        if (!r) {
            alert("Failed to add pipeline to group, please try again.");
            return;
        }
        setPipelineGroups(pgd.groups);
    }

    async function onRemovePipelineFromGroup() {
        if (!selectedPipelineId) {
            alert("Please select a pipeline to remove from the group.");
            return;
        }
        if (tabIdPath.length == 0) {
            alert("Please select a subgroup to add the pipeline to.");
            return;
        }
        let pgd: PipelineGroupsDocument = await Azdo.getOrCreateSharedDocument(collectionId, pipelineGroupsDocumentId, { groups: [] });
        let ggg = getPipelineGroupForPath(pgd.groups, [...tabIdPath]);
        if (!ggg) {
            alert("Failed to find the selected subgroup.");
            return;
        }

        function rrr(fromGroup: PipelineGroup, pid: number): boolean {
            let found = false;
            let idx = fromGroup.pipelines.indexOf(pid);
            if (idx >= 0) {
                fromGroup.pipelines.splice(idx, 1);
                found = true;
            }
            let gggg = fromGroup.groups || [];
            for (let g of gggg) {
                found = rrr(g, pid) || found;
            }
            return found;
        }

        // TODO: RECURSIVE REMOVE
        let found = rrr(ggg, selectedPipelineId);

        // let idx = ggg.pipelines.indexOf(selectedPipelineId);
        // if (idx < 0) {
        if (!found) {
            alert("That pipeline is not in the selected group.");
            return;
        }
        // ggg.pipelines.splice(idx, 1);
        let r = await Azdo.trySaveSharedDocument(collectionId, pipelineGroupsDocumentId, pgd)
        if (!r) {
            alert("Failed to remove pipeline from group, please try again.");
            return;
        }
        setPipelineGroups(pgd.groups);
    }

    return (
        <Page>
            <div className="padding-8 margin-8">
                <div className="padding-8 flex-row flex-baseline rhythm-horizontal-16">
                    <h2>Current Pipelines</h2>
                    <div className="flex-grow"></div>
                    <div className="flex-column rhythm-vertical-8">
                        <div className="flex-row flex-baseline rhythm-horizontal-16">
                            <TextField
                                value={newSubgroupName}
                                onChange={(_e, newValue) => (setNewSubgroupName(newValue))}
                                placeholder="Name"
                                width={TextFieldWidth.tabBar}
                            />
                            <Button
                                disabled={false} // TODO: validation
                                onClick={onAddSubgroup}
                                text="Add Group"
                            />
                            <Button
                                disabled={false} // TODO: validation
                                onClick={onRemoveSubgroup}
                                text="Remove Group"
                            />
                        </div>
                        <div className="flex-row flex-baseline rhythm-horizontal-16">
                            <TextField
                                value={selectedPipelineName}
                                placeholder="Pipeline"
                                width={TextFieldWidth.tabBar}
                                readOnly={true}
                            />
                            <TextField
                                value={selectedPipelineId ? selectedPipelineId.toString() : ""}
                                placeholder="Pipeline"
                                width={TextFieldWidth.tabBar}
                            />
                            <Button
                                disabled={false} // TODO: validation
                                onClick={onAddPipelineToGroup}
                                text="Add to Group"
                            />
                            <Button
                                disabled={false} // TODO: validation
                                onClick={onRemovePipelineFromGroup}
                                text="Remove from Group"
                            />
                        </div>
                    </div>
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
