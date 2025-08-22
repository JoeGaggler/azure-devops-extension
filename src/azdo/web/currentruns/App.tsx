import React from "react";
// import * as joe from "../lib.ts";
import * as SDK from 'azure-devops-extension-sdk';
import * as Azdo from '../azdo/azdo.ts';
// import * as luxon from 'luxon'
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

function App(p: AppProps) {
    console.log("App render", p);

    const [tenantInfo, setTenantInfo] = React.useState<Azdo.TenantInfo>({});
    const [allTopBuilds, setAllTopBuilds] = React.useState<Azdo.TopBuild[]>([]);

    // HACK: force rerendering for server sync
    const [pollHack, setPollHack] = React.useState(Math.random());
    React.useEffect(() => { poll(); }, [pollHack]);
    function Resync() { setPollHack(Math.random()); }

    // state
    let topBuildsQueueItems: Azdo.TopBuild[] = (allTopBuilds || []);
    let topBuildsSelection = new ListSelection(true);


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

        await getTopBuilds();
    }

    async function getTopBuilds() {
        const ti = tenantInfo;
        if (!ti.organization || !ti.project) {
            console.warn("Tenant info not set, skipping getTopBuilds.");
            return;
        }
        let pullRequests = await Azdo.getTopRecentBuilds(ti);
        console.log("Top Recent Builds", pullRequests);

        setAllTopBuilds(pullRequests || []);
    }

    async function activateTopBuild(_: any, evt: any) {
        console.log("activated: ", evt);
        let idx = evt.index;
        let data: Azdo.TopBuild = evt.data;
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

    function renderTopBuildRow(
        index: number,
        topBuild: Azdo.TopBuild,
        details: IListItemDetails<any>,
        key?: string
    ): React.JSX.Element {
        // let extra = "";
        // let className = `scroll-hidden flex-row flex-center rhythm-horizontal-8 flex-grow padding-4 ${extra}`;

        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}
            >
                {/* <div className={className}>
                    <div className="font-size-m flex-row flex-center flex-shrink rhythm-horizontal-8">
                        <div>{index + 1}</div>
                        <div>{topBuild.pipelineId}</div>
                        <div>{topBuild.buildId}</div>
                        <div>{topBuild.buildNumber || "(unknown build number)"}</div>
                        <div>{topBuild.repositoryName || "(unknown repo)"}</div>
                        <div>{topBuild.definitionName || "(unknown definition)"}</div>
                    </div>
                </div> */}
                <Run
                    name={topBuild.buildNumber || "?"}
                    status={topBuildToStatus(topBuild)}
                    comment={`comment`}
                    started={null}
                    isAlternate={false}
                />
            </ListItem>
        );
    };

    function onSelectTopBuilds(list: Azdo.TopBuild[], listSelection: ListSelection) {
        if (list.length == 0) {
            // setSelectedIds([]);
            return;
        }

        // let pids: number[] = []
        for (let selRange of listSelection.value) {
            for (let i = selRange.beginIndex; i <= selRange.endIndex; i++) {
                //         let pr = list[i];
                //         if (pr && pr.pullRequestId) {
                //             pids.push(pr.pullRequestId);
                //         }
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
                            itemProvider={new ArrayItemProvider(topBuildsQueueItems)}
                            selection={topBuildsSelection}
                            onSelect={(_evt, _listRow) => { onSelectTopBuilds(topBuildsQueueItems, topBuildsSelection); }}
                            onActivate={activateTopBuild}
                            renderRow={renderTopBuildRow}
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
