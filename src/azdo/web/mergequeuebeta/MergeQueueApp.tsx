import React from "react";

import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { Header } from "azure-devops-ui/Components/Header/Header";
import { IHeaderCommandBarItem } from "azure-devops-ui/Components/HeaderCommandBar/HeaderCommandBar.Props";
import { IListItemDetails, IListRow } from "azure-devops-ui/Components/List/List.Props";
import { ListItem, ScrollableList } from "azure-devops-ui/Components/List/List";
import { ListSelection } from "azure-devops-ui/Components/List/ListSelection";
import { Page } from "azure-devops-ui/Components/Page/Page";
import { TitleSize } from "azure-devops-ui/Components/Header/Header.Props";
import { Card } from "azure-devops-ui/Card";

import { getAzdoInfo, getGitClient, TenantInfo } from "./azuredevops";

import { GitPullRequestSearchCriteria, PullRequestStatus, PullRequestTimeRangeType } from "azure-devops-extension-api/Git/Git";

export interface MergeQueueAppSingleton {
    bearerToken: string;
    appToken: string;
}

interface PullRequestInfo {
    id: number;
    title: string;
}

interface ReducerState {
    activePullRequests: PullRequestInfo[];
}

interface ReducerAction {
    activePullRequests?: PullRequestInfo[];
}

function reducer(state: ReducerState, action: ReducerAction): ReducerState {
    let next = { ...state };

    if (action.activePullRequests) {
        console.log("MQ: reducer -> updating active pull requests", action.activePullRequests);
        next.activePullRequests = action.activePullRequests;
    }

    return next;
}

export function MergeQueueApp(p: { singleton: MergeQueueAppSingleton }) {
    let _tenantInfo = React.useRef<TenantInfo>();
    let singleton = React.useRef(p.singleton);
    let gitClient = React.useRef(getGitClient());

    const [state, dispatch] = React.useReducer<(state: ReducerState, action: ReducerAction) => ReducerState>(reducer, {
        activePullRequests: [],
        // TODO: initial reducer state
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
            _tenantInfo.current = nextTenantInfo;
        } catch (error) {
            console.error("MQ: init -> error occurred", error);
        }
    }

    // ticktock
    React.useEffect(() => {
        let id = setInterval(() => { ticktock(); }, 5000);
        return () => clearInterval(id);
    }, []);
    async function ticktock() {
        try {
            console.log("MQ: ticktock");

            if (!singleton.current) {
                console.error("MQ: ticktock -> no singleton available");
                return;
            }

            let tenantInfo = _tenantInfo.current;
            if (!tenantInfo) {
                console.error("MQ: ticktock -> no tenant info available");
                return;
            }

            let git = gitClient.current;
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
            let prs1 = await git.getPullRequestsByProject(tenantInfo.project, criteria, undefined, undefined, undefined);
            let prs2 = prs1.map((pr): PullRequestInfo => {
                return {
                    ...pr, // HACK: smuggle full response
                    id: pr.pullRequestId,
                    title: pr.title
                };
            });

            dispatch({ activePullRequests: prs2 });
            console.log("MQ: pull requests", prs2);
        } catch (error) {
            console.error("MQ: ticktock -> error occurred", error);
        }
    }

    function renderPageCommandBarItems(): IHeaderCommandBarItem[] {
        // TODO: return page command bar items
        return [];
    }

    return (
        <Page>
            <Header
                title="Merge Queue Beta"
                titleSize={TitleSize.Large}
                commandBarItems={renderPageCommandBarItems()}
            />

            <Card
                className="padding-8 margin-8"
                contentProps={{ contentPadding: false }}
                titleProps={{ text: "All Pull Requests", className:"", size: TitleSize.Medium }}
                headerClassName=""
                // headerCommandBarItems={[{
                //     id: "runTargetPipeline",
                //     text: "Run",
                //     onActivate: () => { showRunTargetPipelinePanel(); },
                //     isPrimary: true,
                //     important: true,
                //     disabled: !hasSelectedTargetPipeline,
                // }]}
            >
                <PullRequestList pullRequests={state.activePullRequests} />
            </Card>

            <div className="text-neutral-30 flex-row padding-4">
                <div className="flex-grow"></div>
                <div>__MERGEQUEUEVERSION__</div>
            </div>
        </Page>
    );
}

export interface PullRequestListProps {
    pullRequests: any[]; // TODO: pull request type
}

export function PullRequestList({ pullRequests }: PullRequestListProps) {
    console.log("MQ: PullRequestList -> rendering", pullRequests);
    let listSelection = new ListSelection(true);

    function onSelectRow(row: IListRow<any>) { // TODO: pull request type
        console.log("NextRunTab -> targetPipelineSelect", row);
        // dispatch({ selectTargetPipeline: row.data });
    }

    function renderRow(
        index: number,
        item: any, // TODO: pull request type
        details: IListItemDetails<any>,
        key?: string
    ): JSX.Element {
        console.log("MQ: PullRequestList -> renderRow", index, key, item, details);

        if (!item) { return <></> }

        return <ListItem
            key={key || "list-item" + index}
            index={index}
            details={details}
        >
            <div className="flex-row">{item.pullRequestId} - {item.title}</div>
        </ListItem>
    }

    return <>
        <div className="flex-column">
            <ScrollableList
                itemProvider={new ArrayItemProvider(pullRequests || [])}
                selection={listSelection}
                onSelect={(_evt, listRow) => { onSelectRow(listRow); }}
                // onActivate={showRunTargetPipelinePanel}
                renderRow={renderRow}
                width="100%"
            />
        </div>
    </>
}