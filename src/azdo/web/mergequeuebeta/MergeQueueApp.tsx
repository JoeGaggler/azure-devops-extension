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

import { getAzdoInfo, getGitClient, getExtensionManagementClient, TenantInfo } from "./azuredevops";

import { GitPullRequestSearchCriteria, PullRequestStatus, PullRequestTimeRangeType } from "azure-devops-extension-api/Git/Git";
import { ExtensionManagementRestClient } from "azure-devops-extension-api/ExtensionManagement/ExtensionManagementClient";

const publisher = "pingmint";
const extensionName = "pingmint-extension";

export interface MergeQueueAppSingleton {
    bearerToken: string;
    appToken: string;
}

interface PullRequestInfo {
    id: number;
    title: string;
}

interface ReducerState {
    mergeQueuePullRequests: PullRequestInfo[];
    activePullRequests: PullRequestInfo[];
}

interface ReducerAction {
    activePullRequests?: PullRequestInfo[];
    enqueuePullRequests?: PullRequestInfo[];
}

function reducer(state: ReducerState, action: ReducerAction): ReducerState {
    let next = { ...state };

    if (action.activePullRequests) {
        console.log("MQ: reducer -> updating active pull requests", action.activePullRequests);
        next.activePullRequests = action.activePullRequests;
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
        console.error("MQ: getDocument -> error occurred", error);
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

    const [state, dispatch] = React.useReducer<(state: ReducerState, action: ReducerAction) => ReducerState>(reducer, {
        mergeQueuePullRequests: [],
        activePullRequests: [],
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
            let doc = await getDocument(em, "test-collection", "test-document");
            if (!doc) {
                console.log("MQ: init -> document not found, creating new one");
                doc = await createDocument(em, "test-collection", "test-document", { foo: "bar" });
                if (!doc) {
                    console.error("MQ: init -> failed to create document");
                    return;
                }
            }
            console.log("MQ: init -> retrieved document", doc);
            let doc2 = await updateDocument(em, "test-collection", "test-document", doc);
            console.log("MQ: init -> updated document 2", doc2);
            let doc3 = await updateDocument(em, "test-collection", "test-document", doc2);
            console.log("MQ: init -> updated document 3", doc3);

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

            let ti = tenantInfo.current;
            if (!ti) {
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
            let prs1 = await git.getPullRequestsByProject(ti.project, criteria, undefined, undefined, undefined);
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

    function renderMergeQueueCommandBarItems(): IHeaderCommandBarItem[] {
        // TODO: return command bar items for the merge queue
        return [];
    }

    function renderAllPullRequestsCommandBarItems(): IHeaderCommandBarItem[] {
        // TODO: return command bar items for all pull requests
        return [
            {
                id: "enqueue",
                text: "Enqueue",
                onActivate: () => { /* TODO: Implement enqueue functionality */ },
                isPrimary: true,
                important: true,
                disabled: false
            }
        ];
    }

    return (
        <Page className="">
            <Header
                title="Merge Queue Beta"
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
                <PullRequestList pullRequests={state.mergeQueuePullRequests} />
            </Card>

            <Card
                className="padding-8 margin-8"
                contentProps={{ contentPadding: false }}
                titleProps={{ text: "All Pull Requests", className: "", size: TitleSize.Medium }}
                headerClassName=""
                headerCommandBarItems={renderAllPullRequestsCommandBarItems()}
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
        if (!item) { return <></> }

        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}
            >
                <div className="flex-row rhythm-horizontal-8">
                    <div>{item.pullRequestId}</div>
                    <div>{item.title}</div>
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
                // onActivate={showRunTargetPipelinePanel}
                renderRow={renderRow}
                width="100%"
            />
        </div>
    </>
}