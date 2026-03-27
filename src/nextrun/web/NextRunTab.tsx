import React from "react";
import * as Azdo from '../shared/azdo.ts';
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { Card } from "azure-devops-ui/Card";
import { ListItem, ListSelection, type IListItemDetails, type IListRow } from "azure-devops-ui/List";
import { MessageCard, MessageCardSeverity } from "azure-devops-ui/MessageCard";
import { ScrollableList } from "azure-devops-ui/List";

export interface NextRunTabSingleton {
    bearerToken: string;
    appToken: string;

    build: any;
    definition: any;
}

export interface NextRunTabProps {
    singleton: NextRunTabSingleton;
}

interface ReducerState {
    targetPipelines?: TargetPipeline[];
}

interface ReducerAction {
    targetPipelines?: TargetPipeline[];
}

function reducer(state: ReducerState, action: ReducerAction): ReducerState {
    let next = { ...state };

    if (action.targetPipelines) {
        next.targetPipelines = action.targetPipelines;
    }

    return next;
}

export function NextRunTab(p: NextRunTabProps) {
    console.log("NextRunTab render", p);

    const sourcePipelinesCollectionId = "source-pipelines";
    let tenantInfo = React.useRef<Azdo.TenantInfo>();

    const [_state, dispatch] = React.useReducer<(state: ReducerState, action: ReducerAction) => ReducerState>(reducer, {})

    let targetPipelineSelection = new ListSelection(true);
    let targetPipelineItems: TargetPipeline[] = [
        {
            name: "Test Pipeline 1"
        },
        {
            name: "Test Pipeline 2"
        },
        {
            name: "Test Pipeline 3"
        }
    ]
    // joe.applySelections(allSelection, allFilteredPullRequests, i => i.pullRequestId, selectedIds);


    // initialize the app
    React.useEffect(() => { init() }, []);
    async function init() {
        console.log("NextRunTab -> init");

        let info = await Azdo.getAzdoInfo();
        console.log("NextRunTab -> tenant", info);
        tenantInfo.current = info;

        let project = await info.project;
        let defId = p.singleton.definition?.id;
        if (!project || !defId) {
            console.warn("NextRunTab -> missing project or definition id", { project, defId });
            return;
        }

        let docId = `project-${project}-pipeline-${defId}`;
        console.log("NextRunTab -> source pipeline document id", docId);

        let sources: SourcePipelineDocument = {
            targetPipelines: []
        }
        sources = await Azdo.getOrCreateSharedDocument(sourcePipelinesCollectionId, docId, sources)
        console.log("NextRunTab -> got shared document", sources);

        dispatch({ targetPipelines: sources.targetPipelines });
    }

    React.useEffect(() => {
        let id = setInterval(() => { tick(); }, 5000);
        return () => clearInterval(id);
    }, []);
    async function tick() {
        console.log("NextRunTab -> tick");
    }

    function targetPipelineRenderRow(
        index: number,
        item: TargetPipeline,
        details: IListItemDetails<any>,
        key?: string
    ): JSX.Element {
        if (!item.name) { return <></> }

        return <ListItem
            key={key || "list-item" + index}
            index={index}
            details={details}
        >
            <TargetPipelineListItem name={item.name} />
        </ListItem>
    }

    function targetPipelineSelect(row: IListRow<TargetPipeline>) {
        console.log("NextRunTab -> targetPipelineSelect", row);
    }

    return <>
        <MessageCard severity={MessageCardSeverity.Info}>
            Work in progress. Please check back later.
        </MessageCard>

        <Card className="padding-8">
            <div className="flex-column">
                <ScrollableList
                    itemProvider={new ArrayItemProvider(targetPipelineItems)}
                    selection={targetPipelineSelection}
                    onSelect={(_evt, listRow) => { targetPipelineSelect(listRow); }}
                    // onActivate={targetPipelineActivate}
                    renderRow={targetPipelineRenderRow}
                    width="100%"
                />
            </div>
        </Card>

        <p>
            __NEXTRUNVERSION__
        </p>
    </>
}

export interface SourcePipelineDocument {
    targetPipelines: TargetPipeline[];
}

export interface TargetPipeline {
    name?: string;
}

export interface TargetPipelineListItemProps {
    name: string;
}

function TargetPipelineListItem({ name }: TargetPipelineListItemProps) {
    let className = `scroll-hidden flex-row flex-center flex-grow padding-4`;
    return (
        <div className={className}>
            <div className="margin-right-4"></div>
            <div className="font-size-m flex-self-center padding-4 flex-noshrink">{name}</div>
        </div>
    )
}