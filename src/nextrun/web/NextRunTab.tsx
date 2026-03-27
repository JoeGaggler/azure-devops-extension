import React from "react";
import * as Azdo from '../shared/azdo.ts';
import * as Ping from '../shared/lib.ts';
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { Card } from "azure-devops-ui/Card";
import { ListItem, ListSelection, type IListItemDetails, type IListRow } from "azure-devops-ui/List";
import { ScrollableList } from "azure-devops-ui/List";
// import { Page } from "azure-devops-ui/Page";
import { Header, TitleSize } from "azure-devops-ui/Header";

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
    selectedTargetPipelineId?: number;
}

interface ReducerAction {
    targetPipelines?: TargetPipeline[];
    selectTargetPipeline?: TargetPipeline | null;
}

function reducer(state: ReducerState, action: ReducerAction): ReducerState {
    let next = { ...state };

    if (action.targetPipelines) {
        next.targetPipelines = action.targetPipelines;
    }

    if (action.selectTargetPipeline !== undefined) {
        let candidates = next.targetPipelines;
        let selected = action.selectTargetPipeline;
        if (!selected || !candidates) {
            next.selectedTargetPipelineId = undefined;
        } else {
            next.selectedTargetPipelineId = candidates.find(t => t.name === selected.name)?.id || undefined;
        }
    }

    return next;
}

export function NextRunTab(p: NextRunTabProps) {
    console.log("NextRunTab render", p);

    const sourcePipelinesCollectionId = "source-pipelines";
    let tenantInfo = React.useRef<Azdo.TenantInfo>();

    const [state, dispatch] = React.useReducer<(state: ReducerState, action: ReducerAction) => ReducerState>(reducer, {})

    let targetPipelineSelection = new ListSelection(true);
    Ping.applySelection(targetPipelineSelection, state.targetPipelines || [], i => i.id, state.selectedTargetPipelineId);
    const hasSelectedTargetPipeline = state.selectedTargetPipelineId !== undefined;

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

        // TODO: REMOVE HARDCODED TARGET PIPELINES
        let targetPipelineItems: TargetPipeline[] = [
            {
                id: 1,
                name: "Test Pipeline 1"
            },
            {
                id: 2,
                name: "Test Pipeline 2"
            },
            {
                id: 3,
                name: "Test Pipeline 3"
            }
        ];
        sources.targetPipelines = targetPipelineItems;

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
        dispatch({ selectTargetPipeline: row.data });
    }

    function showAddTargetPipelinePanel() {
        console.log("NextRunTab -> showAddTargetPipelinePanel");
    }

    function showRunTargetPipelinePanel() {
        console.log("NextRunTab -> showRunTargetPipelinePanel");
    }

    return (
        <div className="flex-column padding-4">
            <Header
                title={"Pipelines"}
                titleSize={TitleSize.Large}
                // titleIconProps={{ iconName: "Next" }}
                contentClassName='flex-center'
                // backButtonProps={Util.makeHeaderBackButtonProps(p.appNav)}
                commandBarItems={[
                    {
                        id: "addTargetPipeline",
                        // text: "Add Pipeline",
                        iconProps: { iconName: "Add" },
                        onActivate: () => { showAddTargetPipelinePanel(); },
                        isPrimary: false,
                        important: true,
                        disabled: false,
                    },
                    {
                        id: "runTargetPipeline",
                        text: "Run",
                        // iconProps: { iconName: "Add" },
                        onActivate: () => { showRunTargetPipelinePanel(); },
                        isPrimary: true,
                        important: true,
                        disabled: !hasSelectedTargetPipeline,
                    }
                ]}
            />

            <Card className="padding-8">
                <div className="flex-column">
                    <ScrollableList
                        itemProvider={new ArrayItemProvider(state.targetPipelines || [])}
                        selection={targetPipelineSelection}
                        onSelect={(_evt, listRow) => { targetPipelineSelect(listRow); }}
                        onActivate={showRunTargetPipelinePanel}
                        renderRow={targetPipelineRenderRow}
                        width="100%"
                    />
                </div>
            </Card>
            {
                // isAddingHuddle &&
                // <CreateHuddlePanel
                //     onCommit={onCommitNewHuddle}
                //     onCancel={onCancelNewHuddle}
                // />
            }

            <p>
                __NEXTRUNVERSION__
            </p>
        </div>
    )
}

export interface SourcePipelineDocument {
    targetPipelines: TargetPipeline[];
}

export interface TargetPipeline {
    id?: number;
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