import React from "react";
import * as Azdo from '../shared/azdo.ts';
import * as Ping from '../shared/lib.ts';
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { Card } from "azure-devops-ui/Card";
import { ListItem, ListSelection, type IListItemDetails, type IListRow } from "azure-devops-ui/List";
import { ScrollableList } from "azure-devops-ui/List";
// import { Page } from "azure-devops-ui/Page";
import { Header, TitleSize } from "azure-devops-ui/Header";
import { AddPipelinePanel, type AddPipelinePanelValues } from "./AddPipelinePanel.tsx";
import { type IHostNavigationService } from 'azure-devops-extension-api';
import * as SDK from 'azure-devops-extension-sdk';

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
    targetPipelines: TargetPipeline[];
    selectedTargetPipelineId?: number;
    isShowingAddPipelinePanel: boolean;
}

interface ReducerAction {
    targetPipelines?: TargetPipeline[];
    selectTargetPipeline?: TargetPipeline | null;
    showAddPipelinePanel?: boolean;
}

function makeDocId(project: string, definitionId: number) {
    return `project-${project}-pipeline-${definitionId}`;
}

function reducer(state: ReducerState, action: ReducerAction): ReducerState {
    let next = { ...state };

    if (action.targetPipelines !== undefined) {
        let sortedPipelines = action.targetPipelines || [];
        Ping.sortByString(sortedPipelines, i => i.name || "");
        next.targetPipelines = sortedPipelines;
    }

    if (action.selectTargetPipeline !== undefined) {
        let candidates = next.targetPipelines;
        let selected = action.selectTargetPipeline;
        if (!selected || !candidates) {
            next.selectedTargetPipelineId = undefined;
        } else {
            next.selectedTargetPipelineId = selected.id;
            if (undefined === candidates.find(t => t.id === selected.id)) {
                next.selectedTargetPipelineId = undefined;
            }
        }
    }

    if (action.showAddPipelinePanel !== undefined) {
        next.isShowingAddPipelinePanel = action.showAddPipelinePanel;
    }

    return next;
}

export function NextRunTab(p: NextRunTabProps) {
    console.log("NextRunTab render", p);

    const sourcePipelinesCollectionId = "source-pipelines";
    let documentId = React.useRef<string>();
    let tenantInfo = React.useRef<Azdo.TenantInfo>();
    let singleton = React.useRef(p.singleton);

    const [state, dispatch] = React.useReducer<(state: ReducerState, action: ReducerAction) => ReducerState>(reducer, {
        targetPipelines: [],
        selectedTargetPipelineId: undefined,
        isShowingAddPipelinePanel: false,
    })

    let targetPipelineSelection = new ListSelection(true);
    Ping.applySelection(targetPipelineSelection, state.targetPipelines, i => i.id, state.selectedTargetPipelineId);
    const hasSelectedTargetPipeline = state.selectedTargetPipelineId !== undefined;

    // initialize the app
    React.useEffect(() => { init() }, []);
    async function init() {
        console.log("NextRunTab -> init");

        let info = await Azdo.getAzdoInfo();
        console.log("NextRunTab -> tenant", info);
        tenantInfo.current = info;

        let project = info.project;
        let defId = singleton.current.definition?.id;
        if (project === undefined || defId === undefined) {
            console.warn("NextRunTab -> missing project or definition id", { project, defId });
            return;
        }

        let docId = makeDocId(project, defId);
        documentId.current = docId;
        console.log("NextRunTab -> source pipeline document id", docId);

        let sources: SourcePipelineDocument = {
            targetPipelines: []
        }
        sources = await Azdo.getOrCreateSharedDocument(sourcePipelinesCollectionId, docId, sources)
        console.log("NextRunTab -> got shared document", sources);

        dispatch({ targetPipelines: sources.targetPipelines });

        setupPipelineMapping();
    }

    React.useEffect(() => {
        let id = setInterval(() => { tick(); }, 5000);
        return () => clearInterval(id);
    }, []);
    async function tick() {
        console.log("NextRunTab -> tick");
    }

    async function setupPipelineMapping() {
        console.log("NextRunTab -> setupPipelineMapping");
        let info = tenantInfo.current;
        if (info === undefined) { return; }

        let project = info.project;
        if (project === undefined) { return; }

        let targetDefinitionId = singleton.current.definition?.id;
        if (targetDefinitionId === undefined || typeof targetDefinitionId !== "number") { return; }

        let targetDefinitionName = singleton.current.definition?.name;
        if (targetDefinitionName === undefined) { return; }

        let targetRunId = singleton.current.build?.id;
        if (targetRunId === undefined) { return; }

        let targetRunModel = await Azdo.getPipelineRun(info, targetDefinitionId, targetRunId);
        if (targetRunModel === undefined) { return; }

        console.log("NextRunTab -> target run model", targetRunModel);

        let pipelineResources = targetRunModel.resources?.pipelines;
        if (!pipelineResources) { return; }

        console.log("NextRunTab -> source pipelines", pipelineResources);
        for (let pipelineResourceKey in pipelineResources) {
            let pipelineResource = pipelineResources[pipelineResourceKey];

            let sourcePipelineResource = pipelineResource.pipeline;
            if (sourcePipelineResource == null) { continue; }
            console.log("NextRunTab -> source pipeline", pipelineResourceKey, pipelineResource);

            let sourceRunId = sourcePipelineResource.id;
            if (sourceRunId == null) { continue; }

            // let spDefName = spPipeline.name;
            // if (!Ping.isString(spDefName)) { continue; }

            let sourceRunModel = await Azdo.getBuildRun(info, sourceRunId);
            if (sourceRunModel == null) { continue; }
            console.log("NextRunTab -> got source pipeline run model", sourceRunModel);

            let sourceDefinitionId = sourceRunModel.definition.id;
            if (sourceDefinitionId == null) { continue; }

            let docId = makeDocId(project, sourceDefinitionId);
            // await Azdo.deleteSharedDocument(sourcePipelinesCollectionId, docId);
            let doc = await Azdo.getOrCreateSharedDocument(sourcePipelinesCollectionId, docId, { targetPipelines: [] });
            console.log("NextRunTab -> got source pipeline document", docId, doc);

            let docTargetPipelines: Array<TargetPipeline> = doc.targetPipelines || [];
            console.log("NextRunTab -> got source pipeline target pipelines", docId, docTargetPipelines);

            if (docTargetPipelines.find(i => i.id === targetDefinitionId)) {
                console.log("NextRunTab -> target pipeline already mapped in source pipeline document, skipping", docId, targetDefinitionId);
                continue;
            }

            docTargetPipelines.push({
                id: targetDefinitionId,
                name: targetDefinitionName,
                resourceName: pipelineResourceKey,
            });

            let nnn = await Azdo.trySaveSharedDocument(sourcePipelinesCollectionId, docId, doc);
            console.log("NextRunTab -> saved source pipeline document", docId, nnn);
        }
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

    // function showAddTargetPipelinePanel() {
    //     console.log("NextRunTab -> showAddTargetPipelinePanel");
    //     dispatch({ showAddPipelinePanel: true });
    // }

    async function showRunTargetPipelinePanel() {
        console.log("NextRunTab -> showRunTargetPipelinePanel");

        let org = tenantInfo.current?.organization;
        let project = tenantInfo.current?.project;
        if (!org || !project) {
            console.error("NextRunTab -> showRunTargetPipelinePanel -> missing org or project", { org, project });
            return;
        }

        let pipelineId = state.selectedTargetPipelineId;
        if (pipelineId === undefined) {
            console.error("NextRunTab -> showRunTargetPipelinePanel -> missing pipeline id");
            return;
        }

        let pipeline = state.targetPipelines?.find(t => t.id === pipelineId);
        if (!pipeline) {
            console.error("NextRunTab -> showRunTargetPipelinePanel -> selected pipeline not found in state", { pipelineId, pipelines: state.targetPipelines });
            return;
        }

        let pipelineResourceName = pipeline.resourceName;
        if (!pipelineResourceName) {
            console.error("NextRunTab -> showRunTargetPipelinePanel -> selected pipeline missing resource name", pipeline);
            return;
        }

        // https://dev.azure.com/{organization}/{project}/_apis/pipelines/{pipelineId}/runs?api-version=7.2-preview.1
        let url = `https://dev.azure.com/${org}/${project}/_apis/pipelines/${pipelineId}/runs?api-version=7.2-preview.1`;
        let body = {
            previewRun: false,
            stagesToSkip: [],
            resources: {
                repositories: {
                    ["self"]: {
                        refName: singleton.current.build?.sourceBranch,
                    }
                },
                pipelines: {
                    [pipelineResourceName]: {
                        runId: singleton.current.build.id
                    }
                },
            }
        };
        console.log("NextRunTab -> showRunTargetPipelinePanel -> posting to azdo api", url, body);
        let response = await Azdo.postAzdo(url, body, singleton.current.bearerToken);
        console.log("NextRunTab -> showRunTargetPipelinePanel -> got response", response);
        if (response === undefined) {
            return;
        }
        let link = response._links?.web?.href;
        console.log("NextRunTab -> showRunTargetPipelinePanel -> got run link", link);
        const navService = await SDK.getService<IHostNavigationService>("ms.vss-features.host-navigation-service");
        console.log("NextRunTab -> showRunTargetPipelinePanel -> got nav service", navService);
        navService.openNewWindow(link, "");
    }

    async function onCommitNewPipeline(data: AddPipelinePanelValues) {
        console.log("NextRunTab -> onCommitNewPipeline", data);

        // TODO: update shared document

        const docId = documentId.current;
        if (!docId) {
            console.error("NextRunTab -> onCommitNewPipeline -> missing document id");
            return;
        }

        let prevPipelines = state.targetPipelines ?? [];
        // TODO: deduplicate, sort

        let prevPipelinesDoc: SourcePipelineDocument = {
            targetPipelines: prevPipelines
        }
        prevPipelinesDoc = await Azdo.getOrCreateSharedDocument(sourcePipelinesCollectionId, docId, prevPipelinesDoc);

        console.log("What is this?", prevPipelinesDoc.targetPipelines);
        prevPipelinesDoc.targetPipelines = [
            ...(prevPipelinesDoc.targetPipelines || []),
            {
                id: data.id,
                name: data.name,
                resourceName: data.resource,
            }];

        const nextPipelinesDoc = await Azdo.trySaveSharedDocument(sourcePipelinesCollectionId, docId, prevPipelinesDoc);
        if (!nextPipelinesDoc) {
            console.warn("Failed to save document.", prevPipelinesDoc);
            dispatch({
                showAddPipelinePanel: false
            });
        } else {
            console.log("Saved document.", nextPipelinesDoc);
            dispatch({
                targetPipelines: nextPipelinesDoc.targetPipelines || [],
                showAddPipelinePanel: false
            });
        }
    }

    async function onCancelNewPipeline() {
        console.log("NextRunTab -> onCancelNewPipeline");
        dispatch({ showAddPipelinePanel: false });
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
                    // {
                    //     id: "addTargetPipeline",
                    //     // text: "Add Pipeline",
                    //     iconProps: { iconName: "Add" },
                    //     onActivate: () => { showAddTargetPipelinePanel(); },
                    //     isPrimary: false,
                    //     important: true,
                    //     disabled: false,
                    // },
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
                    {
                        (state.targetPipelines && state.targetPipelines.length > 0) ? (
                            <ScrollableList
                                itemProvider={new ArrayItemProvider(state.targetPipelines || [])}
                                selection={targetPipelineSelection}
                                onSelect={(_evt, listRow) => { targetPipelineSelect(listRow); }}
                                onActivate={showRunTargetPipelinePanel}
                                renderRow={targetPipelineRenderRow}
                                width="100%"
                            />
                        ) : (
                            <div className="flex-row flex-center padding-16">
                                <div className="font-size-m text-neutral-70">
                                    No target pipelines configured.
                                </div>
                            </div>
                        )
                    }
                </div>
            </Card>
            {
                state.isShowingAddPipelinePanel &&
                <AddPipelinePanel
                    onCommit={onCommitNewPipeline}
                    onCancel={onCancelNewPipeline}
                />
            }

            <div className="text-neutral-30 flex-row padding-4">
                <div className="flex-grow"></div>
                <div>__NEXTRUNVERSION__</div>
            </div>
        </div>
    )
}

export interface SourcePipelineDocument {
    targetPipelines: TargetPipeline[];
}

export interface TargetPipeline {
    id?: number;
    name?: string;
    resourceName?: string;
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