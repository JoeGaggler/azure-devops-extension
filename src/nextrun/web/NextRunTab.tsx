import React from "react";
import * as Azdo from '../shared/azdo.ts';
import { MessageCard, MessageCardSeverity } from "azure-devops-ui/MessageCard";
// import * as SDK from 'azure-devops-extension-sdk';

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

    return <>
        <MessageCard severity={MessageCardSeverity.Info}>
            Work in progress. Please check back later.
        </MessageCard>

        <p>
            __NEXTRUNVERSION__
        </p>
    </>
}

export interface SourcePipelineDocument {
    targetPipelines: TargetPipeline[];
}

export interface TargetPipeline {
}