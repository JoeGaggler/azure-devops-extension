import { IListItemDetails, IListRow, ListItem, ListSelection, ScrollableList } from "azure-devops-ui/List";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { GetRunStatusType, Run } from "../currentruns/Run";
// import { Icon, IconSize } from "azure-devops-ui/Icon";
// import { VssPersona } from "azure-devops-ui/VssPersona";
// import { PillGroup } from "azure-devops-ui/PillGroup";
// import { Pill, PillSize, PillVariant } from "azure-devops-ui/Pill";
// import { PullRequestAsyncStatus } from "azure-devops-extension-api/Git/Git";

export interface PipelineListItem {
    runId: number;
    runName: string;
    pipelineId: number;
    pipelineName: string;
    sourceVersion: string;
    status: string;
    result: string;
}

export interface PullRequestListProps {
    pipelines: PipelineListItem[];
    selectedIds: number[];
    onSelectPipelines?: (runIds: number[]) => void;
    onActivatePipeline?: (runId: number, pipelineId: number) => void;
}

function applySelections<T, TItem>(listSelection: ListSelection, all: T[], accessor: (t: T) => TItem, ids: TItem[]) {
    listSelection.clear();
    for (let id of ids) {
        let idx = all.findIndex((item) => accessor(item) === id);
        if (idx < 0) { continue; }
        listSelection.select(idx, 1, true, true);
    }
}

export function PipelineList({ pipelines, selectedIds, onSelectPipelines, onActivatePipeline }: PullRequestListProps) {
    let listSelection = new ListSelection(true);

    applySelections(listSelection, pipelines, (p) => p.runId, selectedIds);

    function onSelectRow(row: IListRow<PipelineListItem>) {
        console.log("NextRunTab -> targetPipelineSelect", row);
        onSelectPipelines?.([row.data.pipelineId]);
    }

    function onActivateRow(row: IListRow<PipelineListItem>) {
        console.log("NextRunTab -> targetPipelineActivate", row);
        if (onActivatePipeline) {
            onActivatePipeline(row.data.runId, row.data.pipelineId);
        }
    }

    function renderRow(
        index: number,
        item: PipelineListItem,
        details: IListItemDetails<PipelineListItem>,
        key?: string
    ): JSX.Element {
        if (!item) { return <></> }
        // let extra = "";
        // let className = `scroll-hidden flex-row flex-center rhythm-horizontal-8 flex-grow padding-4 ${extra}`;

        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}
            >
                <Run
                    name={item.runName}
                    definitionName={item.pipelineName}
                    status={GetRunStatusType(item.status, item.result)}
                    comment={""}
                    started={0}
                    isAlternate={false}
                    isKnown={true}
                    knownTags={[]}
                />

                {/* <div className={className}>
                    <div>{item.runId}</div>
                    <div>{item.runName}</div>
                    <div>{item.pipelineId}</div>
                    <div>{item.pipelineName}</div>
                    <div>{item.sourceVersion}</div>
                </div> */}
            </ListItem>
        )
    }

    return <>
        <div className="flex-column">
            <ScrollableList
                itemProvider={new ArrayItemProvider(pipelines || [])}
                selection={listSelection}
                onSelect={(_evt, listRow) => { onSelectRow(listRow); }}
                onActivate={(_evt, listRow) => { onActivateRow(listRow); }}
                renderRow={renderRow}
                width="100%"
            />
        </div>
    </>
}