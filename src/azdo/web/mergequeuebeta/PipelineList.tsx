import { IListItemDetails, IListRow, ListItem, ListSelection, ScrollableList } from "azure-devops-ui/List";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
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
        let extra = "";
        let className = `scroll-hidden flex-row flex-center rhythm-horizontal-8 flex-grow padding-4 ${extra}`;

        // let initialsIdentityProvider = {
        //     getDisplayName() {
        //         return pullRequest.author?.displayName || "?";
        //     },
        //     getIdentityImageUrl(_size: number) {
        //         return pullRequest.author?.imageUrl || undefined;
        //     }
        // }

        // function renderPills(pullRequest: PipelineListItem): JSX.Element {
        //     let mergeStatus = pullRequest.mergeStatus;
        //     let voteStatus = pullRequest.voting?.status;
        //     let voteCount = pullRequest.voting?.count || 0;
        //     let voteCountString = voteCount > 1 ? ` (${voteCount})` : "";

        //     return <PillGroup className="padding-left-16 padding-right-16">
        //         {mergeStatus === PullRequestAsyncStatus.Conflicts && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 192, green: 0, blue: 0 }}>Conflicts</Pill>)}
        //         {mergeStatus === PullRequestAsyncStatus.Failure && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 192, green: 0, blue: 0 }}>Failure</Pill>)}
        //         {mergeStatus === PullRequestAsyncStatus.RejectedByPolicy && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 192, green: 0, blue: 0 }}>Policy</Pill>)}

        //         {voteStatus === "approved" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 64, green: 128, blue: 64 }}>Approved{voteCountString}</Pill>)}
        //         {voteStatus === "suggestions" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 64, green: 64, blue: 128 }}>Suggestions{voteCountString}</Pill>)}
        //         {voteStatus === "waiting" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 169, green: 154, blue: 60 }}>Waiting{voteCountString}</Pill>)}
        //         {voteStatus === "rejected" && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 192, green: 0, blue: 0 }}>Rejected{voteCountString}</Pill>)}

        //         {pullRequest.isDraft && (<Pill size={PillSize.compact}>Draft</Pill>)}
        //         {pullRequest.isAutoComplete && (<Pill size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 92, green: 128, blue: 92 }}>Auto-Complete</Pill>)}

        //         {pullRequest.nonDefaultTargetBranch && (<Pill size={PillSize.compact} variant={PillVariant.outlined}>{pullRequest.nonDefaultTargetBranch}</Pill>)}
        //     </PillGroup>
        // }

        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}
            >
                <div className={className}>
                    <div>{item.runId}</div>
                    <div>{item.runName}</div>
                    <div>{item.pipelineId}</div>
                    <div>{item.pipelineName}</div>
                    <div>{item.sourceVersion}</div>
                    {/* <Icon iconName={pullRequest.icon} size={IconSize.medium} className={pullRequest.iconClassName} />
                    <div className="font-size-m flex-row flex-center flex-shrink">{pullRequest.pullRequestId}</div>
                    <VssPersona size={"extra-small"} identityDetailsProvider={initialsIdentityProvider} />
                    <div className="font-size-m">{pullRequest.repository}</div>
                    <div className="font-size-m italic text-neutral-70 text-ellipsis">{pullRequest.title}</div>
                    <div>{renderPills(pullRequest)}</div>
                    <div className="font-size-m flex-row flex-center flex-grow rhythm-horizontal-8">
                        <div className="flex-grow" />
                        <div>{pullRequest.dateString}</div>
                    </div> */}
                </div>
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