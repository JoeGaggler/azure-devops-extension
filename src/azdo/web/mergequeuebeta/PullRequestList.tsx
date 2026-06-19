import { IListItemDetails, IListRow, ListItem, ListSelection, ScrollableList } from "azure-devops-ui/List";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { Icon, IconSize } from "azure-devops-ui/Icon";
import { VssPersona } from "azure-devops-ui/VssPersona";
import { AuthorInfo } from "./azuredevops";

export interface PullRequestListItem {
    pullRequestId: number;
    repository: string;
    title: string;
    icon: string;
    iconClassName?: string;
    author: AuthorInfo;
    dateString?: string;
}

export interface PullRequestListProps {
    pullRequests: PullRequestListItem[];
    selectedIds: number[];
    onSelectPullRequestIds: (id: number[]) => void;
    onActivatePullRequest?: (id: number, repo: string) => void;
}

function applySelections<T, TItem>(listSelection: ListSelection, all: T[], accessor: (t: T) => TItem, ids: TItem[]) {
    listSelection.clear();
    for (let id of ids) {
        let idx = all.findIndex((item) => accessor(item) === id);
        if (idx < 0) { continue; }
        listSelection.select(idx, 1, true, true);
    }
}

export function PullRequestList({ pullRequests, selectedIds, onSelectPullRequestIds, onActivatePullRequest }: PullRequestListProps) {
    let listSelection = new ListSelection(true);

    applySelections(listSelection, pullRequests, (pr) => pr.pullRequestId, selectedIds);

    function onSelectRow(row: IListRow<PullRequestListItem>) {
        console.log("NextRunTab -> targetPipelineSelect", row);
        onSelectPullRequestIds([row.data.pullRequestId]);
    }

    function onActivateRow(row: IListRow<PullRequestListItem>) {
        console.log("NextRunTab -> targetPipelineActivate", row);
        if (onActivatePullRequest) {
            onActivatePullRequest(row.data.pullRequestId, row.data.repository);
        }
    }

    function renderRow(
        index: number,
        pullRequest: PullRequestListItem,
        details: IListItemDetails<PullRequestListItem>,
        key?: string
    ): JSX.Element {
        if (!pullRequest) { return <></> }
        let extra = "";
        let className = `scroll-hidden flex-row flex-center rhythm-horizontal-8 flex-grow padding-4 ${extra}`;

        let initialsIdentityProvider = {
            getDisplayName() {
                return pullRequest.author?.displayName || "?";
            },
            getIdentityImageUrl(_size: number) {
                return pullRequest.author?.imageUrl || undefined;
            }
        }

        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}
            >
                <div className={className}>
                    <Icon iconName={pullRequest.icon} size={IconSize.medium} className={pullRequest.iconClassName} />
                    <div className="font-size-m flex-row flex-center flex-shrink">{pullRequest.pullRequestId}</div>
                    <VssPersona size={"extra-small"} identityDetailsProvider={initialsIdentityProvider} />
                    <div className="font-size-m">{pullRequest.repository}</div>
                    <div className="font-size-m italic text-neutral-70 text-ellipsis">{pullRequest.title}</div>
                    <div className="font-size-m flex-row flex-center flex-grow rhythm-horizontal-8">
                        <div className="flex-grow" />
                        <div>{pullRequest.dateString}</div>
                    </div>
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
                onActivate={(_evt, listRow) => { onActivateRow(listRow); }}
                renderRow={renderRow}
                width="100%"
            />
        </div>
    </>
}