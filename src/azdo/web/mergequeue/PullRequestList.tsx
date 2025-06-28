import * as SDK from 'azure-devops-extension-sdk';
import * as luxon from 'luxon'
import * as Azdo from '../azdo/azdo.ts';
import { type IHostNavigationService } from 'azure-devops-extension-api';
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { ScrollableList, IListItemDetails, ListSelection, ListItem } from "azure-devops-ui/List";
import { Status, Statuses, StatusSize } from "azure-devops-ui/Status";
import { Pill, PillVariant } from "azure-devops-ui/Pill";
import { PillGroup } from "azure-devops-ui/PillGroup";

interface PullRequestListProps {
    organization?: string;
    project?: string;
    pullRequests: Array<Azdo.PullRequest>;
    selection: ListSelection;
    onSelectionChanged?: (selection: Azdo.PullRequest) => void;
}

// interface PullRequestItemState extends Azdo.PullRequest {
//     isDefaultBranch: boolean;
//     targetBranch: string;
// }

function PullRequestList(p: PullRequestListProps) {

    function renderPullRequestRow(
        index: number,
        pullRequest: any,
        details: IListItemDetails<any>,
        key?: string
    ): React.JSX.Element {
        let extra = "";
        let className = `scroll-hidden flex-row flex-center flex-grow padding-4 ${extra}`;
        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}>

                {
                    <div className={className}>
                        <Status
                            {...(pullRequest.isDraft ? Statuses.Queued : Statuses.Information)}
                            key="information"
                            size={StatusSize.m}
                        />
                        <div className="font-size-m padding-left-8">{pullRequest.repository.name}</div>
                        <div className="font-size-m italic text-neutral-70 text-ellipsis padding-left-8">{pullRequest.title}</div>
                        <PillGroup className="padding-left-16 padding-right-16">
                            {
                                pullRequest.isDraft && (
                                    <Pill>Draft</Pill>
                                )
                            }
                            {
                                !pullRequest.isDefaultBranch && pullRequest.targetBranch && (
                                    <Pill variant={PillVariant.outlined}>{pullRequest.targetBranch}</Pill>
                                )
                            }
                        </PillGroup>
                        <div className="font-size-m flex-row flex-grow"><div className="flex-grow" />
                            <div>{luxon.DateTime.fromISO(pullRequest.creationDate).toRelative()}</div>
                        </div>
                    </div>
                }



            </ListItem>
        );
    };

    function selectPullRequest(_: any, data: any) {
        console.log("selected run: ", data, data.data);
        
        const t = p.selection.value?.[0];
        if (!t) return;
        // let u = filteredList()[t.beginIndex];
        let u = p.pullRequests[t.beginIndex];
        if (!u) return;
        p.onSelectionChanged?.(u);
    }

    function getPullRequests(): Array<any> {
        // let all = [...filteredList()];
        let all = [...p.pullRequests];
        return all.sort((a, b) => {
            let x = a.pullRequestId || 0;
            let y = b.pullRequestId || 0;
            if (x < y) { return 1; }
            else if (x > y) { return -1; }
            else { return 0; }
        })
    }

    async function activatePullRequest(_: any, evt: any) {
        console.log("activated pull request: ", evt);
        let idx = evt.index;
        let data = evt.data;
        console.log("activated pull request2: ", idx, data);
        const navService = await SDK.getService<IHostNavigationService>("ms.vss-features.host-navigation-service");
        let url = `https://dev.azure.com/${p.organization}/${p.project}/_git/${data.repository.name}/pullrequest/${data.pullRequestId}`;
        console.log("url: ", url);
        navService.openNewWindow(url, "");
    }

    return (
        <>
            <ScrollableList
                itemProvider={new ArrayItemProvider(getPullRequests())}
                selection={p.selection}
                onSelect={selectPullRequest}
                onActivate={activatePullRequest}
                renderRow={renderPullRequestRow}
                width="100%" />
        </>
    )
}

export { PullRequestList };
export type { PullRequestListProps };
