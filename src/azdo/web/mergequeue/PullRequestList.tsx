import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { ScrollableList, IListItemDetails, ListSelection, ListItem } from "azure-devops-ui/List";

interface PullRequestListProps {
    pullRequests: Array<any>;
}

function PullRequestList(p: PullRequestListProps) {
    let selection = new ListSelection(true);

    function renderPullRequestRow(
        index: number,
        pullRequest: any,
        details: IListItemDetails<any>,
        key?: string
    ): React.JSX.Element {
        return (
            <ListItem
                key={key || "list-item" + index}
                index={index}
                details={details}>

                <div>PR index: ${index} - ${pullRequest.pullRequestId} - ${pullRequest.title}</div>

            </ListItem>
        );
    };


    function selectPullRequest(_: any, data: any) {
        console.log("selected run: ", data, data.data);
        // setSelectedRunId(data.data.i);
        // setSelectedPipelineId(data.data.p);
        // p.onSelectRun && p.onSelectRun(data.data.i);
    }

    function getPullRequests(): Array<any> {
        return p.pullRequests
    }

    async function activatePullRequest(_: any, evt: any) {
        console.log("activated pull request: ", evt);
        // let idx = evt.index;
        // let data = evt.data;
        // console.log("activated run: ", idx, data);
        // const navService = await SDK.getService<IHostNavigationService>("ms.vss-features.host-navigation-service");
        // console.log("url: ", data.w);
        // navService.openNewWindow(`https://dev.azure.com/Emdat/Emdat/_build/results?buildId=${data.i}`, "");
    }

    return (
        <>
            {
                <ScrollableList
                    itemProvider={new ArrayItemProvider(getPullRequests())}
                    selection={selection}
                    onSelect={selectPullRequest}
                    onActivate={activatePullRequest}
                    renderRow={renderPullRequestRow}
                    width="100%" />
            }
            <br />
            {
                getPullRequests().map(
                    (pullRequest, index) => {
                        return (
                            <div className="flex-row flex-center rhythm-vertical-8">
                                {index} - !{pullRequest.pullRequestId} - {pullRequest.title}
                            </div>
                        )
                    })
            }

            <br />
        </>
    )
}

export { PullRequestList };
export type { PullRequestListProps };
