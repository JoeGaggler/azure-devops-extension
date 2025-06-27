import * as SDK from 'azure-devops-extension-sdk';
import * as luxon from 'luxon'
import { type IHostNavigationService } from 'azure-devops-extension-api';
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import { ScrollableList, IListItemDetails, ListSelection, ListItem } from "azure-devops-ui/List";
import { Status, Statuses, StatusSize } from "azure-devops-ui/Status";
import { Pill, PillVariant } from "azure-devops-ui/Pill";
import { PillGroup } from "azure-devops-ui/PillGroup";

interface PullRequestListProps {
    organization?: string;
    project?: string;
    pullRequests: Array<any>;
    filters: any;
    repos: any;
}

function PullRequestList(p: PullRequestListProps) {
    let selection = new ListSelection(true);

    function filteredList() {
        let all = [...p.pullRequests];
        console.log("Filtering drafts: ", p.filters, p.filters.drafts);
        if (p.filters && (p.filters.drafts as boolean) === false) {
            console.log("Filtering out drafts");
            all = all.filter(pr => !pr.isDraft);
        } else {
            console.log("Not filtering out drafts");
        }

        all = all.map(pr => {
            pr.isDefaultBranch = false;
            pr.targetBranch = undefined;
            let repo = p.repos[pr.repository.name];
            if (repo && pr.targetRefName && repo.defaultBranch) {
                pr.isDefaultBranch = ((pr.targetRefName == repo.defaultBranch) as boolean)
                pr.targetBranch = pr.targetRefName.replace("refs/heads/", "");
            }
            return pr;
        })
        console.log("all map: ", all);
        if (p.filters && (p.filters.allBranches as boolean) == false && p.repos) {
            all = all.filter(pr => { return true === (pr && pr.isDefaultBranch) });
            // console.log("Repo count: ", Object.keys(p.repos).length);
            // for (let pr of all) {
            //     pr.isDefaultBranch = false
            //     let repo = p.repos[pr.repository.name];
            //     if (!repo) { continue }
            //     pr.isDefaultBranch = (pr.targetRefName == repo.defaultBranch)
            // }

            // all = all.filter(pr => {
            //     if (!pr.repository) {
            //         console.warn("Pull request has no repository:", pr);
            //         return false;
            //     }
            //     let repo = p.repos[pr.repository.name];
            //     if (!repo) {
            //         console.warn("No repository found for pull request:", pr);
            //         return false;
            //     }
            //     return (pr.targetRefName == repo.defaultBranch);
            // });
        }

        return all;
    }

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
        // setSelectedRunId(data.data.i);
        // setSelectedPipelineId(data.data.p);
        // p.onSelectRun && p.onSelectRun(data.data.i);
    }

    function getPullRequests(): Array<any> {
        let all = [...filteredList()];
        return all.sort((a, b) => {
            if (a.pullRequestId < b.pullRequestId) {
                return 1;
            } else if (a.pullRequestId > b.pullRequestId) {
                return -1;
            } else {
                return 0;
            }
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
                selection={selection}
                onSelect={selectPullRequest}
                onActivate={activatePullRequest}
                renderRow={renderPullRequestRow}
                width="100%" />
        </>
    )
}

export { PullRequestList };
export type { PullRequestListProps };
